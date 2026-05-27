package internal

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"html"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"math"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"github.com/jung-kurt/gofpdf"
)

const defaultPDFTitle = "Reporte"
const defaultHeaderDesignPath = "internal/diseño.xml"
const defaultHeaderLogoPath = "../logo_gq_transparente.png"

// BuildReportesPDF arma un PDF con dos páginas:
// 1) reporte diario, 2) reporte semanal.
func BuildReportesPDF(ctx context.Context, sh *SheetsReader) ([]byte, string, error) {
	if sh == nil {
		return nil, "", fmt.Errorf("Sheets no configurado")
	}

	diario, err := sh.ReadReporteDiarioRich(ctx)
	if err != nil {
		return nil, "", fmt.Errorf("leer reporte diario: %w", err)
	}
	semanal, err := sh.ReadReporteSemanalRich(ctx)
	if err != nil {
		return nil, "", fmt.Errorf("leer reporte semanal: %w", err)
	}

	pdf := gofpdf.New("P", "mm", "A4", "")
	pdf.SetTitle(defaultPDFTitle, false)
	pdf.SetAuthor("G3Bot", false)
	pdf.SetMargins(12, 12, 12)
	pdf.SetAutoPageBreak(true, 12)

	tr := pdf.UnicodeTranslatorFromDescriptor("")
	// Encabezado en cada página física (incluye saltos automáticos del contenido HTML).
	pdf.SetHeaderFuncMode(func() {
		addPDFHeader(pdf, tr)
	}, false)

	addReportPage(pdf, tr, "Reporte Diario", diario)
	addReportPage(pdf, tr, "Reporte Semanal", semanal)

	var out bytes.Buffer
	if err := pdf.Output(&out); err != nil {
		return nil, "", fmt.Errorf("generar pdf: %w", err)
	}

	name := fmt.Sprintf("reporte_%s.pdf", time.Now().Format("20060102_150405"))
	return out.Bytes(), name, nil
}

func addReportPage(pdf *gofpdf.Fpdf, tr func(string) string, title string, body []RichLine) {
	pdf.AddPage()

	pdf.SetFont("Arial", "B", 16)
	pdf.CellFormat(0, 9, tr(title), "", 1, "L", false, 0, "")
	pdf.Ln(2)

	pdf.SetFont("Arial", "", 10)
	if len(body) == 0 {
		body = []RichLine{{{Text: "(Vacío) No hay datos para este reporte.", Bold: false}}}
	}
	htmlW := pdf.HTMLBasicNew()
	htmlBody := richLinesToHTML(body, tr)
	htmlW.Write(5, htmlBody)
}

func addPDFHeader(pdf *gofpdf.Fpdf, tr func(string) string) {
	left, _, right, _ := pdf.GetMargins()
	pageW, _ := pdf.GetPageSize()
	contentW := pageW - left - right
	topY := pdf.GetY()
	header := readHeaderData(resolveHeaderDesignPath())

	// Texto principal (se pinta directo en PDF para evitar limitaciones de SVG text).
	pdf.SetFont("Arial", "B", 24)
	pdf.SetTextColor(23, 60, 102)
	pdf.SetXY(left, topY+4)
	pdf.CellFormat(contentW*0.62, 10, tr(header.Name), "", 1, "L", false, 0, "")

	pdf.SetFont("Arial", "B", 14)
	pdf.SetTextColor(50, 91, 128)
	pdf.SetXY(left, topY+14)
	pdf.CellFormat(contentW*0.62, 7, tr(header.Title), "", 1, "L", false, 0, "")

	pdf.SetFont("Arial", "", 11)
	pdf.SetTextColor(95, 95, 95)
	pdf.SetXY(left, topY+22)
	pdf.CellFormat(contentW*0.62, 7, tr(fmt.Sprintf("%s | %s", header.Phone, header.Email)), "", 1, "L", false, 0, "")

	// Logo personalizado a la derecha.
	if logoPath := resolveHeaderLogoPath(); logoPath != "" {
		if imgType := detectImageType(logoPath); imgType != "" {
			// Mantener proporción para evitar deformación y conservar nitidez.
			logoW, logoH := resolveLogoBoxSize(logoPath, 22, 18)
			logoX := pageW - right - logoW
			logoY := topY + 3
			pdf.ImageOptions(
				logoPath,
				logoX,
				logoY,
				logoW,
				logoH,
				false,
				gofpdf.ImageOptions{ImageType: imgType, ReadDpi: true},
				0,
				"",
			)
		}
	}

	// Barra inferior: azul completa + acento verde al inicio.
	lineY := topY + 35
	pdf.SetDrawColor(26, 35, 126)
	pdf.SetLineWidth(1.2)
	pdf.Line(left, lineY, pageW-right, lineY)
	pdf.SetDrawColor(76, 175, 80)
	pdf.SetLineWidth(1.6)
	pdf.Line(left, lineY, left+(contentW*0.24), lineY)

	pdf.SetTextColor(0, 0, 0)
	pdf.SetY(lineY + 4)
}

type headerData struct {
	Name  string
	Title string
	Phone string
	Email string
}

func resolveHeaderDesignPath() string {
	if v := strings.TrimSpace(os.Getenv("PDF_HEADER_SVG_PATH")); v != "" {
		if st, err := os.Stat(v); err == nil && !st.IsDir() {
			return v
		}
	}
	candidates := []string{
		defaultHeaderDesignPath,
		filepath.Join("backend", defaultHeaderDesignPath),
	}
	for _, p := range candidates {
		if st, err := os.Stat(p); err == nil && !st.IsDir() {
			return p
		}
	}
	return ""
}

func readHeaderData(path string) headerData {
	h := headerData{
		Name:  "GENESIS QUINTERO",
		Title: "Contadora Publica",
		Phone: "+58-4242564570",
		Email: "genesisdaniqg@gmail.com",
	}
	if path == "" {
		return h
	}

	b, err := os.ReadFile(path)
	if err != nil {
		return h
	}

	type textNode struct {
		Value string `xml:",chardata"`
	}
	type svgNode struct {
		Text []textNode `xml:"text"`
	}
	var doc svgNode
	if err := xml.Unmarshal(b, &doc); err != nil {
		return h
	}
	var values []string
	for _, t := range doc.Text {
		v := strings.TrimSpace(t.Value)
		if v != "" {
			values = append(values, v)
		}
	}
	if len(values) >= 1 {
		h.Name = values[0]
	}
	if len(values) >= 2 {
		h.Title = values[1]
	}
	for _, v := range values {
		if strings.Contains(v, "@") {
			h.Email = v
		}
	}
	phoneRe := regexp.MustCompile(`[\d+\-]{8,}`)
	for _, v := range values {
		if phoneRe.MatchString(v) && !strings.Contains(v, "@") {
			h.Phone = v
			break
		}
	}
	return h
}

func resolveHeaderLogoPath() string {
	if v := strings.TrimSpace(os.Getenv("PDF_HEADER_LOGO_PATH")); v != "" {
		if p := ResolveAssetPath(v); p != "" {
			if st, err := os.Stat(p); err == nil && !st.IsDir() {
				return p
			}
		}
	}
	candidates := []string{
		defaultHeaderLogoPath,
		filepath.Join("backend", defaultHeaderLogoPath),
	}
	for _, p := range candidates {
		if st, err := os.Stat(p); err == nil && !st.IsDir() {
			return p
		}
	}
	return ""
}

func resolveLogoBoxSize(path string, maxW, maxH float64) (float64, float64) {
	f, err := os.Open(path)
	if err != nil {
		return maxW, maxH
	}
	defer f.Close()
	cfg, _, err := image.DecodeConfig(f)
	if err != nil || cfg.Width <= 0 || cfg.Height <= 0 {
		return maxW, maxH
	}
	scaleW := maxW / float64(cfg.Width)
	scaleH := maxH / float64(cfg.Height)
	scale := math.Min(scaleW, scaleH)
	if scale <= 0 {
		return maxW, maxH
	}
	return float64(cfg.Width) * scale, float64(cfg.Height) * scale
}

func detectImageType(path string) string {
	b, err := os.ReadFile(path)
	if err != nil || len(b) < 12 {
		return ""
	}
	switch {
	case bytes.HasPrefix(b, []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}):
		return "PNG"
	case bytes.HasPrefix(b, []byte{0xFF, 0xD8, 0xFF}):
		return "JPG"
	case bytes.HasPrefix(b, []byte("GIF87a")) || bytes.HasPrefix(b, []byte("GIF89a")):
		return "GIF"
	default:
		return ""
	}
}

func richLinesToHTML(lines []RichLine, tr func(string) string) string {
	var b strings.Builder
	for i, line := range lines {
		for _, run := range line {
			txt := formatRunTextForHTML(run.Text, run.Bold, tr)
			b.WriteString(txt)
		}
		if i != len(lines)-1 {
			b.WriteString("<br>")
		}
	}
	return b.String()
}

// formatRunTextForHTML prepara texto del Sheet para HTMLBasic de gofpdf:
// - tabs → espacios (evita &nbsp; literal que el motor no interpreta)
// - saltos entre actividades en una misma celda (p. ej. "... ) - Otra tarea")
// - *texto* del sheet → negrita real
func formatRunTextForHTML(s string, sheetBold bool, tr func(string) string) string {
	s = strings.ReplaceAll(s, "\t", "    ")
	s = strings.ReplaceAll(s, ") - ", ")\n- ")
	s = strings.ReplaceAll(s, " (Completado) - ", " (Completado)\n- ")
	s = strings.ReplaceAll(s, " (completado) - ", " (completado)\n- ")
	s = strings.ReplaceAll(s, ": - ", ":\n- ")
	// Viñetas pegadas con punto medio (listas en una sola celda)
	s = strings.ReplaceAll(s, "· - ", "·\n- ")

	inner := markdownAsteriskToBoldHTML(s, tr)
	if sheetBold && inner != "" {
		return "<b>" + inner + "</b>"
	}
	return inner
}

// markdownAsteriskToBoldHTML convierte *título* del sheet a <b>título</b> y escapa el resto.
func markdownAsteriskToBoldHTML(s string, tr func(string) string) string {
	if s == "" {
		return ""
	}
	re := regexp.MustCompile(`\*([^*]+)\*`)
	last := 0
	var b strings.Builder
	for _, loc := range re.FindAllStringSubmatchIndex(s, -1) {
		b.WriteString(html.EscapeString(tr(s[last:loc[0]])))
		b.WriteString("<b>")
		b.WriteString(html.EscapeString(tr(s[loc[2]:loc[3]])))
		b.WriteString("</b>")
		last = loc[1]
	}
	b.WriteString(html.EscapeString(tr(s[last:])))
	return strings.ReplaceAll(b.String(), "\n", "<br>")
}
