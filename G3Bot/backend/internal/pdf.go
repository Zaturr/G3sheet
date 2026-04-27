package internal

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"github.com/jung-kurt/gofpdf"
)

const defaultPDFTitle = "Reporte"
const defaultHeaderDesignPath = "internal/diseño.xml"
const defaultHeaderLogoPath = `C:\Users\bdsyc\.cursor\projects\c-Users-bdsyc-OneDrive-Escritorio-telegram-botcito-G3sheet-G3Bot-backend\assets\c__Users_bdsyc_AppData_Roaming_Cursor_User_workspaceStorage_a145905ef5f376731132aed751c0aa79_images_Gemini_Generated_Image_k80in3k80in3k80i-b6573866-6b65-4dc9-bd1a-1bc27e039724.png`

// BuildReportesPDF arma un PDF con dos páginas:
// 1) reporte diario, 2) reporte semanal.
func BuildReportesPDF(ctx context.Context, sh *SheetsReader) ([]byte, string, error) {
	if sh == nil {
		return nil, "", fmt.Errorf("Sheets no configurado")
	}

	diario, err := sh.ReadReporteDiarioTexto(ctx)
	if err != nil {
		return nil, "", fmt.Errorf("leer reporte diario: %w", err)
	}
	semanal, err := sh.ReadReporteSemanalTexto(ctx)
	if err != nil {
		return nil, "", fmt.Errorf("leer reporte semanal: %w", err)
	}

	pdf := gofpdf.New("P", "mm", "A4", "")
	pdf.SetTitle(defaultPDFTitle, false)
	pdf.SetAuthor("G3Bot", false)
	pdf.SetMargins(12, 12, 12)
	pdf.SetAutoPageBreak(true, 12)

	addReportPage(pdf, "Reporte Diario", diario)
	addReportPage(pdf, "Reporte Semanal", semanal)

	var out bytes.Buffer
	if err := pdf.Output(&out); err != nil {
		return nil, "", fmt.Errorf("generar pdf: %w", err)
	}

	name := fmt.Sprintf("reporte_%s.pdf", time.Now().Format("20060102_150405"))
	return out.Bytes(), name, nil
}

func addReportPage(pdf *gofpdf.Fpdf, title, body string) {
	tr := pdf.UnicodeTranslatorFromDescriptor("")

	pdf.AddPage()
	addPDFHeader(pdf, tr)

	pdf.SetFont("Arial", "B", 16)
	pdf.CellFormat(0, 9, tr(title), "", 1, "L", false, 0, "")
	pdf.Ln(2)

	pdf.SetFont("Arial", "", 10)
	txt := strings.TrimSpace(body)
	if txt == "" {
		txt = "(Vacío) No hay datos para este reporte."
	}
	// MultiCell soporta saltos de línea y reparte texto largo en varias líneas.
	pdf.MultiCell(0, 5, tr(txt), "", "L", false)
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
			logoW := 56.0
			logoH := 30.0
			logoX := pageW - right - logoW
			logoY := topY + 2
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
		if st, err := os.Stat(v); err == nil && !st.IsDir() {
			return v
		}
	}
	if st, err := os.Stat(defaultHeaderLogoPath); err == nil && !st.IsDir() {
		return defaultHeaderLogoPath
	}
	return ""
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
