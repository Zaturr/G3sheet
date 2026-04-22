package internal

import (
	"bytes"
	"context"
	"fmt"
	"hash/fnv"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"unicode/utf8"

	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

// SheetsReader lee un rango y arma notificaciones legibles para Telegram.
type SheetsReader struct {
	srv           *sheets.Service
	spreadsheetID string
	readRange     string
	firstRowHeader bool
}

// ResolveGoogleCredentials obtiene los bytes del JSON de servicio: inline (Config) o archivo.
func ResolveGoogleCredentials(cfg Config) ([]byte, error) {
	if s := strings.TrimSpace(cfg.GoogleCredentialsJSON); s != "" {
		return []byte(s), nil
	}
	p := strings.TrimSpace(cfg.GoogleCredentialsPath)
	if p == "" {
		return nil, fmt.Errorf("define GOOGLE_CREDENTIALS_JSON o GOOGLE_APPLICATION_CREDENTIALS")
	}
	b, err := readCredentialFile(p)
	if err != nil {
		return nil, fmt.Errorf("leer credenciales google (%s): %w", p, err)
	}
	return b, nil
}

func readCredentialFile(p string) ([]byte, error) {
	if filepath.IsAbs(p) {
		return os.ReadFile(p)
	}
	add := func(list *[]string, s string) {
		s = filepath.Clean(s)
		for _, x := range *list {
			if x == s {
				return
			}
		}
		*list = append(*list, s)
	}
	var candidates []string
	add(&candidates, p)
	if wd, err := os.Getwd(); err == nil {
		add(&candidates, filepath.Join(wd, p))
		add(&candidates, filepath.Join(wd, "..", p))
	}
	if exe, err := os.Executable(); err == nil {
		dir := filepath.Dir(exe)
		add(&candidates, filepath.Join(dir, p))
		add(&candidates, filepath.Join(dir, "..", p))
	}
	var lastErr error
	for _, c := range candidates {
		b, err := os.ReadFile(c)
		if err == nil {
			return b, nil
		}
		lastErr = err
	}
	if lastErr != nil {
		return nil, lastErr
	}
	return nil, fmt.Errorf("archivo no encontrado: %s", p)
}

// NewSheetsReader usa credenciales de cuenta de servicio (JSON en bytes) y el scope de solo lectura.
func NewSheetsReader(ctx context.Context, credJSON []byte, spreadsheetID, readRange string, firstRowHeader bool) (*SheetsReader, error) {
	spreadsheetID = strings.TrimSpace(spreadsheetID)
	if spreadsheetID == "" {
		return nil, fmt.Errorf("spreadsheet id vacío")
	}
	credJSON = bytes.TrimSpace(credJSON)
	if len(credJSON) == 0 {
		return nil, fmt.Errorf("credenciales google vacías")
	}
	creds, err := google.CredentialsFromJSON(ctx, credJSON, sheets.SpreadsheetsReadonlyScope)
	if err != nil {
		return nil, fmt.Errorf("parsear credenciales google: %w", err)
	}
	srv, err := sheets.NewService(ctx, option.WithCredentials(creds))
	if err != nil {
		return nil, fmt.Errorf("cliente sheets: %w", err)
	}
	if strings.TrimSpace(readRange) == "" {
		readRange = "A1:Z200"
	}
	return &SheetsReader{
		srv:            srv,
		spreadsheetID:  spreadsheetID,
		readRange:      readRange,
		firstRowHeader: firstRowHeader,
	}, nil
}

// Pestañas del libro CXP_Control_Semanal: 6 = diario, 7 = semanal (nombres exactos en Google Sheets).
// /reporte_diario → SheetTabReporteDiario | /reporte_semanal → SheetTabReporteSemanal
const (
	SheetTabReporteDiario  = "6 Reporte diario texto"
	SheetTabReporteSemanal = "7 Reporte semanal texto"
	SheetTabPendienteIndex = 8 // base 0: 8 => novena hoja
	reportTextoMaxRow      = 800
)

// quotedSheetRange arma un A1 con la pestaña entre comillas simples (API de Sheets).
func quotedSheetRange(tab, colFrom, colTo string, row1, row2 int) string {
	esc := strings.ReplaceAll(tab, "'", "''")
	return fmt.Sprintf("'%s'!%s%d:%s%d", esc, colFrom, row1, colTo, row2)
}

// ReadRangeConcatenated lee un rango y arma texto como en la hoja: columnas con tab, filas vacías conservadas.
func (r *SheetsReader) ReadRangeConcatenated(ctx context.Context, a1Range string) (string, error) {
	resp, err := r.srv.Spreadsheets.Values.Get(r.spreadsheetID, a1Range).Context(ctx).Do()
	if err != nil {
		return "", err
	}
	var lines []string
	for _, row := range resp.Values {
		lines = append(lines, formatReportRow(row))
	}
	// Quita solo líneas vacías consecutivas al inicio/final del bloque (no las del medio).
	for len(lines) > 0 && lines[0] == "" {
		lines = lines[1:]
	}
	for len(lines) > 0 && lines[len(lines)-1] == "" {
		lines = lines[:len(lines)-1]
	}
	return strings.Join(lines, "\n"), nil
}

// formatReportRow: columnas no vacías hasta la última con dato, unidas con tab; mantiene \n dentro de una celda.
// Fila totalmente vacía → "" (línea en blanco en el mensaje).
func formatReportRow(row []interface{}) string {
	cells := cellsToStrings(row)
	if len(cells) == 0 {
		return ""
	}
	last := -1
	for i := len(cells) - 1; i >= 0; i-- {
		if strings.TrimSpace(cells[i]) != "" {
			last = i
			break
		}
	}
	if last < 0 {
		return ""
	}
	cells = cells[:last+1]
	parts := make([]string, len(cells))
	for i, c := range cells {
		parts[i] = strings.TrimSpace(c)
	}
	return strings.Join(parts, "\t")
}

// ReadReporteDiarioTexto lee A:Z (layout del reporte: títulos en A, fechas en otras columnas, etc.).
func (r *SheetsReader) ReadReporteDiarioTexto(ctx context.Context) (string, error) {
	rng := quotedSheetRange(SheetTabReporteDiario, "A", "Z", 1, reportTextoMaxRow)
	return r.ReadRangeConcatenated(ctx, rng)
}

// ReadReporteSemanalTexto igual que diario, pestaña 7.
func (r *SheetsReader) ReadReporteSemanalTexto(ctx context.Context) (string, error) {
	rng := quotedSheetRange(SheetTabReporteSemanal, "A", "Z", 1, reportTextoMaxRow)
	return r.ReadRangeConcatenated(ctx, rng)
}

// ReadPendienteTexto lee A:Z de la novena hoja del spreadsheet.
func (r *SheetsReader) ReadPendienteTexto(ctx context.Context) (string, error) {
	tab, err := r.readSheetTitleByIndex(ctx, SheetTabPendienteIndex)
	if err != nil {
		return "", err
	}
	rng := quotedSheetRange(tab, "A", "Z", 1, reportTextoMaxRow)
	return r.ReadRangeConcatenated(ctx, rng)
}

func (r *SheetsReader) readSheetTitleByIndex(ctx context.Context, idx int) (string, error) {
	if idx < 0 {
		return "", fmt.Errorf("índice de hoja inválido: %d", idx)
	}
	doc, err := r.srv.Spreadsheets.Get(r.spreadsheetID).Fields("sheets.properties.title").Context(ctx).Do()
	if err != nil {
		return "", fmt.Errorf("leer metadata de hojas: %w", err)
	}
	if len(doc.Sheets) <= idx {
		return "", fmt.Errorf("la hoja %d no existe (solo hay %d hojas)", idx+1, len(doc.Sheets))
	}
	title := strings.TrimSpace(doc.Sheets[idx].Properties.Title)
	if title == "" {
		return "", fmt.Errorf("la hoja %d no tiene título", idx+1)
	}
	return title, nil
}

// AsNotificationSource adapta el lector al tipo que usa RunDispatchLoop.
func (r *SheetsReader) AsNotificationSource() NotificationSource {
	return r.FetchNotifications
}

// FetchNotifications lee el rango y devuelve una notificación por fila de datos no vacía.
func (r *SheetsReader) FetchNotifications(ctx context.Context) ([]Notification, error) {
	resp, err := r.srv.Spreadsheets.Values.Get(r.spreadsheetID, r.readRange).Context(ctx).Do()
	if err != nil {
		return nil, fmt.Errorf("sheets get %q: %w", r.readRange, err)
	}
	if len(resp.Values) == 0 {
		return nil, nil
	}

	var header []string
	rows := resp.Values
	startRowOneBased := 1
	if r.firstRowHeader && len(rows) > 0 {
		header = cellsToStrings(rows[0])
		rows = rows[1:]
		startRowOneBased = 2
	}

	out := make([]Notification, 0, len(rows))
	for i, raw := range rows {
		cells := cellsToStrings(raw)
		if rowIsEmpty(cells) {
			continue
		}
		sheetRow := startRowOneBased + i
		n, ok := buildNotificationFromSheetRow(r.spreadsheetID, sheetRow, header, cells)
		if ok {
			out = append(out, n)
		}
	}
	return out, nil
}

// buildNotificationFromSheetRow arma un aviso compacto: una viñeta y valores separados por · (sin "Col 2:").
func buildNotificationFromSheetRow(spreadsheetID string, sheetRowOneBased int, header, cells []string) (Notification, bool) {
	_ = header // reservado si más adelante querés otra plantilla con cabeceras reales
	if rowIsEmpty(cells) {
		return Notification{}, false
	}
	parts := make([]string, 0, len(cells))
	for _, c := range cells {
		if s := strings.TrimSpace(c); s != "" {
			parts = append(parts, s)
		}
	}
	if len(parts) == 0 {
		return Notification{}, false
	}
	body := "• " + strings.Join(parts, " · ")
	body = truncateRunes(body, 3900)

	key := fmt.Sprintf("%s:r%d:%08x", spreadsheetID, sheetRowOneBased, rowFingerprint(cells))
	return Notification{
		Subject:   "",
		Body:      body,
		DedupeKey: key,
	}, true
}

func cellsToStrings(row []interface{}) []string {
	out := make([]string, 0, len(row))
	for _, v := range row {
		switch t := v.(type) {
		case string:
			out = append(out, t)
		case float64:
			if t == float64(int64(t)) {
				out = append(out, strconv.FormatInt(int64(t), 10))
			} else {
				out = append(out, fmt.Sprint(t))
			}
		case bool:
			out = append(out, strconv.FormatBool(t))
		default:
			out = append(out, fmt.Sprint(t))
		}
	}
	return out
}

func rowIsEmpty(cells []string) bool {
	for _, c := range cells {
		if strings.TrimSpace(c) != "" {
			return false
		}
	}
	return true
}

func rowFingerprint(cells []string) uint32 {
	h := fnv.New32a()
	for _, c := range cells {
		h.Write([]byte(strings.TrimSpace(c)))
		h.Write([]byte{0})
	}
	return h.Sum32()
}

func truncateRunes(s string, max int) string {
	if max <= 0 {
		return ""
	}
	if utf8.RuneCountInString(s) <= max {
		return s
	}
	r := []rune(s)
	if max == 1 {
		return string(r[:1]) + "…"
	}
	return string(r[:max-1]) + "…"
}

// EmptyNotificationSource no envía nada (útil cuando Sheets está solo para comandos /reporte_*).
func EmptyNotificationSource() NotificationSource {
	return func(ctx context.Context) ([]Notification, error) {
		return nil, nil
	}
}

// WithDedupe envuelve una fuente y solo emite cada DedupeKey una vez por vida del proceso.
// Si cambia el contenido de una fila, la clave cambia y se puede volver a notificar.
func WithDedupe(inner NotificationSource) NotificationSource {
	var mu sync.Mutex
	seen := make(map[string]struct{})
	return func(ctx context.Context) ([]Notification, error) {
		items, err := inner(ctx)
		if err != nil {
			return nil, err
		}
		mu.Lock()
		defer mu.Unlock()
		out := make([]Notification, 0, len(items))
		for _, n := range items {
			k := strings.TrimSpace(n.DedupeKey)
			if k == "" {
				out = append(out, n)
				continue
			}
			if _, ok := seen[k]; ok {
				continue
			}
			seen[k] = struct{}{}
			out = append(out, n)
		}
		return out, nil
	}
}
