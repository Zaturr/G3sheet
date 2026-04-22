package internal

import "time"

// Config agrupa parámetros del servicio leídos de entorno.
type Config struct {
	TelegramBotToken string
	TelegramChatID   string
	// DispatchEvery es cada cuánto se vuelve a evaluar si hay algo que enviar.
	DispatchEvery time.Duration

	// Sheets (opcional). Si SpreadsheetID está vacío, no se usa la API de Sheets.
	SpreadsheetID string
	SheetsRange   string
	// GoogleCredentialsPath: ruta al JSON de cuenta de servicio (GOOGLE_APPLICATION_CREDENTIALS).
	GoogleCredentialsPath string
	// GoogleCredentialsJSON: mismo JSON pegado en el .env (GOOGLE_CREDENTIALS_JSON). Si ambos están, gana este.
	GoogleCredentialsJSON string
	// SheetsFirstRowHeader: la fila 1 del rango son títulos de columnas.
	SheetsFirstRowHeader bool
	// SheetsAutoNotify: avisos periódicos desde SHEETS_RANGE (por defecto off en LoadConfig; activar con SHEETS_AUTO_NOTIFY=true).
	SheetsAutoNotify bool
}

// Notification es una unidad de mensaje lista para Telegram (viene de Sheets u otra fuente).
type Notification struct {
	Subject string
	Body    string
	// DedupeKey identifica el mismo aviso entre ejecuciones (útil cuando conectes Sheets).
	DedupeKey string
}
