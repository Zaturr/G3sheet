package internal

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"mime/multipart"
	"net/http"
	"net/url"
	"os"
	"strings"
	"time"
)

// NotificationSource devuelve los mensajes que deben enviarse en un ciclo.
// Más adelante aquí conectarás la lectura de Google Sheets y filtrarás “solo lo nuevo”.
type NotificationSource func(ctx context.Context) ([]Notification, error)

// Client envía mensajes con la Bot API (HTTP, sin librerías extra).
type Client struct {
	http   *http.Client
	token  string
	chatID string
}

func NewClient(token, chatID string) *Client {
	return &Client{
		http:   &http.Client{Timeout: 30 * time.Second},
		token:  strings.TrimSpace(token),
		chatID: strings.TrimSpace(chatID),
	}
}

// LoadConfig lee TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID y opcional DISPATCH_EVERY (ej. 5m, 1h).
func LoadConfig() (Config, error) {
	token := strings.TrimSpace(os.Getenv("TELEGRAM_BOT_TOKEN"))
	if token == "" {
		return Config{}, errors.New("TELEGRAM_BOT_TOKEN is required")
	}
	chatID := strings.TrimSpace(os.Getenv("TELEGRAM_CHAT_ID"))
	if chatID == "" {
		return Config{}, errors.New("TELEGRAM_CHAT_ID is required")
	}
	every := 5 * time.Minute
	if v := strings.TrimSpace(os.Getenv("DISPATCH_EVERY")); v != "" {
		d, err := time.ParseDuration(v)
		if err != nil {
			return Config{}, fmt.Errorf("DISPATCH_EVERY: %w", err)
		}
		if d < time.Second {
			return Config{}, errors.New("DISPATCH_EVERY must be at least 1s")
		}
		every = d
	}

	sheetRange := strings.TrimSpace(os.Getenv("SHEETS_RANGE"))
	if sheetRange == "" {
		sheetRange = "A1:Z200"
	}
	header := true
	if v := strings.TrimSpace(strings.ToLower(os.Getenv("SHEETS_FIRST_ROW_HEADER"))); v == "0" || v == "false" || v == "no" {
		header = false
	}

	// Por defecto off: los avisos por rango SHEETS_RANGE suelen ser ruido si solo querés /reporte_*.
	auto := false
	if v := strings.TrimSpace(strings.ToLower(os.Getenv("SHEETS_AUTO_NOTIFY"))); v == "1" || v == "true" || v == "yes" {
		auto = true
	}

	return Config{
		TelegramBotToken:      token,
		TelegramChatID:        chatID,
		DispatchEvery:         every,
		SpreadsheetID:         strings.TrimSpace(os.Getenv("SPREADSHEET_ID")),
		SheetsRange:           sheetRange,
		GoogleCredentialsPath: strings.TrimSpace(os.Getenv("GOOGLE_APPLICATION_CREDENTIALS")),
		GoogleCredentialsJSON: strings.TrimSpace(os.Getenv("GOOGLE_CREDENTIALS_JSON")),
		SheetsFirstRowHeader:  header,
		SheetsAutoNotify:      auto,
	}, nil
}

// FormatNotification arma el texto que verá el usuario en Telegram.
func FormatNotification(n Notification) string {
	sub := strings.TrimSpace(n.Subject)
	body := strings.TrimSpace(n.Body)
	switch {
	case sub != "" && body != "":
		return sub + "\n\n" + body
	case sub != "":
		return sub
	default:
		return body
	}
}

// SendMessage envía al chat por defecto de configuración (TELEGRAM_CHAT_ID).
func (c *Client) SendMessage(ctx context.Context, text string) error {
	if strings.TrimSpace(c.chatID) == "" {
		return errors.New("telegram client: missing default chat id")
	}
	return c.SendMessageTo(ctx, c.chatID, text)
}

// SendMessageTo envía un mensaje de texto a un chat concreto (id de usuario, grupo, etc.).
func (c *Client) SendMessageTo(ctx context.Context, chatID string, text string) error {
	if strings.TrimSpace(c.token) == "" {
		return errors.New("telegram client: missing token")
	}
	chatID = strings.TrimSpace(chatID)
	if chatID == "" {
		return errors.New("telegram: empty chat_id")
	}
	text = strings.TrimSpace(text)
	if text == "" {
		return errors.New("telegram: empty message")
	}

	endpoint := "https://api.telegram.org/bot" + c.token + "/sendMessage"
	form := url.Values{}
	form.Set("chat_id", chatID)
	form.Set("text", text)
	form.Set("disable_web_page_preview", "true")

	req, err := http.NewRequestWithContext(ctx, http.MethodPost, endpoint, strings.NewReader(form.Encode()))
	if err != nil {
		return err
	}
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")

	res, err := c.http.Do(req)
	if err != nil {
		return err
	}
	defer res.Body.Close()

	body, err := io.ReadAll(io.LimitReader(res.Body, 1<<20))
	if err != nil {
		return err
	}

	var parsed struct {
		OK          bool   `json:"ok"`
		Description string `json:"description"`
	}
	if err := json.Unmarshal(body, &parsed); err != nil {
		return fmt.Errorf("telegram: invalid json (%s): %w", res.Status, err)
	}
	if !parsed.OK {
		return fmt.Errorf("telegram: %s — %s", res.Status, parsed.Description)
	}
	return nil
}

// SendDocumentTo envía un archivo (document) al chat indicado.
func (c *Client) SendDocumentTo(ctx context.Context, chatID, filename string, content []byte) error {
	if strings.TrimSpace(c.token) == "" {
		return errors.New("telegram client: missing token")
	}
	chatID = strings.TrimSpace(chatID)
	if chatID == "" {
		return errors.New("telegram: empty chat_id")
	}
	filename = strings.TrimSpace(filename)
	if filename == "" {
		return errors.New("telegram: empty filename")
	}
	if len(content) == 0 {
		return errors.New("telegram: empty document content")
	}

	endpoint := "https://api.telegram.org/bot" + c.token + "/sendDocument"
	var body bytes.Buffer
	writer := multipart.NewWriter(&body)

	if err := writer.WriteField("chat_id", chatID); err != nil {
		return err
	}
	part, err := writer.CreateFormFile("document", filename)
	if err != nil {
		return err
	}
	if _, err := part.Write(content); err != nil {
		return err
	}
	if err := writer.Close(); err != nil {
		return err
	}

	req, err := http.NewRequestWithContext(ctx, http.MethodPost, endpoint, &body)
	if err != nil {
		return err
	}
	req.Header.Set("Content-Type", writer.FormDataContentType())

	res, err := c.http.Do(req)
	if err != nil {
		return err
	}
	defer res.Body.Close()

	respBody, err := io.ReadAll(io.LimitReader(res.Body, 1<<20))
	if err != nil {
		return err
	}
	var parsed struct {
		OK          bool   `json:"ok"`
		Description string `json:"description"`
	}
	if err := json.Unmarshal(respBody, &parsed); err != nil {
		return fmt.Errorf("telegram: invalid json (%s): %w", res.Status, err)
	}
	if !parsed.OK {
		return fmt.Errorf("telegram: %s — %s", res.Status, parsed.Description)
	}
	return nil
}

// SendMessageChunksTo parte texto largo en varios sendMessage (límite ~4096 de Telegram).
func (c *Client) SendMessageChunksTo(ctx context.Context, chatID string, text string) error {
	chunks := chunkTelegramText(text, 3800)
	for i, ch := range chunks {
		if err := c.SendMessageTo(ctx, chatID, ch); err != nil {
			return fmt.Errorf("parte %d/%d: %w", i+1, len(chunks), err)
		}
		if i < len(chunks)-1 {
			time.Sleep(400 * time.Millisecond)
		}
	}
	return nil
}

func chunkTelegramText(s string, maxRunes int) []string {
	if maxRunes < 200 {
		maxRunes = 3800
	}
	runes := []rune(s)
	if len(runes) <= maxRunes {
		if len(s) == 0 {
			return nil
		}
		return []string{s}
	}
	var out []string
	for len(runes) > 0 {
		if len(runes) <= maxRunes {
			out = append(out, string(runes))
			break
		}
		cut := maxRunes
		minCut := maxRunes * 2 / 3
		if minCut < 1 {
			minCut = 1
		}
		for i := maxRunes - 1; i >= minCut; i-- {
			if runes[i] == '\n' {
				cut = i + 1
				break
			}
		}
		out = append(out, string(runes[:cut]))
		runes = runes[cut:]
	}
	return out
}

func (c *Client) SendNotification(ctx context.Context, n Notification) error {
	return c.SendMessage(ctx, FormatNotification(n))
}

// RunDispatchLoop ejecuta cada DispatchEvery: pide notificaciones a src y las envía una a una.
// Cancela ctx para detener el bucle después del ciclo actual.
func RunDispatchLoop(ctx context.Context, cfg Config, client *Client, src NotificationSource) error {
	if src == nil {
		return errors.New("RunDispatchLoop: nil NotificationSource")
	}
	t := time.NewTicker(cfg.DispatchEvery)
	defer t.Stop()

	runOnce := func() error {
		items, err := src(ctx)
		if err != nil {
			return err
		}
		for _, n := range items {
			if err := client.SendNotification(ctx, n); err != nil {
				return err
			}
		}
		return nil
	}

	if err := runOnce(); err != nil {
		return err
	}
	for {
		select {
		case <-ctx.Done():
			return ctx.Err()
		case <-t.C:
			if err := runOnce(); err != nil {
				return err
			}
		}
	}
}

// LogWriter returns an io.Writer that sends each Write as a Telegram message (for quick debugging).
func (c *Client) LogWriter(ctx context.Context) io.Writer {
	return &telegramLogWriter{ctx: ctx, c: c}
}

type telegramLogWriter struct {
	ctx context.Context
	c   *Client
	buf bytes.Buffer
}

func (w *telegramLogWriter) Write(p []byte) (int, error) {
	n, _ := w.buf.Write(p)
	for {
		b := w.buf.Bytes()
		i := bytes.IndexByte(b, '\n')
		if i < 0 {
			break
		}
		line := strings.TrimSpace(string(b[:i]))
		w.buf.Next(i + 1)
		if line != "" {
			_ = w.c.SendMessage(w.ctx, line)
		}
	}
	return n, nil
}
