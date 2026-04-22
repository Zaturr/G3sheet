package internal

import (
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"strconv"
	"strings"
	"time"
)

const (
	cmdAnniversary     = "/20-12"
	replyAnniversary   = "te amo mi amorcito"
	cmdReporteDiario   = "/reporte_diario"
	cmdReporteSemanal  = "/reporte_semanal"
	cmdPendiente       = "/pendiente"
	msgSheetsNoConfig  = "Sheets no configurado: revisá SPREADSHEET_ID y GOOGLE_CREDENTIALS_JSON (o GOOGLE_APPLICATION_CREDENTIALS) en .env."
	msgReporteVacioTpl = "(Vacío) No hubo líneas con texto en columna A de «%s»."
)

// RunUpdatesLoop hace long polling a getUpdates y responde comandos en el mismo chat.
// sh puede ser nil: los comandos de reporte avisarán que falta configuración.
func RunUpdatesLoop(ctx context.Context, c *Client, sh *SheetsReader) error {
	if c == nil || strings.TrimSpace(c.token) == "" {
		return fmt.Errorf("RunUpdatesLoop: client without token")
	}

	var offset int64
	longClient := &http.Client{Timeout: 55 * time.Second}

	for {
		if err := ctx.Err(); err != nil {
			return err
		}

		u := fmt.Sprintf(
			"https://api.telegram.org/bot%s/getUpdates?timeout=50&offset=%d",
			c.token, offset,
		)
		req, err := http.NewRequestWithContext(ctx, http.MethodGet, u, nil)
		if err != nil {
			return err
		}

		res, err := longClient.Do(req)
		if err != nil {
			if ctx.Err() != nil {
				return ctx.Err()
			}
			time.Sleep(2 * time.Second)
			continue
		}

		body, err := io.ReadAll(io.LimitReader(res.Body, 8<<20))
		res.Body.Close()
		if err != nil {
			time.Sleep(time.Second)
			continue
		}

		var wrap struct {
			OK          bool            `json:"ok"`
			Result      json.RawMessage `json:"result"`
			Description string          `json:"description"`
		}
		if err := json.Unmarshal(body, &wrap); err != nil || !wrap.OK {
			time.Sleep(2 * time.Second)
			continue
		}

		var updates []struct {
			UpdateID int64 `json:"update_id"`
			Message  *struct {
				Chat struct {
					ID int64 `json:"id"`
				} `json:"chat"`
				Text string `json:"text"`
			} `json:"message"`
		}
		if err := json.Unmarshal(wrap.Result, &updates); err != nil {
			continue
		}

		for _, up := range updates {
			if up.UpdateID+1 > offset {
				offset = up.UpdateID + 1
			}
			msg := up.Message
			if msg == nil {
				continue
			}
			chat := strconv.FormatInt(msg.Chat.ID, 10)
			cmd := normalizeCommand(msg.Text)
			switch cmd {
			case cmdAnniversary:
				_ = c.SendMessageTo(ctx, chat, replyAnniversary)
			case cmdReporteDiario:
				handleReporteDiario(ctx, c, sh, chat)
			case cmdReporteSemanal:
				handleReporteSemanal(ctx, c, sh, chat)
			case cmdPendiente:
				handlePendiente(ctx, c, sh, chat)
			default:
				// ignorar otros mensajes
			}
		}
	}
}

func handleReporteDiario(ctx context.Context, c *Client, sh *SheetsReader, chat string) {
	if sh == nil {
		_ = c.SendMessageTo(ctx, chat, msgSheetsNoConfig)
		return
	}
	txt, err := sh.ReadReporteDiarioTexto(ctx)
	if err != nil {
		_ = c.SendMessageTo(ctx, chat, "Error leyendo Sheets: "+err.Error())
		return
	}
	if strings.TrimSpace(txt) == "" {
		_ = c.SendMessageTo(ctx, chat, fmt.Sprintf(msgReporteVacioTpl, SheetTabReporteDiario))
		return
	}
	_ = c.SendMessageChunksTo(ctx, chat, txt)
}

func handleReporteSemanal(ctx context.Context, c *Client, sh *SheetsReader, chat string) {
	if sh == nil {
		_ = c.SendMessageTo(ctx, chat, msgSheetsNoConfig)
		return
	}
	txt, err := sh.ReadReporteSemanalTexto(ctx)
	if err != nil {
		_ = c.SendMessageTo(ctx, chat, "Error leyendo Sheets: "+err.Error())
		return
	}
	if strings.TrimSpace(txt) == "" {
		_ = c.SendMessageTo(ctx, chat, fmt.Sprintf(msgReporteVacioTpl, SheetTabReporteSemanal))
		return
	}
	_ = c.SendMessageChunksTo(ctx, chat, txt)
}

func handlePendiente(ctx context.Context, c *Client, sh *SheetsReader, chat string) {
	if sh == nil {
		_ = c.SendMessageTo(ctx, chat, msgSheetsNoConfig)
		return
	}
	txt, err := sh.ReadPendienteTexto(ctx)
	if err != nil {
		_ = c.SendMessageTo(ctx, chat, "Error leyendo Sheets: "+err.Error())
		return
	}
	if strings.TrimSpace(txt) == "" {
		_ = c.SendMessageTo(ctx, chat, "(Vacío) No hubo líneas con texto en la novena hoja.")
		return
	}
	_ = c.SendMessageChunksTo(ctx, chat, txt)
}

// normalizeCommand deja solo el comando base: "/reporte_diario" desde "/reporte_diario@G3D4_bot hola".
func normalizeCommand(text string) string {
	text = strings.TrimSpace(text)
	if text == "" {
		return ""
	}
	if i := strings.IndexByte(text, ' '); i >= 0 {
		text = text[:i]
	}
	if i := strings.IndexByte(text, '\n'); i >= 0 {
		text = text[:i]
	}
	if i := strings.IndexByte(text, '\t'); i >= 0 {
		text = text[:i]
	}
	if i := strings.Index(text, "@"); i >= 0 {
		text = text[:i]
	}
	return text
}
