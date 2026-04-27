package main

import (
	"context"
	"log"
	"os"
	"os/signal"
	"path/filepath"
	"sync"
	"syscall"

	"g3bot/backend/internal"

	"github.com/joho/godotenv"
)

func main() {
	loadDotEnv()

	cfg, err := internal.LoadConfig()
	if err != nil {
		log.Fatalf("config: %v", err)
	}

	client := internal.NewClient(cfg.TelegramBotToken, cfg.TelegramChatID)

	sheetsReader, err := initSheetsReader(cfg)
	if err != nil {
		log.Fatalf("sheets: %v", err)
	}
	src := pickNotificationSource(cfg, sheetsReader)

	ctx, stop := signal.NotifyContext(context.Background(), os.Interrupt, syscall.SIGTERM)
	defer stop()

	var wg sync.WaitGroup
	wg.Add(2)

	go func() {
		defer wg.Done()
		log.Printf("bot listo: envíos cada %s (DISPATCH_EVERY)", cfg.DispatchEvery)
		if err := internal.RunDispatchLoop(ctx, cfg, client, src); err != nil && err != context.Canceled {
			log.Printf("dispatch: %v", err)
		}
	}()

	go func() {
		defer wg.Done()
		log.Println("comandos: /20-12, /reporte_diario, /reporte_semanal, /pdf, /pendiente")
		if err := internal.RunUpdatesLoop(ctx, client, sheetsReader); err != nil && err != context.Canceled {
			log.Printf("updates: %v", err)
		}
	}()

	<-ctx.Done()
	log.Println("cerrando…")
	wg.Wait()
	log.Println("terminado")
}

// loadDotEnv carga el primer .env que exista: junto al ejecutable, en el cwd, o cwd/backend.
// godotenv.Load() solo miraba el directorio desde el que corres el programa; por eso fallaba si no era backend/.
func loadDotEnv() {
	var paths []string
	if exe, err := os.Executable(); err == nil {
		paths = append(paths, filepath.Join(filepath.Dir(exe), ".env"))
	}
	if wd, err := os.Getwd(); err == nil {
		paths = append(paths, filepath.Join(wd, ".env"), filepath.Join(wd, "backend", ".env"))
	}
	paths = append(paths, ".env")

	seen := make(map[string]struct{})
	for _, p := range paths {
		p = filepath.Clean(p)
		if _, dup := seen[p]; dup {
			continue
		}
		seen[p] = struct{}{}
		if st, err := os.Stat(p); err != nil || st.IsDir() {
			continue
		}
		if err := godotenv.Load(p); err != nil {
			log.Fatalf("env: no se pudo leer %s: %v", p, err)
		}
		log.Printf("config: variables desde %s", p)
		return
	}
	log.Print("config: no se encontró archivo .env (se usarán solo variables del sistema si existen)")
}

func initSheetsReader(cfg internal.Config) (*internal.SheetsReader, error) {
	hasGoogle := cfg.GoogleCredentialsJSON != "" || cfg.GoogleCredentialsPath != ""
	if cfg.SpreadsheetID == "" || !hasGoogle {
		return nil, nil
	}
	credBytes, err := internal.ResolveGoogleCredentials(cfg)
	if err != nil {
		return nil, err
	}
	sr, err := internal.NewSheetsReader(
		context.Background(),
		credBytes,
		cfg.SpreadsheetID,
		cfg.SheetsRange,
		cfg.SheetsFirstRowHeader,
	)
	if err != nil {
		return nil, err
	}
	log.Printf("sheets: spreadsheet=%s rango dispatch=%s", cfg.SpreadsheetID, cfg.SheetsRange)
	return sr, nil
}

func pickNotificationSource(cfg internal.Config, sr *internal.SheetsReader) internal.NotificationSource {
	if sr != nil {
		if cfg.SheetsAutoNotify {
			log.Print("sheets: avisos automáticos desde SHEETS_RANGE activos (SHEETS_AUTO_NOTIFY=true)")
			return internal.WithDedupe(sr.AsNotificationSource())
		}
		log.Print("sheets: sin avisos automáticos (SHEETS_AUTO_NOTIFY=false o no definido); usá /reporte_diario y /reporte_semanal")
		return internal.EmptyNotificationSource()
	}
	log.Print("sheets: sin SPREADSHEET_ID o credenciales; dispatch demo (comandos /reporte_* necesitan Sheets)")
	return demoNotifications()
}

// demoNotifications devuelve un aviso solo la primera vez; sirve para probar sin spamear.
func demoNotifications() internal.NotificationSource {
	var sent bool
	return func(ctx context.Context) ([]internal.Notification, error) {
		if sent {
			return nil, nil
		}
		sent = true
		return []internal.Notification{
			{
				Subject:   "G3Bot",
				Body:      "Conexión OK. Sustituye esta fuente por datos de Google Sheets.",
				DedupeKey: "demo-startup",
			},
		}, nil
	}
}
