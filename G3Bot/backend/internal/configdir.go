package internal

import (
	"os"
	"path/filepath"
	"strings"
)

const configDirEnv = "G3BOT_CONFIG_DIR"
const configEnvFile = ".env"

var configDir string

// ConfigDir devuelve la carpeta de configuración detectada (vacía si no hay).
func ConfigDir() string {
	return configDir
}

// InitConfigDir localiza G3Bot/config (o la ruta en G3BOT_CONFIG_DIR).
func InitConfigDir() string {
	if d := strings.TrimSpace(os.Getenv(configDirEnv)); d != "" {
		if isConfigDir(d) {
			configDir = filepath.Clean(d)
			return configDir
		}
	}
	for _, c := range configDirCandidates() {
		if isConfigDir(c) {
			configDir = c
			return configDir
		}
	}
	return ""
}

func isConfigDir(dir string) bool {
	dir = filepath.Clean(dir)
	st, err := os.Stat(filepath.Join(dir, configEnvFile))
	return err == nil && !st.IsDir()
}

func configDirCandidates() []string {
	var out []string
	add := func(list *[]string, s string) {
		s = filepath.Clean(s)
		for _, x := range *list {
			if x == s {
				return
			}
		}
		*list = append(*list, s)
	}
	if exe, err := os.Executable(); err == nil {
		dir := filepath.Dir(exe)
		add(&out, filepath.Join(dir, "config"))
		add(&out, filepath.Join(dir, "..", "config"))
		add(&out, filepath.Join(dir, "..", "..", "config"))
	}
	if wd, err := os.Getwd(); err == nil {
		add(&out, filepath.Join(wd, "config"))
		add(&out, filepath.Join(wd, "..", "config"))
		add(&out, filepath.Join(wd, "G3Bot", "config"))
		add(&out, filepath.Join(wd, "..", "G3Bot", "config"))
	}
	return out
}

// ResolveConfigPath resuelve rutas relativas respecto a la carpeta config.
func ResolveConfigPath(p string) string {
	p = strings.TrimSpace(p)
	if p == "" || configDir == "" || filepath.IsAbs(p) {
		return p
	}
	return filepath.Join(configDir, p)
}

// ResolveAssetPath busca un archivo relativo (config, cwd, exe).
func ResolveAssetPath(p string) string {
	p = strings.TrimSpace(p)
	if p == "" {
		return ""
	}
	if filepath.IsAbs(p) {
		return p
	}
	candidates := []string{p, ResolveConfigPath(p)}
	if wd, err := os.Getwd(); err == nil {
		candidates = append(candidates,
			filepath.Join(wd, p),
			filepath.Join(wd, "..", p),
		)
	}
	if exe, err := os.Executable(); err == nil {
		dir := filepath.Dir(exe)
		candidates = append(candidates,
			filepath.Join(dir, p),
			filepath.Join(dir, "..", p),
		)
	}
	seen := make(map[string]struct{})
	for _, c := range candidates {
		c = filepath.Clean(c)
		if _, ok := seen[c]; ok {
			continue
		}
		seen[c] = struct{}{}
		if st, err := os.Stat(c); err == nil && !st.IsDir() {
			return c
		}
	}
	return p
}
