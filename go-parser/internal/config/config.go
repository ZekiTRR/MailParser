// Загрузка конфигурации из JSON-файла
package config

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
)

// Config хранит пути к входной и выходной папкам
type Config struct {
	InputDir  string `json:"input_dir"`
	OutputDir string `json:"output_dir"`
}

// Load читает JSON-конфиг и возвращает структуру с путями.
// Относительные пути резолвятся относительно папки с конфигом.
func Load(configPath string) (*Config, error) {
	data, err := os.ReadFile(configPath)
	if err != nil {
		return nil, fmt.Errorf("чтение конфига: %w", err)
	}

	cfg := &Config{}
	if err := json.Unmarshal(data, cfg); err != nil {
		return nil, fmt.Errorf("разбор конфига: %w", err)
	}

	// Папка конфига — база для относительных путей
	configDir := filepath.Dir(configPath)

	// Дефолты, если в JSON не указаны
	if cfg.InputDir == "" {
		cfg.InputDir = filepath.Join("input", "eml")
	}
	if cfg.OutputDir == "" {
		cfg.OutputDir = "output"
	}

	// Превращаем относительные пути в абсолютные
	if !filepath.IsAbs(cfg.InputDir) {
		cfg.InputDir = filepath.Join(configDir, cfg.InputDir)
	}
	if !filepath.IsAbs(cfg.OutputDir) {
		cfg.OutputDir = filepath.Join(configDir, cfg.OutputDir)
	}

	return cfg, nil
}
