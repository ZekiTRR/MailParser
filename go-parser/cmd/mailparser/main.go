package main

import (
	"flag"
	"fmt"
	"log"
	"path/filepath"
	"strings"

	"MailParser/internal/config"
	"MailParser/internal/parser"
	"MailParser/internal/storage"
)

func main() {
	// Флаг -config задаёт путь к parser.json (дефолт — ../configs/parser.json)
	configPath := flag.String("config", "../configs/parser.json", "путь к файлу конфигурации")
	flag.Parse()

	cfg, err := config.Load(*configPath)
	if err != nil {
		log.Fatalf("ошибка загрузки конфига: %v", err)
	}

	fmt.Printf("Входная папка:   %s\n", cfg.InputDir)
	fmt.Printf("Выходная папка:  %s\n", cfg.OutputDir)
	fmt.Println()

	// Ищем все .eml файлы во входной папке
	files, err := filepath.Glob(filepath.Join(cfg.InputDir, "*.eml"))
	if err != nil {
		log.Fatalf("ошибка чтения папки: %v", err)
	}

	if len(files) == 0 {
		fmt.Println("EML файлы не найдены в", cfg.InputDir)
		return
	}

	fmt.Printf("Найдено EML файлов: %d\n\n", len(files))

	// Обрабатываем каждый EML-файл
	for _, path := range files {
		fmt.Println(strings.Repeat("=", 60))
		fmt.Println("Обработка файла:", filepath.Base(path))
		fmt.Println(strings.Repeat("=", 60))

		// Парсим EML → структура с заголовками и телом
		email, err := parser.ParseEML(path)
		if err != nil {
			log.Printf("ошибка парсинга %s: %v", path, err)
			continue
		}

		// Выводим метаданные
		fmt.Printf("Отправитель: %s\n", email.From)
		fmt.Printf("Получатель:  %s\n", email.To)
		fmt.Printf("Тема письма: %s\n", email.Subject)
		fmt.Println(strings.Repeat("-", 40))

		// Приоритет: plain text → HTML → пусто
		if email.TextBody != "" {
			fmt.Printf("Текстовое сообщение (Plain Text):\n%s\n", email.TextBody)
		} else if email.HTMLBody != "" {
			fmt.Printf("HTML сообщение (вывод сырого кода):\n%s\n", email.HTMLBody)
		} else {
			fmt.Println("[Тело письма пустое или содержит только неподдерживаемый формат]")
		}

		// Дописываем тело письма в cleaned.json
		writeSuccess := storage.Write_to_json(email.TextBody, cfg.OutputDir)
		if writeSuccess {
			fmt.Println("Данные письма успешно записаны в cleaned.json")
		} else {
			fmt.Println("Ошибка при записи данных письма в cleaned.json")
		}

		// Информация о вложениях
		if len(email.Attachments) > 0 {
			fmt.Println(strings.Repeat("-", 40))
			fmt.Printf("Найдено вложений: %d\n", len(email.Attachments))
			for i, attach := range email.Attachments {
				fmt.Printf("  %d. %s (%s) — %d байт\n", i+1, attach.Filename, attach.ContentType, len(attach.Data))
			}
		}
		fmt.Println()
	}
}
