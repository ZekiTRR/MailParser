package main

import (
	"fmt"
	"log"
	"path/filepath"
	"strings"
	"MailParser/internal/storage"

	"MailParser/internal/parser"
)

func main() {
	folder := "D:\\Code\\MailParser\\input\\eml\\"

	// Находим все EML файлы в папке
	files, err := filepath.Glob(filepath.Join(folder, "*.eml"))
	if err != nil {
		log.Fatalf("ошибка чтения папки: %v", err)
	}

	for _, path := range files { // Отбрасываем первый аргумент
		fmt.Println(strings.Repeat("=", 60))
		fmt.Println("Обработка файла:", filepath.Base(path))
		fmt.Println(strings.Repeat("=", 60))

		// Вызываем ОДНУ функцию вместо двух. Она делает всю работу за один сеанс I/O.
		email, err := parser.ParseEML(path)
		if err != nil {
			log.Printf("ошибка парсинга %s: %v", path, err)
			continue
		}

		// Выводим метаданные письма
		fmt.Printf("Отправитель: %s\n", email.From)
		fmt.Printf("Получатель:  %s\n", email.To)
		fmt.Printf("Тема письма: %s\n", email.Subject)
		fmt.Println(strings.Repeat("-", 40))

		// Логика отображения тела: приоритет отдаем чистому тексту, если его нет — выводим HTML
		if email.TextBody != "" {
			fmt.Printf("Текстовое сообщение (Plain Text):\n%s\n", email.TextBody)
		} else if email.HTMLBody != "" {
			fmt.Printf("HTML сообщение (вывод сырого кода):\n%s\n", email.HTMLBody)
		} else {
			fmt.Println("[Тело письма пустое или содержит только неподдерживаемый формат]")
		}

		writeSuccess := storage.Write_to_json(email.TextBody)
		if writeSuccess {
			fmt.Println("Данные письма успешно записаны в cleaned.json")
		} else {
			fmt.Println("Ошибка при записи данных письма в cleaned.json")
		}

		// Бонус: выводим информацию о вложениях, если они есть
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
