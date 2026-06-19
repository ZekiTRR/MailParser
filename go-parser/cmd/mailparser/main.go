// Старт программы, запуск всего пайплайна
package main

import (
	"fmt"
	"log"
	"path/filepath"
	"strings"

	"MailParser/internal/parser"
)

func main() {
	folder := "D:\\Code\\MailParser\\input\\eml\\"

	files, err := filepath.Glob(filepath.Join(folder, "*.eml"))
	if err != nil {
		log.Fatalf("ошибка чтения папки: %v", err)
	}

	for _, path := range files {
		fmt.Println(strings.Repeat("=", 60))
		fmt.Println("Файл:", path)

		from, to, _, err := parser.ReadEML(path)
		if err != nil {
			log.Printf("ошибка парсинга %s: %v", path, err)
			continue
		}

		body, err := parser.DecodeBody(path)
		if err != nil {
			log.Printf("ошибка декодирования %s: %v", path, err)
			continue
		}

		fmt.Printf("Отправитель: %s\n", from)
		fmt.Printf("Получатель:  %s\n", to)
		fmt.Printf("Сообщение:\n%s\n", body)
		fmt.Println()
	}
}
