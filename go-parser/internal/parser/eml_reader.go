// Чтение .eml файлов из папки
package parser

import (
	"fmt"
	"os"
	"strings"

	"github.com/DusanKasan/parsemail"
)

// ReadEML читает и парсит EML-файл, возвращая отправителя, получателей и текст сообщения
func ReadEML(path string) (from string, to string, body string, err error) {
	file, err := os.Open(path) // Открывает файл
	if err != nil {
		return "", "", "", fmt.Errorf("открытие файла: %w", err)
	}

	defer file.Close()

	email, err := parsemail.Parse(file)
	if err != nil {
		return "", "", "", fmt.Errorf("парсинг письма: %w", err)
	}

	// Отправитель
	if len(email.From) > 0 {
		from = email.From[0].String()
	}

	// Получатели
	var toAddrs []string
	for _, addr := range email.To {
		toAddrs = append(toAddrs, addr.String())
	}
	to = strings.Join(toAddrs, ", ")

	// Текст сообщения
	if email.TextBody != "" {
		body = email.TextBody
	} else {
		body = email.HTMLBody
	}

	return from, to, body, nil
}
