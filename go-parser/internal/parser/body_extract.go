// Извлечение TextBody / HTMLBody через net/mail и mime
package parser

import (
	"bufio"
	"encoding/base64"
	"fmt"
	"io"
	"mime"
	"mime/multipart"
	"net/mail"
	"os"
	"strings"
)

// DecodeBody читает EML-файл и возвращает расшифрованное тело письма
func DecodeBody(path string) (string, error) {
	file, err := os.Open(path)
	if err != nil {
		return "", fmt.Errorf("открытие файла: %w", err)
	}
	defer file.Close()

	msg, err := mail.ReadMessage(bufio.NewReader(file))
	if err != nil {
		return "", fmt.Errorf("чтение заголовков: %w", err)
	}

	ct := msg.Header.Get("Content-Type")
	enc := msg.Header.Get("Content-Transfer-Encoding")

	body, err := io.ReadAll(msg.Body)
	if err != nil {
		return "", fmt.Errorf("чтение тела: %w", err)
	}

	body, err = decodeTransferEncoding(body, enc)
	if err != nil {
		return "", err
	}

	if strings.HasPrefix(ct, "multipart/") {
		return extractMultipart(body, ct)
	}

	return decodeCharset(body, ct)
}

func extractMultipart(body []byte, ct string) (string, error) {
	_, params, err := mime.ParseMediaType(ct)
	if err != nil {
		return "", fmt.Errorf("разбор Content-Type: %w", err)
	}

	boundary := params["boundary"]
	if boundary == "" {
		return "", fmt.Errorf("boundary не найден в %s", ct)
	}

	reader := multipart.NewReader(strings.NewReader(string(body)), boundary)
	var textPlain, fallback string

	for {
		part, err := reader.NextPart()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", fmt.Errorf("чтение части: %w", err)
		}

		partBody, err := io.ReadAll(part)
		if err != nil {
			return "", fmt.Errorf("чтение тела части: %w", err)
		}

		partCT := part.Header.Get("Content-Type")
		partEnc := part.Header.Get("Content-Transfer-Encoding")

		partBody, err = decodeTransferEncoding(partBody, partEnc)
		if err != nil {
			continue
		}

		text, err := decodeCharset(partBody, partCT)
		if err != nil {
			continue
		}

		if strings.HasPrefix(partCT, "text/plain") && textPlain == "" {
			textPlain = text
		}
		if strings.HasPrefix(partCT, "text/html") && fallback == "" {
			fallback = text
		}
	}

	if textPlain != "" {
		return textPlain, nil
	}
	return fallback, nil
}

func decodeTransferEncoding(data []byte, encoding string) ([]byte, error) {
	switch strings.ToLower(encoding) {
	case "base64":
		decoded := make([]byte, base64.StdEncoding.DecodedLen(len(data)))
		n, err := base64.StdEncoding.Decode(decoded, data)
		if err != nil {
			return data, err
		}
		return decoded[:n], nil
	default:
		return data, nil
	}
}

func decodeCharset(data []byte, ct string) (string, error) {
	_, params, _ := mime.ParseMediaType(ct)
	charset := strings.ToLower(params["charset"])

	switch charset {
	case "utf-8", "us-ascii", "":
		return string(data), nil
	default:
		return string(data), nil
	}
}
