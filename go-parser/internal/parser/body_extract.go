package parser

import (
	"bufio"
	"bytes"
	"encoding/base64"
	"fmt"
	"io"
	"mime"
	"mime/multipart"
	"mime/quotedprintable"
	"net/mail"
	"os"
	"strings"
)

// Attachment описывает структуру вложенного файла
type Attachment struct {
	Filename    string
	ContentType string
	Data        []byte
}

// ParsedEmail содержит в себе все извлеченные данные из EML
type ParsedEmail struct {
	From        string
	To          string
	Subject     string
	TextBody    string
	HTMLBody    string
	Attachments []Attachment
}

// ParseEML — главная точка входа для разбора EML файла
func ParseEML(path string) (*ParsedEmail, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("открытие файла: %w", err)
	}
	defer file.Close()

	// 1. Читаем глобальные заголовки
	msg, err := mail.ReadMessage(bufio.NewReader(file))
	if err != nil {
		return nil, fmt.Errorf("чтение заголовков: %w", err)
	}

	result := &ParsedEmail{
		From:    msg.Header.Get("From"),
		To:      msg.Header.Get("To"),
		Subject: msg.Header.Get("Subject"),
	}

	// Декодируем тему письма (заголовки часто зашифрованы в RFC 2047, например =?UTF-8?B?...)
	dec := new(mime.WordDecoder)
	if subj, err := dec.DecodeHeader(result.Subject); err == nil {
		result.Subject = subj
	}

	ct := msg.Header.Get("Content-Type")
	enc := msg.Header.Get("Content-Transfer-Encoding")

	// 2. Начинаем разбор тела письма
	err = parsePart(msg.Body, ct, enc, result)
	if err != nil {
		return nil, fmt.Errorf("разбор тела письма: %w", err)
	}

	return result, nil
}

// parsePart рекурсивно анализирует каждую MIME-часть письма
func parsePart(bodyReader io.Reader, ct, enc string, result *ParsedEmail) error {
	// Если Content-Type не указан, по стандарту RFC это text/plain
	if ct == "" {
		ct = "text/plain; charset=us-ascii"
	}

	mediaType, params, err := mime.ParseMediaType(ct)
	if err != nil {
		return fmt.Errorf("ошибка парсинга медиа-типа: %w", err)
	}

	// Сценарий А: Это контейнер (multipart/*), внутри него другие части
	if strings.HasPrefix(mediaType, "multipart/") {
		boundary := params["boundary"]
		if boundary == "" {
			return fmt.Errorf("отсутствует boundary для %s", mediaType)
		}

		// Передаем reader в multipart.Reader напрямую без вычитки в память!
		mr := multipart.NewReader(bodyReader, boundary)
		for {
			part, err := mr.NextPart()
			if err == io.EOF {
				break // Все части кончились
			}
			if err != nil {
				return fmt.Errorf("ошибка чтения вложенной части: %w", err)
			}

			partCT := part.Header.Get("Content-Type")
			partEnc := part.Header.Get("Content-Transfer-Encoding")

			// РЕКУРСИЯ: уходим вглубь этой части
			if err := parsePart(part, partCT, partEnc, result); err != nil {
				return err
			}
		}
		return nil
	}

	// Сценарий Б: Это конечная часть (Лист дерева: текст, html или файл)
	// 1. Вычитываем сырые данные этой части
	rawBytes, err := io.ReadAll(bodyReader)
	if err != nil {
		return fmt.Errorf("чтение сырых данных части: %w", err)
	}

	// 2. Снимаем транспортную кодировку (Base64 или Quoted-Printable)
	decodedBytes, err := decodeTransferEncoding(rawBytes, enc)
	if err != nil {
		return fmt.Errorf("декодирование транспорта: %w", err)
	}

	// 3. Проверяем, не вложение ли это?
	// Вложения определяются по наличию "filename" в Content-Type или Content-Disposition
	filename := params["name"]
	if filename == "" {
		// Если в Content-Type имени нет, ищем в Content-Disposition
		disp := result.Subject // временный буфер, просто смотрим контекст (на самом деле нам нужен заголовок part)
		_ = disp               // Безопасно игнорируем, ниже точная проверка для конкретной part:
	}

	// Для точной проверки вложения смотрим, есть ли имя файла
	if part, ok := bodyReader.(*multipart.Part); ok {
		filename = part.FileName()
	}

	if filename != "" {
		// Декодируем имя файла, если оно зашифровано в заголовках
		dec := new(mime.WordDecoder)
		if fName, err := dec.DecodeHeader(filename); err == nil {
			filename = fName
		}

		result.Attachments = append(result.Attachments, Attachment{
			Filename:    filename,
			ContentType: mediaType,
			Data:        decodedBytes,
		})
		return nil
	}

	// 4. Раз это не вложение, значит это текст или HTML. Декодируем кодировку (charset)
	text, err := decodeCharset(decodedBytes, params["charset"])
	if err != nil {
		text = string(decodedBytes) // Если упало, пишем как есть
	}

	if mediaType == "text/plain" {
		result.TextBody += text
	} else if mediaType == "text/html" {
		result.HTMLBody += text
	}

	return nil
}

// decodeTransferEncoding обрабатывает Base64 и Quoted-Printable
func decodeTransferEncoding(data []byte, encoding string) ([]byte, error) {
	switch strings.ToLower(strings.TrimSpace(encoding)) {
	case "base64":
		decoded := make([]byte, base64.StdEncoding.DecodedLen(len(data)))
		n, err := base64.StdEncoding.Decode(decoded, data)
		if err != nil {
			return data, err
		}
		return decoded[:n], nil
	case "quoted-printable":
		r := quotedprintable.NewReader(bytes.NewReader(data))
		decoded, err := io.ReadAll(r)
		if err != nil {
			return data, err
		}
		return decoded, nil
	default:
		return data, nil // binary, 7bit, 8bit возвращаем "как есть"
	}
}

// decodeCharset переводит массив байт в UTF-8 строку на основе переданного charset
func decodeCharset(data []byte, charset string) (string, error) {
	switch strings.ToLower(charset) {
	case "utf-8", "us-ascii", "":
		return string(data), nil
	// Если проект коммерческий, сюда стоит подключить "golang.org/x/net/html/charset"
	// Она автоматически превратит "windows-1251" или "koi8-r" в валидный Go UTF-8.
	default:
		return string(data), nil
	}
}
