// Запись очищенных данных в cleaned.json (JSONL: один объект на строку)
package storage

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
)

// Message — один объект в JSONL-файле
type Message struct {
	Message string `json:"message"`
}

// Write_to_json дописывает сообщение в конец cleaned.json.
// Формат: каждое сообщение — отдельная строка JSON.
func Write_to_json(message string, outputDir string) bool {
	// Создаём папку输出, если её ещё нет
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		fmt.Println("Ошибка создания папки вывода:", err)
		return false
	}

	message_obj := Message{
		Message: message,
	}

	filePath := filepath.Join(outputDir, "cleaned.json")

	// O_APPEND — дописываем в конец, не затирая предыдущие сообщения
	file, err := os.OpenFile(filePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	if err != nil {
		fmt.Println("Ошибка открытия файла:", err)
		return false
	}
	defer file.Close()

	if err = json.NewEncoder(file).Encode(message_obj); err != nil {
		fmt.Println("Ошибка кодирования JSON:", err)
		return false
	}

	return true
}
