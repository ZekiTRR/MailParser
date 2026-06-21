// Запись очищенных данных в cleaned.json
package storage

import (
	"encoding/json"
	"fmt"
	"os"
)


type Message struct {
	Message string `json:"message"`
}

func Write_to_json(message string) bool {

	message_obj := Message{
		Message: message,
	}


	// Если файла нет, O_CREATE его создаст.
	file, err := os.OpenFile("cleaned.json", os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	if err != nil {
		fmt.Println("Error opening file:", err)
		return false
	}
	defer file.Close() // Файл закроется автоматически при выходе из функции

	// Кодируем структуру в JSON и записываем в файл
	if err = json.NewEncoder(file).Encode(message_obj); err != nil {
		fmt.Println("Error encoding JSON:", err)
		return false
	}

	return true
}
