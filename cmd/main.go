package main

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	inputDir := "./input"
	outputDir := "./output"
	os.MkdirAll(outputDir, 0755)

	// Собираем данные по городам
	cityData := make(map[string][][]string)

	err := filepath.Walk(inputDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() && strings.HasSuffix(strings.ToLower(info.Name()), ".xlsx") {
			fmt.Printf("📂 Обработка файла: %s\n", path)
			processFile(path, cityData)
		}
		return nil
	})
	if err != nil {
		fmt.Println("❌ Ошибка обхода:", err)
		return
	}

	// Сохраняем результат по городам
	for city, rows := range cityData {
		filename := fmt.Sprintf("%s.xlsx", city)
		outPath := filepath.Join(outputDir, filename)

		f := excelize.NewFile()
		sheet := f.GetSheetName(0)

		for i, row := range rows {
			cell, _ := excelize.CoordinatesToCellName(1, i+1)
			f.SetSheetRow(sheet, cell, &row)
		}

		if err := f.SaveAs(outPath); err != nil {
			fmt.Printf("❌ Ошибка сохранения %s: %v\n", filename, err)
		} else {
			fmt.Printf("✅ Сохранено: %s\n", filename)
		}
	}

	fmt.Println("🎉 Все файлы собраны по городам!")
}

func processFile(filePath string, cityData map[string][][]string) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println("❌ Ошибка открытия файла:", err)
		return
	}
	defer f.Close()

	sheet := "all"
	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Printf("❌ Ошибка чтения строк в %s: %v\n", filePath, err)
		return
	}

	for i, row := range rows {
		if len(row) < 9 {
			continue
		}

		raw := row[8] // колонка I, город
		city := extractCityName(raw)
		if city == "" {
			continue
		}

		cityData[city] = append(cityData[city], row)

		// Заголовки (1-я строка) добавляются только 1 раз
		if i == 0 {
			continue
		}
	}
}

// Функция для извлечения названия города
func extractCityName(s string) string {
	// Примеры: "МБОУ «Средняя ОШ № 32» Кировского р-на г. Казани"
	// → "Казань", "г. Краснодар" → "Краснодар"
	re := regexp.MustCompile(`(?i)(г\.?|город)\s*([А-Яа-яA-Za-z-]+)`)
	matches := re.FindStringSubmatch(s)
	if len(matches) >= 3 {
		return normalizeCity(matches[2])
	}

	// Альтернатива: попытка извлечь последнее слово в строке
	parts := strings.Fields(s)
	if len(parts) > 0 {
		last := parts[len(parts)-1]
		return normalizeCity(last)
	}

	return ""
}

// Приводим к нормальной форме
func normalizeCity(city string) string {
	city = strings.ToLower(city)
	city = strings.Trim(city, ".,«»\"'")
	city = strings.Title(city)
	return city
}
