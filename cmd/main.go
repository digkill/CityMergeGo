package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	inputDir := "./input"
	outputDir := "./output"
	os.MkdirAll(outputDir, 0755)

	// Для каждого города накапливаем все строки
	cityData := make(map[string][][]string)

	// Сканируем все файлы
	err := filepath.Walk(inputDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() && strings.HasSuffix(strings.ToLower(info.Name()), ".xlsx") {
			fmt.Printf("📂 Читаем файл: %s\n", path)
			processFile(path, cityData)
		}
		return nil
	})
	if err != nil {
		fmt.Println("❌ Ошибка обхода папки:", err)
		return
	}

	// Сохраняем строго по названию города из E4
	for city, rows := range cityData {
		outPath := filepath.Join(outputDir, fmt.Sprintf("%s.xlsx", city))
		file := excelize.NewFile()
		sheet := file.GetSheetName(0)
		for i, row := range rows {
			cell, _ := excelize.CoordinatesToCellName(1, i+1)
			file.SetSheetRow(sheet, cell, &row)
		}
		if err := file.SaveAs(outPath); err != nil {
			fmt.Printf("❌ Ошибка при сохранении %s: %v\n", outPath, err)
		} else {
			fmt.Printf("✅ Сохранено как: %s\n", outPath)
		}
	}

	fmt.Println("🎉 Все файлы успешно объединены по названию из E4")
}

func processFile(filePath string, cityData map[string][][]string) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println("❌ Ошибка открытия файла:", err)
		return
	}
	defer f.Close()

	sheet := f.GetSheetName(0)
	// Читаем значение из ячейки E4
	city, err := f.GetCellValue(sheet, "E4")
	if err != nil || strings.TrimSpace(city) == "" {
		fmt.Printf("⚠️ В файле %s не удалось найти значение в E4, пропускаем\n", filePath)
		return
	}
	city = strings.TrimSpace(city)

	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println("❌ Ошибка чтения строк:", err)
		return
	}

	// Просто добавляем все строки файла в список города
	cityData[city] = append(cityData[city], rows...)
}
