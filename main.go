package main

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Configuration struct {
	CSVDir  string `json:"csv_dir"`
	XLSXDir string `json:"xlsx_dir"`
}

func main() {
	configFile, err := os.Open("config.json")
	if err != nil {
		log.Fatalf("error opening config file: %s", err)
	}
	defer configFile.Close()

	var config Configuration
	err = json.NewDecoder(configFile).Decode(&config)
	if err != nil {
		log.Fatalf("error decoding config file: %s", err)
	}

	errorsFile, err := os.OpenFile("errors.log", os.O_RDWR|os.O_CREATE|os.O_APPEND, 0666)
	if err != nil {
		log.Fatalf("error opening errors log file: %s", err)
	}
	defer errorsFile.Close()

	err = filepath.Walk(config.CSVDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			log.Fatalf("error walking filepath: %s", err)
		}

		if strings.ToLower(filepath.Ext(path)) == ".csv" {
			csvFile, err := os.Open(path)
			if err != nil {
				log.Fatalf("error opening CSV file %s: %s", path, err)
			}
			defer csvFile.Close()

			xlsxFile := filepath.Join(config.XLSXDir, strings.TrimSuffix(info.Name(), filepath.Ext(info.Name()))+".xlsx")
			xlsx := excelize.NewFile()

			reader := csv.NewReader(csvFile)
			rowNum := 1
			for {
				record, err := reader.Read()
				if err == io.EOF {
					break
				} else if err != nil {
					log.Fatalf("error reading CSV file %s: %s\n", path, err)
					break
				}
				colNum := 1
				for _, value := range record {
					colName, _ := excelize.ColumnNumberToName(colNum)
					xlsx.SetCellValue("Sheet1", colName+fmt.Sprint(rowNum), value)
					colNum++
				}
				rowNum++
			}

			err = xlsx.SaveAs(xlsxFile)
			if err != nil {
				log.Fatalf("error saving XLSX file %s: %s\n", xlsxFile, err)
			}
		}

		return nil
	})

	if err != nil {
		log.Fatalf("error walking through CSV directory: %s", err)
	}
}
