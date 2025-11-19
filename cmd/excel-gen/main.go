package main

import (
	"fmt"
	"os"

	"dsb-excel-generator/pkg/excel"
)

func main() {
	filename := "dsb-mock-data-excel.xlsx"
	if err := excel.Generate(filename); err != nil {
		fmt.Printf("Error generating Excel file: %v\n", err)
		os.Exit(1)
	}
}
