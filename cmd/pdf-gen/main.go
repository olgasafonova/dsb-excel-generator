package main

import (
	"fmt"
	"os"

	"dsb-excel-generator/pkg/pdf"
)

func main() {
	// Generate a limited number of PDFs for testing (e.g., 10)
	// Change limit to 2000 for full generation, or 0 for all rows
	if err := pdf.GeneratePDFs("dsb-mock-data-excel.xlsx", "output_pdfs", 10); err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}
}
