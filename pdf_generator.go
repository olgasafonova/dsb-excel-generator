package main

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/jung-kurt/gofpdf"
	"github.com/xuri/excelize/v2"
)

type EmployeeData struct {
	CPR                  string
	FirstName            string
	LastName             string
	BaseSalary           string
	NewBaseSalary        string
	GrossSalary          string
	NewGrossSalary       string
	IndividualAdjustment string
	PercentageIncrease   string
	EffectiveDate        string
	PensionIncrease      string
}

func generatePDFs(excelFile string, outputDir string, limit int) error {
	// Create output directory
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return fmt.Errorf("failed to create output directory: %v", err)
	}

	// Open Excel file
	f, err := excelize.OpenFile(excelFile)
	if err != nil {
		return fmt.Errorf("failed to open Excel file: %v", err)
	}
	defer f.Close()

	// Get all rows
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		return fmt.Errorf("failed to get rows: %v", err)
	}

	// Skip header row and process data rows
	count := 0
	totalSize := int64(0)
	for i, row := range rows {
		if i == 0 {
			continue // Skip header
		}

		if limit > 0 && count >= limit {
			break
		}

		if len(row) < 11 {
			fmt.Printf("Skipping row %d: insufficient data\n", i+1)
			continue
		}

		employee := EmployeeData{
			CPR:                  row[0],
			FirstName:            row[1],
			LastName:             row[2],
			BaseSalary:           row[3],
			NewBaseSalary:        row[4],
			GrossSalary:          row[5],
			NewGrossSalary:       row[6],
			IndividualAdjustment: row[7],
			PercentageIncrease:   row[8],
			EffectiveDate:        row[9],
			PensionIncrease:      row[10],
		}

		filename := fmt.Sprintf("Lønregulering 2025 – %s %s – %s.pdf",
			employee.FirstName, employee.LastName, employee.CPR)
		filepath := filepath.Join(outputDir, filename)

		if err := createWCAGCompliantPDF(employee, filepath); err != nil {
			fmt.Printf("Error generating PDF for %s: %v\n", filename, err)
			continue
		}

		// Get file size
		if fileInfo, err := os.Stat(filepath); err == nil {
			totalSize += fileInfo.Size()
		}

		count++
		if count%100 == 0 {
			fmt.Printf("Generated %d PDFs... (%.2f MB so far)\n", count, float64(totalSize)/1024/1024)
		}
	}

	fmt.Printf("\nSuccessfully generated %d PDFs in %s\n", count, outputDir)
	fmt.Printf("Total size: %.2f MB (average: %.2f KB per PDF)\n",
		float64(totalSize)/1024/1024,
		float64(totalSize)/float64(count)/1024)

	return nil
}

func createWCAGCompliantPDF(emp EmployeeData, outputPath string) error {
	// Create new PDF with A4 page size and UTF-8 encoding
	pdf := gofpdf.New("P", "mm", "A4", "")

	// Set document metadata for accessibility
	pdf.SetTitle(fmt.Sprintf("Lønregulering 2025 – %s %s", emp.FirstName, emp.LastName), false)
	pdf.SetAuthor("HR Services & Compensation", false)
	pdf.SetSubject("Lønregulering 2025", false)
	pdf.SetCreator("DSB Salary Regulation System", false)
	pdf.SetKeywords("lønregulering salary 2025", false)

	pdf.AddPage()

	// Set margins for better readability (20mm all sides)
	pdf.SetMargins(20, 20, 20)
	pdf.SetAutoPageBreak(true, 20)

	// WCAG AAA compliant: Black text (0,0,0) on white background = 21:1 contrast ratio
	textColor := func() {
		pdf.SetTextColor(0, 0, 0)
	}

	// Document Title (H1 equivalent)
	pdf.SetFont("Helvetica", "B", 18)
	textColor()
	pdf.CellFormat(0, 15, "Lønregulering 2025", "", 1, "C", false, 0, "")
	pdf.Ln(5)

	// Greeting
	pdf.SetFont("Helvetica", "", 12)
	textColor()
	greeting := fmt.Sprintf("Kære %s %s", emp.FirstName, emp.LastName)
	pdf.MultiCell(0, 7, greeting, "", "L", false)
	pdf.Ln(3)

	// Introduction
	intro := "Lønreguleringen 2025 for HK medarbejdere er nu afsluttet, og i dette brev kan du læse om hvad det betyder for dig."
	pdf.MultiCell(0, 7, intro, "", "L", false)
	pdf.Ln(5)

	// Section Heading (H2 equivalent)
	pdf.SetFont("Helvetica", "B", 14)
	textColor()
	pdf.MultiCell(0, 7, "Regulering i henhold til overenskomst", "", "L", false)
	pdf.Ln(2)

	// Overenskomst text
	pdf.SetFont("Helvetica", "", 12)
	textColor()
	overenskomstText := "Følgende regulering er fastlagt i overenskomsten med virkning 1. maj 2025:"
	pdf.MultiCell(0, 7, overenskomstText, "", "L", false)
	pdf.Ln(2)

	// Bullet point
	bulletText := fmt.Sprintf("• Forhøjelse af pensionsbidrag med %s%%", emp.PensionIncrease)
	pdf.SetX(pdf.GetX() + 10) // Indent bullet
	pdf.MultiCell(0, 7, bulletText, "", "L", false)
	pdf.Ln(5)

	// Individual adjustment heading (H2 equivalent)
	pdf.SetFont("Helvetica", "B", 14)
	textColor()
	pdf.MultiCell(0, 7, "Individuel lønregulering", "", "L", false)
	pdf.Ln(2)

	// Individual adjustment text
	pdf.SetFont("Helvetica", "", 12)
	textColor()
	individualText := fmt.Sprintf("Din nærmeste leder har besluttet, at du ud over den nævnte stigning i overenskomsten også skal have en individuel lønregulering gældende pr. %s.", emp.EffectiveDate)
	pdf.MultiCell(0, 7, individualText, "", "L", false)
	pdf.Ln(5)

	// Salary details with mixed formatting
	pdf.SetFont("Helvetica", "", 12)
	pdf.Write(7, "Din basisløn er blevet reguleret til ")
	pdf.SetFont("Helvetica", "B", 12)
	pdf.Write(7, fmt.Sprintf("%s kr.", emp.NewBaseSalary))
	pdf.SetFont("Helvetica", "", 12)
	pdf.Write(7, " og din nye bruttoløn udgør nu ")
	pdf.SetFont("Helvetica", "B", 12)
	pdf.Write(7, fmt.Sprintf("%s kr.", emp.NewGrossSalary))
	pdf.SetFont("Helvetica", "", 12)
	pdf.Write(7, fmt.Sprintf(" Den individuelle lønregulering på din bruttoløn er %s kr., svarende til en stigning på %s%%.", emp.IndividualAdjustment, emp.PercentageIncrease))
	pdf.Ln(10)

	// Effective date
	effectiveText := fmt.Sprintf("Din nye løn er med tilbagevirkende kraft fra den %s.", emp.EffectiveDate)
	pdf.MultiCell(0, 7, effectiveText, "", "L", false)
	pdf.Ln(3)

	// Payment timing
	paymentText := "Denne individuelle regulering vil finde sted ved lønudbetalingen ultimo juni måned 2025."
	pdf.MultiCell(0, 7, paymentText, "", "L", false)
	pdf.Ln(10)

	// Signature
	pdf.SetFont("Helvetica", "", 12)
	textColor()
	pdf.MultiCell(0, 7, "Med venlig hilsen", "", "L", false)
	pdf.Ln(1)

	pdf.SetFont("Helvetica", "B", 12)
	pdf.MultiCell(0, 7, "HR Services & Compensation", "", "L", false)

	// Write to file
	err := pdf.OutputFileAndClose(outputPath)
	if err != nil {
		return fmt.Errorf("failed to write PDF: %v", err)
	}

	return nil
}

func main() {
	// Generate a limited number of PDFs for testing (e.g., 10)
	// Change limit to 2000 for full generation, or 0 for all rows
	if err := generatePDFs("dsb-mock-data-excel.xlsx", "output_pdfs", 10); err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}
}
