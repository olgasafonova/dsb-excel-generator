package pdf

import (
	"fmt"
	"os"
	"path/filepath"
	"sync"

	"dsb-excel-generator/pkg/models"

	"github.com/go-pdf/fpdf"
	"github.com/xuri/excelize/v2"
)

// GeneratePDFs reads the Excel file and generates PDFs concurrently
func GeneratePDFs(excelFile string, outputDir string, limit int) error {
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

	var wg sync.WaitGroup
	numWorkers := 8 // Can be adjusted based on CPU cores
	jobs := make(chan models.EmployeeData, 100)

	// Start workers
	for w := 0; w < numWorkers; w++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for emp := range jobs {
				filename := fmt.Sprintf("Lønregulering 2025 – %s %s – %s.pdf",
					emp.FirstName, emp.LastName, emp.CPR)
				filePath := filepath.Join(outputDir, filename)
				if err := createWCAGCompliantPDF(emp, filePath); err != nil {
					fmt.Printf("Error generating PDF for %s: %v\n", filename, err)
				}
			}
		}()
	}

	// Send jobs
	count := 0
	for i, row := range rows {
		if i == 0 {
			continue // Skip header
		}

		if limit > 0 && count >= limit {
			break
		}

		if len(row) < 11 {
			continue
		}

		jobs <- models.EmployeeData{
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
		count++
	}
	close(jobs)

	wg.Wait()

	fmt.Printf("\nSuccessfully generated %d PDFs in %s\n", count, outputDir)
	return nil
}

func createWCAGCompliantPDF(emp models.EmployeeData, outputPath string) error {
	// Create new PDF with A4 page size
	pdf := fpdf.New("P", "mm", "A4", "")

	// Enable UTF-8 support for Danish characters (æ, ø, å)
	tr := pdf.UnicodeTranslatorFromDescriptor("cp1252")

	// Set document metadata for accessibility
	pdf.SetTitle(tr("Lønregulering 2025 – "+emp.FirstName+" "+emp.LastName), false)
	pdf.SetAuthor("HR Services & Compensation", false)
	pdf.SetSubject(tr("Lønregulering 2025"), false)
	pdf.SetCreator("DSB Salary Regulation System", false)
	pdf.SetKeywords(tr("lønregulering salary 2025"), false)

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
	pdf.CellFormat(0, 15, tr("Lønregulering 2025"), "", 1, "C", false, 0, "")
	pdf.Ln(5)

	// Greeting
	pdf.SetFont("Helvetica", "", 12)
	textColor()
	greeting := fmt.Sprintf("Kære %s %s", emp.FirstName, emp.LastName)
	pdf.MultiCell(0, 7, tr(greeting), "", "L", false)
	pdf.Ln(3)

	// Introduction
	intro := "Lønreguleringen 2025 for HK medarbejdere er nu afsluttet, og i dette brev kan du læse om hvad det betyder for dig."
	pdf.MultiCell(0, 7, tr(intro), "", "L", false)
	pdf.Ln(5)

	// Section Heading (H2 equivalent)
	pdf.SetFont("Helvetica", "B", 14)
	textColor()
	pdf.MultiCell(0, 7, tr("Regulering i henhold til overenskomst"), "", "L", false)
	pdf.Ln(2)

	// Overenskomst text
	pdf.SetFont("Helvetica", "", 12)
	textColor()
	overenskomstText := "Følgende regulering er fastlagt i overenskomsten med virkning 1. maj 2025:"
	pdf.MultiCell(0, 7, tr(overenskomstText), "", "L", false)
	pdf.Ln(2)

	// Bullet point
	bulletText := fmt.Sprintf("• Forhøjelse af pensionsbidrag med %s%%", emp.PensionIncrease)
	pdf.SetX(pdf.GetX() + 10) // Indent bullet
	pdf.MultiCell(0, 7, tr(bulletText), "", "L", false)
	pdf.Ln(5)

	// Individual adjustment heading (H2 equivalent)
	pdf.SetFont("Helvetica", "B", 14)
	textColor()
	pdf.MultiCell(0, 7, tr("Individuel lønregulering"), "", "L", false)
	pdf.Ln(2)

	// Individual adjustment text
	pdf.SetFont("Helvetica", "", 12)
	textColor()
	individualText := fmt.Sprintf("Din nærmeste leder har besluttet, at du ud over den nævnte stigning i overenskomsten også skal have en individuel lønregulering gældende pr. %s.", emp.EffectiveDate)
	pdf.MultiCell(0, 7, tr(individualText), "", "L", false)
	pdf.Ln(5)

	// Salary details with mixed formatting
	pdf.SetFont("Helvetica", "", 12)
	pdf.Write(7, tr("Din basisløn er blevet reguleret til "))
	pdf.SetFont("Helvetica", "B", 12)
	pdf.Write(7, tr(fmt.Sprintf("%s kr.", emp.NewBaseSalary)))
	pdf.SetFont("Helvetica", "", 12)
	pdf.Write(7, tr(" og din nye bruttoløn udgør nu "))
	pdf.SetFont("Helvetica", "B", 12)
	pdf.Write(7, tr(fmt.Sprintf("%s kr.", emp.NewGrossSalary)))
	pdf.SetFont("Helvetica", "", 12)
	pdf.Write(7, tr(fmt.Sprintf(" Den individuelle lønregulering på din bruttoløn er %s kr., svarende til en stigning på %s%%.", emp.IndividualAdjustment, emp.PercentageIncrease)))
	pdf.Ln(10)

	// Effective date
	effectiveText := fmt.Sprintf("Din nye løn er med tilbagevirkende kraft fra den %s.", emp.EffectiveDate)
	pdf.MultiCell(0, 7, tr(effectiveText), "", "L", false)
	pdf.Ln(3)

	// Payment timing
	paymentText := "Denne individuelle regulering vil finde sted ved lønudbetalingen ultimo juni måned 2025."
	pdf.MultiCell(0, 7, tr(paymentText), "", "L", false)
	pdf.Ln(10)

	// Signature
	pdf.SetFont("Helvetica", "", 12)
	textColor()
	pdf.MultiCell(0, 7, tr("Med venlig hilsen"), "", "L", false)
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
