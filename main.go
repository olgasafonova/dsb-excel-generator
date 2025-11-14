package main

import (
	"fmt"
	"math/rand"
	"time"

	"github.com/xuri/excelize/v2"
)

// Danish first names
var danishFirstNames = []string{
	"Anders", "Anne", "Bent", "Birthe", "Carl", "Charlotte", "Christian", "Christine",
	"Emma", "Erik", "Finn", "Freja", "Hans", "Hanne", "Henrik", "Ida", "Jens", "Julie",
	"Karen", "Kasper", "Lars", "Laura", "Lone", "Mads", "Maria", "Martin", "Mette",
	"Michael", "Morten", "Niels", "Ole", "Peter", "Pia", "Rasmus", "Sofie", "Soren",
	"Susanne", "Thomas", "Tina", "Torben", "William", "Sofia", "Noah", "Oliver", "Ella",
	"Lucas", "Victor", "Alma", "Clara", "Alfred", "Oscar", "Agnes", "Karl", "Viggo",
	"Anna", "Jakob", "Mathilde", "Magnus", "Isabella", "Alexander", "Josefine", "Sebastian",
	"Caroline", "Frederik", "Emilie", "Mikkel", "Katrine", "Tobias", "Louise", "Jonas",
	"Camilla", "Andreas", "Cecilie", "Nikolaj", "Sara", "Kristian", "Maja", "Simon",
	"Nanna", "Daniel", "Signe", "Jesper", "Lærke", "Mathias", "Astrid", "Philip",
	"Ellen", "Benjamin", "Liv", "Anton", "Marie", "Gustav", "Johanne", "Valdemar",
}

// Danish last names
var danishLastNames = []string{
	"Jensen", "Nielsen", "Hansen", "Pedersen", "Andersen", "Christensen", "Larsen",
	"Sorensen", "Rasmussen", "Jorgensen", "Petersen", "Madsen", "Kristensen",
	"Olsen", "Thomsen", "Christiansen", "Poulsen", "Johansen", "Moller", "Mortensen",
	"Knudsen", "Jacobsen", "Frederiksen", "Lund", "Eriksen", "Schmidt", "Holm",
	"Bertelsen", "Andreasen", "Iversen", "Laursen", "Berg", "Christoffersen",
	"Clausen", "Simonsen", "Henriksen", "Svendsen", "Vestergaard", "Ostergaard",
	"Dahl", "Mogensen", "Villadsen", "Frandsen", "Mikkelsen", "Lorentzen", "Bruun",
	"Kofoed", "Danielsen", "Thygesen", "Nygaard", "Winther", "Holst", "Rosendahl",
}

func main() {
	rand.Seed(time.Now().UnixNano())

	f := excelize.NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Set column headers
	headers := []string{"CPR", "FirstName", "LastName", "BaseSalary", "NewBaseSalary",
		"GrossSalary", "NewGrossSalary", "IndividualAdjustment", "PercentageIncrease",
		"EffectiveDate", "PensionIncrease"}

	// Write headers
	for i, header := range headers {
		cell := fmt.Sprintf("%s1", string(rune('A'+i)))
		f.SetCellValue(sheetName, cell, header)
	}

	// Track used CPR numbers to ensure uniqueness
	usedCPRs := make(map[string]bool)

	// Generate 2000 rows of data
	for row := 2; row <= 2001; row++ {
		// Generate unique CPR number (DDMMYY-XXXX)
		var cpr string
		for {
			cpr = generateCPR()
			if !usedCPRs[cpr] {
				usedCPRs[cpr] = true
				break
			}
		}

		// Generate name
		firstName := danishFirstNames[rand.Intn(len(danishFirstNames))]
		lastName := danishLastNames[rand.Intn(len(danishLastNames))]

		// Generate salaries (25,000 - 75,000 kr)
		// Add more realistic variation with decimal places
		baseSalary := float64(rand.Intn(50000)+25000) + rand.Float64()*1000

		// Calculate individual adjustment (0.5% to 5% increase)
		// Some employees get higher increases
		percentageIncrease := 0.5 + rand.Float64()*4.5
		// Round to 2 decimal places for realism
		percentageIncrease = float64(int(percentageIncrease*100)) / 100

		individualAdjustment := baseSalary * (percentageIncrease / 100)
		newBaseSalary := baseSalary + individualAdjustment

		// Gross salary includes some additional compensation (about 10-25% more)
		// Variation depends on seniority/role
		additionalComp := 1.1 + rand.Float64()*0.15
		grossSalary := baseSalary * additionalComp
		newGrossSalary := newBaseSalary * additionalComp

		// Pension increase - mostly 1.00% but some variation for different agreements
		pensionIncrease := 1.00
		if rand.Float64() < 0.1 { // 10% get different pension rates
			pensionIncrease = 0.5 + rand.Float64()*1.5 // 0.5% to 2%
			pensionIncrease = float64(int(pensionIncrease*100)) / 100
		}

		// Effective date - most are 1. marts 2025, but some have different dates
		effectiveDates := []string{
			"1. marts 2025", "1. marts 2025", "1. marts 2025", "1. marts 2025",
			"1. april 2025", "1. maj 2025", "1. februar 2025",
		}
		effectiveDate := effectiveDates[rand.Intn(len(effectiveDates))]

		// Write data to cells
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), cpr)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), firstName)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), lastName)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), fmt.Sprintf("%.2f", baseSalary))
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", row), fmt.Sprintf("%.2f", newBaseSalary))
		f.SetCellValue(sheetName, fmt.Sprintf("F%d", row), fmt.Sprintf("%.2f", grossSalary))
		f.SetCellValue(sheetName, fmt.Sprintf("G%d", row), fmt.Sprintf("%.2f", newGrossSalary))
		f.SetCellValue(sheetName, fmt.Sprintf("H%d", row), fmt.Sprintf("%.2f", individualAdjustment))
		f.SetCellValue(sheetName, fmt.Sprintf("I%d", row), fmt.Sprintf("%.2f", percentageIncrease))
		f.SetCellValue(sheetName, fmt.Sprintf("J%d", row), effectiveDate)
		f.SetCellValue(sheetName, fmt.Sprintf("K%d", row), fmt.Sprintf("%.2f", pensionIncrease))

		if row%100 == 0 {
			fmt.Printf("Generated %d rows...\n", row-1)
		}
	}

	// Auto-fit columns for better readability
	for i := 0; i < len(headers); i++ {
		col := string(rune('A' + i))
		f.SetColWidth(sheetName, col, col, 18)
	}

	// Save the file
	filename := "dsb-mock-data-excel.xlsx"
	if err := f.SaveAs(filename); err != nil {
		fmt.Printf("Error saving file: %v\n", err)
		return
	}

	fmt.Printf("\nSuccessfully generated %s with 2000 rows of data!\n", filename)
}

// generateCPR generates a fake Danish CPR number in format DDMMYY-XXXX
func generateCPR() string {
	// Generate a random date between 1960 and 2005
	year := rand.Intn(46) + 60  // 60-105 (representing 1960-2005)
	month := rand.Intn(12) + 1
	day := rand.Intn(28) + 1 // Keep it simple, avoid month-specific day validation

	// Generate random 4-digit sequence
	sequence := rand.Intn(9000) + 1000

	return fmt.Sprintf("%02d%02d%02d-%04d", day, month, year, sequence)
}
