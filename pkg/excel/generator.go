package excel

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

// Letter types
var letterTypes = []string{
	"Salary Regulation 2025",
	"Pension Change",
	"Contract Amendment",
	"Annual Salary Review",
}

// Departments
var departments = []string{
	"Operations", "Finance", "HR", "IT", "Customer Service",
	"Marketing", "Sales", "Logistics", "Administration", "Legal",
}

// Manager names (using the same name lists)
func getRandomManagerName() string {
	return danishFirstNames[rand.Intn(len(danishFirstNames))] + " " +
		danishLastNames[rand.Intn(len(danishLastNames))]
}

// Document types for P360
var documentTypes = []string{
	"Salary Letter", "Contract Amendment", "Pension Notice", "HR Communication",
}

// Security levels for P360
var securityLevels = []string{
	"Internal", "Confidential", "Strictly Confidential",
}

// Generate creates the Excel file with mock data
func Generate(filename string) error {
	rand.Seed(time.Now().UnixNano())

	f := excelize.NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Set column headers - expanded for P360 integration
	headers := []string{
		"CPR", "FirstName", "LastName", "EmployeeNumber", "Department",
		"BaseSalary", "NewBaseSalary", "GrossSalary", "NewGrossSalary",
		"IndividualAdjustment", "PercentageIncrease", "EffectiveDate", "PensionIncrease",
		"LetterType", "ChangeDescription", "ManagerName", "AdditionalNotes",
		"DocumentType", "CaseNumber", "SecurityLevel", "LetterContent",
	}

	// Write headers
	for i, header := range headers {
		col := getExcelColumn(i)
		f.SetCellValue(sheetName, col+"1", header)
	}

	// Track used CPR numbers to ensure uniqueness
	usedCPRs := make(map[string]bool)

	// Generate 3000 rows of data
	for row := 2; row <= 3001; row++ {
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

		// Generate additional fields
		employeeNumber := fmt.Sprintf("EMP%05d", row-1)
		department := departments[rand.Intn(len(departments))]
		letterType := letterTypes[rand.Intn(len(letterTypes))]
		managerName := getRandomManagerName()
		documentType := documentTypes[rand.Intn(len(documentTypes))]
		caseNumber := fmt.Sprintf("2025-%05d", rand.Intn(99999)+1)
		securityLevel := securityLevels[rand.Intn(len(securityLevels))]

		// Generate change description based on letter type
		var changeDescription string
		switch letterType {
		case "Salary Regulation 2025":
			changeDescription = fmt.Sprintf("Individual salary increase of %.2f%% effective %s", percentageIncrease, effectiveDate)
		case "Pension Change":
			changeDescription = fmt.Sprintf("Pension contribution increase to %.2f%%", pensionIncrease)
		case "Contract Amendment":
			changeDescription = fmt.Sprintf("Contract update with new salary terms from %s", effectiveDate)
		case "Annual Salary Review":
			changeDescription = fmt.Sprintf("Annual review resulting in %.2f%% increase", percentageIncrease)
		}

		// Additional notes - 30% of employees get notes
		additionalNotes := ""
		if rand.Float64() < 0.3 {
			notes := []string{
				"Please confirm receipt by signing and returning this letter",
				"Questions? Contact HR at hr@company.dk",
				"This change was approved by your department manager",
				"No action required from your side",
				"Tax implications will be detailed in your next payslip",
			}
			additionalNotes = notes[rand.Intn(len(notes))]
		}

		// Generate full letter content
		letterContent := generateLetterContent(
			firstName, lastName,
			fmt.Sprintf("%.2f", baseSalary),
			fmt.Sprintf("%.2f", newBaseSalary),
			fmt.Sprintf("%.2f", grossSalary),
			fmt.Sprintf("%.2f", newGrossSalary),
			fmt.Sprintf("%.2f", individualAdjustment),
			fmt.Sprintf("%.2f", percentageIncrease),
			effectiveDate,
			fmt.Sprintf("%.2f", pensionIncrease),
			letterType,
		)

		// Write data to cells using column helper
		col := 0
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), cpr)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), firstName)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), lastName)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), employeeNumber)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), department)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), fmt.Sprintf("%.2f", baseSalary))
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), fmt.Sprintf("%.2f", newBaseSalary))
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), fmt.Sprintf("%.2f", grossSalary))
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), fmt.Sprintf("%.2f", newGrossSalary))
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), fmt.Sprintf("%.2f", individualAdjustment))
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), fmt.Sprintf("%.2f", percentageIncrease))
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), effectiveDate)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), fmt.Sprintf("%.2f", pensionIncrease))
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), letterType)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), changeDescription)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), managerName)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), additionalNotes)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), documentType)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), caseNumber)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), securityLevel)
		col++
		f.SetCellValue(sheetName, getExcelColumn(col)+fmt.Sprintf("%d", row), letterContent)

		if row%100 == 0 {
			fmt.Printf("Generated %d rows...\n", row-1)
		}
	}

	// Set column widths for better readability
	for i := 0; i < len(headers); i++ {
		col := getExcelColumn(i)
		width := 18.0

		// Wider columns for text-heavy fields
		if headers[i] == "LetterContent" {
			width = 80.0
		} else if headers[i] == "ChangeDescription" || headers[i] == "AdditionalNotes" {
			width = 40.0
		}

		f.SetColWidth(sheetName, col, col, width)
	}

	// Save the file
	if err := f.SaveAs(filename); err != nil {
		return fmt.Errorf("error saving file: %v", err)
	}

	fmt.Printf("\nSuccessfully generated %s with 3000 rows of data!\n", filename)
	return nil
}

// getExcelColumn converts column index to Excel column letter(s)
func getExcelColumn(index int) string {
	column := ""
	for index >= 0 {
		column = string(rune('A'+(index%26))) + column
		index = index/26 - 1
	}
	return column
}

// generateCPR generates a fake Danish CPR number in format DDMMYY-XXXX
func generateCPR() string {
	// Generate a random date between 1960 and 2005
	year := rand.Intn(46) + 60 // 60-105 (representing 1960-2005)
	month := rand.Intn(12) + 1
	day := rand.Intn(28) + 1 // Keep it simple, avoid month-specific day validation

	// Generate random 4-digit sequence
	sequence := rand.Intn(9000) + 1000

	return fmt.Sprintf("%02d%02d%02d-%04d", day, month, year, sequence)
}

// generateLetterContent creates the full personalized letter text
func generateLetterContent(firstName, lastName, baseSalary, newBaseSalary, grossSalary, newGrossSalary,
	individualAdjustment, percentageIncrease, effectiveDate, pensionIncrease, letterType string) string {

	fullName := firstName + " " + lastName

	switch letterType {
	case "Salary Regulation 2025":
		return fmt.Sprintf(`Lønregulering 2025

Kære %s

Lønreguleringen 2025 for HK medarbejdere er nu afsluttet, og i dette brev kan du læse om hvad det betyder for dig.

Følgende regulering er fastlagt i overenskomsten med virkning 1. maj 2025:
• Forhøjelse af pensionsbidrag med %s%%

Din nærmeste leder har besluttet, at du ud over den nævnte stigning i overenskomsten også skal have en individuel lønregulering gældende pr. %s.

Din basisløn er blevet reguleret til %s kr. og din nye bruttoløn udgør nu %s kr. Den individuelle lønregulering på din bruttoløn er %s kr., svarende til en stigning på %s%%.

Din nye løn er med tilbagevirkende kraft fra den %s.

Denne individuelle regulering vil finde sted ved lønudbetalingen ultimo juni måned 2025.

Med venlig hilsen
HR Services & Compensation`, fullName, pensionIncrease, effectiveDate, newBaseSalary, newGrossSalary, individualAdjustment, percentageIncrease, effectiveDate)

	case "Pension Change":
		return fmt.Sprintf(`Ændring af pensionsbidrag

Kære %s

Vi ønsker at informere dig om en ændring i dit pensionsbidrag.

Med virkning fra %s vil dit pensionsbidrag blive forhøjet med %s%%.

Din nuværende bruttoløn på %s kr. forbliver uændret. Ændringen påvirker kun pensionsbidraget.

Ændringen er en del af den nye overenskomst og vil fremgå af din næste lønseddel.

Med venlig hilsen
HR Services & Compensation`, fullName, effectiveDate, pensionIncrease, grossSalary)

	case "Contract Amendment":
		return fmt.Sprintf(`Tillæg til ansættelseskontrakt

Kære %s

Dette brev bekræfter ændringer til din ansættelseskontrakt med virkning fra %s.

Dine lønvilkår opdateres som følger:
- Ny basisløn: %s kr.
- Ny bruttoløn: %s kr.

Alle andre vilkår i din ansættelseskontrakt forbliver uændrede.

Med venlig hilsen
HR Services & Compensation`, fullName, effectiveDate, newBaseSalary, newGrossSalary)

	case "Annual Salary Review":
		return fmt.Sprintf(`Årlig lønregulering

Kære %s

Som en del af vores årlige lønregulering har vi glæden af at meddele dig følgende ændringer med virkning fra %s.

Din basisløn forhøjes fra %s kr. til %s kr., hvilket svarer til en stigning på %s%%.

Din nye bruttoløn vil udgøre %s kr.

Denne stigning er baseret på din præstation og udvikling i det forløbne år.

Med venlig hilsen
HR Services & Compensation`, fullName, effectiveDate, baseSalary, newBaseSalary, percentageIncrease, newGrossSalary)

	default:
		return "Letter content not available"
	}
}
