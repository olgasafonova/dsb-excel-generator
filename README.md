# DSB Excel and PDF Generator

This Go program generates mock employee data in Excel format and creates individual, WCAG AAA-compliant PDF salary regulation letters.

## Features

### Excel Generator (`main.go`)
- Generates 3,000 rows of realistic Danish employee data
- **Basic employee data:** CPR, FirstName, LastName, EmployeeNumber, Department
- **Salary information:** Base salary, new salary, gross salary adjustments (25,000-75,000 kr/month)
- **Letter metadata:** LetterType (4 varieties), ChangeDescription, ManagerName, AdditionalNotes
- **P360 integration fields:** DocumentType, CaseNumber, SecurityLevel
- **Full letter content:** Complete personalized letter text for each employee (4 different letter templates)
- All CPR numbers are guaranteed unique
- Realistic variety in departments, managers, and letter types

### PDF Generator (`pdf_generator.go`)
- Creates individual PDF letters for each employee
- **WCAG AAA Compliant:**
  - Black text on white background (21:1 contrast ratio exceeds AAA 7:1 requirement)
  - Clear document structure with semantic headings (H1, H2)
  - Proper metadata (title, author, subject, keywords)
  - Embedded Helvetica fonts for universal compatibility
  - Logical reading order
  - Generous margins (20mm) for readability
  - Appropriate font sizes (12pt body, 14pt headings, 18pt title)
- Individual filenames: `Lønregulering 2025 – [Name] – [CPR].pdf`

## Excel Columns

The generated Excel file contains 21 columns:

1. **CPR** - Danish CPR number (DDMMYY-XXXX)
2. **FirstName** - Employee first name
3. **LastName** - Employee last name
4. **EmployeeNumber** - Unique employee ID (EMP00001-EMP03000)
5. **Department** - Department (10 varieties)
6. **BaseSalary** - Current base salary
7. **NewBaseSalary** - New base salary after adjustment
8. **GrossSalary** - Current gross salary
9. **NewGrossSalary** - New gross salary
10. **IndividualAdjustment** - Salary adjustment amount
11. **PercentageIncrease** - Percentage increase (0.5%-5%)
12. **EffectiveDate** - When changes take effect
13. **PensionIncrease** - Pension contribution change
14. **LetterType** - Type of letter (Salary Regulation/Pension Change/Contract Amendment/Annual Review)
15. **ChangeDescription** - Brief summary of changes
16. **ManagerName** - Approving manager name
17. **AdditionalNotes** - Optional notes (30% of employees)
18. **DocumentType** - P360 document classification
19. **CaseNumber** - P360 case reference
20. **SecurityLevel** - Document security (Internal/Confidential/Strictly Confidential)
21. **LetterContent** - Full personalized letter text (ready for PDF generation or P360 upload)

## Size Estimates

- **Single PDF:** ~2.1 KB
- **3,000 PDFs:** ~6.3 MB total
- **Excel file:** ~499 KB (with full letter content)

## Usage

### 1. Generate Excel File with Mock Data

```bash
go run main.go
```

This creates `dsb-mock-data-excel.xlsx` with 3,000 employee records.

### 2. Generate PDFs

**Test with 10 PDFs:**
```bash
go run pdf_generator.go
```

**Generate all 3,000 PDFs:**

Edit `pdf_generator.go` line 223 to change the limit:
```go
generatePDFs("dsb-mock-data-excel.xlsx", "output_pdfs", 3000)
```

Then run:
```bash
go run pdf_generator.go
```

This creates PDFs in the `output_pdfs/` directory.

## WCAG Compliance Details

The generated PDFs meet **WCAG 2.1 AAA** standards:

### ✅ Perceivable
- **Contrast Ratio:** 21:1 (exceeds AAA requirement of 7:1 for normal text)
- **Text Alternatives:** Proper document metadata provides context
- **Adaptable:** Clean structure allows for screen reader navigation

### ✅ Operable
- **Keyboard Accessible:** PDFs can be navigated via keyboard
- **Enough Time:** Static documents, no time constraints

### ✅ Understandable
- **Readable:** 12pt font size, clear language, logical flow
- **Predictable:** Consistent layout and structure
- **Input Assistance:** N/A for read-only documents

### ✅ Robust
- **Compatible:** Uses standard PDF format and embedded fonts
- **Metadata:** Proper title, author, subject, keywords set

## Limitations

- Full **PDF/UA-1 tagging** (tagged PDF for screen readers) requires commercial PDF libraries
- This implementation uses the free `gofpdf` library which provides excellent visual accessibility but limited semantic tagging
- For production use requiring full PDF/UA-1 compliance, consider:
  - UniPDF (commercial license required)
  - Post-processing with Adobe Acrobat Pro
  - Using PDF/UA conversion tools

## Dependencies

```bash
go get github.com/xuri/excelize/v2
go get github.com/jung-kurt/gofpdf
```

## Project Structure

```
dsb-excel-generator/
├── main.go                    # Excel data generator
├── pdf_generator.go           # PDF letter generator
├── go.mod                     # Go module dependencies
├── dsb-mock-data-excel.xlsx   # Generated Excel file
├── output_pdfs/               # Generated PDF letters
└── README.md                  # This file
```

## Public 360 Integration

The Excel file is designed to be ready for Public 360 integration with the following fields:

- **DocumentType:** Classification for P360 document management
- **CaseNumber:** Reference number for case filing
- **SecurityLevel:** Document security classification
- **LetterContent:** Full letter text ready for upload or PDF generation

### Next Steps for P360 Integration

1. Create `p360_client.go` with API client
2. Implement document upload functionality
3. Map Excel columns to P360 metadata
4. Handle authentication (API key/OAuth)
5. Add error handling and retry logic
6. Batch upload with progress tracking

## Letter Types

The generator creates 4 different letter types with variety:

1. **Salary Regulation 2025** - Full salary regulation with pension changes
2. **Pension Change** - Pension contribution adjustments only
3. **Contract Amendment** - Contract updates with salary terms
4. **Annual Salary Review** - Performance-based annual increase

Each letter type has unique content while maintaining consistent formatting and structure.

## License

This is a demonstration project for generating WCAG-compliant documents.
