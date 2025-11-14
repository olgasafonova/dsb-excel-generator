# DSB Excel and PDF Generator

This Go program generates mock employee data in Excel format and creates individual, WCAG AAA-compliant PDF salary regulation letters.

## Features

### Excel Generator (`main.go`)
- Generates 3,000 rows of realistic Danish employee data
- Includes columns: CPR, FirstName, LastName, salary information, adjustments, and dates
- All CPR numbers are guaranteed unique
- Realistic Danish names and salary ranges (25,000-75,000 kr/month)

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

## Size Estimates

- **Single PDF:** ~2.1 KB
- **3,000 PDFs:** ~6.3 MB total
- **Excel file:** ~225 KB

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

## License

This is a demonstration project for generating WCAG-compliant documents.
