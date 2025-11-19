package models

// EmployeeData represents the data for a single employee
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
