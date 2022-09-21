package common

var MIN_PHONE_NUMBER_LENGTH = 9  // ex) 10 123 4567
var MAX_PHONE_NUMBER_LENGTH = 11 // (trim제거 후) 01012345678

var FILE_NAME string = "resultFile.xlsx"
var ORG_SHEET_NAME = "org"
var EMPTY_SHEET_NAME = "empty"
var DUPLICATE_SHEET_NAME = "dup"
var VALIDATED_SHEET_NAME = "validated"
var UN_VALIDATED_SHEET_NAME = "un-validated"

type ExcelInfo struct {
	Name        string
	PhoneNumber string
	Locate      string

	ExcelIndex     int
	NewPhoneNumber string
	Desc           string
}

var ExcelInfos = make(map[string]ExcelInfo)
