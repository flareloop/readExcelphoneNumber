package excelcontroller

import (
	"fmt"
	"log"
	"os"
	"regexp"
	"strconv"
	"strings"
	"utils"

	"common"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var resultFile *excelize.File

func Init() {
	// 기존 파일 제거
	_, err := os.Stat(common.FILE_NAME)
	if !os.IsNotExist(err) {
		fmt.Printf("the result file(%s) is already exist. remove it.\n", common.FILE_NAME)
		os.Remove(common.FILE_NAME)
	}

	// 결과 excel 파일 생성
	if resultFile == nil {
		resultFile = excelize.NewFile()
	}

	// default sheet의 이름을 "orig"로 변경
	sheetIdx := resultFile.GetActiveSheetIndex()
	resultFile.SetSheetName(resultFile.GetSheetName(sheetIdx), common.ORG_SHEET_NAME)

	setSheetName(resultFile, common.ORG_SHEET_NAME)
	setSheetName(resultFile, common.EMPTY_SHEET_NAME)
	setSheetName(resultFile, common.DUPLICATE_SHEET_NAME)
	setSheetName(resultFile, common.VALIDATED_SHEET_NAME)
	setSheetName(resultFile, common.UN_VALIDATED_SHEET_NAME)

	resultFile.NewSheet(common.ORG_SHEET_NAME)
	resultFile.NewSheet(common.EMPTY_SHEET_NAME)
	resultFile.NewSheet(common.DUPLICATE_SHEET_NAME)
	resultFile.NewSheet(common.VALIDATED_SHEET_NAME)
	resultFile.NewSheet(common.UN_VALIDATED_SHEET_NAME)
}

func setSheetName(resultFile *excelize.File, sheetName string) {
	if sheetName == common.ORG_SHEET_NAME {
		resultFile.SetCellValue(sheetName, "A1", "이름")
		resultFile.SetCellValue(sheetName, "B1", "휴대폰번호")
		resultFile.SetCellValue(sheetName, "C1", "지역")
		resultFile.SetCellValue(sheetName, "D1", "변경 후 번호")
		resultFile.SetCellValue(sheetName, "E1", "Desc")
	} else {
		resultFile.SetCellValue(sheetName, "A1", "기존 액셀 index")
		resultFile.SetCellValue(sheetName, "B1", "이름")
		resultFile.SetCellValue(sheetName, "C1", "휴대폰번호")
		resultFile.SetCellValue(sheetName, "D1", "지역")
		resultFile.SetCellValue(sheetName, "E1", "변경 후 번호")
		resultFile.SetCellValue(sheetName, "F1", "Desc")
	}
}

func SaveExcel() {
	err := resultFile.SaveAs(common.FILE_NAME)
	if err != nil {
		fmt.Println(err)
	}
}

// orig 소스 파일에 대해, 중복 & 번호가 없는 것을 제외한 모든 정보 가져옴
func ReadExcel(orgFileName string) {
	var titleRow int
	var weiredCount int
	var emptycount int
	var dupCount int

	var emptyList []common.ExcelInfo      // 번호 없는 정보
	var duplicatedList []common.ExcelInfo // 중복 번호
	var validatedNumberList []common.ExcelInfo
	var unValidatedNumberList []common.ExcelInfo // 검증통과 못한 번호 리스트

	orgExcelFile, err := excelize.OpenFile(orgFileName)
	if err != nil {
		log.Fatal(err)
	}

	// TODO : sheet name은 나중에 param으로 받을 것! 단, 입력 없을 경우 default로 'Sheet1' 사용
	readRows := orgExcelFile.GetRows("Sheet1")

	totalCount := 0
	for rowNum, row := range readRows {
		totalCount++
		var excelInfo common.ExcelInfo
		if rowNum == 0 {
			titleRow++
			continue
		}

		// phoneNumber(row[1]) 정보 없는 row skip & emptyList 에 따로 보관
		if len(row) == 1 || len(row[1]) == 0 {
			emptycount++
			//fmt.Printf("[excel row index:%d][name:%s] is empty(PhoneNumber)\n", rowNum, row[0])

			convertRowToExcelInfo(row, &excelInfo, rowNum)
			emptyList = append(emptyList, excelInfo)

			continue
		}

		for idx, v := range row {
			if idx == 0 {
				excelInfo.Name = v
			} else if idx == 1 {
				excelInfo.PhoneNumber = v
			} else {
				excelInfo.Locate = v
			}
		}
		excelInfo.ExcelIndex = rowNum + 1

		// validation check & replace number
		isValidated := setNewPhoneNumber(&excelInfo)

		if len(excelInfo.NewPhoneNumber) == 0 {
			weiredCount++
			unValidatedNumberList = append(unValidatedNumberList, excelInfo)

			continue
		}

		// 중복정보 duplicatedList 에 따로 보관
		if IsDuplicated(&excelInfo, &dupCount) {

			utils.SetDesc(&excelInfo, "Dup")
			duplicatedList = append(duplicatedList, excelInfo)
			WriteOneRowExcelSheet(common.ORG_SHEET_NAME, excelInfo, rowNum)

			//WriteOneRowExcelSheet("duplicated", excelInfo, dupCount)
		} else {
			if isValidated {
				common.ExcelInfos[excelInfo.NewPhoneNumber] = excelInfo
				validatedNumberList = append(validatedNumberList, excelInfo)
				WriteOneRowExcelSheet(common.ORG_SHEET_NAME, excelInfo, rowNum)
			} else {
				weiredCount++
				unValidatedNumberList = append(unValidatedNumberList, excelInfo)
				//fmt.Printf("weired(%s/%s)\n", excelInfo.PhoneNumber, excelInfo.NewPhoneNumber)
			}
		}
	}

	utils.Sort(emptyList)
	WriteExcelSheet(common.EMPTY_SHEET_NAME, emptyList)

	utils.Sort(duplicatedList)
	WriteExcelSheet(common.DUPLICATE_SHEET_NAME, duplicatedList)

	utils.Sort(validatedNumberList)
	WriteExcelSheet(common.VALIDATED_SHEET_NAME, validatedNumberList)

	utils.Sort(unValidatedNumberList)
	WriteExcelSheet(common.UN_VALIDATED_SHEET_NAME, unValidatedNumberList)

	SaveExcel()

	fmt.Printf("All count(%d/%d) = saved(%d) + duplicated(%d) + unValidated(%d) + empty(%d) + titleRow(%d)\n",
		len(common.ExcelInfos)+dupCount+emptycount+titleRow+weiredCount,
		totalCount, len(common.ExcelInfos), dupCount, weiredCount, emptycount, titleRow)
	fmt.Println("-------------------------------------------------------------------------")
}

func setNewPhoneNumber(excelInfo *common.ExcelInfo) bool {
	var newKey string
	var isChanged bool

	checkingNumber := excelInfo.PhoneNumber

	// step1. [key change] trim
	trimmedPhoneNumber := strings.Trim(checkingNumber, " ")
	if trimmedPhoneNumber != checkingNumber {
		excelInfo.NewPhoneNumber = trimmedPhoneNumber
		checkingNumber = excelInfo.NewPhoneNumber
	}

	//step2. [key change] "-" 제거
	regExpRemoveDash := regexp.MustCompile(`-`)
	newKey, isChanged = changeByRegexp(regExpRemoveDash, checkingNumber, "") //, &deleteDashCnt)
	if isChanged {
		excelInfo.NewPhoneNumber = newKey
		checkingNumber = excelInfo.NewPhoneNumber
		utils.SetDesc(excelInfo, "dash")
	}

	//step3. [un-validated] // 숫자이외의 번호가 포함된 번호
	regExpincludingStr := regexp.MustCompile(`[^0-9]`)
	if regExpincludingStr.Match([]byte(checkingNumber)) {
		utils.SetDesc(excelInfo, "string")
		//fmt.Printf("[^0-9](%s)\n", excelInfo.PhoneNumber)
		return false
	}

	//step4-1. [key change] "10"으로 시작하는 번호 -> "010" 으로 변경
	regExpStart10ChgString := "010"
	regExpStart10 := regexp.MustCompile(`^10`) // "10"으로 시작하는 번호
	newKey, isChanged = changeByRegexp(regExpStart10, checkingNumber, regExpStart10ChgString /*, &change010Cnt*/)
	if isChanged {
		excelInfo.NewPhoneNumber = newKey
		checkingNumber = excelInfo.NewPhoneNumber
		utils.SetDesc(excelInfo, "10")
	}

	//step4-2. [key change] "11"으로 시작하는 번호 -> "011" 으로 변경
	regExpStart11ChgString := "011"
	regExpStart11 := regexp.MustCompile(`^11`) // "11"으로 시작하는 번호
	newKey, isChanged = changeByRegexp(regExpStart11, checkingNumber, regExpStart11ChgString /*, &change011Cnt*/)
	if isChanged {
		excelInfo.NewPhoneNumber = newKey
		checkingNumber = excelInfo.NewPhoneNumber
		utils.SetDesc(excelInfo, "11")
		//fmt.Printf("011(%s)\n", excelInfo.NewPhoneNumber)
	}

	//step4-3. [key change] "16"으로 시작하는 번호 -> "016" 으로 변경
	//var change016Cnt int
	regExpStart16ChgString := "016"
	regExpStart16 := regexp.MustCompile(`^16`) // "16"으로 시작하는 번호
	newKey, isChanged = changeByRegexp(regExpStart16, checkingNumber, regExpStart16ChgString /*, &change016Cnt*/)
	if isChanged {
		excelInfo.NewPhoneNumber = newKey
		checkingNumber = excelInfo.NewPhoneNumber
		utils.SetDesc(excelInfo, "16")
		//fmt.Printf("016(%s)\n", excelInfo.NewPhoneNumber)
	}

	//step4-4. [key change] "17"으로 시작하는 번호 -> "017" 으로 변경
	//var change017Cnt int
	regExpStart17ChgString := "017"
	regExpStart17 := regexp.MustCompile(`^17`) // "17"으로 시작하는 번호
	newKey, isChanged = changeByRegexp(regExpStart17, checkingNumber, regExpStart17ChgString /*, &change017Cnt*/)
	if isChanged {
		excelInfo.NewPhoneNumber = newKey
		checkingNumber = excelInfo.NewPhoneNumber
		utils.SetDesc(excelInfo, "17")
		//fmt.Printf("017(%s)\n", excelInfo.NewPhoneNumber)
	}

	//step4-5. [key change] "18"으로 시작하는 번호 -> "018" 으로 변경
	//var change018Cnt int
	regExpStart18ChgString := "018"
	regExpStart18 := regexp.MustCompile(`^18`) // "18"으로 시작하는 번호
	newKey, isChanged = changeByRegexp(regExpStart18, checkingNumber, regExpStart18ChgString /*, &change018Cnt*/)
	if isChanged {
		excelInfo.NewPhoneNumber = newKey
		checkingNumber = excelInfo.NewPhoneNumber
		utils.SetDesc(excelInfo, "18")
		//fmt.Printf("018(%s)\n", excelInfo.NewPhoneNumber)
	}

	//step4-6. [key change] "19"으로 시작하는 번호 -> "019" 으로 변경
	//var change019Cnt int
	regExpStart19ChgString := "019"
	regExpStart19 := regexp.MustCompile(`^19`) // "19"으로 시작하는 번호
	newKey, isChanged = changeByRegexp(regExpStart19, checkingNumber, regExpStart19ChgString /*, &change019Cnt*/)
	if isChanged {
		excelInfo.NewPhoneNumber = newKey
		checkingNumber = excelInfo.NewPhoneNumber
		utils.SetDesc(excelInfo, "19")
		//fmt.Printf("019(%s)\n", excelInfo.NewPhoneNumber)
	}

	//step4-7. [key change] "70"으로 시작하는 번호 -> "070" 으로 변경
	//var change070Cnt int
	regExpStart70ChgString := "070"
	regExpStart70 := regexp.MustCompile(`^70`) // "70"으로 시작하는 번호
	newKey, isChanged = changeByRegexp(regExpStart70, checkingNumber, regExpStart70ChgString /*, &change070Cnt*/)
	if isChanged {
		excelInfo.NewPhoneNumber = newKey
		checkingNumber = excelInfo.NewPhoneNumber
		utils.SetDesc(excelInfo, "70")
		//fmt.Printf("70(%s)\n", excelInfo.NewPhoneNumber)
	}

	// step5. [delete] 최소길이 미만 skip
	if len(checkingNumber) < common.MIN_PHONE_NUMBER_LENGTH {
		utils.SetDesc(excelInfo, "min")
		//lessLenCnt++
		//fmt.Printf("MinLen(%4d)(%s)\n", lessLenCnt, info.PhoneNumber)
		return false
	}

	// step6. [delete] 최대길이 초과 skip
	if len(checkingNumber) > common.MAX_PHONE_NUMBER_LENGTH {
		//overLenCnt++
		//fmt.Printf("MaxLen(%4d)(%s)\n", len(checkingNumber), checkingNumber)
		utils.SetDesc(excelInfo, "max")
		return false
	}

	// step7. [delete] "0" 으로 시작하지 않는 번호(10/11/16/17/18/19로 시작하는 번호들은 위에서 모두 01x로 바꿨음)
	regExpNotStartX := regexp.MustCompile(`^[^0]`) // "0" 으로 시작하지 않는 번호
	if regExpNotStartX.Match([]byte(checkingNumber)) {
		//notMobileCnt++
		//fmt.Printf("notMobile(CheckingNumber:%s) -> (%s)\n" /*notMobileCnt,*/, checkingNumber, excelInfo.NewPhoneNumber)
		utils.SetDesc(excelInfo, "notMobile")
		return false
	}

	// step8. [delete] "01" 로 시작하지 않는 번호(ex: 031, 032 등 지역번호)
	regExpNotStart01X := regexp.MustCompile(`^0[^17]`) // "01" 로 시작하지 않는 번호(ex: 031, 032 등 지역번호)
	if regExpNotStart01X.Match([]byte(checkingNumber)) {
		//notMobileCnt++
		//fmt.Printf("notMobile(CheckingNumber:%s) -> (%s)\n" /*notMobileCnt,*/, checkingNumber, excelInfo.NewPhoneNumber)
		utils.SetDesc(excelInfo, "notMobile")
		return false
	}

	//step9. [key change] "0000000" 제거
	if strings.Contains(checkingNumber, "0000000") {
		utils.SetDesc(excelInfo, "notMobile")
		return false
	}

	excelInfo.NewPhoneNumber = checkingNumber

	return true
}

// 추가하려는 정보가 이미 map에 있는지 확인
func IsDuplicated(addInfo *common.ExcelInfo, count *int) bool {
	info, isExisted := common.ExcelInfos[addInfo.NewPhoneNumber]
	if isExisted {
		*count++

		// 이미 저장된 정보에서 이름이 없고, 중복된 정보에서 이름 정보 있을 경우 이름정보 가져옴
		if len(info.Name) == 0 && len(addInfo.Name) > 0 {
			info.Name = addInfo.Name
		}

		// 위와 같은 조건으로 위치정보 가져옴
		if len(info.Locate) == 0 && len(addInfo.Locate) > 0 {
			info.Locate = addInfo.Locate
		}

		//fmt.Printf("saved:%s, dup:%s\n", checkingNumber, info.PhoneNumber)

		return true
	}

	return false
}

// source file 은 validation check 여부와 상관없이 row 단위로 모두 저장한다.
func WriteOneRowExcelSheet(sheetName string, excelInfo common.ExcelInfo, rowIdx int) {

	cellNameA := "A" + strconv.Itoa(rowIdx+1)
	cellNameB := "B" + strconv.Itoa(rowIdx+1)
	cellNameC := "C" + strconv.Itoa(rowIdx+1)
	cellNameD := "D" + strconv.Itoa(rowIdx+1)
	cellNameE := "E" + strconv.Itoa(rowIdx+1)

	resultFile.SetCellValue(sheetName, cellNameA, excelInfo.Name)
	resultFile.SetCellValue(sheetName, cellNameB, excelInfo.PhoneNumber)
	resultFile.SetCellValue(sheetName, cellNameC, excelInfo.Locate)
	resultFile.SetCellValue(sheetName, cellNameD, excelInfo.NewPhoneNumber)
	resultFile.SetCellValue(sheetName, cellNameE, excelInfo.Desc)
}

func WriteExcelSheet(sheetName string, list []common.ExcelInfo) {
	var excelInfo common.ExcelInfo

	for i := 0; i < len(list); i++ {
		excelInfo = list[i]

		cellNameA := "A" + strconv.Itoa(i+1)
		cellNameB := "B" + strconv.Itoa(i+1)
		cellNameC := "C" + strconv.Itoa(i+1)
		cellNameD := "D" + strconv.Itoa(i+1)
		cellNameE := "E" + strconv.Itoa(i+1)
		cellNameF := "F" + strconv.Itoa(i+1)

		resultFile.SetCellValue(sheetName, cellNameA, excelInfo.ExcelIndex)
		resultFile.SetCellValue(sheetName, cellNameB, excelInfo.Name)
		resultFile.SetCellValue(sheetName, cellNameC, excelInfo.PhoneNumber)
		resultFile.SetCellValue(sheetName, cellNameD, excelInfo.Locate)
		resultFile.SetCellValue(sheetName, cellNameE, excelInfo.NewPhoneNumber)
		resultFile.SetCellValue(sheetName, cellNameF, excelInfo.Desc)
	}
}

func convertRowToExcelInfo(row []string, excelInfo *common.ExcelInfo, idx int) {
	excelInfo.Name = row[0]
	excelInfo.PhoneNumber = row[1]
	excelInfo.Locate = row[2]

	// extra info
	excelInfo.ExcelIndex = idx
}

// return value : oldnumber, isChanged
func changeByRegexp(regExp *regexp.Regexp, checkingNumber string, chgStr string /*, cnt *int*/) (string, bool) {

	newPhoneNumber := regExp.ReplaceAllString(checkingNumber, chgStr)
	if newPhoneNumber != checkingNumber {
		return newPhoneNumber, true
	}

	return newPhoneNumber, false
}
