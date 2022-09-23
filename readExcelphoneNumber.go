package main

import (
	"common"
	"excelcontroller"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"regexp"
	"time"
)

func getInputFileName() string {
	currentPath, _ := os.Getwd()
	files, err := ioutil.ReadDir(currentPath)
	if err != nil {
		log.Fatal(err)
	}

	var excelFiles []string
	regExpExcel := regexp.MustCompile(`.xls.*$`) // ".xlsx" , ".xls" 으로 끝나는 string
	for _, file := range files {

		if regExpExcel.Match([]byte(file.Name())) {
			excelFiles = append(excelFiles, file.Name())
		}
	}
	if len(excelFiles) == 1 {
		return excelFiles[0]
	}

	var pick int
	var pickFileName string
	fmt.Println("There are excel files in current directory.")
	for idx, fileName := range excelFiles {
		fmt.Printf("\t[%d]%s\n", idx+1, fileName)
	}
	fmt.Printf("choose one(1~%d): ", len(excelFiles))
	fmt.Scanf("%d", &pick)
	if pick < 1 || pick > len(excelFiles) {
		fmt.Printf("[Usage] pick range is 1 ~ %d\n", len(excelFiles)+1)
	}

	for _, file := range files {
		if file.Name() == excelFiles[pick-1] {
			pickFileName = file.Name()
		}
	}

	return pickFileName
}

func main() {
	iputFileName := getInputFileName()
	fmt.Printf("%s is reading...\n", iputFileName)

	excelcontroller.Init()

	startTime := time.Now()
	excelcontroller.ReadExcel(iputFileName)
	elapsedTime := time.Since(startTime)

	fmt.Printf("Final validated Phone number:%d\n", len(common.ExcelInfos))
	fmt.Printf("(%.2f)seconds-----------------------------------------------------------\n", elapsedTime.Seconds())
}
