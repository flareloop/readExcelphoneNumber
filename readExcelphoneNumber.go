package main

import (
	"bufio"
	"common"
	"excelcontroller"
	"fmt"
	"log"
	"os"
	"strings"
	"time"
)

func getInputFileName() string {
	fmt.Printf("input source file name(relative path):")

	in := bufio.NewReader(os.Stdin)
	inputStr, err := in.ReadString('\n')
	if err != nil {
		log.Fatal(err)
	}
	inputStr = strings.TrimSuffix(inputStr, "\n")
	inputStr = strings.TrimSuffix(inputStr, "\r")

	path, _ := os.Getwd()

	iputFileName := path + "\\" + inputStr
	fmt.Printf("File Path:%s\n", iputFileName)

	return iputFileName
}

func main() {
	iputFileName := getInputFileName()

	excelcontroller.Init()

	startTime := time.Now()
	excelcontroller.ReadExcel(iputFileName)
	elapsedTime := time.Since(startTime)

	fmt.Printf("Final validated Phone number:%d\n", len(common.ExcelInfos))
	fmt.Printf("(%.2f)seconds------------------------------------------------------------------\n", elapsedTime.Seconds())
}
