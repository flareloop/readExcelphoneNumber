package utils

import (
	"common"
	"sort"
)

type ExcelInfoSlice []common.ExcelInfo

// Len implements sort.Interface
func (info ExcelInfoSlice) Len() int {
	return len(info)
}

// Less implements sort.Interface
func (info ExcelInfoSlice) Less(i int, j int) bool {
	return info[i].NewPhoneNumber < info[j].NewPhoneNumber
}

// Swap implements sort.Interface
func (info ExcelInfoSlice) Swap(i int, j int) {
	info[i], info[j] = info[j], info[i]
}

func Sort(validatedNumberList ExcelInfoSlice) {
	sort.Sort(ExcelInfoSlice(validatedNumberList))
}

func SetDesc(info *common.ExcelInfo, appendStr string) {
	if len(info.Desc) > 0 {
		info.Desc += ", " + appendStr
	} else {
		info.Desc = appendStr
	}
}
