package handleExcelUtils

import (
	"testing"
)

const tempFile = "D:\\桌面\\main\\6666666666\\会议_8778.xlsx"

func TestExcel(t *testing.T) {
	data, err := GetExcelOnlyListData(tempFile, "Sheet1", 1, 1)
	if err != nil {
		t.Fatal(err)
	}
	t.Log(data)
	data1, err := GetExcelMultiListData(tempFile, "Sheet1", 0, 1)
	if err != nil {
		t.Fatal(err)
	}
	t.Log(data1)
}
