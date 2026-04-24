package handleExcelUtils

import (
	"log"
	"testing"

	"github.com/stretchr/testify/require"
)

const tempFile = "D:\\桌面\\main\\6666666666\\会议_6368.xlsx"

func TestExcel(t *testing.T) {
	data, _, err := GetExcelOnlyListData(tempFile, "Sheet1", 1, 1, false)
	if err != nil {
		t.Fatal(err)
	}
	t.Log(data)
	data1, _, err := GetExcelMultiListData(tempFile, "Sheet1", 0, 1, false)
	if err != nil {
		t.Fatal(err)
	}
	t.Log(data1)
}

func TestWriter_AddSheet(t *testing.T) {
	w := NewWriter(WithSheetName("用户")) // 创建第一个工作表
	defer w.Close()

	w.WriteRow([]string{"姓名", "年龄"})
	w.WriteRow([]string{"张三", "30"})

	// 添加第二个工作表
	if err := w.AddSheet("订单"); err != nil {
		log.Fatal(err)
	}
	w.WriteRow([]string{"订单号", "金额"})
	w.WriteRow([]string{"1001", "99.00"})

	w.SaveAs("D:\\桌面\\main\\6666666666\\多工作表.xlsx")

	w1 := NewWriter(WithSheetName("用户")) // 创建第一个工作表
	defer w.Close()

	w1.EnableStreamMode()
	w1.WriteRows(1, [][]string{{"姓名", "年龄"}, {"张三", "30"}})
	w1.Flush()

	// 添加第二个工作表
	if err := w1.AddSheet("订单"); err != nil {
		log.Fatal(err)
	}
	w1.EnableStreamMode()
	w1.WriteRows(1, [][]string{{"订单号", "金额"}, {"1001", "99.00"}})
	w1.Flush()

	w1.SaveAs("D:\\桌面\\main\\6666666666\\多工作表1.xlsx")
}

func TestSplitTableMergeSameKey(t *testing.T) {
	data, header, err := GetExcelMultiListData(tempFile, "Sheet1", 8, 1, true)
	if err != nil {
		t.Fatal(err)
	}
	t.Log(data)
	t.Log(header)
	err = SplitTableMergeSameKey(data, header, true, "D:\\桌面\\main\\6666666666\\会议_6368-11.xlsx")
	require.NoError(t, err)
}

func TestMergeTablesAppendList(t *testing.T) {
	data, header, err := GetExcelOnlyListData("D:\\桌面\\main\\6666666666\\x.xlsx", "Sheet1", 0, 2, true)
	if err != nil {
		t.Fatal(err)
	}
	data1, header1, err := GetExcelOnlyListData("D:\\桌面\\main\\6666666666\\y.xlsx", "Sheet1", 0, 2, true)
	if err != nil {
		t.Fatal(err)
	}
	var inputs []MergeTableRequest
	inputs = append(inputs, MergeTableRequest{Data: data, Header: header, Original: false})
	inputs = append(inputs, MergeTableRequest{Data: data1, Header: header1, Original: true})
	err = MergeTablesAppendList(true, "D:\\桌面\\main\\6666666666\\合并.xlsx", inputs...)
	require.NoError(t, err)
}

func TestMergeTablesReplaceList(t *testing.T) {
	data, header, err := GetExcelOnlyListData("D:\\桌面\\main\\6666666666\\会议_6368.xlsx", "Sheet1", 1, 1, true)
	if err != nil {
		t.Fatal(err)
	}
	data1, header1, err := GetExcelOnlyListData("D:\\桌面\\main\\6666666666\\会议_6368-23.xlsx", "Sheet1", 1, 1, true)
	if err != nil {
		t.Fatal(err)
	}
	var inputs []MergeTableRequest
	inputs = append(inputs, MergeTableRequest{Data: data, Header: header, Original: true})
	inputs = append(inputs, MergeTableRequest{Data: data1, Header: header1, Original: false})
	err = MergeTablesReplaceList(true, "D:\\桌面\\main\\6666666666\\合并.xlsx", inputs...)
	require.NoError(t, err)
}
