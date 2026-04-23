package handleExcelUtils

import (
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// 创建临时 Excel 用于测试合并函数
func createTestExcel(t *testing.T, path string, sheet string, headerRows int, data [][]string) {
	w, err := NewWriter(WithSheetName(sheet))
	require.NoError(t, err)
	defer w.Close()

	for _, row := range data {
		_, err := w.WriteRow(row)
		require.NoError(t, err)
	}
	err = w.SaveAs(path)
	require.NoError(t, err)
}

func TestSplitTableMergeSameKey(t *testing.T) {
	data := map[string][][]string{
		"Alice": {{"Alice", "30"}, {"Alice", "31"}},
		"Bob":   {{"Bob", "25"}},
	}
	header := [][]string{{"Name", "Age"}}

	tmpDir := t.TempDir()
	outPath := filepath.Join(tmpDir, "split.xlsx")

	err := SplitTableMergeSameKey(data, header, false, outPath)
	require.NoError(t, err)

	r, err := OpenReader(outPath)
	require.NoError(t, err)
	defer r.Close()

	sheets := r.SheetNames()
	assert.ElementsMatch(t, []string{"Alice", "Bob"}, sheets)

	aliceRows, _ := r.GetRows("Alice")
	assert.Len(t, aliceRows, 3) // header + 2 data
}

func TestMergeTablesAppendList(t *testing.T) {
	baseData := map[string][]string{
		"1": {"1", "baseA"},
		"2": {"2", "baseB"},
	}
	baseHeader := [][]string{{"ID", "Value"}}

	otherData := map[string][]string{
		"1": {"extra1"},
		"2": {"extra2"},
	}
	otherHeader := [][]string{{"Extra"}}

	input := []MergeTableRequest{
		{Header: baseHeader, Data: baseData, Original: true},
		{Header: otherHeader, Data: otherData, Original: false},
	}

	tmpDir := t.TempDir()
	outPath := filepath.Join(tmpDir, "append.xlsx")

	err := MergeTablesAppendList(false, false, outPath, input...)
	require.NoError(t, err)

	r, err := OpenReader(outPath)
	require.NoError(t, err)
	defer r.Close()

	rows, err := r.GetRows("")
	require.NoError(t, err)
	assert.Len(t, rows, 3) // header + 2 data
	// 验证追加
	assert.Equal(t, []string{"ID", "Value", "Extra"}, rows[0])
	assert.Equal(t, []string{"1", "baseA", "extra1"}, rows[1])
	assert.Equal(t, []string{"2", "baseB", "extra2"}, rows[2])
}

func TestMergeTablesReplaceList(t *testing.T) {
	baseData := map[string][]string{
		"1": {"1", "oldName", "oldCity"},
		"2": {"2", "keep", "oldCity2"},
	}
	baseHeader := [][]string{{"ID", "Name", "City"}}

	replaceData := map[string][]string{
		"1": {"", "newName", ""}, // 只替换 Name，保留 City
	}
	replaceHeader := [][]string{{"", "Name", "City"}}

	input := []MergeTableRequest{
		{Header: baseHeader, Data: baseData, Original: true},
		{Header: replaceHeader, Data: replaceData, Original: false},
	}

	tmpDir := t.TempDir()
	outPath := filepath.Join(tmpDir, "replace.xlsx")

	err := MergeTablesReplaceList(false, false, outPath, input...)
	require.NoError(t, err)

	r, _ := OpenReader(outPath)
	defer r.Close()
	rows, _ := r.GetRows("")
	assert.Equal(t, []string{"1", "newName", "oldCity"}, rows[1])
	assert.Equal(t, []string{"2", "keep", "oldCity2"}, rows[2])
}
