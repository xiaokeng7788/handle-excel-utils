package handleExcelUtils

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// ────────────── 1. 基于表头名称的操作 ──────────────

func TestGetColumnIndex(t *testing.T) {
	header := []string{"ID", "Name", "Score"}
	assert.Equal(t, 1, GetColumnIndex(header, "Name"))
	assert.Equal(t, -1, GetColumnIndex(header, "Missing"))
}

func TestTransformRowsByName(t *testing.T) {
	rows := [][]string{
		{"时间戳", "日期"},
		{"1777017316", "1777017316"},
	}

	transformers := map[string]TransformFunc{
		"时间戳": TimestampToDate("2006-01-02"),
		"日期":  TimestampToDate("2006-01-02 15:04:05"),
	}
	result := TransformRowsByName(rows, 1, transformers)
	expected := [][]string{
		{"时间戳", "日期"},
		{"2026-04-24", "2026-04-24 15:55:16"},
	}
	assert.Equal(t, expected, result)
}

func TestFilterColumnsByName(t *testing.T) {
	rows := [][]string{
		{"A", "B", "C", "D"},
		{"1", "2", "3", "4"},
	}
	result := FilterColumnsByName(rows, 1, "B", "D")
	expected := [][]string{
		{"A", "C"},
		{"1", "3"},
	}
	assert.Equal(t, expected, result)
}

// ────────────── 2. 自动列宽 ──────────────

func TestAutoFitColumns(t *testing.T) {
	tmpDir := t.TempDir()
	outPath := filepath.Join(tmpDir, "autofit.xlsx")

	w, err := NewWriter(WithSheetName("Test"))
	require.NoError(t, err)
	defer w.Close()

	w.WriteRow([]string{"短", "很长很长的列标题", "普通"})
	w.WriteRow([]string{"a", "bb", "ccc"})

	err = w.AutoFitColumns(30)
	require.NoError(t, err)
	err = w.SaveAs(outPath)
	require.NoError(t, err)

	// 可以再次打开验证列宽是否设置（可选），此处仅保证无错误
}

// ────────────── 3. 按条件过滤行 ──────────────

func TestFilterRows(t *testing.T) {
	rows := [][]string{
		{"Name", "Age"},
		{"Alice", "30"},
		{"Bob", ""},
		{"Charlie", "25"},
	}
	filtered := FilterRows(rows, 1, func(row []string) bool {
		return row[1] != "" // 保留 Age 非空的行
	})
	expected := [][]string{
		{"Name", "Age"},
		{"Alice", "30"},
		{"Charlie", "25"},
	}
	assert.Equal(t, expected, filtered)
}

// ────────────── 4. 打开已有文件编辑 ──────────────

func TestNewWriterFromFile(t *testing.T) {
	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "existing.xlsx")
	w1, err := NewWriter(WithSheetName("Data"))
	require.NoError(t, err)
	w1.WriteRow([]string{"Name"})
	w1.WriteRow([]string{"Alice"})
	err = w1.SaveAs(path)
	require.NoError(t, err)
	w1.Close()

	// 重新打开编辑
	w2, err := NewWriterFromFile(path)
	require.NoError(t, err)
	defer w2.Close()

	// 追加一行
	_, err = w2.WriteRow([]string{"Bob"})
	require.NoError(t, err)
	err = w2.SaveAs(path)
	require.NoError(t, err)

	// 读取验证
	r, err := OpenReader(path)
	require.NoError(t, err)
	defer r.Close()
	rows, err := r.GetRows("Data")
	require.NoError(t, err)
	assert.Len(t, rows, 3)
	assert.Equal(t, "Bob", rows[2][0])
}

func TestEditWriterSetCell(t *testing.T) {
	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "setcell.xlsx")
	w, _ := NewWriter()
	w.WriteRow([]string{"A", "B"})
	w.WriteRow([]string{"1", "2"})
	w.SaveAs(path)
	w.Close()

	// 修改单元格
	w2, _ := NewWriterFromFile(path)
	defer w2.Close()
	err := w2.SetCellValue(2, 1, "修改后")
	require.NoError(t, err)
	w2.SaveAs(path)

	r, _ := OpenReader(path)
	defer r.Close()
	rows, _ := r.GetRows("Sheet1")
	assert.Equal(t, "修改后", rows[1][0])
}

// ────────────── 6. 行列转置 ──────────────

func TestTranspose(t *testing.T) {
	rows := [][]string{
		{"A", "B", "C"},
		{"1", "2", "3"},
	}
	transposed := Transpose(rows)
	expected := [][]string{
		{"A", "1"},
		{"B", "2"},
		{"C", "3"},
	}
	assert.Equal(t, expected, transposed)
}

// ────────────── 7. 多文件/多Sheet合并 ──────────────

func TestAppendExcelFiles(t *testing.T) {
	tmpDir := t.TempDir()
	srcPath := filepath.Join(tmpDir, "src.xlsx")
	dstPath := filepath.Join(tmpDir, "dst.xlsx")

	// 创建源文件，包含两个Sheet
	ws, _ := NewWriter(WithSheetName("SheetA"))
	ws.WriteRow([]string{"A"})
	ws.AddSheet("SheetB")
	ws.WriteRow([]string{"B"})
	ws.SaveAs(srcPath)
	ws.Close()

	// 创建目标文件，已存在SheetA
	wd, _ := NewWriter(WithSheetName("SheetA"))
	wd.WriteRow([]string{"Old"})
	wd.SaveAs(dstPath)
	wd.Close()

	// 合并
	err := AppendExcelFiles(srcPath, dstPath)
	require.NoError(t, err)

	// 验证
	r, _ := OpenReader(dstPath)
	defer r.Close()
	// SheetA 应该有 Old + A
	rows, _ := r.GetRows("SheetA")
	assert.Len(t, rows, 2)
	assert.Equal(t, "Old", rows[0][0])
	assert.Equal(t, "A", rows[1][0])
	// SheetB 应该有 B
	rowsB, _ := r.GetRows("SheetB")
	assert.Len(t, rowsB, 1)
	assert.Equal(t, "B", rowsB[0][0])
}

func TestMergeSheets(t *testing.T) {
	tmpDir := t.TempDir()
	srcPath := filepath.Join(tmpDir, "merge_src.xlsx")
	dstPath := filepath.Join(tmpDir, "merge_dst.xlsx")

	w1, _ := NewWriter(WithSheetName("Data"))
	w1.WriteRow([]string{"Src1"})
	w1.SaveAs(srcPath)
	w1.Close()

	w2, _ := NewWriter(WithSheetName("Data"))
	w2.WriteRow([]string{"Dst1"})
	w2.SaveAs(dstPath)
	w2.Close()

	err := MergeSheets(srcPath, dstPath, "Data", "Data")
	require.NoError(t, err)

	r, _ := OpenReader(dstPath)
	defer r.Close()
	rows, _ := r.GetRows("Data")
	assert.Len(t, rows, 2)
	assert.Equal(t, "Dst1", rows[0][0])
	assert.Equal(t, "Src1", rows[1][0])
}

// ────────────── 8. Map写入Excel ──────────────

func TestWriteFromMaps(t *testing.T) {
	w, _ := NewWriter()
	defer w.Close()

	data := []map[string]string{
		{"Name": "Alice", "Age": "30"},
		{"Name": "Bob", "Age": "25"},
	}
	err := w.WriteFromMaps(data)
	require.NoError(t, err)

	// 提取工作表数据
	rows, _ := w.file.GetRows("Sheet1")
	assert.Len(t, rows, 3)
	// 表头按字母排序：Age, Name
	assert.Equal(t, []string{"Age", "Name"}, rows[0])
	assert.Equal(t, "30", rows[1][0])
	assert.Equal(t, "Alice", rows[1][1])
}

// ────────────── 10. CSV互转 ──────────────

func TestCSVToExcelAndBack(t *testing.T) {
	tmpDir := t.TempDir()
	csvPath := filepath.Join(tmpDir, "test.csv")
	xlsxPath := filepath.Join(tmpDir, "test.xlsx")
	csvBackPath := filepath.Join(tmpDir, "back.csv")

	// 准备 CSV
	content := "Name,Age\nAlice,30\nBob,25\n"
	os.WriteFile(csvPath, []byte(content), 0644)

	// CSV -> Excel
	err := CSVToExcel(csvPath, xlsxPath, "People")
	require.NoError(t, err)

	// Excel -> CSV
	err = ExcelToCSV(xlsxPath, csvBackPath, "People")
	require.NoError(t, err)

	// 验证内容
	back, _ := os.ReadFile(csvBackPath)
	// 注意：CSV 写入可能不会保留原始逗号风格，但数据应一致
	lines := strings.Split(strings.TrimSpace(string(back)), "\n")
	assert.Len(t, lines, 3)
	assert.Contains(t, lines[0], "Name")
	assert.Contains(t, lines[1], "Alice")
}
