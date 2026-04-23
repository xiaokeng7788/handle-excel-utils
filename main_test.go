package handleExcelUtils

import (
	"path/filepath"
	"strconv"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestWriterAndReader(t *testing.T) {
	tmpDir := t.TempDir()
	outPath := filepath.Join(tmpDir, "test.xlsx")

	// 创建写入器并写入数据
	w := NewWriter(WithSheetName("Data"))
	defer w.Close()

	headers := [][]string{{"ID", "Name", "Score"}}
	data := [][]string{
		{"1", "Alice", "95"},
		{"2", "Bob", "87"},
	}

	// 写入表头
	_, err := w.WriteRow(headers[0])
	require.NoError(t, err)

	// 写入数据
	startRow := 2
	err = w.WriteRows(startRow, data)
	require.NoError(t, err)

	err = w.SaveAs(outPath)
	require.NoError(t, err)

	// 读取并验证
	r, err := OpenReader(outPath)
	require.NoError(t, err)
	defer r.Close()

	rows, err := r.GetRows("Data")
	require.NoError(t, err)
	assert.Len(t, rows, 3)
	assert.Equal(t, []string{"1", "Alice", "95"}, rows[1])
	assert.Equal(t, []string{"2", "Bob", "87"}, rows[2])
}

func TestStreamWriter(t *testing.T) {
	tmpDir := t.TempDir()
	outPath := filepath.Join(tmpDir, "stream.xlsx")

	w := NewWriter(WithSheetName("Stream"))
	defer w.Close()

	err := w.EnableStreamMode()
	require.NoError(t, err)

	// 写入大量行（模拟）
	rows := make([][]string, 1000)
	for i := 0; i < 1000; i++ {
		rows[i] = []string{strconv.Itoa(i), "value"}
	}
	err = w.WriteRows(1, rows)
	require.NoError(t, err)

	err = w.Flush()
	require.NoError(t, err)

	err = w.SaveAs(outPath)
	require.NoError(t, err)

	// 验证行数 - 显式指定工作表名
	r, err := OpenReader(outPath)
	require.NoError(t, err)
	defer r.Close()

	allRows, err := r.GetRows("Stream")
	require.NoError(t, err)
	assert.Len(t, allRows, 1000)
}

func TestReaderFromStream(t *testing.T) {
	// 先创建一个文件写入 buffer
	w := NewWriter()
	_, _ = w.WriteRow([]string{"A", "B"})
	buf, err := w.WriteToBuffer()
	require.NoError(t, err)
	w.Close()

	// 从 buffer 读取
	r, err := OpenReaderFromStream(buf)
	require.NoError(t, err)
	defer r.Close()

	rows, err := r.GetRows("")
	require.NoError(t, err)
	assert.Equal(t, [][]string{{"A", "B"}}, rows)
}

func TestRowIterator(t *testing.T) {
	tmpDir := t.TempDir()
	outPath := filepath.Join(tmpDir, "iterator.xlsx")

	w := NewWriter()
	_, _ = w.WriteRow([]string{"H1", "H2"})
	_, _ = w.WriteRow([]string{"1", "a"})
	_, _ = w.WriteRow([]string{"2", "b"})
	_ = w.SaveAs(outPath)
	w.Close()

	r, err := OpenReader(outPath)
	require.NoError(t, err)
	defer r.Close()

	it, err := r.RowIterator("")
	require.NoError(t, err)
	defer it.Close()

	var rows [][]string
	for it.Next() {
		cols, err := it.Columns()
		require.NoError(t, err)
		rows = append(rows, cols)
	}
	assert.Len(t, rows, 3)
}

func TestReadRowsAsMap(t *testing.T) {
	rows := [][]string{
		{"ID", "Name"},
		{"1", "Alice"},
		{"2", "Bob"},
		{"3", "Charlie"},
	}

	m, err := ReadRowsAsMap(rows, 1, 0)
	require.NoError(t, err)
	assert.Len(t, m, 3)
	assert.Equal(t, []string{"1", "Alice"}, m["1"])
	assert.Equal(t, []string{"2", "Bob"}, m["2"])
}

func TestReadRowsAsMapDuplicateKey(t *testing.T) {
	rows := [][]string{
		{"ID", "Name"},
		{"1", "Alice"},
		{"1", "Duplicate"},
	}
	_, err := ReadRowsAsMap(rows, 1, 0)
	assert.ErrorContains(t, err, "duplicate key")
}

func TestReadRowsAsMultiMap(t *testing.T) {
	rows := [][]string{
		{"ID", "Name"},
		{"1", "Alice"},
		{"1", "Alicia"},
		{"2", "Bob"},
	}
	m, err := ReadRowsAsMultiMap(rows, 1, 0)
	require.NoError(t, err)
	assert.Len(t, m["1"], 2)
	assert.Len(t, m["2"], 1)
}

func TestMergeStringSlices(t *testing.T) {
	dst := map[string][]string{"a": {"1"}}
	src := map[string][]string{"a": {"2"}, "b": {"3"}}
	MergeStringSlices(dst, src, true)
	assert.Equal(t, []string{"2"}, dst["a"])
	assert.Equal(t, []string{"3"}, dst["b"])

	dst2 := map[string][]string{"a": {"1"}}
	MergeStringSlices(dst2, src, false)
	assert.Equal(t, []string{"1"}, dst2["a"])
}

func TestSortMapByNumericKey(t *testing.T) {
	m := map[string][]string{
		"10": {"ten"},
		"2":  {"two"},
		"1":  {"one"},
	}
	keys := SortMapByNumericKey(m)
	assert.Equal(t, []string{"1", "2", "10"}, keys)
}
