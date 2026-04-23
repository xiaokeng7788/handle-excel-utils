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

	w, err := NewWriter(WithSheetName("Data"))
	require.NoError(t, err)
	defer w.Close()

	_, err = w.WriteRow([]string{"ID", "Name", "Score"})
	require.NoError(t, err)

	data := [][]string{
		{"1", "Alice", "95"},
		{"2", "Bob", "87"},
	}
	err = w.WriteRows(2, data)
	require.NoError(t, err)
	err = w.SaveAs(outPath)
	require.NoError(t, err)

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

	w, err := NewWriter(WithSheetName("Stream"))
	require.NoError(t, err)
	defer w.Close()

	err = w.EnableStreamMode()
	require.NoError(t, err)

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

	r, err := OpenReader(outPath)
	require.NoError(t, err)
	defer r.Close()
	allRows, err := r.GetRows("Stream")
	require.NoError(t, err)
	assert.Len(t, allRows, 1000)
}

func TestReaderFromStream(t *testing.T) {
	w, err := NewWriter()
	require.NoError(t, err)
	_, err = w.WriteRow([]string{"A", "B"})
	require.NoError(t, err)
	buf, err := w.WriteToBuffer()
	require.NoError(t, err)
	w.Close()

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

	w, _ := NewWriter()
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
