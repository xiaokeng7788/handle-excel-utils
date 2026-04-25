package handleExcelUtils

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestReadRowsAsMap(t *testing.T) {
	rows := [][]string{
		{"ID", "Name"},
		{"1", "Alice"},
		{"2", "Bob"},
	}
	m, err := ReadRowsAsMap(rows, 1, 0)
	require.NoError(t, err)
	assert.Len(t, m, 2)
	assert.Equal(t, []string{"1", "Alice"}, m["1"])
}

func TestReadRowsAsMapDuplicate(t *testing.T) {
	rows := [][]string{
		{"ID", "Name"},
		{"1", "Alice"},
		{"1", "Dup"},
	}
	_, err := ReadRowsAsMap(rows, 1, 0)
	assert.ErrorContains(t, err, "第 3 行发现重复的键 \"1\"")
}

func TestReadRowsAsMultiMap(t *testing.T) {
	rows := [][]string{
		{"ID", "Name"},
		{"1", "Alice"},
		{"1", "Alicia"},
	}
	m, err := ReadRowsAsMultiMap(rows, 1, 0)
	require.NoError(t, err)
	assert.Len(t, m["1"], 2)
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

func TestFilterColumns(t *testing.T) {
	rows := [][]string{
		{"A", "B", "C", "D"},
		{"1", "2", "3", "4"},
		{"5", "6", "7", "8"},
	}

	// 移除第1列和第3列
	result := FilterColumns(rows, 1, 3)
	expected := [][]string{
		{"A", "C"},
		{"1", "3"},
		{"5", "7"},
	}
	assert.Equal(t, expected, result)

	// 移除不存在的列（索引 10）
	result2 := FilterColumns(rows, 10)
	assert.Equal(t, rows, result2)

	// 空输入
	assert.Nil(t, FilterColumns(nil, 0))
	assert.Empty(t, FilterColumns([][]string{}, 0))

	// 不移除任何列
	result3 := FilterColumns(rows)
	assert.Equal(t, rows, result3)
}

func TestReadRowsAsMapByName(t *testing.T) {
	rows := [][]string{
		{"ID", "Name"},
		{"1", "Alice"},
		{"2", "Bob"},
	}
	m, err := ReadRowsAsMapByName(rows, 1, "ID")
	require.NoError(t, err)
	assert.Len(t, m, 2)
	assert.Equal(t, []string{"1", "Alice"}, m["1"])

	// 不存在的列名
	_, err = ReadRowsAsMapByName(rows, 1, "Missing")
	assert.Error(t, err)
}
