package handleExcelUtils

import (
	"errors"
	"fmt"
	"sort"
	"strconv"
)

// ReadRowsAsMap 将二维行转换为以 keyCol 列为键的 map[string][]string。
// headerRows：表头行数（跳过）。若 keyCol<0 使用第0列。
// 重复键会返回错误。
func ReadRowsAsMap(rows [][]string, headerRows int, keyCol int) (map[string][]string, error) {
	if headerRows < 0 {
		headerRows = 0
	}
	if keyCol < 0 {
		keyCol = 0
	}
	if len(rows) <= headerRows {
		return nil, errors.New("表头后没有数据行")
	}
	result := make(map[string][]string, len(rows)-headerRows)
	for i := headerRows; i < len(rows); i++ {
		row := rows[i]
		if len(row) <= keyCol {
			continue
		}
		key := row[keyCol]
		if key == "" {
			continue
		}
		if _, exists := result[key]; exists {
			return nil, fmt.Errorf("第 %d 行发现重复的键 %q", i+1, key)
		}
		result[key] = row
	}
	return result, nil
}

// ReadRowsAsMultiMap 转换为允许重复键的一对多 map[string][][]string。
func ReadRowsAsMultiMap(rows [][]string, headerRows int, keyCol int) (map[string][][]string, error) {
	if headerRows < 0 {
		headerRows = 0
	}
	if keyCol < 0 {
		keyCol = 0
	}
	if len(rows) <= headerRows {
		return nil, errors.New("表头后没有数据行")
	}
	result := make(map[string][][]string)
	for i := headerRows; i < len(rows); i++ {
		row := rows[i]
		if len(row) <= keyCol {
			continue
		}
		key := row[keyCol]
		if key == "" {
			continue
		}
		result[key] = append(result[key], row)
	}
	return result, nil
}

// MergeStringSlices 合并两个 map[string][]string，overwrite 为 true 时源覆盖目标。
func MergeStringSlices(dst, src map[string][]string, overwrite bool) {
	for k, v := range src {
		if _, exists := dst[k]; !exists || overwrite {
			dst[k] = v
		}
	}
}

// MergeStringSliceMulti 合并两个 map[string][][]string，相同键的值会追加。
func MergeStringSliceMulti(dst, src map[string][][]string) {
	for k, v := range src {
		dst[k] = append(dst[k], v...)
	}
}

// ExtractHeaders 提取前 headerRows 行作为表头。
func ExtractHeaders(rows [][]string, headerRows int) [][]string {
	if headerRows <= 0 {
		return nil
	}
	if len(rows) < headerRows {
		return rows
	}
	return rows[:headerRows]
}

// ToInterfaceSlice 将 []string 转换为 []interface{}（方便 excelize 操作）。
func ToInterfaceSlice(s []string) []interface{} {
	res := make([]interface{}, len(s))
	for i, v := range s {
		res[i] = v
	}
	return res
}

// SortMapByNumericKey 对键为数字字符串的 map 按键数值排序，返回有序键切片。
func SortMapByNumericKey(m map[string][]string) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, k)
	}
	sort.Slice(keys, func(i, j int) bool {
		ni, _ := strconv.Atoi(keys[i])
		nj, _ := strconv.Atoi(keys[j])
		return ni < nj
	})
	return keys
}

// GetExcelMap 读取 Excel 并返回以 keyCol 为唯一键的 map。title 为 true 时同时返回表头。
// 取代旧的 GetExcelOnlyListData。
func GetExcelMap(filePath, sheetName string, keyCol, headerRows int, title bool) (map[string][]string, [][]string, error) {
	reader, err := OpenReader(filePath)
	if err != nil {
		return nil, nil, err
	}
	defer reader.Close()
	rows, err := reader.GetRows(sheetName)
	if err != nil {
		return nil, nil, err
	}
	m, err := ReadRowsAsMap(rows, headerRows, keyCol)
	if err != nil {
		return nil, nil, err
	}
	if title {
		return m, ExtractHeaders(rows, headerRows), nil
	}
	return m, nil, nil
}

// GetExcelMultiMap 读取 Excel 并返回允许重复键的 map。title 控制是否返回表头。
// 取代旧的 GetExcelMultiListData。
func GetExcelMultiMap(filePath, sheetName string, keyCol, headerRows int, title bool) (map[string][][]string, [][]string, error) {
	reader, err := OpenReader(filePath)
	if err != nil {
		return nil, nil, err
	}
	defer reader.Close()
	rows, err := reader.GetRows(sheetName)
	if err != nil {
		return nil, nil, err
	}
	m, err := ReadRowsAsMultiMap(rows, headerRows, keyCol)
	if err != nil {
		return nil, nil, err
	}
	if title {
		return m, ExtractHeaders(rows, headerRows), nil
	}
	return m, nil, nil
}
