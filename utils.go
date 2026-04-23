package handleExcelUtils

import (
	"errors"
	"fmt"
	"strconv"
)

// ReadRowsAsMap 将二维行数据转换为以指定列为键的 map。
// headerRows 指定表头所占行数，这些行不会作为数据处理。
// keyCol 指定作为键的列索引（从0开始）。若 keyCol < 0 则使用第一列。
// 返回的 map 中值为该行的完整切片（长度可能不同）。
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
			continue // 忽略列数不足的行，可根据需求修改为报错
		}
		key := row[keyCol]
		if key == "" {
			continue // 忽略键为空的行
		}
		if _, exists := result[key]; exists {
			return nil, fmt.Errorf("在 %d 行发现重复的键 '%s'", i+1, key)
		}
		result[key] = row
	}
	return result, nil
}

// ReadRowsAsMultiMap 将二维行数据转换为一对多的 map（允许重复键）。
// 参数含义同 ReadRowsAsMap。
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
	result := make(map[string][][]string, len(rows)-headerRows)
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

// MergeStringSlices 合并两个以字符串为键、值为 []string 的 map。
// 若 overwrite 为 true，则源映射中的键会覆盖目标映射。
func MergeStringSlices(dst, src map[string][]string, overwrite bool) {
	for k, v := range src {
		if _, exists := dst[k]; !exists || overwrite {
			dst[k] = v
		}
	}
}

// MergeStringSliceMulti 合并两个以字符串为键、值为 [][]string 的 map。
// 对于相同键，源映射中的值会追加到目标映射。
func MergeStringSliceMulti(dst, src map[string][][]string) {
	for k, v := range src {
		dst[k] = append(dst[k], v...)
	}
}

// ExtractHeaders 从二维行数据中提取表头行（前 headerRows 行）。
func ExtractHeaders(rows [][]string, headerRows int) [][]string {
	if headerRows <= 0 {
		return nil
	}
	if len(rows) < headerRows {
		return rows
	}
	return rows[:headerRows]
}

// ToInterfaceSlice 将 []string 转换为 []interface{}，方便写入 excelize。
func ToInterfaceSlice(s []string) []interface{} {
	res := make([]interface{}, len(s))
	for i, v := range s {
		res[i] = v
	}
	return res
}

// SortMapByNumericKey 对键为数字字符串的 map[string][]string 按数值大小排序，
// 返回有序的键切片。
func SortMapByNumericKey(m map[string][]string) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, k)
	}
	// 简单排序：按字符串数值比较
	for i := 0; i < len(keys)-1; i++ {
		for j := i + 1; j < len(keys); j++ {
			ni, _ := strconv.Atoi(keys[i])
			nj, _ := strconv.Atoi(keys[j])
			if ni > nj {
				keys[i], keys[j] = keys[j], keys[i]
			}
		}
	}
	return keys
}

// 适用于指定键 且表格中无重复数据
// 快速获取指定文件和指定表头的表格数据并返回整理好的map集合
//
// filePaths: 文件路径 必填 sheetName: 表单名称 必填 keyCol: 键所在列 必填 headerRows: 表头行数 必填
func GetExcelOnlyListData(filePaths, sheetName string, keyCol, headerRows int) (map[string][]string, error) {
	reader, err := OpenReader(filePaths)
	if err != nil {
		return nil, err
	}
	defer reader.Close()
	rows, err := reader.GetRows(sheetName)
	if err != nil {
		return nil, err
	}
	asMap, err := ReadRowsAsMap(rows, headerRows, keyCol)
	if err != nil {
		return nil, err
	}
	return asMap, nil
}

// 适用于指定键 且表格中可能存在重复数据
// 快速获取指定文件和指定表头的表格数据并返回整理好的map集合
//
// filePaths: 文件路径 必填 sheetName: 表单名称 必填 keyCol: 键所在列 必填 headerRows: 表头行数 必填
func GetExcelMultiListData(filePaths, sheetName string, keyCol, headerRows int) (map[string][][]string, error) {
	reader, err := OpenReader(filePaths)
	if err != nil {
		return nil, err
	}
	defer reader.Close()
	rows, err := reader.GetRows(sheetName)
	if err != nil {
		return nil, err
	}
	asMap, err := ReadRowsAsMultiMap(rows, headerRows, keyCol)
	if err != nil {
		return nil, err
	}
	return asMap, nil
}
