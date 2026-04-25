package handleExcelUtils

import (
	"encoding/csv"
	"errors"
	"fmt"
	"os"
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

// ReadRowsAsMapByName 类似于 ReadRowsAsMap，但使用表头名称指定键列。
// headerRows 必须 >= 1，使用最后一行表头（rows[headerRows-1]）查找列名。
func ReadRowsAsMapByName(rows [][]string, headerRows int, keyColName string) (map[string][]string, error) {
	if headerRows < 1 {
		return nil, errors.New("使用列名时表头行数必须至少为 1")
	}
	if headerRows > len(rows) {
		return nil, errors.New("表头行数超出总行数")
	}
	header := rows[headerRows-1]
	keyCol := GetColumnIndex(header, keyColName)
	if keyCol == -1 {
		return nil, fmt.Errorf("在表头中未找到列 %q", keyColName)
	}
	return ReadRowsAsMap(rows, headerRows, keyCol)
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

// ReadRowsAsMultiMapByName 通过列名生成允许重复键的 map。
func ReadRowsAsMultiMapByName(rows [][]string, headerRows int, keyColName string) (map[string][][]string, error) {
	if headerRows < 1 {
		return nil, errors.New("使用列名时表头行数必须至少为 1")
	}
	if headerRows > len(rows) {
		return nil, errors.New("表头行数超出总行数")
	}
	header := rows[headerRows-1]
	keyCol := GetColumnIndex(header, keyColName)
	if keyCol == -1 {
		return nil, fmt.Errorf("在表头中未找到列 %q", keyColName)
	}
	return ReadRowsAsMultiMap(rows, headerRows, keyCol)
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

// GetExcelMapByName 从 Excel 文件读取，并通过表头名称指定键列。
func GetExcelMapByName(filePath, sheetName string, keyColName string, headerRows int, title bool) (map[string][]string, [][]string, error) {
	reader, err := OpenReader(filePath)
	if err != nil {
		return nil, nil, err
	}
	defer reader.Close()
	rows, err := reader.GetRows(sheetName)
	if err != nil {
		return nil, nil, err
	}
	m, err := ReadRowsAsMapByName(rows, headerRows, keyColName)
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

// GetExcelMultiMapByName 从 Excel 文件读取，通过列名生成允许重复键的 map。
func GetExcelMultiMapByName(filePath, sheetName string, keyColName string, headerRows int, title bool) (map[string][][]string, [][]string, error) {
	reader, err := OpenReader(filePath)
	if err != nil {
		return nil, nil, err
	}
	defer reader.Close()
	rows, err := reader.GetRows(sheetName)
	if err != nil {
		return nil, nil, err
	}
	m, err := ReadRowsAsMultiMapByName(rows, headerRows, keyColName)
	if err != nil {
		return nil, nil, err
	}
	if title {
		return m, ExtractHeaders(rows, headerRows), nil
	}
	return m, nil, nil
}

// FilterColumns 从二维行数据中移除指定的列索引（0-based），返回新的切片。
// excludeCols 中的索引超出列范围时会被忽略。原数据不会被修改。
func FilterColumns(rows [][]string, excludeCols ...int) [][]string {
	// nil 输入返回 nil
	if rows == nil {
		return nil
	}
	if len(rows) == 0 || len(excludeCols) == 0 {
		newRows := make([][]string, len(rows))
		copy(newRows, rows)
		return newRows
	}

	excludeSet := make(map[int]bool, len(excludeCols))
	for _, col := range excludeCols {
		excludeSet[col] = true
	}

	newRows := make([][]string, len(rows))
	for i, row := range rows {
		newRow := make([]string, 0, len(row))
		for j, cell := range row {
			if !excludeSet[j] {
				newRow = append(newRow, cell)
			}
		}
		newRows[i] = newRow
	}
	return newRows
}

// FilterColumnsByName 过滤指定名称的列，保留其他列。
func FilterColumnsByName(rows [][]string, headerRows int, excludeColNames ...string) [][]string {
	if headerRows <= 0 || len(excludeColNames) == 0 || len(rows) == 0 {
		// 无表头或无需过滤
		newRows := make([][]string, len(rows))
		copy(newRows, rows)
		return newRows
	}
	headerRow := rows[headerRows-1]
	excludeSet := make(map[int]bool)
	for _, name := range excludeColNames {
		idx := GetColumnIndex(headerRow, name)
		if idx != -1 {
			excludeSet[idx] = true
		}
	}
	excludeSlice := make([]int, 0, len(excludeSet))
	for idx := range excludeSet {
		excludeSlice = append(excludeSlice, idx)
	}
	return FilterColumns(rows, excludeSlice...)
}

// FilterRows 根据 predicate 过滤数据行（表头行始终保留）。
func FilterRows(rows [][]string, headerRows int, predicate func(row []string) bool) [][]string {
	if headerRows <= 0 {
		headerRows = 0
	}
	filtered := make([][]string, 0, len(rows))
	// 保留表头
	for i := 0; i < headerRows && i < len(rows); i++ {
		filtered = append(filtered, rows[i])
	}
	// 过滤数据行
	for i := headerRows; i < len(rows); i++ {
		if predicate(rows[i]) {
			filtered = append(filtered, rows[i])
		}
	}
	return filtered
}

// GetColumnIndex 根据表头行和列名返回列索引（-1 表示不存在）。
func GetColumnIndex(headerRow []string, colName string) int {
	for i, v := range headerRow {
		if v == colName {
			return i
		}
	}
	return -1
}

// Transpose 将二维切片进行转置。
func Transpose(rows [][]string) [][]string {
	if len(rows) == 0 {
		return nil
	}
	colCount := 0
	for _, row := range rows {
		if len(row) > colCount {
			colCount = len(row)
		}
	}
	result := make([][]string, colCount)
	for i := 0; i < colCount; i++ {
		result[i] = make([]string, len(rows))
		for j := 0; j < len(rows); j++ {
			if i < len(rows[j]) {
				result[i][j] = rows[j][i]
			}
		}
	}
	return result
}

// CSVToExcel 将 CSV 文件转换为 Excel 文件。
func CSVToExcel(csvPath, xlsxPath, sheetName string) error {
	f, err := os.Open(csvPath)
	if err != nil {
		return fmt.Errorf("打开CSV文件失败: %w", err)
	}
	defer f.Close()

	reader := csv.NewReader(f)
	reader.LazyQuotes = true
	reader.TrimLeadingSpace = true

	records, err := reader.ReadAll()
	if err != nil {
		return fmt.Errorf("读取CSV失败: %w", err)
	}

	w, err := NewWriter(WithSheetName(sheetName))
	if err != nil {
		return err
	}
	defer w.Close()

	if err := w.WriteAll(records); err != nil {
		return err
	}
	return w.SaveAs(xlsxPath)
}

// ExcelToCSV 将 Excel 文件的指定工作表转换为 CSV。
func ExcelToCSV(xlsxPath, csvPath, sheetName string) error {
	r, err := OpenReader(xlsxPath)
	if err != nil {
		return err
	}
	defer r.Close()

	rows, err := r.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("读取Excel工作表失败: %w", err)
	}

	f, err := os.Create(csvPath)
	if err != nil {
		return fmt.Errorf("创建CSV文件失败: %w", err)
	}
	defer f.Close()

	writer := csv.NewWriter(f)
	defer writer.Flush()

	for _, row := range rows {
		if err := writer.Write(row); err != nil {
			return err
		}
	}
	return nil
}
