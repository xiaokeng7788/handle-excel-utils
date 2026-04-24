package handleExcelUtils

import (
	"errors"
	"fmt"
	"sort"
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

// GetExcelOnlyListData 适用于指定键 且表格中无重复数据
// 快速获取指定文件和指定表头的表格数据并返回整理好的map集合
//
// filePaths: 文件路径 必填 sheetName: 表单名称 必填 keyCol: 键所在列 必填 headerRows: 表头行数 必填 title: 是否返回包含表头
func GetExcelOnlyListData(filePaths, sheetName string, keyCol, headerRows int, title bool) (map[string][]string, [][]string, error) {
	reader, err := OpenReader(filePaths)
	if err != nil {
		return nil, nil, err
	}
	defer reader.Close()
	rows, err := reader.GetRows(sheetName)
	if err != nil {
		return nil, nil, err
	}
	asMap, err := ReadRowsAsMap(rows, headerRows, keyCol)
	if err != nil {
		return nil, nil, err
	}
	if title {
		return asMap, ExtractHeaders(rows, headerRows), nil
	}
	return asMap, nil, nil
}

// GetExcelMultiListData 适用于指定键 且表格中可能存在重复数据
// 快速获取指定文件和指定表头的表格数据并返回整理好的map集合
//
// filePaths: 文件路径 必填 sheetName: 表单名称 必填 keyCol: 键所在列 必填 headerRows: 表头行数 必填 title: 是否返回包含表头
func GetExcelMultiListData(filePaths, sheetName string, keyCol, headerRows int, title bool) (map[string][][]string, [][]string, error) {
	reader, err := OpenReader(filePaths)
	if err != nil {
		return nil, nil, err
	}
	defer reader.Close()
	rows, err := reader.GetRows(sheetName)
	if err != nil {
		return nil, nil, err
	}
	asMap, err := ReadRowsAsMultiMap(rows, headerRows, keyCol)
	if err != nil {
		return nil, nil, err
	}
	if title {
		return asMap, ExtractHeaders(rows, headerRows), nil
	}
	return asMap, nil, nil
}

// SplitTableMergeSameKey 快速拆分指定列相同数据 创建统一文件的多工作铺模式
// 适用于需要表格分发给不同人员
//
// data: 待处理数据 header:表头数据如果存在将写入表头 stream: 是否开启流式写入模式 outFile: 输出文件路径包含文件名
func SplitTableMergeSameKey(data map[string][][]string, header [][]string, stream bool, outFile string) error {
	if len(data) == 0 {
		return errors.New("待处理数据为空")
	}

	w := NewWriter()
	defer w.Close()

	for sheetName, rows := range data {
		if err := w.AddSheet(sheetName); err != nil {
			return fmt.Errorf("add sheet %s failed: %w", sheetName, err)
		}
		// 组合表头和数据
		allRows := make([][]string, 0, len(header)+len(rows))
		allRows = append(allRows, header...)
		allRows = append(allRows, rows...)

		if err := w.WriteAllRows(allRows, stream); err != nil {
			return err
		}
	}
	return w.SaveAs(outFile)
}

// 多组数据表合并的请求结构体
type MergeTableRequest struct {
	Header   [][]string          // 当前数据表的表头
	Data     map[string][]string // 当前数据表数据 除了表头
	Original bool                // 是否为以此表为原始数据表进行数据合并
}

// MergeTablesAppendList 将多组数据按相同索引列拼接成一张宽表。
// 基础表（Original=true）作为骨架，其他表的行若索引相同则追加在对应行后面。
// stream: 是否启用流式写入（推荐大数据量时开启）
// outFile: 输出文件完整路径
func MergeTablesAppendList(stream bool, outFile string, input ...MergeTableRequest) error {
	if len(input) < 2 {
		return errors.New("至少需要两个数据表")
	}

	// 1. 查找原始表（Original=true），且必须唯一
	var baseReq *MergeTableRequest
	for i := range input {
		if input[i].Original {
			if baseReq != nil {
				return errors.New("只能有一个原始表（Original=true）")
			}
			baseReq = &input[i]
		}
	}
	if baseReq == nil {
		return errors.New("缺少原始表，请将一个请求的 Original 设置为 true")
	}

	// 2. 初始化合并数据：复制原始表的数据
	merged := make(map[string][]string, len(baseReq.Data))
	for k, v := range baseReq.Data {
		merged[k] = append([]string{}, v...) // 复制切片，避免污染原数据
	}

	// 3. 将其它表的行追加到对应索引的行后面
	for _, req := range input {
		if req.Original {
			continue
		}
		for key, row := range req.Data {
			if exist, ok := merged[key]; ok {
				merged[key] = append(exist, row...)
			}
			// 如果原始表中没有这个 key，则忽略（根据需求）
		}
	}

	// 4. 对原始表的 key 排序（保持输出顺序）
	keys := make([]string, 0, len(baseReq.Data))
	for k := range baseReq.Data {
		keys = append(keys, k)
	}
	sort.Strings(keys) // 简单字典序，如需数值排序可替换为 SortMapByNumericKey

	// 5. 构建输出行：表头 + 数据
	var finalRows [][]string
	finalRows = append(finalRows, mergeHeaders(input...)...) // 合并表头
	for _, k := range keys {
		if row, ok := merged[k]; ok {
			finalRows = append(finalRows, row)
		}
	}

	// 6. 写入 Excel
	w := NewWriter()
	defer w.Close()

	if err := w.WriteAllRows(finalRows, stream); err != nil {
		return err
	}
	return w.SaveAs(outFile)
}

// mergeHeaders 合并所有请求的表头，保证每行的列数等于所有表列数之和
func mergeHeaders(requests ...MergeTableRequest) [][]string {
	// 1. 计算每个请求的列数（优先用 Header 第一行长度，否则用 Data 第一行长度）
	type reqMeta struct {
		colCount int
		headers  [][]string // 已补齐到最大行数，每行为该请求在该行的标题片段
	}
	metas := make([]reqMeta, len(requests))
	maxRows := 0
	for i, req := range requests {
		colCount := 0
		if len(req.Header) > 0 {
			colCount = len(req.Header[0])
		} else {
			// 尝试从 Data 获取列数
			for _, row := range req.Data {
				colCount = len(row)
				break
			}
		}
		metas[i].colCount = colCount
		metas[i].headers = req.Header
		if len(req.Header) > maxRows {
			maxRows = len(req.Header)
		}
	}

	// 2. 构建结果表头：maxRows 行，每行依次拼接所有请求在该行的标题片段（不足部分补空）
	result := make([][]string, maxRows)
	for r := 0; r < maxRows; r++ {
		row := make([]string, 0)
		for _, meta := range metas {
			if r < len(meta.headers) {
				row = append(row, meta.headers[r]...)
			} else {
				// 缺失的行，填充 colCount 个空字符串
				for c := 0; c < meta.colCount; c++ {
					row = append(row, "")
				}
			}
		}
		result[r] = row
	}
	return result
}

// MergeTablesReplaceList 将多组数据按索引键替换非空值，最终合并为一张表。
// 原始表（Original=true）决定最终输出的行集合和结构，其他表的非空单元格会覆盖对应位置。
// 若原始表中不存在的键，则直接忽略。
// stream: 是否启用流式写入
// outFile: 输出文件完整路径
func MergeTablesReplaceList(stream bool, outFile string, input ...MergeTableRequest) error {
	if len(input) < 2 {
		return errors.New("至少需要两个数据表")
	}

	// 查找原始表
	var baseReq *MergeTableRequest
	for i := range input {
		if input[i].Original {
			if baseReq != nil {
				return errors.New("只能有一个原始表（Original=true）")
			}
			baseReq = &input[i]
		}
	}
	if baseReq == nil {
		return errors.New("缺少原始表，请将一个请求的 Original 设置为 true")
	}

	// 确定列数：优先使用原始表头第一行的长度，否则从数据中推测最大列数
	numCols := 0
	if len(baseReq.Header) > 0 && len(baseReq.Header[0]) > 0 {
		numCols = len(baseReq.Header[0])
	} else {
		for _, row := range baseReq.Data {
			if len(row) > numCols {
				numCols = len(row)
			}
		}
	}
	if numCols == 0 {
		return errors.New("无法确定数据列数")
	}

	// 复制原始数据并补齐到 numCols（防止行尾空列丢失）
	merged := make(map[string][]string, len(baseReq.Data))
	for k, v := range baseReq.Data {
		row := make([]string, numCols)
		copy(row, v) // 不足的部分自动为空字符串
		merged[k] = row
	}

	// 用其他表非空值覆盖
	for _, req := range input {
		if req.Original {
			continue
		}
		for key, newRow := range req.Data {
			origRow, ok := merged[key]
			if !ok {
				continue
			}
			// 逐列覆盖：新值非空则更新，但不超出原始列数
			for i := 0; i < numCols && i < len(newRow); i++ {
				if newRow[i] != "" {
					origRow[i] = newRow[i]
				}
			}
		}
	}

	// 确定输出顺序
	var keys []string
	// 默认按字典序排序（保证输出稳定）
	keys = make([]string, 0, len(baseReq.Data))
	for k := range baseReq.Data {
		keys = append(keys, k)
	}
	sort.Strings(keys)

	// 构建最终行（表头 + 数据）
	finalRows := make([][]string, 0, len(baseReq.Header)+len(keys))
	finalRows = append(finalRows, baseReq.Header...)
	for _, k := range keys {
		if row, ok := merged[k]; ok {
			finalRows = append(finalRows, row)
		}
	}

	// 写入 Excel
	w := NewWriter(WithSheetName("Merged"))
	defer w.Close()

	if err := w.WriteAllRows(finalRows, stream); err != nil {
		return err
	}
	return w.SaveAs(outFile)
}
