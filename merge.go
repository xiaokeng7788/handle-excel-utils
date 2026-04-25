package handleExcelUtils

import (
	"errors"
	"fmt"
	"sort"
	"strconv"
)

// MergeTableRequest 描述一个待合并的数据表。
type MergeTableRequest struct {
	Header   [][]string          // 表头（可多行）
	Data     map[string][]string // 数据（键为索引列的值）
	Original bool                // 是否为原始表（骨架表），有且仅有一个
}

// SplitTableMergeSameKey 按键拆分一对多数据为多工作表，每个工作表以键命名。
// stream：是否使用流式写入（适合大数据量）。
func SplitTableMergeSameKey(data map[string][][]string, header [][]string, stream bool, outFile string) error {
	if len(data) == 0 {
		return errors.New("待处理数据为空")
	}

	w, err := NewWriter()
	if err != nil {
		return err
	}
	defer w.Close()

	// 获取排序后的目标工作表名列表
	sheetNames := make([]string, 0, len(data))
	for name := range data {
		sheetNames = append(sheetNames, name)
	}
	sort.Strings(sheetNames)

	// 将默认的 Sheet1 重命名为第一个目标工作表（避免残留空白表）
	if err := w.file.SetSheetName("Sheet1", sheetNames[0]); err != nil {
		return fmt.Errorf("重命名默认工作表失败: %w", err)
	}
	w.sheetName = sheetNames[0]

	// 写入第一个工作表的数据
	allRows := make([][]string, 0, len(header)+len(data[sheetNames[0]]))
	allRows = append(allRows, header...)
	allRows = append(allRows, data[sheetNames[0]]...)
	if err := w.writeAll(stream, allRows); err != nil {
		return err
	}

	// 添加并写入其余工作表
	for _, name := range sheetNames[1:] {
		if err := w.AddSheet(name); err != nil {
			return fmt.Errorf("添加工作表 %q 失败: %w", name, err)
		}
		allRows = make([][]string, 0, len(header)+len(data[name]))
		allRows = append(allRows, header...)
		allRows = append(allRows, data[name]...)
		if err := w.writeAll(stream, allRows); err != nil {
			return err
		}
	}

	return w.SaveAs(outFile)
}

// MergeTablesAppendList 将多组数据按索引列拼接成一张宽表。
// 原始表（Original=true）决定行的存在与顺序，其他表的同键行会被追加到对应行尾部。
// sortNumeric：true 则按键的数值排序，false 则按字典序排序。
func MergeTablesAppendList(stream bool, sortNumeric bool, outFile string, input ...MergeTableRequest) error {
	if len(input) < 2 {
		return errors.New("至少需要两个数据表")
	}

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

	// 复制原始数据
	merged := make(map[string][]string, len(baseReq.Data))
	for k, v := range baseReq.Data {
		merged[k] = append([]string{}, v...)
	}

	// 拼接其他表的数据
	for _, req := range input {
		if req.Original {
			continue
		}
		for key, row := range req.Data {
			if exist, ok := merged[key]; ok {
				merged[key] = append(exist, row...)
			}
		}
	}

	// 排序键
	keys := make([]string, 0, len(baseReq.Data))
	for k := range baseReq.Data {
		keys = append(keys, k)
	}
	sortKeys(keys, sortNumeric)

	// 组装最终行
	var finalRows [][]string
	finalRows = append(finalRows, mergeHeaders(input...)...)
	for _, k := range keys {
		if row, ok := merged[k]; ok {
			finalRows = append(finalRows, row)
		}
	}

	w, err := NewWriter()
	if err != nil {
		return err
	}
	defer w.Close()

	if err := w.writeAll(stream, finalRows); err != nil {
		return err
	}
	return w.SaveAs(outFile)
}

// MergeTablesReplaceList 将多组数据按索引键替换非空单元格，合并为一张表。
// 原始表决定最终输出的行和列数，其他表的非空值会覆盖原始表对应位置。
func MergeTablesReplaceList(stream bool, sortNumeric bool, outFile string, input ...MergeTableRequest) error {
	if len(input) < 2 {
		return errors.New("至少需要两个数据表")
	}

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

	// 确定列数
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

	// 复制并补齐原始数据
	merged := make(map[string][]string, len(baseReq.Data))
	for k, v := range baseReq.Data {
		row := make([]string, numCols)
		copy(row, v)
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
			for i := 0; i < numCols && i < len(newRow); i++ {
				if newRow[i] != "" {
					origRow[i] = newRow[i]
				}
			}
		}
	}

	// 排序
	keys := make([]string, 0, len(baseReq.Data))
	for k := range baseReq.Data {
		keys = append(keys, k)
	}
	sortKeys(keys, sortNumeric)

	finalRows := make([][]string, 0, len(baseReq.Header)+len(keys))
	finalRows = append(finalRows, baseReq.Header...)
	for _, k := range keys {
		if row, ok := merged[k]; ok {
			finalRows = append(finalRows, row)
		}
	}

	w, err := NewWriter()
	if err != nil {
		return err
	}
	defer w.Close()

	if err := w.writeAll(stream, finalRows); err != nil {
		return err
	}
	return w.SaveAs(outFile)
}

// mergeHeaders 合并多个请求的表头，保证每行列数为各请求列数之和。
func mergeHeaders(requests ...MergeTableRequest) [][]string {
	type reqMeta struct {
		colCount int
		headers  [][]string
	}
	metas := make([]reqMeta, len(requests))
	maxRows := 0
	for i, req := range requests {
		colCount := 0
		if len(req.Header) > 0 {
			colCount = len(req.Header[0])
		} else {
			for _, row := range req.Data {
				colCount = len(row)
				break
			}
		}
		metas[i] = reqMeta{colCount: colCount, headers: req.Header}
		if len(req.Header) > maxRows {
			maxRows = len(req.Header)
		}
	}

	result := make([][]string, maxRows)
	for r := 0; r < maxRows; r++ {
		row := make([]string, 0)
		for _, meta := range metas {
			if r < len(meta.headers) {
				row = append(row, meta.headers[r]...)
			} else {
				for c := 0; c < meta.colCount; c++ {
					row = append(row, "")
				}
			}
		}
		result[r] = row
	}
	return result
}

func sortKeys(keys []string, numeric bool) {
	if numeric {
		sort.Slice(keys, func(i, j int) bool {
			ni, _ := strconv.Atoi(keys[i])
			nj, _ := strconv.Atoi(keys[j])
			return ni < nj
		})
	} else {
		sort.Strings(keys)
	}
}

// AppendExcelFiles 将源文件的所有工作表追加到目标文件末尾。
// 若目标文件不存在则创建；相同名称的工作表会合并行（追加在最后）。
func AppendExcelFiles(sourcePath, destPath string) error {
	srcReader, err := OpenReader(sourcePath)
	if err != nil {
		return err
	}
	defer srcReader.Close()

	// 打开或创建目标写入器
	var w *Writer
	if fileExists(destPath) {
		w, err = NewWriterFromFile(destPath)
		if err != nil {
			return err
		}
	} else {
		w, err = NewWriter()
		if err != nil {
			return err
		}
	}
	defer w.Close()

	for _, sheet := range srcReader.SheetNames() {
		rows, err := srcReader.GetRows(sheet)
		if err != nil {
			return fmt.Errorf("读取源文件工作表 %s 失败: %w", sheet, err)
		}
		// 检查目标文件是否已存在该工作表
		sheets := w.file.GetSheetList()
		exists := false
		for _, s := range sheets {
			if s == sheet {
				exists = true
				break
			}
		}
		if !exists {
			if err := w.AddSheet(sheet); err != nil {
				return err
			}
		}
		// 确保 Writer 当前操作工作表与目标工作表一致
		w.sheetName = sheet
		// 获取当前工作表已有行数（如果存在）
		existingRows, _ := w.file.GetRows(sheet)
		startRow := len(existingRows) + 1
		if err := w.WriteRows(startRow, rows); err != nil {
			return err
		}
	}
	return w.SaveAs(destPath)
}

// MergeSheets 将源文件的指定工作表追加到目标文件的指定工作表。
func MergeSheets(sourcePath, destPath, sourceSheet, destSheet string) error {
	srcReader, err := OpenReader(sourcePath)
	if err != nil {
		return err
	}
	defer srcReader.Close()

	rows, err := srcReader.GetRows(sourceSheet)
	if err != nil {
		return fmt.Errorf("读取源工作表失败: %w", err)
	}

	var w *Writer
	if fileExists(destPath) {
		w, err = NewWriterFromFile(destPath)
		if err != nil {
			return err
		}
	} else {
		w, err = NewWriter()
		if err != nil {
			return err
		}
	}
	defer w.Close()

	// 确保目标工作表存在
	sheets := w.file.GetSheetList()
	exists := false
	for _, s := range sheets {
		if s == destSheet {
			exists = true
			break
		}
	}
	if !exists {
		if err := w.AddSheet(destSheet); err != nil {
			return err
		}
	}
	// 显式设置当前工作表
	w.sheetName = destSheet
	existingRows, _ := w.file.GetRows(destSheet)
	startRow := len(existingRows) + 1
	if err := w.WriteRows(startRow, rows); err != nil {
		return err
	}
	return w.SaveAs(destPath)
}
