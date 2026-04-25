package handleExcelUtils

import (
	"bytes"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

// Writer 提供 Excel 文件的写入操作，支持流式与非流式。
type Writer struct {
	file       *excelize.File
	sheetName  string
	streamMode bool
	stream     *excelize.StreamWriter
}

// NewWriter 创建一个新的 Excel 写入器。默认包含一个工作表 "Sheet1"，
// 可通过 WithSheetName 重命名。返回的 Writer 必须在使用后调用 Close()。
func NewWriter(opts ...WriterOption) (*Writer, error) {
	f := excelize.NewFile()
	w := &Writer{
		file:      f,
		sheetName: "Sheet1",
	}
	for _, opt := range opts {
		opt(w)
	}
	if w.sheetName != "Sheet1" {
		if err := f.SetSheetName("Sheet1", w.sheetName); err != nil {
			return nil, fmt.Errorf("重命名默认工作表失败: %w", err)
		}
	}
	return w, nil
}

// AddSheet 添加一个命名工作表并切换为当前操作对象。
// 如果工作表已存在则返回错误。
func (w *Writer) AddSheet(name string) error {
	// 确保之前的流写入已刷盘
	if w.stream != nil {
		if err := w.stream.Flush(); err != nil {
			return err
		}
		w.stream = nil
		w.streamMode = false
	}

	// 检查是否已存在同名工作表
	for _, s := range w.file.GetSheetList() {
		if s == name {
			return fmt.Errorf("工作表 %q 已存在", name)
		}
	}

	idx, err := w.file.NewSheet(name)
	if err != nil {
		return fmt.Errorf("创建工作表失败: %w", err)
	}
	w.file.SetActiveSheet(idx)
	w.sheetName = name
	w.streamMode = false
	w.stream = nil
	return nil
}

// EnableStreamMode 启用流式写入模式，用于写入海量数据。
// 如果之前已有未刷新的流，会先自动 Flush。
func (w *Writer) EnableStreamMode() error {
	if w.stream != nil {
		if err := w.stream.Flush(); err != nil {
			return err
		}
		w.stream = nil
	}
	sw, err := w.file.NewStreamWriter(w.sheetName)
	if err != nil {
		return fmt.Errorf("启用流式模式失败: %w", err)
	}
	w.streamMode = true
	w.stream = sw
	return nil
}

// WriteRow 在非流式模式下追加一行数据，返回行号（从1开始）。
func (w *Writer) WriteRow(row []string) (int, error) {
	if w.streamMode {
		return 0, errors.New("已启用流式模式，请使用 WriteRows 或先调用 Flush")
	}
	rows, err := w.file.GetRows(w.sheetName)
	if err != nil {
		return 0, fmt.Errorf("获取现有行数据失败: %w", err)
	}
	rowNum := len(rows) + 1
	cell, _ := excelize.CoordinatesToCellName(1, rowNum)
	if err := w.file.SetSheetRow(w.sheetName, cell, &row); err != nil {
		return 0, fmt.Errorf("写入行失败: %w", err)
	}
	return rowNum, nil
}

// WriteRows 批量写入多行，从 startRow 开始（1-based）。
// 流式模式：写入缓冲区，需调用 Flush 持久化。
// 非流式模式：逐行立即写入。
func (w *Writer) WriteRows(startRow int, rows [][]string) error {
	if w.streamMode && w.stream != nil {
		for i, row := range rows {
			cell, _ := excelize.CoordinatesToCellName(1, startRow+i)
			vals := make([]interface{}, len(row))
			for j, v := range row {
				vals[j] = v
			}
			if err := w.stream.SetRow(cell, vals); err != nil {
				return fmt.Errorf("流式设置第 %d 行失败: %w", startRow+i, err)
			}
		}
		return nil
	}
	// 非流式
	for i, row := range rows {
		cell, _ := excelize.CoordinatesToCellName(1, startRow+i)
		if err := w.file.SetSheetRow(w.sheetName, cell, &row); err != nil {
			return fmt.Errorf("设置第 %d 行失败: %w", startRow+i, err)
		}
	}
	return nil
}

// Flush 将流式缓冲区的数据刷写到文件。非流式模式无操作。
func (w *Writer) Flush() error {
	if w.stream != nil {
		return w.stream.Flush()
	}
	return nil
}

// WriteAll 非流式批量写入全部数据（追加在现有数据之后）。
func (w *Writer) WriteAll(data [][]string) error {
	for _, row := range data {
		if _, err := w.WriteRow(row); err != nil {
			return err
		}
	}
	return nil
}

// WriteAllStream 流式批量写入全部数据（从第1行开始，覆盖现有内容）。
// 适用于新工作表或需要重写整个工作表的情况。
func (w *Writer) WriteAllStream(data [][]string) error {
	if err := w.EnableStreamMode(); err != nil {
		return err
	}
	if err := w.WriteRows(1, data); err != nil {
		return err
	}
	return w.Flush()
}

// SaveAs 将 Excel 文件保存到指定路径，会自动创建目录。
func (w *Writer) SaveAs(filePath string) error {
	if w.stream != nil {
		if err := w.stream.Flush(); err != nil {
			return err
		}
	}
	dir := filepath.Dir(filePath)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return fmt.Errorf("创建目录失败: %w", err)
	}
	return w.file.SaveAs(filePath)
}

// WriteToBuffer 将 Excel 数据写入内存缓冲区，用于网络传输等场景。
func (w *Writer) WriteToBuffer() (*bytes.Buffer, error) {
	if w.stream != nil {
		if err := w.stream.Flush(); err != nil {
			return nil, err
		}
	}
	buf := new(bytes.Buffer)
	if _, err := w.file.WriteTo(buf); err != nil {
		return nil, fmt.Errorf("写入缓冲区失败: %w", err)
	}
	return buf, nil
}

// Close 关闭写入器，释放资源。会自动 Flush 流式数据。
func (w *Writer) Close() error {
	if w.stream != nil {
		if err := w.stream.Flush(); err != nil {
			w.file.Close()
			return err
		}
	}
	return w.file.Close()
}

// writeAll 内部方法，根据参数选择流式或非流式写入。
func (w *Writer) writeAll(stream bool, data [][]string) error {
	if stream {
		return w.WriteAllStream(data)
	}
	return w.WriteAll(data)
}

// AutoFitColumns 根据当前工作表的内容自动调整列宽。
// maxWidth 为列宽上限（单位：字符数），0 表示不设上限。
// 默认使用 excelize 的默认最小宽度。
func (w *Writer) AutoFitColumns(maxWidth float64) error {
	rows, err := w.file.GetRows(w.sheetName)
	if err != nil {
		return fmt.Errorf("获取行数据失败: %w", err)
	}
	if len(rows) == 0 {
		return nil
	}
	// 找出最大列数
	colCount := 0
	for _, row := range rows {
		if len(row) > colCount {
			colCount = len(row)
		}
	}
	for col := 0; col < colCount; col++ {
		maxLen := 0
		for _, row := range rows {
			if col < len(row) {
				cell := row[col]
				// 粗略计算中文字符宽度（1中文≈2英文字符宽度）
				length := 0
				for _, r := range cell {
					if r > 127 || utf8.RuneLen(r) > 1 {
						length += 2
					} else {
						length++
					}
				}
				if length > maxLen {
					maxLen = length
				}
			}
		}
		width := float64(maxLen) + 2 // 加一点边距
		if maxWidth > 0 && width > maxWidth {
			width = maxWidth
		}
		if width < 10 {
			width = 10
		}
		colName, _ := excelize.ColumnNumberToName(col + 1)
		if err := w.file.SetColWidth(w.sheetName, colName, colName, width); err != nil {
			return err
		}
	}
	return nil
}

// NewWriterFromFile 打开已有 Excel 文件进行追加或修改。
// 默认操作工作表为第一个 sheet，可通过 WithSheetName 选项指定。
func NewWriterFromFile(filePath string, opts ...WriterOption) (*Writer, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("打开文件失败: %w", err)
	}
	w := &Writer{
		file: f,
	}
	// 应用选项
	for _, opt := range opts {
		opt(w)
	}
	// 若未指定工作表，默认使用第一个工作表
	if w.sheetName == "" {
		sheets := f.GetSheetList()
		if len(sheets) > 0 {
			w.sheetName = sheets[0]
		} else {
			w.sheetName = "Sheet1"
		}
	}
	return w, nil
}

// SetCellValue 设置指定单元格的值。
func (w *Writer) SetCellValue(row, col int, value interface{}) error {
	cell, _ := excelize.CoordinatesToCellName(col, row)
	return w.file.SetCellValue(w.sheetName, cell, value)
}

// AutoFilter 为指定区域设置自动筛选。
func (w *Writer) AutoFilter(startRow, startCol, endRow, endCol int) error {
	startCell, _ := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, _ := excelize.CoordinatesToCellName(endCol, endRow)
	return w.file.AutoFilter(w.sheetName, fmt.Sprintf("%s:%s", startCell, endCell), nil)
}

// WriteFromMaps 将 []map[string]string 写入当前工作表。
// 如果 headers 为空，则使用 map 的键（按字母排序）作为表头。
func (w *Writer) WriteFromMaps(data []map[string]string, headers ...string) error {
	if len(data) == 0 {
		return nil
	}
	var headerRow []string
	if len(headers) > 0 && len(headers[0]) > 0 {
		headerRow = headers
	} else {
		// 从第一个 map 提取 key 并排序
		first := data[0]
		keys := make([]string, 0, len(first))
		for k := range first {
			keys = append(keys, k)
		}
		sort.Strings(keys)
		headerRow = keys
	}
	// 写入表头
	if _, err := w.WriteRow(headerRow); err != nil {
		return err
	}
	// 写入数据行
	for _, record := range data {
		row := make([]string, len(headerRow))
		for i, key := range headerRow {
			row[i] = record[key]
		}
		if _, err := w.WriteRow(row); err != nil {
			return err
		}
	}
	return nil
}
