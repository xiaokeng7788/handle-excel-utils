package handleExcelUtils

import (
	"bytes"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

// Reader 提供对 Excel 文件的只读访问。
// 使用完毕后必须调用 Close() 释放资源。
type Reader struct {
	file *excelize.File
	path string
}

// OpenReader 从文件路径打开 Excel 读取器。
// 若文件不存在或无法打开则返回错误。
func OpenReader(filePath string) (*Reader, error) {
	if !fileExists(filePath) {
		return nil, fmt.Errorf("文件不存在: %s", filePath)
	}
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("打开文件失败: %w", err)
	}
	return &Reader{file: f, path: filePath}, nil
}

// OpenReaderFromStream 从 io.Reader 打开 Excel（例如 HTTP 上传）。
// 调用者需确保流包含有效的 Excel 数据。
func OpenReaderFromStream(r io.Reader) (*Reader, error) {
	f, err := excelize.OpenReader(r)
	if err != nil {
		return nil, fmt.Errorf("打开数据流失败: %w", err)
	}
	return &Reader{file: f}, nil
}

// Close 关闭底层文件，释放资源。
func (r *Reader) Close() error {
	if r.file != nil {
		return r.file.Close()
	}
	return nil
}

// SheetNames 返回工作簿中所有工作表的名称。
func (r *Reader) SheetNames() []string {
	return r.file.GetSheetList()
}

// GetRows 获取指定工作表的所有行数据。
// 若 sheetName 为空字符串，则使用第一个工作表。
func (r *Reader) GetRows(sheetName string) ([][]string, error) {
	if sheetName == "" {
		sheets := r.SheetNames()
		if len(sheets) == 0 {
			return nil, errors.New("工作簿中不包含任何工作表")
		}
		sheetName = sheets[0]
	}
	return r.file.GetRows(sheetName)
}

// RowIterator 返回一个行迭代器，用于逐行读取大型工作表。
// 迭代器使用完毕需调用 Close()。
func (r *Reader) RowIterator(sheetName string) (*RowIterator, error) {
	if sheetName == "" {
		sheets := r.SheetNames()
		if len(sheets) == 0 {
			return nil, errors.New("工作簿中不包含任何工作表")
		}
		sheetName = sheets[0]
	}
	rows, err := r.file.Rows(sheetName)
	if err != nil {
		return nil, err
	}
	return &RowIterator{rows: rows}, nil
}

// RowIterator 封装 excelize.Rows，提供逐行读取功能。
type RowIterator struct {
	rows *excelize.Rows
}

// Next 准备下一行供读取，返回是否还有数据。
func (it *RowIterator) Next() bool {
	return it.rows.Next()
}

// Columns 返回当前行的所有单元格值。
func (it *RowIterator) Columns() ([]string, error) {
	return it.rows.Columns()
}

// Close 关闭迭代器，释放资源。
func (it *RowIterator) Close() error {
	if it.rows != nil {
		return it.rows.Close()
	}
	return nil
}

// Writer 提供对 Excel 文件的写入操作，支持流式与非流式。
type Writer struct {
	file       *excelize.File
	sheetName  string
	streamMode bool
	stream     *excelize.StreamWriter
}

// NewWriter 创建一个新的 Excel 写入器。
// 默认使用工作表 "Sheet1"，可通过选项自定义。
func NewWriter(opts ...WriterOption) *Writer {
	w := &Writer{
		file:      excelize.NewFile(),
		sheetName: "Sheet1",
	}
	for _, opt := range opts {
		opt(w)
	}
	// 如果用户指定了非默认工作表，则删除默认的 Sheet1 并创建新工作表
	if w.sheetName != "Sheet1" {
		// 删除默认工作表（忽略错误，因为默认一定存在）
		_ = w.file.DeleteSheet("Sheet1")
		_, _ = w.file.NewSheet(w.sheetName)
	}
	return w
}

// EnableStreamMode 启用流式写入模式，适用于写入海量数据。
// 启用后必须使用 WriteRows 并指定起始行，最后调用 Flush。
func (w *Writer) EnableStreamMode() error {
	sw, err := w.file.NewStreamWriter(w.sheetName)
	if err != nil {
		return fmt.Errorf("启用流式模式失败: %w", err)
	}
	w.streamMode = true
	w.stream = sw
	return nil
}

// WriteRow 在非流式模式下追加一行数据。
// 返回写入的行号（从1开始）。
func (w *Writer) WriteRow(row []string) (int, error) {
	if w.streamMode {
		return 0, errors.New("已启用流式模式，请使用 WriteRows 并指定起始行")
	}
	rows, err := w.file.GetRows(w.sheetName)
	if err != nil {
		return 0, fmt.Errorf("获取现有行数据失败: %w", err)
	}
	rowNum := len(rows) + 1
	cell, _ := excelize.CoordinatesToCellName(1, rowNum)
	if err := w.file.SetSheetRow(w.sheetName, cell, &row); err != nil {
		return 0, fmt.Errorf("设置工作表行失败: %w", err)
	}
	return rowNum, nil
}

// WriteRows 批量写入多行数据，从指定起始行号开始。
// 流式模式下会一次性写入缓冲区，非流式模式下逐行写入。
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

// Flush 将流式缓冲区的数据刷写到文件。
// 流式模式下必须在保存前调用。
func (w *Writer) Flush() error {
	if w.stream != nil {
		return w.stream.Flush()
	}
	return nil
}

// SaveAs 将 Excel 文件保存到指定路径。
// 会自动创建不存在的目录。
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

// WriteToBuffer 将 Excel 数据写入内存缓冲区，用于网络传输或进一步处理。
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

// Close 关闭写入器，释放资源。
func (w *Writer) Close() error {
	if w.stream != nil {
		_ = w.stream.Flush()
	}
	return w.file.Close()
}

// 内部辅助：判断文件是否存在（非目录）
func fileExists(path string) bool {
	info, err := os.Stat(path)
	if os.IsNotExist(err) {
		return false
	}
	return err == nil && !info.IsDir()
}
