package handleExcelUtils

import (
	"errors"
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

// Reader 提供对 Excel 文件的只读访问。
// 使用完毕后必须调用 Close() 释放资源。
type Reader struct {
	file *excelize.File
}

// OpenReader 从文件路径打开 Excel 读取器。
func OpenReader(filePath string) (*Reader, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("打开文件失败: %w", err)
	}
	return &Reader{file: f}, nil
}

// OpenReaderFromStream 从 io.Reader 打开 Excel（例如 HTTP 上传）。
func OpenReaderFromStream(r io.Reader) (*Reader, error) {
	f, err := excelize.OpenReader(r)
	if err != nil {
		return nil, fmt.Errorf("打开数据流失败: %w", err)
	}
	return &Reader{file: f}, nil
}

// Close 关闭底层文件，释放资源。
func (r *Reader) Close() error {
	return r.file.Close()
}

// SheetNames 返回工作簿中所有工作表的名称。
func (r *Reader) SheetNames() []string {
	return r.file.GetSheetList()
}

// GetRows 获取指定工作表的所有行。sheetName 为空时使用第一个工作表。
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

// RowIterator 返回行迭代器，用于逐行读取大文件。使用完毕需调用 Close()。
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

// RowIterator 封装 excelize.Rows，提供逐行读取。
type RowIterator struct {
	rows *excelize.Rows
}

// Next 准备下一行。
func (it *RowIterator) Next() bool {
	return it.rows.Next()
}

// Columns 返回当前行的所有单元格值。
func (it *RowIterator) Columns() ([]string, error) {
	return it.rows.Columns()
}

// Close 关闭迭代器。
func (it *RowIterator) Close() error {
	return it.rows.Close()
}
