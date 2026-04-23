package handleExcelUtils

import (
	"bytes"
	"errors"
	"fmt"
	"os"
	"path/filepath"

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
