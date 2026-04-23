package handleExcelUtils

// WriterOption 是配置 Writer 的函数选项。
type WriterOption func(*Writer)

// WithSheetName 设置写入器默认工作表的名称。
// 新文件将把默认的 "Sheet1" 重命名为该名称。
func WithSheetName(name string) WriterOption {
	return func(w *Writer) {
		if name != "" {
			w.sheetName = name
		}
	}
}

// WithDefaultSheet 显式声明使用默认工作表（已为默认行为，无需调用）。
func WithDefaultSheet() WriterOption {
	return func(w *Writer) {}
}
