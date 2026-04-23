package handleExcelUtils

// WriterOption 是配置 Writer 的函数选项。
type WriterOption func(*Writer)

// WithSheetName 设置写入器使用的工作表名称。
// 若名称为空则忽略。
func WithSheetName(name string) WriterOption {
	return func(w *Writer) {
		if name != "" {
			w.sheetName = name
		}
	}
}

// WithDefaultSheet 创建写入器时同时创建默认工作表（"Sheet1"已被默认创建）。
// 此选项可用于显式声明。
func WithDefaultSheet() WriterOption {
	return func(w *Writer) {
		// 默认行为已创建，无需额外操作
	}
}
