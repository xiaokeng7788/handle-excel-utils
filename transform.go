package handleExcelUtils

import (
	"strconv"
	"time"
)

// TransformFunc 列转换函数：输入原始字符串，返回转换后字符串。
type TransformFunc func(string) string

// TransformRows 对二维行数据应用列转换，跳过 headerRows 个表头行。
// transformers 的键为列索引（0-based）。返回新的切片，原数据不受影响。
func TransformRows(rows [][]string, headerRows int, transformers map[int]TransformFunc) [][]string {
	newRows := make([][]string, len(rows))
	for i, row := range rows {
		newRow := make([]string, len(row))
		copy(newRow, row)
		if i >= headerRows {
			for colIdx, fn := range transformers {
				if colIdx < len(newRow) {
					newRow[colIdx] = fn(newRow[colIdx])
				}
			}
		}
		newRows[i] = newRow
	}
	return newRows
}

// TransformRowsByName 类似 TransformRows，但通过列名指定转换函数。
// 内部自动在 rows[headerRows-1]（最后一行表头）中查找列名。
// 若 headerRows <= 0，则不进行任何转换（无表头可参考）。
func TransformRowsByName(rows [][]string, headerRows int, transformers map[string]TransformFunc) [][]string {
	if headerRows <= 0 || len(transformers) == 0 || len(rows) == 0 {
		// 无表头或转换规则，退化为普通拷贝
		return TransformRows(rows, headerRows, nil)
	}
	// 取最后一行为表头
	headerRow := rows[headerRows-1]
	colIndexMap := make(map[int]TransformFunc, len(transformers))
	for name, fn := range transformers {
		idx := GetColumnIndex(headerRow, name)
		if idx != -1 {
			colIndexMap[idx] = fn
		}
	}
	return TransformRows(rows, headerRows, colIndexMap)
}

// TimestampToDate 将秒级 Unix 时间戳转换为指定格式的日期字符串。
// layout 示例："2006-01-02" 或 "2006-01-02 15:04:05"。
func TimestampToDate(layout string) TransformFunc {
	return func(s string) string {
		ts, err := strconv.ParseInt(s, 10, 64)
		if err != nil {
			return s
		}
		return time.Unix(ts, 0).Format(layout)
	}
}

// MultiplyBy 将列值乘以一个浮点因子（用于单位转换等）。
func MultiplyBy(factor float64) TransformFunc {
	return func(s string) string {
		val, err := strconv.ParseFloat(s, 64)
		if err != nil {
			return s
		}
		return strconv.FormatFloat(val*factor, 'f', -1, 64)
	}
}

// AppendSuffix 在单元格末尾追加后缀。
func AppendSuffix(suffix string) TransformFunc {
	return func(s string) string {
		return s + suffix
	}
}
