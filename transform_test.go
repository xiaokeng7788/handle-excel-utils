package handleExcelUtils

import (
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func TestTransformRows(t *testing.T) {
	rows := [][]string{
		{"时间戳", "日期"},
		{"1777017316", "1777017316"},
		{"1777017316", "1777017316"},
	}

	transformers := map[int]TransformFunc{
		0: TimestampToDate(time.DateOnly),
		1: TimestampToDate(time.DateTime),
	}
	result := TransformRows(rows, 1, transformers)

	expected := [][]string{
		{"时间戳", "日期"},
		{"2026-04-24", "2026-04-24 15:55:16"},
		{"2026-04-24", "2026-04-24 15:55:16"},
	}
	assert.Equal(t, expected, result)
}

func TestMultiplyBy(t *testing.T) {
	fn := MultiplyBy(2.5)
	assert.Equal(t, "25", fn("10"))
	assert.Equal(t, "abc", fn("abc"))
}

func TestAppendSuffix(t *testing.T) {
	fn := AppendSuffix("kg")
	assert.Equal(t, "100kg", fn("100"))
}
