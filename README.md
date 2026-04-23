# handle-excel-utils

一个轻量、安全的 Excel 读写工具库，基于 [excelize](https://github.com/xuri/excelize/v2) 封装。

## 特性

- **清晰的读写分离**：`Reader` 用于读取，`Writer` 用于写入。
- **显式资源管理**：必须调用 `Close()` 释放文件句柄。
- **流式写入支持**：处理海量数据时内存占用低。
- **灵活的配置**：通过选项模式设置工作表名称。
- **实用的转换函数**：将二维行数据转换为 Map 结构，便于业务处理。
- **完善的错误处理**：不忽略任何错误，符合 Go 惯例。

## 安装

```bash
go get github.com/xiaokeng7788/handle-excel-utils