# Writer.AI

Writer.AI 是面向 LibreOffice Writer 的 AI 辅助格式化扩展。它将自然语言
指令转换为经过校验的格式化计划，供用户确认后执行，并将整次修改作为一个
可撤销操作处理。

## 功能

- 格式化当前文档或当前选区。
- 格式化标题、各级标题、正文、字体、颜色、对齐方式和段首缩进。
- 按段落序号、关键词、表格名称、行和列定位内容。
- 设置表头、背景、边框、字体、对齐、行高和列宽。
- 支持奇偶行颜色、数字/日期自动对齐、首列加粗、单元格合并和表格编号标题。
- 支持重复表头、表格跨页和表格与后续内容保持在一起。
- 支持格式化预览、整次撤销、请求取消和 60 秒超时。
- 支持 Kimi K3 及其他 OpenAI 兼容 API 服务商。

## 运行要求

- LibreOffice 25.8.4 或更高版本；
- 已支持服务商的 API Key；
- 能够访问所选 API 接口的网络环境。

## 安装

1. 构建或下载 `writer.ai.oxt`。
2. 打开 LibreOffice，选择 **工具 > 扩展管理器**。
3. 点击 **添加**，选择 `writer.ai.oxt`。
4. 如果 Writer.AI 菜单没有立即出现，请重启 LibreOffice。

## 配置

打开 **Writer.AI > AI Formatter > Setting**，配置服务商、模型名称和 API Key。
选择预设时程序会自动填写 Base URL，只有选择高级的自定义 OpenAI 兼容接口时
才会显示 Base URL 输入框。默认配置为阿里云百炼的 Kimi K3。API Key 保存在
LibreOffice 密码容器中，不会写入项目配置文件或应用日志。配置完成后点击
“保存配置”，以后打开 Writer 会自动读取这些设置。

## 使用

1. 打开 Writer 文档。
2. 选择 **Writer.AI > AI Formatter**。
3. 输入格式化指令，例如：`将每个段落段首缩进 2 个字符`。
4. 检查格式化计划，选择 **Yes** 应用。
5. 如需恢复，确认撤销提示即可撤销本次完整修改。

请求执行期间，Writer 状态栏会显示分析状态。可以通过 **Writer.AI >
AI Formatter > Cancel Formatting** 取消请求。

## 开发

构建扩展包：

```sh
./build.sh
```

运行完整测试：

```sh
make test
```

测试包含真实无头 LibreOffice 文档测试、DOCX 往返测试、表格格式化测试和
API 返回内容校验测试。

英文文档请参阅 [README.md](README.md)，版本记录请参阅
[CHANGELOG.md](CHANGELOG.md)。

## 版权所有与使用限制

Copyright (c) 2026 Anna Wu. All rights reserved.

官方发行包允许用户安装并用于个人、非商业用途。复制源代码、修改、再发布、
更名、转售以及商业使用，均须事先获得版权所有者的书面许可。完整声明请参阅
[LICENSE](LICENSE)。
