# Writer.AI

Writer.AI is an AI-assisted formatting extension for LibreOffice Writer. It
converts natural-language instructions into a validated formatting plan and
applies changes as one undoable operation.

## Features

- Format the current document or selection.
- Format titles, headings, body paragraphs, fonts, colors, alignment, and indentation.
- Locate content by paragraph number, keyword, table name, row, and column.
- Format table headers, backgrounds, borders, fonts, alignment, row height, and column widths.
- Support zebra rows, automatic numeric/date alignment, first-column emphasis, cell merging, and numbered captions.
- Control table pagination with repeated headers, table splitting, and keep-together behavior.
- Apply validated formatting plans automatically, undo the complete operation,
  and use a 30-second timeout.
- Recover JSON plans from fenced or prose-wrapped replies, validate numeric
  ranges, and roll back changes if a formatting command fails.
- Request JSON output mode for Qwen models on DashScope.
- Confirm plans that insert or replace text or merge cells.
- Use a loopback OpenAI-compatible endpoint without an API key; remote endpoints
  require HTTPS and an API key.
- Work with Kimi K3 and other OpenAI-compatible API providers.

## Requirements

- LibreOffice 25.8.4 or later.
- An API key for a supported provider.
- Network access to the selected API endpoint.

## Installation

1. Build or download `writer.ai.oxt`.
2. Open LibreOffice and select **Tools > Extension Manager**.
3. Select **Add** and choose `writer.ai.oxt`.
4. Restart LibreOffice if the Writer.AI menu does not appear immediately.

## Configuration

Open **Writer.AI > AI Formatter > Setting** and configure the provider, model
name, and API key. Provider presets fill the Base URL automatically; the Base
URL field is shown only for the advanced custom OpenAI-compatible provider.
For a local provider such as Ollama, choose the custom preset and enter its
loopback API URL and model name; the API key can be empty.
The default preset uses Kimi K3 through Alibaba Cloud Bailian. The API key is
stored in LibreOffice's password container and is not written to the project
configuration file or application logs. Use **Test Connection** to check the
current fields without saving them. Click **Save Settings** to keep the
configuration for future Writer sessions.

## Usage

1. Open a Writer document.
2. Select **Writer.AI > AI Formatter**.
3. Enter an instruction, such as `Indent every paragraph by two characters`.
4. The validated plan is applied automatically.
   Plans that change text or merge cells ask for confirmation first.
   A completion message reports the number of planned properties.
5. Use LibreOffice's Ctrl+Z to revert the complete formatting operation if needed.

While a request is running, the Writer status bar shows the analysis state.
Requests use a 30-second timeout. If the document changes while a request is
pending, its plan is discarded. Selection instructions also require the
original selection to remain in place. Ambiguous paragraph matches are
rejected before formatting. When an operation fails, Writer.AI attempts to
roll back changes in its undo context; recovery also depends on LibreOffice's
undo support for the specific operation. Use `Ctrl+Z` to undo a successful request.

## Development

Build the extension package:

```sh
./build.sh
```

Run the complete test suite:

```sh
make test
```

The suite includes real headless LibreOffice document tests, DOCX round-trip
tests, table formatting tests, and API response validation tests.
GitHub Actions runs the tests and builds an OXT on pushes and pull requests;
tagged pushes also attach the package to a GitHub release.

See [README.zh-CN.md](README.zh-CN.md) for the Chinese documentation and
[CHANGELOG.md](CHANGELOG.md) for release history.

## Copyright and usage

Copyright (c) 2026 Anna Wu. All rights reserved.

The official release package may be installed and used for personal,
non-commercial purposes. Copying source code, modification, redistribution,
rebranding, resale, and commercial use require prior written permission. See
[LICENSE](LICENSE) for the full notice.
