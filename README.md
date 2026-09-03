# Writer.AI

Writer.AI is an AI-assisted formatting extension for LibreOffice Writer. It
converts natural-language instructions into a validated formatting plan, shows
the plan for confirmation, and applies changes as one undoable operation.

## Features

- Format the current document or selection.
- Format titles, headings, body paragraphs, fonts, colors, alignment, and indentation.
- Locate content by paragraph number, keyword, table name, row, and column.
- Format table headers, backgrounds, borders, fonts, alignment, row height, and column widths.
- Support zebra rows, automatic numeric/date alignment, first-column emphasis, cell merging, and numbered captions.
- Control table pagination with repeated headers, table splitting, and keep-together behavior.
- Preview formatting plans, undo the complete operation, cancel requests, and use a 60-second timeout.
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
The default preset uses Kimi K3 through Alibaba Cloud Bailian. The API key is
stored in LibreOffice's password container and is not written to the project
configuration file or application logs. Click **Save Settings** to keep the
configuration for future Writer sessions.

## Usage

1. Open a Writer document.
2. Select **Writer.AI > AI Formatter**.
3. Enter an instruction, such as `Indent every paragraph by two characters`.
4. Review the validated plan and choose **Yes** to apply it.
5. Use LibreOffice's `Ctrl+Z` to revert the complete formatting operation if needed.

While a request is running, the Writer status bar shows the analysis state.
Requests use a 60-second timeout. This does not undo formatting that has
already been applied; use `Ctrl+Z` to undo applied changes.

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

See [README.zh-CN.md](README.zh-CN.md) for the Chinese documentation and
[CHANGELOG.md](CHANGELOG.md) for release history.

## Copyright and usage

Copyright (c) 2026 Anna Wu. All rights reserved.

The official release package may be installed and used for personal,
non-commercial purposes. Copying source code, modification, redistribution,
rebranding, resale, and commercial use require prior written permission. See
[LICENSE](LICENSE) for the full notice.
