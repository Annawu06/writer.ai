# Changelog

## 1.1.1 - 2026-09-03

- Enable persistent LibreOffice password storage for API keys.
- Enforce the current formatting JSON structure in the model prompt.
- Show a clear error when no executable formatting instructions are returned.

## 1.1.0 - 2026-09-03

- Correctly remove bold, italic, and underline formatting when requested.
- Add regression coverage for formatting removal.
- Remove the unused request-cancellation menu command.

## 1.0.0 - 2026-09-03

First release of Writer.AI for LibreOffice Writer.

- Format titles, headings, body paragraphs, fonts, colors, alignment, and indentation.
- Format tables with headers, colors, borders, widths, heights, zebra rows, merging, and automatic numeric alignment.
- Add numbered table captions and control table pagination.
- Support configurable OpenAI-compatible providers, models, Base URLs, and persistent API keys.
- Run model requests asynchronously with timeout, cancellation, validation preview, and undo.
