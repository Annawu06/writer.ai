# Changelog

## Unreleased

- Reject stale AI plans if the document changes while a request is pending;
  selection instructions also require the original selection.
- Check paragraph and table targets before editing, and reject ambiguous
  paragraph text matches.
- Reject the entire model plan when validation drops invalid instructions.
- Clarify that numbered paragraphs and visual page lines are different targets.
- Add guarded JSON parsing, Qwen JSON mode, a connection test, failure recovery
  tests, and automated OXT builds.
- Preserve multi-paragraph ranges while setting alignment; remove the
  document/selection substring heuristic and input example prefill.
- Confirm content-changing plans, allow keyless loopback providers, require
  HTTPS for remote providers, and refresh extension descriptions and test setup.

## 1.1.2 - 2026-09-04

- Queue consecutive formatting requests and apply them in submission order.
- Remove the formatting-plan confirmation dialog.

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
- Run model requests asynchronously with timeout and undo.
