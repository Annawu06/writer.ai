# writer.ai

AI formatting extension for LibreOffice Writer.

## Install

1. Build `writer.ai.oxt` with `./build.sh`.
2. In LibreOffice, open **Tools > Extension Manager**.
3. Select **Add** and choose `writer.ai.oxt`.

The extension adds Writer AI formatting commands to the Tools menu.

### **Function**

user can choose to format current file or choose other files

support certain paragraph,page and whole document

Font style

1. font name (even the semantic one) ✅

2. Bold ✅ or remove

3.  Italic ✅ or remove 

4. underline (user-defined color and different format) ✅ or remove

5. font size ✅

6. font color✅ or remove

7. Highlight (user-defined color) or remove ✅

   

Paragraph

1. alignment (left,right,center and justify) ✅

2. insert Text before or after certain line ✅and replace text ✅

   

Remove all format ✅



What i need to do :

1. achieve  formating title and paragraph ✅
2. table formating ✅
3. thinking about

### API configuration

The settings dialog supports a provider preset or a custom OpenAI-compatible
API. Users can enter their own API key, model name, and Base URL. The default
configuration is Kimi K3 through Alibaba Cloud Bailian.

The API key is stored in LibreOffice's password container. Provider, Base URL,
and model can be changed independently in Settings. The extension uses a
60-second request timeout, runs requests in the background, and supports
cancellation from the Tools menu.

## Development checks

Run the complete test suite with:

```sh
make test
```

This includes real headless LibreOffice document tests and DOCX round-trip
tests. Build the release package with `./build.sh`.
