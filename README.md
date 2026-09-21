# Word efn2Footnotes

A Microsoft Word VBA macro that converts tagged text into standard Word footnotes.

For example:

```text
This is some text.[efn_note]This explanatory text becomes a footnote.[/efn_note]
```

becomes normal document text followed by a Word footnote reference number. The text between the tags is moved into the footnote area, and Word provides its normal footnote navigation/hyperlink behavior.

## Compatibility

- Microsoft Word 2010
- VBA macro-enabled Word documents (`.docm`)
- Later desktop versions of Word should also be compatible

## Installation

1. Download `Efn2Footnotes.bas` from this repository.
2. Open Word and either open an existing document or create a new one.
3. Press **Alt+F11** to open the Visual Basic Editor.
4. Select **File > Import File...**.
5. Select `Efn2Footnotes.bas`.
   - Alternatively, choose **Insert > Module** and paste the contents of the `.bas` file into the new module.
6. Save the document as a **Word Macro-Enabled Document (*.docm)** if the macro is stored in that document.
7. Close the Visual Basic Editor.

If Word displays a security warning, click **Enable Content** only when you trust the document and its source. Macro security can also be configured under **File > Options > Trust Center > Trust Center Settings > Macro Settings**.

## Usage

1. Make a backup copy of the document before conversion.
2. Ensure each note uses this exact format:

   ```text
   [efn_note]Footnote text goes here.[/efn_note]
   ```

3. Press **Alt+F8** in Word.
4. Select `ConvertEfnTagsToFootnotes`.
5. Click **Run**.

The macro processes tagged blocks in document order. Each complete block is replaced by a standard Word footnote reference. The macro displays a count when it finishes.

## Important behavior

- Tags are case-sensitive: use `[efn_note]` and `[/efn_note]` exactly.
- The first closing tag after an opening tag is treated as its matching closing tag.
- Multiple tagged blocks are supported.
- Line breaks inside a tagged block are converted to spaces in the footnote.
- If an opening tag has no closing tag, the macro stops safely, leaves that unmatched content unchanged, and reports the problem.
- The macro searches the main document body. It does not intentionally process headers, footers, text boxes, comments, or existing footnote text.
- Footnote numbering, placement, and formatting use Word's current footnote settings.

## License

MIT License. See `LICENSE` in the repository.
