# Fast Excel Hex Analyzer

Fast Excel Hex Analyzer is a program for Microsoft Excel that allows you to view files in HEX format, quickly navigate through data, and perform color highlighting on desired bytes.

---

## Features

- View and navigate files in HEX format.
- Jump to the desired offset.
- Support for color highlighting (tags) for specific bytes.

---

## Limitations

Please note that the program is intended for working with very small files (up to ~5 MB), as loading larger files can take significant time and may cause Excel to hang.

---

## Requirements

- Microsoft Excel (with macros enabled).
- Windows (recommended for full compatibility with VBA macros).

---

## Installation

1. Download the Excel file from the repository.
2. Open the file in Excel.
3. Enable macros when prompted.
4. Start working with the data by clicking the **"Open hex tool"** button.

---

## Interface

### Sheet

- #### Offset Panel
  - **First Column (A)**: Displays the offset of the data row in HEX format.

- #### Edit Panel
  - **Columns C-R**: Contain data bytes (16 bytes per row) in HEX format.

- #### Decoded Text Panel
  - **Columns T-AI**: Represent the decoded text corresponding to the bytes in the row.

### Main Form

- #### Selected Tags
  - Shows the selected tags.

- #### Info
  - Displays the current offset and the number of selected byte cells.

---

## How to Use

1. Click the **"Open hex tool"** button on the current sheet in Excel.
2. In the main application form, click the **"Load file"** button to load data from a file in HEX format. In the opened window, select the desired file and wait for the data to fully load into the Excel table until the message **"File loaded successfully."** appears (this may take considerable time depending on the file size).
3. To delete the data and all associated tags, click the **"Delete data"** button.
4. After loading the data, you can add tags (see the next section) for quick byte highlighting; they will be displayed in both the Edit Panel and the Decoded Text Panel.

---

## Go To

Navigating by byte offsets is done using a separate window that can be opened by clicking the **"GoTo"** button on the main form. Select a single cell, then choose where to start counting the new offset—from the beginning of the file or from the position of the currently selected cell. Enter in the text field how many bytes to move in HEX format (the value in DEC format will be displayed above it). Click the **"Move to"** button to jump to the new position.

---

## Tags

1. Select a range of cells, either in the Edit Panel or the Decoded Text Panel.
2. ### To Create a Tag:
   - In the main frame, enter your tag description in the **(Enter the tag description)** field.
   - Choose a color for displaying the tag by entering the necessary RGB values or in HEX format.
   - Click the **"Set tag"** button to apply the tag, or **"Delete tag"** to remove the tag from the selected cells.

---

## License

```plaintext
MIT License

Copyright (c) 2024 [Nazuha26]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
