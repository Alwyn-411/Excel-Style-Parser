# Excel Style Parser

## Overview

Excel Style Parser is a React-based web application that allows users to upload Excel files and analyze their cell styles and formatting. Built with TypeScript, Vite, and ExcelJS, this tool provides a detailed view of cell styles, including font properties, fill colors, alignments, and more.

## Features

- 📄 Excel File Upload
- 🔍 Detailed Cell Style Parsing
- 📊 Pagination of Parsed Styles
- 🎨 Rich Styling Visualization
- 💻 Supports .xlsx and .xls file formats

## Key Components

### ParseWithStyles (src/utils/ParseWithStyles.ts)

A utility class for extracting and parsing Excel file styles:

- Extracts detailed cell styles including:

  - Font properties (name, size, color, bold, italic)
  - Fill colors
  - Cell alignments
  - Border styles

- Methods:
  - `parseExcelStyles`: Parses entire Excel worksheet
  - `filterStyles`: Optional filtering of cell styles
  - `extractCellStyle`: Detailed style extraction logic

### ExcelReader (src/components/ExcelReader.tsx)

A React component that handles:

- File upload interface
- Excel file parsing
- Rendering parsed cell styles
- Pagination of style results
- Dynamic style rendering for different cell value types

## Technologies Used

- React
- TypeScript
- Vite
- ExcelJS
- React Hooks

## Installation

1. Clone the repository

```bash
git clone https://github.com/Alwyn-411/Excel-Style-Parser.git
```

2. Install dependencies

```bash
npm install
```

3. Run the development server

```bash
npm run dev
```

## Usage

1. Click "Choose Excel File"
2. Select an .xlsx or .xls file
3. View parsed cell styles with detailed formatting information

## Example

```typescript
// Parsing Excel styles
const stylesData = await ParseWithStyles.parseExcelStyles(file);

// Filtering bold cells
const boldCells = ParseWithStyles.filterStyles(
  stylesData,
  (cellData) => cellData.style?.font?.bold === true
);
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the GNU General Public License v3.0 (GPL-3.0).

The GNU General Public License is a free, copyleft license for software and other kinds of works. It ensures that the software remains free and that any modifications or extensions are also shared under the same terms.

### Key Provisions:

- You are free to use, modify, and distribute this software
- Any derivative work must also be distributed under the GPL-3.0
- You must include the original copyright notice
- Changes must be documented
- The source code must be made available

### Full License Text

For the complete license text, see [GNU General Public License v3.0](https://www.gnu.org/licenses/gpl-3.0.en.html)

### How to Apply This License

1. Include a copy of the GPL-3.0 in your project
2. Add license headers to source files
3. Create a COPYING file with the full license text

### Contributions

By contributing to this project, you agree to license your contributions under the GPL-3.0.

````

To formally implement this:

1. Create a COPYING file in the project root with the full GPL-3.0 text
2. Add license headers to source files, like:

```typescript
/**
 * Excel Style Parser - A tool for parsing Excel file styles
 * Copyright (C) [Year] [Your Name]
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */
````
