# Excel File Comparer

A modern web application built with Next.js that allows you to upload and compare two Excel files (.xlsx and .xls) to identify differences across all sheets.

## Features

- **Drag & Drop Upload**: Easy file upload with drag and drop functionality
- **Multi-Sheet Comparison**: Compares all sheets in both Excel files
- **Visual Difference Display**: Shows differences with color-coded indicators:
  - 🟢 **Green**: Added content (exists only in second file)
  - 🔴 **Red**: Removed content (exists only in first file)
  - 🟡 **Yellow**: Modified content (different values in same cell)
- **Tabbed Interface**: Navigate between different sheets easily
- **Cell Reference**: Shows Excel cell references (e.g., A1, B2, etc.)
- **Summary Statistics**: Overview of total differences and file information

## Getting Started

### Prerequisites

- Node.js 18+ 
- npm or yarn

### Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd xlsx_comparer
```

2. Install dependencies:
```bash
npm install
```

3. Run the development server:
```bash
npm run dev
```

4. Open [http://localhost:3000](http://localhost:3000) in your browser.

## How to Use

1. **Upload Files**: Drag and drop or click to select two Excel files (.xlsx or .xls)
2. **Compare**: Click the "Compare Files" button once both files are uploaded
3. **Review Results**: 
   - View the summary showing total differences
   - Navigate between sheets using the tabs
   - Examine individual cell differences in the detailed table

## Supported File Formats

- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)

## Technology Stack

- **Framework**: Next.js 15 with App Router
- **Language**: TypeScript
- **Styling**: Tailwind CSS
- **Excel Processing**: SheetJS (xlsx)
- **File Upload**: react-dropzone

## Development

### Project Structure

```
src/
├── app/
│   ├── page.tsx          # Main application page
│   ├── layout.tsx        # Root layout
│   └── globals.css       # Global styles
└── components/
    └── ExcelComparer.tsx # Excel comparison logic and UI
```

### Available Scripts

- `npm run dev` - Start development server
- `npm run build` - Build for production
- `npm run start` - Start production server
- `npm run lint` - Run ESLint

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is open source and available under the [MIT License](LICENSE).
