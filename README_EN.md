# BfTaskBoard

A powerful Windows desktop task management and data table application with support for multiple data types, AI-powered table generation, multilingual interface, and theme switching.

## Key Features

### Core Features
- **Multi-tab Management**: Create, edit, and sort multiple data table tabs
- **Multiple Data Types**:
  - Text: Plain text content
  - Single Select: Options with colored markers
  - Image: Clipboard image support (Ctrl+V)
  - TodoList: Checkable task lists
  - TextArea: Long text editing with auto-save
- **AI Table Generation**: Automatic table structure generation using OpenAI, DeepSeek, or Claude
- **Data Import/Export**: Excel file export (.xlsx format)

### Advanced Features
- **Global Search**: Cross-tab content search with location navigation (Ctrl+S)
- **Data Processing**:
  - **Sorting**: Text columns support ascending/descending sort via column header icon
  - **Filtering**: Single select columns support option-based filtering, preserves original data
  - **Column Sum**: Automatically calculate sum for numeric columns via right-click menu
  - **One-Click Split**: Quickly split long text by delimiter into multiple rows
  - **JSON Editor**: Built-in JSON viewer and editor with formatting and validation
- **Batch Operations**:
  - Batch delete selected rows
  - Batch copy and paste
  - Column data batch fill
- **Interface Customization**:
  - Bilingual support (Chinese/English) with automatic system language detection
  - Theme switching (dark/light) with real-time preview
  - Colored tab markers with 6 preset colors
  - Auto-adjust column width
- **Keyboard Shortcuts**:
  - Ctrl+T: New tab
  - Ctrl+Q: Tab management
  - Ctrl+S: Global search
  - Ctrl+V: Paste image (in image columns)
  - Delete: Delete selected rows

### Data Management
- **Auto-save**: Real-time data persistence
- **Performance Optimization**: Asynchronous processing for large datasets
- **Data Persistence**: Local JSON storage
- **File Monitoring**: Display last modification time

## Usage

### Basic Operations
1. **Create Tab**: Click the "+" button or use Ctrl+T
2. **Edit Cell**: Double-click to enter edit mode
3. **Add Row**: Click "Add Row" button at the bottom
4. **Manage Columns**: Right-click column header for edit/delete options
5. **AI Table Generation**: Click robot icon, enter requirements for automatic generation

### Advanced Operations
1. **Sort Data**:
   - Hover over text column header, click sort icon
   - Toggle between ascending and descending
   - Sorting doesn't modify original data
   
2. **Filter Data**:
   - Hover over single select column header, click filter icon
   - Select options to display
   - Supports multi-select filtering
   
3. **Column Sum**:
   - Right-click numeric column
   - Select "Column Sum"
   - Automatically calculates and displays total
   
4. **One-Click Split**:
   - Select cell containing delimited text
   - Right-click and select "One-Click Split"
   - Enter delimiter (comma, semicolon, etc.)
   - Automatically splits text into multiple rows
   
5. **JSON Editing**:
   - Double-click cell containing JSON
   - Automatically opens JSON editor
   - Supports formatting, validation, and editing

### Data Type Usage
- **Text**: Direct text input
- **Single Select**: Click to show options list with custom options and colors
- **Image**: Click to upload or paste directly (Ctrl+V)
- **TodoList**: Click to add tasks, check to complete
- **TextArea**: Click to open editor for long text

## Technical Details

### Tech Stack
- **Framework**: .NET 6.0 + Windows Forms
- **Language**: C# 10.0
- **Data Storage**: JSON (Newtonsoft.Json)
- **Excel Processing**: EPPlus
- **HTTP Client**: HttpClient (for AI API calls)

### Architecture
- **MVC Pattern**: Separation of view and data
- **Service Layer**:
  - `DataService`: Data persistence management
  - `LanguageService`: Multilingual support
  - `ThemeService`: Theme management
  - `AIService`: AI interface integration
- **Custom Controls**:
  - `ModernDataGridView`: Enhanced data grid
  - `TabButton`: Custom tab buttons
  - `TodoListControl`: Todo list control

### Performance Optimization
- Asynchronous file saving
- Debounce handling (textarea auto-save)
- Large dataset pagination suggestions
- JSON compression storage

## Build Instructions

### Requirements
- Windows 10/11
- .NET 6.0 SDK or higher
- Visual Studio 2022 (recommended) or VS Code

### Build Steps
1. Clone repository
```bash
git clone https://github.com/yourusername/BfTaskBoard.git
cd BfTaskBoard
```

2. Build project
```bash
build.bat
```
Or using dotnet command:
```bash
dotnet build -c Release
```

3. Run application
```bash
bin\Release\net6.0-windows\BfTaskBoard.exe
```

### Single File Publishing
```bash
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

## Configuration

### AI Service Configuration
The application supports three AI providers:
- **OpenAI**: Requires OpenAI API key
- **DeepSeek**: Requires DeepSeek API key
- **Claude**: Requires Anthropic API key

API keys are entered when using the AI table generation feature and are automatically saved.

### Data Storage Locations
- Application data: `%APPDATA%\BfTaskBoard\appdata.json`
- Theme settings: `%APPDATA%\BfTaskBoard\theme.json`
- AI logs: `%APPDATA%\BfTaskBoard\ai_logs\`

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Issues and Pull Requests are welcome!

### Development Guidelines
1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Changelog

### v1.0.0 (2024-01)
- Initial release
- Multiple data type support
- AI table generation
- Multilingual and theme support
- Excel export functionality

## Contact

- Project Homepage: [https://github.com/yourusername/BfTaskBoard](https://github.com/yourusername/BfTaskBoard)
- Bug Reports: [Issues](https://github.com/yourusername/BfTaskBoard/issues)

## Acknowledgments

Thanks to all contributors who have helped with this project!