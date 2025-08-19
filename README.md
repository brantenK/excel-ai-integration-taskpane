# Excel AI Task Pane

Advanced Excel AI Integration with bidirectional data operations, xlwings backend, and comprehensive targeting support.

## Recent Updates

- **Enhanced API Architecture**: Unified endpoints with HTTPS proxy support
- **Performance Optimizations**: Smart cell mapping with size guards and robust JSON serialization
- **Workbook/Sheet Targeting**: Precise targeting for both read and write operations
- **Improved Security**: HTTPS development server with proper CORS configuration
- **Robust Data Handling**: Support for numpy, datetime, and decimal types
- **Comprehensive Testing**: Full pytest coverage with xlwings mocking
- **Clean UI**: Streamlined chat interface without Excel data context clutter

## Prerequisites

- **Node.js** (v14 or later)
- **Python** (3.8 or later)
- **Microsoft Excel** (Office 365 or Excel 2016+)
- **xlwings** Python package
- **OpenRouter API Key** for AI functionality

## Installation

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd excel-ai-integration-taskpane
   ```

2. **Install Node.js dependencies**:
   ```bash
   npm install
   ```

3. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure xlwings** (see [XLWINGS_SETUP.md](XLWINGS_SETUP.md) for detailed instructions):
   ```bash
   xlwings addin install
   ```

## How to Start the App

1. **Start the development server** (HTTPS enabled):
   ```bash
   npm run dev-server
   ```
   This starts the frontend at `https://localhost:3000`

2. **Start the Python backend**:
   ```bash
   python app.py
   ```
   This starts the Flask server at `http://localhost:3001`

3. **Start the HTTPS proxy** (in a separate terminal):
   ```bash
   npm run python-server
   ```
   This proxies Python backend through HTTPS at `https://localhost:3000/api/*` and `https://localhost:3000/health`

4. **Sideload the add-in** in Excel:
   - Open Excel
   - Go to Insert > My Add-ins > Upload My Add-in
   - Select the `manifest.xml` file
   - The task pane will appear on the right side

## How to Read and Write Excel Data

### Reading Excel Data

The application can read Excel data in multiple ways:

- **Current Selection**: Automatically reads the currently selected range
- **Entire Used Range**: Reads all data in the active worksheet
- **Specific Cells**: Target individual cells (e.g., "A1", "B5")
- **Specific Ranges**: Target cell ranges (e.g., "A1:C10")
- **Workbook/Sheet Targeting**: Read from specific workbooks and sheets

#### Performance Features:
- **Smart Cell Mapping**: Automatically disabled for large datasets (>5000 cells) to prevent memory issues
- **Configurable Mapping**: Use `include_cell_mapping` parameter to control cell-by-cell mapping
- **Robust Serialization**: Handles numpy arrays, datetime objects, and decimal types seamlessly

### Writing Excel Data

The AI can write data back to Excel using structured JSON operations:

#### Enhanced Targeting for Write Operations:
- **Workbook Targeting**: Specify `workbook` parameter to target specific Excel files
- **Sheet Targeting**: Specify `sheet` parameter to target specific worksheets
- **Flexible Operations**: Support for cells, ranges, and formulas

#### Supported Write Operations:
```json
{
  "write_operations": [
    {
      "type": "write_cell",
      "cell": "A1",
      "value": "Hello World"
    },
    {
      "type": "write_range",
      "range": "B1:D3",
      "values": [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
    },
    {
      "type": "insert_formula",
      "cell": "E1",
      "formula": "=SUM(A1:D1)"
    }
  ]
}
```

## Model Selection

The application supports multiple AI models through OpenRouter:
- **Qwen 3 Coder** (default) - Optimized for code and data operations
- **Claude 3.5 Sonnet** - Advanced reasoning and analysis
- **GPT-4** - General-purpose AI assistance
- **Custom Models** - Add your preferred OpenRouter models

## Troubleshooting

### "Unhealthy" Python Health

If the Python health indicator shows "Unhealthy":

1. **Check Python Server**: Ensure `python app.py` is running
2. **Check HTTPS Proxy**: Ensure `npm run python-server` is running
3. **Verify Excel Connection**: Make sure Excel is open with an active workbook
4. **Restart Services**: Use the "Restart Python" button in the task pane
5. **Check Ports**: Ensure ports 3000 and 3001 are available

### Common Issues:
- **HTTPS Certificate Warnings**: Accept the self-signed certificate in your browser
- **xlwings Connection**: Ensure xlwings add-in is properly installed
- **API Key**: Verify your OpenRouter API key is correctly entered

## Summary of Data Transfer to Excel

The application uses a sophisticated data pipeline:

1. **Frontend Request**: User interacts with the task pane
2. **HTTPS Proxy**: Requests are proxied through secure HTTPS
3. **Python Backend**: Flask server processes requests using xlwings
4. **Excel Integration**: Direct bidirectional communication with Excel
5. **AI Processing**: OpenRouter API processes natural language requests
6. **Structured Output**: AI returns JSON operations for Excel manipulation

## Testing

Run the comprehensive test suite:

```bash
# Run all tests
pytest

# Run with verbose output
pytest -v

# Run specific test file
pytest test_app.py

# Run with coverage
pytest --cov=app
```

The test suite includes:
- **Health Check Tests**: Verify Excel connectivity
- **Data Reading Tests**: Test various read scenarios
- **Write Operation Tests**: Test all write operation types
- **Error Handling Tests**: Test edge cases and error conditions
- **Serialization Tests**: Test numpy, datetime, and decimal handling
- **Targeting Tests**: Test workbook and sheet targeting

## API Reference

### Health Check
- **Endpoint**: `GET /health`
- **Purpose**: Verify Excel connectivity and system status
- **Response**: System health information including active workbook details

### Read Excel Data
- **Endpoint**: `POST /api/excel-data`
- **Purpose**: Read data from Excel with flexible targeting
- **Parameters**:
  - `workbook` (optional): Target specific workbook
  - `sheet` (optional): Target specific worksheet
  - `specific_cells` (optional): Array of cell addresses
  - `specific_ranges` (optional): Array of range addresses
  - `include_cell_mapping` (optional): Enable/disable cell mapping
  - `max_cells_for_mapping` (optional): Threshold for cell mapping
  - `force_recalc` (optional): Force Excel recalculation

### Write Excel Operations
- **Endpoint**: `POST /api/write-excel`
- **Purpose**: Execute write operations in Excel
- **Parameters**:
  - `workbook` (optional): Target specific workbook
  - `sheet` (optional): Target specific worksheet
  - `operations` (required): Array of write operations
- **Response**: Results array with operation outcomes

---

*This Excel AI Task Pane provides a powerful, production-ready solution for AI-driven Excel automation with enterprise-grade features and comprehensive error handling.*