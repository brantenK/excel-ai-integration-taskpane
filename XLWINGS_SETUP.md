# Excel AI Task Pane with xlwings Integration

## Setup Instructions

### Prerequisites
1. **Python 3.8+** installed on your system
2. **Microsoft Excel** (Windows or Mac)
3. **Node.js 16+** for the frontend development server
4. **OpenRouter API key** for AI functionality

### Installation Steps

1. **Install Python Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Install Node.js Dependencies**
   ```bash
   npm install
   ```

3. **Start the Development Environment**
   
   **Option A: Using npm scripts (Recommended)**
   ```bash
   # Terminal 1: Start the frontend dev server
   npm run dev-server
   
   # Terminal 2: Start the Python backend
   npm run python-server
   ```
   
   **Option B: Manual startup**
   ```bash
   # Terminal 1: Frontend (HTTPS on port 3000)
   npm run dev-server
   
   # Terminal 2: Python backend (HTTP on port 3001)
   python app.py
   ```

4. **Load the Add-in in Excel**
   - Open Excel
   - Go to Insert > My Add-ins > Upload My Add-in
   - Select the `manifest.xml` file
   - The add-in will appear in the Home tab

## How It Works

### Enhanced Excel Integration
This application creates a sophisticated bridge between Excel and AI through multiple layers:

1. **Frontend (TypeScript/HTML)**: Modern task pane interface running on HTTPS
2. **Python Backend (Flask)**: Handles Excel operations via xlwings
3. **AI Integration**: OpenRouter API for multiple AI model access
4. **Proxy Layer**: Webpack dev server proxies API calls for seamless integration

### Architecture
```
Excel ↔ Task Pane (HTTPS:3000) ↔ Python Backend (HTTP:3001) ↔ xlwings ↔ Excel
                    ↕
              OpenRouter AI API
```

### Data Flow

#### Reading Data
1. User asks a question in the task pane
2. Frontend calls `/api/excel-data` (proxied to Python backend)
3. Python uses xlwings to read Excel data
4. Data is formatted and sent to AI model with user query
5. AI response is displayed in the chat interface

#### Writing Data
1. User enables "Allow Write Operations" and requests data modification
2. AI generates structured JSON operations
3. Frontend automatically detects and executes write operations
4. Python backend receives `/api/write-excel` request
5. xlwings performs the actual Excel modifications

## Features

### Excel Data Operations
- **Read Operations**: Fetch data from any workbook/sheet with optional targeting
- **Write Operations**: Modify cells and ranges with comprehensive error handling
- **Cell Mapping**: Automatic cell reference mapping for smaller datasets
- **Performance Optimization**: Smart payload management for large sheets

### API Enhancements
- **Unified Error Handling**: Consistent JSON error responses across all endpoints
- **Health Monitoring**: Real-time Python backend status with auto-polling
- **Robust Serialization**: Handles numpy, datetime, and decimal types
- **CORS Configuration**: Secure cross-origin setup for development

### User Experience
- **Real-time Health Monitoring**: Visual indicators for Python backend status
- **Workbook/Sheet Targeting**: Optional specific targeting or auto-detection
- **Write Operation Safety**: Explicit user consent required for data modifications
- **Chat History Persistence**: Conversations saved in Office storage
- **Model Selection**: Multiple AI models available through OpenRouter

## Troubleshooting

### "Python: Unhealthy" Status
1. **Check if Python backend is running**:
   ```bash
   python app.py
   ```
   Should show: `Running on http://127.0.0.1:3001`

2. **Verify xlwings installation**:
   ```bash
   python -c "import xlwings; print('xlwings OK')"
   ```

3. **Test Excel connection**:
   - Open Excel with a workbook containing data
   - Check if the health status improves

4. **Check port conflicts**:
   - Ensure port 3001 is not used by other applications
   - Try restarting both frontend and backend servers

### Common Issues
- **CORS Errors**: Make sure both servers are running and proxy is configured
- **API Key Issues**: Verify your OpenRouter API key is saved correctly
- **Excel Not Detected**: Ensure Excel is open with an active workbook
- **Write Operations Failing**: Check that "Allow Write Operations" is enabled

## File Structure
```
├── app.py                 # Flask backend with xlwings integration
├── config.py             # Configuration settings
├── requirements.txt      # Python dependencies
├── package.json          # Node.js dependencies and scripts
├── manifest.xml          # Office Add-in manifest
├── webpack.config.js     # Build configuration with proxy setup
├── python-server.js      # Node.js script to run Python backend
├── src/
│   ├── taskpane/
│   │   ├── taskpane.ts   # Main TypeScript logic
│   │   └── taskpane.html # Task pane interface
│   └── commands/
│       ├── commands.ts   # Office commands
│       └── commands.html # Commands page
└── test_app.py          # Comprehensive test suite
```

## Testing

Run the comprehensive test suite:
```bash
pytest test_app.py -v
```

Tests cover:
- Health endpoint functionality
- Excel data reading with various scenarios
- Write operations with error handling
- JSON serialization of complex data types
- Workbook and sheet targeting
- Cell mapping behavior

## Usage Tips

1. **Start with Health Check**: Always ensure the Python backend is healthy before using AI features
2. **Use Specific Targeting**: For complex workbooks, specify exact workbook/sheet names
3. **Enable Write Carefully**: Only enable write operations when you want to modify data
4. **Monitor Performance**: Large datasets automatically disable cell mapping for better performance
5. **Save API Key**: Your OpenRouter API key is securely stored in Office storage

## Security Notes

- API keys are stored securely using Office's storage API
- CORS is configured to only allow requests from the task pane origin
- Write operations require explicit user consent
- All data processing happens locally through xlwings
- No Excel data is sent to external services except as part of AI queries

## Advanced Configuration

### Custom Model Selection
Modify the model dropdown in `taskpane.html` to add or remove AI models available through OpenRouter.

### Proxy Configuration
The webpack configuration includes proxy settings that route `/api/*` and `/health` requests to the Python backend. This can be customized in `webpack.config.js`.

### Health Polling
Health checking occurs every 5 seconds by default. This can be adjusted in the `startHealthPolling()` function in `taskpane.ts`.