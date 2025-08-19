import pytest
import json
from unittest.mock import Mock, patch, MagicMock
import pandas as pd
import numpy as np
from datetime import datetime
from decimal import Decimal

# Import the Flask app
from app import app, RobustJSONEncoder

@pytest.fixture
def client():
    """Create a test client for the Flask app."""
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client

@pytest.fixture
def mock_excel_app():
    """Mock Excel application."""
    mock_app = Mock()
    mock_app.books = Mock()
    return mock_app

@pytest.fixture
def mock_workbook():
    """Mock Excel workbook."""
    mock_wb = Mock()
    mock_wb.name = "TestWorkbook.xlsx"
    mock_wb.sheets = Mock()
    return mock_wb

@pytest.fixture
def mock_worksheet():
    """Mock Excel worksheet."""
    mock_ws = Mock()
    mock_ws.name = "Sheet1"
    mock_ws.used_range = Mock()
    mock_ws.range = Mock()
    return mock_ws

@pytest.fixture
def mock_used_range():
    """Mock Excel used range."""
    mock_range = Mock()
    mock_range.row = 1
    mock_range.column = 1
    return mock_range

@pytest.fixture
def sample_dataframe():
    """Sample DataFrame for testing."""
    return pd.DataFrame({
        'Name': ['Alice', 'Bob', 'Charlie'],
        'Age': [25, 30, 35],
        'Score': [95.5, 87.2, 92.8]
    })

class TestHealthEndpoint:
    """Test cases for the /health endpoint."""
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    def test_health_check_success(self, mock_get_workbook, mock_get_app, client, mock_excel_app, mock_workbook):
        """Test successful health check."""
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = mock_workbook
        
        response = client.get('/health')
        
        assert response.status_code == 200
        data = json.loads(response.data)
        assert data['status'] == 'healthy'
        assert data['workbook'] == 'TestWorkbook.xlsx'
    
    @patch('app.get_excel_app')
    def test_health_check_no_excel_app(self, mock_get_app, client):
        """Test health check when no Excel application is running."""
        mock_get_app.return_value = None
        
        response = client.get('/health')
        
        assert response.status_code == 503
        data = json.loads(response.data)
        assert data['status'] == 'unhealthy'
        assert 'No Excel application running' in data['error']
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    def test_health_check_no_workbook(self, mock_get_workbook, mock_get_app, client, mock_excel_app):
        """Test health check when no active workbook is found."""
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = None
        
        response = client.get('/health')
        
        assert response.status_code == 404
        data = json.loads(response.data)
        assert data['status'] == 'unhealthy'
        assert 'No active workbook found' in data['error']

class TestExcelDataEndpoint:
    """Test cases for the /api/excel-data endpoint."""
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    @patch('app.get_worksheet')
    def test_get_excel_data_success(self, mock_get_worksheet, mock_get_workbook, mock_get_app, 
                                   client, mock_excel_app, mock_workbook, mock_worksheet, 
                                   mock_used_range, sample_dataframe):
        """Test successful data retrieval."""
        # Setup mocks
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = mock_workbook
        mock_get_worksheet.return_value = mock_worksheet
        mock_worksheet.used_range = mock_used_range
        mock_used_range.options.return_value.value = sample_dataframe
        
        response = client.get('/api/excel-data')
        
        assert response.status_code == 200
        data = json.loads(response.data)
        assert data['workbook'] == 'TestWorkbook.xlsx'
        assert data['sheet'] == 'Sheet1'
        assert len(data['data']) == 3
        assert data['shape'] == [3, 3]
    
    @patch('app.get_excel_app')
    def test_get_excel_data_no_app(self, mock_get_app, client):
        """Test data retrieval when no Excel application is running."""
        mock_get_app.return_value = None
        
        response = client.get('/api/excel-data')
        
        assert response.status_code == 503
        data = json.loads(response.data)
        assert 'No Excel application running' in data['error']
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    @patch('app.get_worksheet')
    def test_get_excel_data_no_data(self, mock_get_worksheet, mock_get_workbook, mock_get_app, 
                                   client, mock_excel_app, mock_workbook, mock_worksheet):
        """Test data retrieval when no data is found."""
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = mock_workbook
        mock_get_worksheet.return_value = mock_worksheet
        mock_worksheet.used_range = None
        
        response = client.get('/api/excel-data')
        
        assert response.status_code == 404
        data = json.loads(response.data)
        assert 'No data found in worksheet' in data['error']

class TestWriteExcelEndpoint:
    """Test cases for the /api/write-excel endpoint."""
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    @patch('app.get_worksheet')
    def test_write_excel_success(self, mock_get_worksheet, mock_get_workbook, mock_get_app, 
                                client, mock_excel_app, mock_workbook, mock_worksheet):
        """Test successful write operation."""
        # Setup mocks
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = mock_workbook
        mock_get_worksheet.return_value = mock_worksheet
        mock_range = Mock()
        mock_worksheet.range.return_value = mock_range
        
        # Test data
        test_data = {
            'operations': [
                {'type': 'write_cell', 'cell': 'A1', 'value': 'Test Value'}
            ]
        }
        
        response = client.post('/api/write-excel', 
                              data=json.dumps(test_data),
                              content_type='application/json')
        
        assert response.status_code == 200
        data = json.loads(response.data)
        assert 'results' in data
        assert len(data['results']) == 1
        assert 'success' in data['results'][0]
    
    @patch('app.get_excel_app')
    def test_write_excel_no_app(self, mock_get_app, client):
        """Test write operation when no Excel application is running."""
        mock_get_app.return_value = None
        
        test_data = {
            'operations': [
                {'type': 'write_cell', 'cell': 'A1', 'value': 'Test Value'}
            ]
        }
        
        response = client.post('/api/write-excel', 
                              data=json.dumps(test_data),
                              content_type='application/json')
        
        assert response.status_code == 503
        data = json.loads(response.data)
        assert 'No Excel application running' in data['error']
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    @patch('app.get_worksheet')
    def test_write_excel_unknown_operation(self, mock_get_worksheet, mock_get_workbook, mock_get_app, 
                                          client, mock_excel_app, mock_workbook, mock_worksheet):
        """Test write operation with unknown operation type."""
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = mock_workbook
        mock_get_worksheet.return_value = mock_worksheet
        
        test_data = {
            'operations': [
                {'type': 'unknown_operation', 'cell': 'A1', 'value': 'Test Value'}
            ]
        }
        
        response = client.post('/api/write-excel', 
                              data=json.dumps(test_data),
                              content_type='application/json')
        
        assert response.status_code == 200
        data = json.loads(response.data)
        assert 'results' in data
        assert 'error' in data['results'][0]
        assert 'Unknown operation type' in data['results'][0]['error']

class TestRobustJSONEncoder:
    """Test cases for the RobustJSONEncoder."""
    
    def test_numpy_types(self):
        """Test encoding of numpy types."""
        encoder = RobustJSONEncoder()
        
        # Test numpy integer
        np_int = np.int64(42)
        assert encoder.default(np_int) == 42
        
        # Test numpy float
        np_float = np.float64(3.14)
        assert encoder.default(np_float) == 3.14
        
        # Test numpy array
        np_array = np.array([1, 2, 3])
        assert encoder.default(np_array) == [1, 2, 3]
    
    def test_datetime_types(self):
        """Test encoding of datetime types."""
        encoder = RobustJSONEncoder()
        
        # Test datetime
        dt = datetime(2023, 1, 1, 12, 0, 0)
        result = encoder.default(dt)
        assert result == '2023-01-01T12:00:00'
        
        # Test pandas timestamp
        pd_ts = pd.Timestamp('2023-01-01 12:00:00')
        result = encoder.default(pd_ts)
        assert '2023-01-01T12:00:00' in result
    
    def test_decimal_type(self):
        """Test encoding of Decimal type."""
        encoder = RobustJSONEncoder()
        
        decimal_val = Decimal('3.14159')
        assert encoder.default(decimal_val) == 3.14159

class TestCellMapping:
    """Test cases for cell mapping functionality."""
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    @patch('app.get_worksheet')
    @patch('app.xw.utils.int_to_col_letter')
    def test_cell_mapping_included_small_dataset(self, mock_col_letter, mock_get_worksheet, 
                                                mock_get_workbook, mock_get_app, client, 
                                                mock_excel_app, mock_workbook, mock_worksheet, 
                                                mock_used_range):
        """Test that cell mapping is included for small datasets."""
        # Setup mocks
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = mock_workbook
        mock_get_worksheet.return_value = mock_worksheet
        mock_worksheet.used_range = mock_used_range
        mock_col_letter.side_effect = lambda x: chr(64 + x)  # A, B, C, etc.
        
        # Small dataset (3 rows)
        small_df = pd.DataFrame({
            'Name': ['Alice', 'Bob', 'Charlie'],
            'Age': [25, 30, 35]
        })
        mock_used_range.options.return_value.value = small_df
        
        response = client.get('/api/excel-data?include_cell_mapping=true')
        
        assert response.status_code == 200
        data = json.loads(response.data)
        assert 'cell_mapping' in data
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    @patch('app.get_worksheet')
    def test_cell_mapping_excluded_large_dataset(self, mock_get_worksheet, mock_get_workbook, 
                                                 mock_get_app, client, mock_excel_app, 
                                                 mock_workbook, mock_worksheet, mock_used_range):
        """Test that cell mapping is excluded for large datasets."""
        # Setup mocks
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = mock_workbook
        mock_get_worksheet.return_value = mock_worksheet
        mock_worksheet.used_range = mock_used_range
        
        # Large dataset (150 rows)
        large_data = {'Name': [f'Person{i}' for i in range(150)], 
                     'Age': list(range(150))}
        large_df = pd.DataFrame(large_data)
        mock_used_range.options.return_value.value = large_df
        
        response = client.get('/api/excel-data?include_cell_mapping=true')
        
        assert response.status_code == 200
        data = json.loads(response.data)
        assert 'cell_mapping' not in data
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    @patch('app.get_worksheet')
    def test_cell_mapping_disabled(self, mock_get_worksheet, mock_get_workbook, mock_get_app, 
                                   client, mock_excel_app, mock_workbook, mock_worksheet, 
                                   mock_used_range, sample_dataframe):
        """Test that cell mapping can be explicitly disabled."""
        # Setup mocks
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = mock_workbook
        mock_get_worksheet.return_value = mock_worksheet
        mock_worksheet.used_range = mock_used_range
        mock_used_range.options.return_value.value = sample_dataframe
        
        response = client.get('/api/excel-data?include_cell_mapping=false')
        
        assert response.status_code == 200
        data = json.loads(response.data)
        assert 'cell_mapping' not in data

class TestWorkbookSheetTargeting:
    """Test cases for workbook and sheet targeting functionality."""
    
    @patch('app.get_excel_app')
    def test_specific_workbook_targeting(self, mock_get_app, client, mock_excel_app):
        """Test targeting a specific workbook."""
        # Setup mocks
        mock_get_app.return_value = mock_excel_app
        mock_workbook = Mock()
        mock_workbook.name = "SpecificWorkbook.xlsx"
        mock_excel_app.books = {'SpecificWorkbook.xlsx': mock_workbook}
        
        # Mock worksheet and data
        mock_worksheet = Mock()
        mock_worksheet.name = "Sheet1"
        mock_worksheet.used_range = None  # Will trigger 404
        mock_workbook.sheets = Mock()
        mock_workbook.sheets.active = mock_worksheet
        
        response = client.get('/api/excel-data?workbook=SpecificWorkbook.xlsx')
        
        # Should get 404 because no data, but workbook was found
        assert response.status_code == 404
        data = json.loads(response.data)
        assert 'No data found in worksheet' in data['error']
    
    @patch('app.get_excel_app')
    def test_nonexistent_workbook(self, mock_get_app, client, mock_excel_app):
        """Test targeting a non-existent workbook."""
        mock_get_app.return_value = mock_excel_app
        mock_excel_app.books = {}
        mock_excel_app.books.__getitem__ = Mock(side_effect=KeyError())
        
        response = client.get('/api/excel-data?workbook=NonExistent.xlsx')
        
        assert response.status_code == 404
        data = json.loads(response.data)
        assert "Workbook 'NonExistent.xlsx' not found" in data['error']
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    def test_specific_sheet_targeting(self, mock_get_workbook, mock_get_app, client, 
                                     mock_excel_app, mock_workbook):
        """Test targeting a specific sheet."""
        # Setup mocks
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = mock_workbook
        
        mock_worksheet = Mock()
        mock_worksheet.name = "SpecificSheet"
        mock_worksheet.used_range = None  # Will trigger 404
        mock_workbook.sheets = {'SpecificSheet': mock_worksheet}
        
        response = client.get('/api/excel-data?sheet=SpecificSheet')
        
        # Should get 404 because no data, but sheet was found
        assert response.status_code == 404
        data = json.loads(response.data)
        assert 'No data found in worksheet' in data['error']
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    def test_nonexistent_sheet(self, mock_get_workbook, mock_get_app, client, 
                              mock_excel_app, mock_workbook):
        """Test targeting a non-existent sheet."""
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = mock_workbook
        mock_workbook.sheets = {}
        mock_workbook.sheets.__getitem__ = Mock(side_effect=KeyError())
        
        response = client.get('/api/excel-data?sheet=NonExistentSheet')
        
        assert response.status_code == 404
        data = json.loads(response.data)
        assert "Sheet 'NonExistentSheet' not found" in data['error']

class TestWriteOperationsTargeting:
    """Test cases for write operations with workbook/sheet targeting."""
    
    @patch('app.get_excel_app')
    def test_write_with_specific_workbook_and_sheet(self, mock_get_app, client, mock_excel_app):
        """Test write operations with specific workbook and sheet targeting."""
        # Setup mocks
        mock_get_app.return_value = mock_excel_app
        
        mock_workbook = Mock()
        mock_workbook.name = "TargetWorkbook.xlsx"
        mock_excel_app.books = {'TargetWorkbook.xlsx': mock_workbook}
        
        mock_worksheet = Mock()
        mock_worksheet.name = "TargetSheet"
        mock_workbook.sheets = {'TargetSheet': mock_worksheet}
        
        mock_range = Mock()
        mock_worksheet.range.return_value = mock_range
        
        test_data = {
            'workbook': 'TargetWorkbook.xlsx',
            'sheet': 'TargetSheet',
            'operations': [
                {'type': 'write_cell', 'cell': 'A1', 'value': 'Test Value'}
            ]
        }
        
        response = client.post('/api/write-excel', 
                              data=json.dumps(test_data),
                              content_type='application/json')
        
        assert response.status_code == 200
        data = json.loads(response.data)
        assert 'results' in data
        assert 'success' in data['results'][0]
    
    @patch('app.get_excel_app')
    def test_write_with_nonexistent_workbook(self, mock_get_app, client, mock_excel_app):
        """Test write operations with non-existent workbook."""
        mock_get_app.return_value = mock_excel_app
        mock_excel_app.books = {}
        mock_excel_app.books.__getitem__ = Mock(side_effect=KeyError())
        
        test_data = {
            'workbook': 'NonExistent.xlsx',
            'operations': [
                {'type': 'write_cell', 'cell': 'A1', 'value': 'Test Value'}
            ]
        }
        
        response = client.post('/api/write-excel', 
                              data=json.dumps(test_data),
                              content_type='application/json')
        
        assert response.status_code == 404
        data = json.loads(response.data)
        assert "Workbook 'NonExistent.xlsx' not found" in data['error']
    
    @patch('app.get_excel_app')
    @patch('app.get_active_workbook')
    def test_write_with_nonexistent_sheet(self, mock_get_workbook, mock_get_app, client, 
                                         mock_excel_app, mock_workbook):
        """Test write operations with non-existent sheet."""
        mock_get_app.return_value = mock_excel_app
        mock_get_workbook.return_value = mock_workbook
        mock_workbook.sheets = {}
        mock_workbook.sheets.__getitem__ = Mock(side_effect=KeyError())
        
        test_data = {
            'sheet': 'NonExistentSheet',
            'operations': [
                {'type': 'write_cell', 'cell': 'A1', 'value': 'Test Value'}
            ]
        }
        
        response = client.post('/api/write-excel', 
                              data=json.dumps(test_data),
                              content_type='application/json')
        
        assert response.status_code == 404
        data = json.loads(response.data)
        assert "Sheet 'NonExistentSheet' not found" in data['error']