// Excel AI Integration Task Pane
// Advanced Excel AI assistant with bidirectional data operations

interface ChatMessage {
    role: 'user' | 'assistant' | 'system';
    content: string;
    timestamp?: number;
}

// Global controller for canceling requests
let currentController: AbortController | null = null;

// Storage keys
const STORAGE_KEYS = {
    API_KEY: 'openrouter_api_key',
    CHAT_HISTORY: 'chat_history'
};

// Initialize messages with system prompt
let messages: ChatMessage[] = [
    {
        role: 'system',
        content: `You are Qwen 3 Coder, an advanced AI assistant specialized in Excel data operations and analysis. You have access to Excel data and can perform both read and write operations.

Key capabilities:
- Read Excel data from any workbook/sheet
- Write data to specific cells or ranges
- Analyze data patterns and provide insights
- Generate formulas and calculations
- Help with data visualization recommendations

Be concise and practical in your responses. When performing write operations, always use this exact JSON format:

{
  "operations": [
    {
      "type": "write_cell",
      "cell": "A1",
      "value": "Your Value"
    },
    {
      "type": "write_range",
      "range": "B1:D3",
      "values": [["Row1Col1", "Row1Col2", "Row1Col3"], ["Row2Col1", "Row2Col2", "Row2Col3"]]
    }
  ]
}

Only include write operations when the user explicitly requests data modifications.`
    }
];

// Helper function for DOM element selection
function $(selector: string): HTMLElement | null {
    return document.querySelector(selector);
}

// Parallax effect setup
function setupParallax(): void {
    const hero = $('.hero-section');
    if (!hero) return;
    
    window.addEventListener('scroll', () => {
        const scrolled = window.pageYOffset;
        const rate = scrolled * -0.5;
        hero.style.transform = `translateY(${rate}px)`;
    });
}

// Render chat messages
function renderChat(): void {
    const chatContainer = $('#chat-container');
    if (!chatContainer) return;
    
    chatContainer.innerHTML = '';
    
    // Only show user and assistant messages (skip system message)
    const visibleMessages = messages.filter(msg => msg.role !== 'system');
    
    visibleMessages.forEach(message => {
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${message.role}-message`;
        
        const content = document.createElement('div');
        content.className = 'message-content';
        content.textContent = message.content;
        
        const timestamp = document.createElement('div');
        timestamp.className = 'message-timestamp';
        timestamp.textContent = message.timestamp ? 
            new Date(message.timestamp).toLocaleTimeString() : 
            new Date().toLocaleTimeString();
        
        messageDiv.appendChild(content);
        messageDiv.appendChild(timestamp);
        chatContainer.appendChild(messageDiv);
    });
    
    // Scroll to bottom
    chatContainer.scrollTop = chatContainer.scrollTop + 1000;
}

// Load state from storage
async function loadState(): Promise<void> {
    try {
        // Load API key
        const apiKey = await OfficeRuntime.storage.getItem(STORAGE_KEYS.API_KEY);
        if (apiKey) {
            const apiKeyInput = $('#api-key') as HTMLInputElement;
            if (apiKeyInput) {
                apiKeyInput.value = apiKey;
            }
        }
        
        // Load chat history
        const chatHistory = await OfficeRuntime.storage.getItem(STORAGE_KEYS.CHAT_HISTORY);
        if (chatHistory) {
            const parsedHistory = JSON.parse(chatHistory);
            // Keep system message and add loaded history
            messages = [messages[0], ...parsedHistory];
            renderChat();
        }
    } catch (error) {
        console.error('Error loading state:', error);
    }
}

// Save API key
async function saveApiKey(): Promise<void> {
    const apiKeyInput = $('#api-key') as HTMLInputElement;
    if (!apiKeyInput) return;
    
    const apiKey = apiKeyInput.value.trim();
    if (!apiKey) {
        alert('Please enter an API key');
        return;
    }
    
    try {
        await OfficeRuntime.storage.setItem(STORAGE_KEYS.API_KEY, apiKey);
        alert('API key saved successfully!');
    } catch (error) {
        console.error('Error saving API key:', error);
        alert('Failed to save API key');
    }
}

// Set loading state
function setLoading(loading: boolean): void {
    const askButton = $('#ask-btn') as HTMLButtonElement;
    const cancelButton = $('#cancel-btn') as HTMLButtonElement;
    const loadingIndicator = $('#loading-indicator');
    
    if (askButton) askButton.disabled = loading;
    if (cancelButton) cancelButton.style.display = loading ? 'block' : 'none';
    if (loadingIndicator) loadingIndicator.style.display = loading ? 'block' : 'none';
}

// Call OpenRouter API
async function callOpenRouter(userMessage: string, model: string): Promise<string> {
    const apiKey = await OfficeRuntime.storage.getItem(STORAGE_KEYS.API_KEY);
    if (!apiKey) {
        throw new Error('Please set your OpenRouter API key first');
    }
    
    // Create new abort controller
    currentController = new AbortController();
    
    const response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${apiKey}`,
            'Content-Type': 'application/json',
            'HTTP-Referer': window.location.origin,
            'X-Title': 'Excel AI Integration'
        },
        body: JSON.stringify({
            model: model,
            messages: [...messages, { role: 'user', content: userMessage }],
            temperature: 0.7,
            max_tokens: 2000
        }),
        signal: currentController.signal
    });
    
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(`OpenRouter API error: ${response.status} - ${errorData.error?.message || response.statusText}`);
    }
    
    const data = await response.json();
    return data.choices[0]?.message?.content || 'No response received';
}

// Python health check functions
async function fetchPythonHealth(): Promise<{status: string, workbook?: string, error?: string}> {
    try {
        const response = await fetch('/health', {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json'
            }
        });
        
        const data = await response.json();
        return data;
    } catch (error) {
        return {
            status: 'unhealthy',
            error: `Connection failed: ${error instanceof Error ? error.message : 'Unknown error'}`
        };
    }
}

async function pythonHealthCheck(): Promise<void> {
    const healthData = await fetchPythonHealth();
    updateHealthPillDOM(healthData);
}

function updateHealthPillDOM(healthData: {status: string, workbook?: string, error?: string}): void {
    const healthPill = $('#python-health-pill');
    if (!healthPill) return;
    
    // Remove existing status classes
    healthPill.classList.remove('healthy', 'unhealthy');
    
    if (healthData.status === 'healthy') {
        healthPill.classList.add('healthy');
        healthPill.textContent = `Python: Healthy${healthData.workbook ? ` (${healthData.workbook})` : ''}`;
        healthPill.title = 'Python backend is running and Excel is accessible';
    } else {
        healthPill.classList.add('unhealthy');
        healthPill.textContent = 'Python: Unhealthy';
        healthPill.title = healthData.error || 'Python backend is not accessible';
    }
}

// Health polling management
let healthPollingInterval: number | null = null;
let isHealthPollingPaused = false;

function startHealthPolling(): void {
    if (healthPollingInterval) return; // Already running
    
    // Initial check
    pythonHealthCheck();
    
    // Set up interval
    healthPollingInterval = window.setInterval(() => {
        if (!isHealthPollingPaused) {
            pythonHealthCheck();
        }
    }, 5000); // Check every 5 seconds
}

function stopHealthPolling(): void {
    if (healthPollingInterval) {
        clearInterval(healthPollingInterval);
        healthPollingInterval = null;
    }
}

function pauseHealthPolling(): void {
    isHealthPollingPaused = true;
}

function resumeHealthPolling(): void {
    isHealthPollingPaused = false;
}

function pollHealthOnce(): Promise<void> {
    return pythonHealthCheck();
}

// Get workbook and sheet parameters
function getWorkbookSheetParams(): { workbook?: string, sheet?: string } {
    const workbookInput = $('#workbook-name') as HTMLInputElement;
    const sheetInput = $('#sheet-name') as HTMLInputElement;
    
    const params: { workbook?: string, sheet?: string } = {};
    
    if (workbookInput?.value.trim()) {
        params.workbook = workbookInput.value.trim();
    }
    
    if (sheetInput?.value.trim()) {
        params.sheet = sheetInput.value.trim();
    }
    
    return params;
}

// Call Python script for reading Excel data
async function callPythonScriptAdvanced(): Promise<any> {
    const params = getWorkbookSheetParams();
    const queryParams = new URLSearchParams();
    
    if (params.workbook) queryParams.append('workbook', params.workbook);
    if (params.sheet) queryParams.append('sheet', params.sheet);
    
    const url = `/api/excel-data${queryParams.toString() ? '?' + queryParams.toString() : ''}`;
    
    const response = await fetch(url, {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json'
        }
    });
    
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || `HTTP ${response.status}: ${response.statusText}`);
    }
    
    return await response.json();
}

// Extract cell references from text
function extractCellRefs(text: string): string[] {
    const cellRefPattern = /\b[A-Z]+\d+\b/g;
    return text.match(cellRefPattern) || [];
}

// Get Excel data for prompt
async function getExcelDataForPrompt(): Promise<string> {
    const maxRetries = 3;
    let lastError: Error | null = null;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            // Check Python health first
            const healthData = await fetchPythonHealth();
            if (healthData.status !== 'healthy') {
                throw new Error(`Python backend unhealthy: ${healthData.error || 'Unknown error'}`);
            }
            
            const data = await callPythonScriptAdvanced();
            
            let context = `Current Excel Data (${data.workbook}/${data.sheet}):\n`;
            context += `Shape: ${data.shape[0]} rows × ${data.shape[1]} columns\n\n`;
            
            if (data.data && data.data.length > 0) {
                // Show first few rows as sample
                const sampleSize = Math.min(5, data.data.length);
                context += `Sample data (first ${sampleSize} rows):\n`;
                
                const headers = Object.keys(data.data[0]);
                context += headers.join('\t') + '\n';
                
                for (let i = 0; i < sampleSize; i++) {
                    const row = data.data[i];
                    const values = headers.map(h => row[h] ?? '').join('\t');
                    context += values + '\n';
                }
                
                if (data.data.length > sampleSize) {
                    context += `... and ${data.data.length - sampleSize} more rows\n`;
                }
            }
            
            return context;
            
        } catch (error) {
            lastError = error as Error;
            console.warn(`Attempt ${attempt} failed:`, error);
            
            if (attempt < maxRetries) {
                // Wait before retry (exponential backoff)
                await new Promise(resolve => setTimeout(resolve, Math.pow(2, attempt) * 1000));
            }
        }
    }
    
    // All attempts failed
    throw lastError || new Error('Failed to get Excel data after multiple attempts');
}

// Get current selection address
async function getSelectionAddress(): Promise<string> {
    return new Promise((resolve) => {
        Excel.run(async (context) => {
            const selection = context.workbook.getSelectedRange();
            selection.load('address');
            await context.sync();
            resolve(selection.address);
        }).catch(() => {
            resolve('Unable to get selection');
        });
    });
}

// Detect if user query implies write operations
function detectWriteIntent(query: string): boolean {
    const writeKeywords = [
        'write', 'set', 'put', 'insert', 'add', 'update', 'change', 'modify',
        'fill', 'populate', 'create', 'generate', 'calculate and put',
        'save to', 'store in', 'place in'
    ];
    
    const lowerQuery = query.toLowerCase();
    return writeKeywords.some(keyword => lowerQuery.includes(keyword));
}

// Main ask function
async function onAsk(): Promise<void> {
    const userInput = $('#user-input') as HTMLTextAreaElement;
    const modelSelect = $('#model-select') as HTMLSelectElement;
    const allowWriteCheckbox = $('#allow-write') as HTMLInputElement;
    
    if (!userInput || !modelSelect) return;
    
    const userText = userInput.value.trim();
    if (!userText) return;
    
    const selectedModel = modelSelect.value;
    const allowWrite = allowWriteCheckbox?.checked || false;
    
    try {
        setLoading(true);
        
        // Add user message
        const userMessage: ChatMessage = {
            role: 'user',
            content: userText,
            timestamp: Date.now()
        };
        messages.push(userMessage);
        renderChat();
        
        // Clear input
        userInput.value = '';
        
        // Prepare prompt with Excel context
        let fullPrompt = userText;
        
        try {
            const excelContext = await getExcelDataForPrompt();
            const selection = await getSelectionAddress();
            
            fullPrompt = `${excelContext}\n\nCurrent Selection: ${selection}\n\nUser Query: ${userText}`;
            
            // Add write operation guidance if allowed and detected
            if (allowWrite && detectWriteIntent(userText)) {
                fullPrompt += `\n\nNote: User has enabled write operations. If they want to modify Excel data, provide the appropriate JSON operations format.`;
            } else if (detectWriteIntent(userText) && !allowWrite) {
                fullPrompt += `\n\nNote: User query suggests write operations, but write mode is disabled. Explain what could be done if write operations were enabled.`;
            }
            
        } catch (error) {
            console.warn('Could not get Excel context:', error);
            fullPrompt = `Excel data unavailable (${error instanceof Error ? error.message : 'Unknown error'}).\n\nUser Query: ${userText}`;
        }
        
        // Call OpenRouter API
        const response = await callOpenRouter(fullPrompt, selectedModel);
        
        // Add assistant message
        const assistantMessage: ChatMessage = {
            role: 'assistant',
            content: response,
            timestamp: Date.now()
        };
        messages.push(assistantMessage);
        renderChat();
        
        // Save chat history (excluding system message)
        const chatHistory = messages.filter(msg => msg.role !== 'system');
        await OfficeRuntime.storage.setItem(STORAGE_KEYS.CHAT_HISTORY, JSON.stringify(chatHistory));
        
        // Try to apply write operations if enabled
        if (allowWrite) {
            await maybeApplyWriteOperations(response);
        }
        
    } catch (error) {
        console.error('Error in onAsk:', error);
        
        if (error instanceof Error && error.name === 'AbortError') {
            // Request was cancelled
            const cancelMessage: ChatMessage = {
                role: 'assistant',
                content: 'Request cancelled by user.',
                timestamp: Date.now()
            };
            messages.push(cancelMessage);
            renderChat();
        } else {
            // Other error
            const errorMessage: ChatMessage = {
                role: 'assistant',
                content: `Error: ${error instanceof Error ? error.message : 'Unknown error occurred'}`,
                timestamp: Date.now()
            };
            messages.push(errorMessage);
            renderChat();
        }
    } finally {
        setLoading(false);
        currentController = null;
    }
}

// Restart Python bridge
async function restartPythonBridge(): Promise<void> {
    try {
        // Just do a health check to see current status
        await pythonHealthCheck();
        alert('Python bridge status refreshed. Check the health indicator.');
    } catch (error) {
        console.error('Error checking Python bridge:', error);
        alert('Failed to check Python bridge status.');
    }
}

// Event listeners
Office.onReady(() => {
    // Load saved state
    loadState();
    
    // Setup parallax effect
    setupParallax();
    
    // Start health polling
    startHealthPolling();
    
    // Ask AI button
    const askButton = $('#ask-btn');
    if (askButton) {
        askButton.addEventListener('click', onAsk);
    }
    
    // Enter key in textarea
    const userInput = $('#user-input') as HTMLTextAreaElement;
    if (userInput) {
        userInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                onAsk();
            }
        });
    }
    
    // Clear chat button
    const clearButton = $('#clear-chat-btn');
    if (clearButton) {
        clearButton.addEventListener('click', async () => {
            messages = [messages[0]]; // Keep only system message
            renderChat();
            await OfficeRuntime.storage.removeItem(STORAGE_KEYS.CHAT_HISTORY);
        });
    }
    
    // Save API key button
    const saveKeyButton = $('#save-key-btn');
    if (saveKeyButton) {
        saveKeyButton.addEventListener('click', saveApiKey);
    }
    
    // Cancel AI button
    const cancelButton = $('#cancel-btn');
    if (cancelButton) {
        cancelButton.addEventListener('click', () => {
            if (currentController) {
                currentController.abort();
            }
        });
    }
    
    // Restart Python bridge button
    const restartButton = $('#restart-python-btn');
    if (restartButton) {
        restartButton.addEventListener('click', restartPythonBridge);
    }
    
    // Health polling toggle
    const healthToggle = $('#health-polling-toggle') as HTMLInputElement;
    if (healthToggle) {
        healthToggle.checked = true; // Default to enabled
        healthToggle.addEventListener('change', () => {
            if (healthToggle.checked) {
                startHealthPolling();
            } else {
                stopHealthPolling();
            }
        });
    }
});

// Excel integration
function run() {
    Excel.run(async (context) => {
        // This function can be used for Excel-specific operations
        console.log('Excel context available');
    });
}

// Apply write operations from AI response
async function maybeApplyWriteOperations(response: string): Promise<void> {
    try {
        // Pause health polling during write operations to avoid contention
        pauseHealthPolling();
        
        // Look for JSON operations in the response
        const jsonMatch = response.match(/```json\s*({[\s\S]*?})\s*```/) || 
                         response.match(/({\s*"operations"[\s\S]*?})/); 
        
        if (!jsonMatch) {
            // Try to find JSON without code blocks
            const lines = response.split('\n');
            for (const line of lines) {
                if (line.trim().startsWith('{') && line.includes('operations')) {
                    try {
                        const operations = JSON.parse(line.trim());
                        if (operations.operations) {
                            await executeWriteOperations(operations);
                            return;
                        }
                    } catch (e) {
                        // Continue searching
                    }
                }
            }
            return; // No operations found
        }
        
        const operationsData = JSON.parse(jsonMatch[1]);
        if (operationsData.operations && Array.isArray(operationsData.operations)) {
            await executeWriteOperations(operationsData);
        }
        
    } catch (error) {
        console.error('Error applying write operations:', error);
        // Don't show error to user as this is automatic
    } finally {
        // Resume health polling after write operations
        setTimeout(() => {
            resumeHealthPolling();
        }, 2000); // Wait 2 seconds before resuming
    }
}

// Execute write operations
async function executeWriteOperations(operationsData: any): Promise<void> {
    const params = getWorkbookSheetParams();
    
    const requestBody = {
        ...operationsData,
        ...params // Add workbook/sheet targeting if specified
    };
    
    const response = await fetch('/api/write-excel', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
    });
    
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || `HTTP ${response.status}: ${response.statusText}`);
    }
    
    const result = await response.json();
    console.log('Write operations completed:', result);
    
    // Show success message
    const successCount = result.results?.filter((r: any) => r.success).length || 0;
    const errorCount = result.results?.filter((r: any) => r.error).length || 0;
    
    if (successCount > 0) {
        console.log(`✅ ${successCount} write operation(s) completed successfully`);
    }
    if (errorCount > 0) {
        console.warn(`❌ ${errorCount} write operation(s) failed`);
    }
}