const { spawn } = require('child_process');
const path = require('path');

// Start Python Flask server
const pythonProcess = spawn('python', ['app.py'], {
    cwd: __dirname,
    stdio: 'inherit',
    env: {
        ...process.env,
        PYTHONUNBUFFERED: '1'
    }
});

pythonProcess.on('error', (error) => {
    console.error('Failed to start Python server:', error);
    process.exit(1);
});

pythonProcess.on('close', (code) => {
    console.log(`Python server exited with code ${code}`);
    process.exit(code);
});

// Handle process termination
process.on('SIGINT', () => {
    console.log('Shutting down Python server...');
    pythonProcess.kill('SIGINT');
});

process.on('SIGTERM', () => {
    console.log('Shutting down Python server...');
    pythonProcess.kill('SIGTERM');
});