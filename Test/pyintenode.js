const { spawn } = require('child_process');

// Run Python script
const pythonProcess = spawn('python', ['TemplateResultReader.py','null']);

// Print output from Python script
pythonProcess.stdout.on('data', (data) => {
    console.log(`Python script output: ${data}`);
});

// Handle Python script errors
pythonProcess.stderr.on('data', (data) => {
    console.error(`Python script error: ${data}`);
});

console.log("done");