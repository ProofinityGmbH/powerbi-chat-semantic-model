const { app, BrowserWindow, ipcMain, session } = require('electron');
const path = require('path');
const fs = require('fs');

// Add native library paths to PATH before requiring electron-edge-js
// This ensures all native DLLs (VC++ runtime, Whisper, etc.) can be found
if (process.platform === 'win32') {
  const electronVersion = process.versions.electron.split('.')[0];
  const arch = process.arch;

  // In development: __dirname points to project root
  // In production (asar): __dirname points to app.asar, but unpacked files are in app.asar.unpacked
  // In production (resources): __dirname might point to resources/app.asar
  let basePath = __dirname;
  if (__dirname.includes('.asar')) {
    basePath = __dirname.replace('app.asar', 'app.asar.unpacked');
  }

  // Get resources path for extraResources (XmlaBridge DLLs)
  const resourcesPath = app.isPackaged ? process.resourcesPath : __dirname;

  const pathsToAdd = [];

  // 1. electron-edge-js native DLLs (VC++ runtime, edge_nativeclr.node)
  const edgeNativePath = path.join(basePath, 'node_modules', 'electron-edge-js', 'lib', 'native', 'win32', arch, electronVersion);
  if (fs.existsSync(edgeNativePath)) {
    pathsToAdd.push(edgeNativePath);
    const dlls = fs.readdirSync(edgeNativePath).filter(f => f.endsWith('.dll') || f.endsWith('.node'));
    console.log('[Setup] ✓ Found electron-edge-js native files:', dlls.join(', '));
  } else {
    console.error('[Setup] ✗ electron-edge-js native path not found:', edgeNativePath);
  }

  // 2. Whisper native DLLs (ggml-whisper.dll, whisper.dll) - in runtimes/win-x64
  const whisperNativePath = path.join(resourcesPath, 'XmlaBridge', 'bin', 'Release', 'net48', 'runtimes', 'win-x64');
  if (fs.existsSync(whisperNativePath)) {
    pathsToAdd.push(whisperNativePath);
    const dlls = fs.readdirSync(whisperNativePath).filter(f => f.endsWith('.dll'));
    console.log('[Setup] ✓ Found Whisper native DLLs in runtimes/win-x64:', dlls.join(', '));
  } else {
    console.error('[Setup] ✗ Whisper native path not found:', whisperNativePath);
  }

  // 2b. Additional native DLLs in runtimes/win-x64/native (e.g., msalruntime.dll)
  const whisperNativePath2 = path.join(resourcesPath, 'XmlaBridge', 'bin', 'Release', 'net48', 'runtimes', 'win-x64', 'native');
  if (fs.existsSync(whisperNativePath2)) {
    pathsToAdd.push(whisperNativePath2);
    const dlls = fs.readdirSync(whisperNativePath2).filter(f => f.endsWith('.dll'));
    console.log('[Setup] ✓ Found additional native DLLs:', dlls.join(', '));
  } else {
    console.error('[Setup] ✗ Additional native path not found:', whisperNativePath2);
  }

  // 3. XmlaBridge main directory (all .NET DLLs)
  const xmlaBridgePath = path.join(resourcesPath, 'XmlaBridge', 'bin', 'Release', 'net48');
  if (fs.existsSync(xmlaBridgePath)) {
    pathsToAdd.push(xmlaBridgePath);
    const dlls = fs.readdirSync(xmlaBridgePath).filter(f => f.endsWith('.dll') && !fs.statSync(path.join(xmlaBridgePath, f)).isDirectory());
    console.log('[Setup] ✓ Found XmlaBridge DLLs:', dlls.length, 'files');
  } else {
    console.error('[Setup] ✗ XmlaBridge path not found:', xmlaBridgePath);
  }

  // Add all paths to PATH environment variable
  if (pathsToAdd.length > 0) {
    const currentPath = process.env.PATH || '';
    process.env.PATH = pathsToAdd.join(path.delimiter) + path.delimiter + currentPath;
    console.log('[Setup] ✓ Added', pathsToAdd.length, 'native library paths to PATH');
  } else {
    console.error('[Setup] ✗ No native library paths found!');
  }
}

const edge = require('electron-edge-js');
const CONFIG = require('./config');
const { escapeXml, unescapeXml } = require('./utils');

// Handle creating/removing shortcuts on Windows when installing/uninstalling
if (require('electron-squirrel-startup')) {
  app.quit();
}

let mainWindow;

/**
 * Parse command-line arguments from Power BI
 * Power BI passes connection details as command-line arguments
 */
const args = process.argv.slice(1);
let serverName = '';
let databaseName = '';

// Power BI passes arguments like: "Server=localhost:12345;Database=abc123;ApplicationName=PowerBI"
args.forEach((arg) => {
  if (arg.includes('Server=') || arg.includes('Database=')) {
    const match = arg.match(/Server=([^;]+)/);
    if (match) serverName = match[1];

    const dbMatch = arg.match(/Database=([^;]+)/);
    if (dbMatch) databaseName = dbMatch[1];
  }
});

// For testing without Power BI, allow manual server/database input
if (!serverName) {
  console.log('No Power BI connection detected. App will start in standalone mode.');
}

/**
 * Creates the main application window
 * Sets up the BrowserWindow with security settings and loads the UI
 */
function createWindow() {
  mainWindow = new BrowserWindow({
    width: CONFIG.WINDOW.WIDTH,
    height: CONFIG.WINDOW.HEIGHT,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    icon: path.join(__dirname, 'robot_icon_64.png'),
    title: CONFIG.WINDOW.TITLE
  });

  const indexPath = path.join(__dirname, 'index.html');
  console.log('[DEBUG] __dirname:', __dirname);
  console.log('[DEBUG] indexPath:', indexPath);
  console.log('[DEBUG] index.html exists:', fs.existsSync(indexPath));

  mainWindow.loadFile(indexPath);

  // Send connection info to renderer process after page loads
  mainWindow.webContents.on('did-finish-load', () => {
    mainWindow.webContents.send('connection-info', {
      server: serverName,
      database: databaseName
    });
  });

  // Forward main process console logs to renderer
  const originalConsoleLog = console.log;
  const originalConsoleError = console.error;

  console.log = (...args) => {
    originalConsoleLog(...args);
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('main-log', { type: 'log', message: args.join(' ') });
    }
  };

  console.error = (...args) => {
    originalConsoleError(...args);
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('main-log', { type: 'error', message: args.join(' ') });
    }
  };

  // Open DevTools in development
  if (CONFIG.DEV.OPEN_DEVTOOLS) {
    mainWindow.webContents.openDevTools();
  }
}

app.whenReady().then(() => {
  // Configure CSP to allow Web Speech API and microphone access
  // Allow any HTTPS connection to support custom API endpoints
  session.defaultSession.webRequest.onHeadersReceived((details, callback) => {
    callback({
      responseHeaders: {
        ...details.responseHeaders,
        'Content-Security-Policy': [
          "default-src 'self'; " +
          "script-src 'self' 'unsafe-inline'; " +
          "style-src 'self' 'unsafe-inline'; " +
          "connect-src 'self' https: http://localhost:* ws://localhost:* http://127.0.0.1:* ws://127.0.0.1:* wss://localhost:* wss://127.0.0.1:*; " +
          "media-src 'self' blob: mediastream:; " +
          "img-src 'self' data: blob:;"
        ]
      }
    });
  });

  // Grant permissions for microphone and media devices
  session.defaultSession.setPermissionRequestHandler((webContents, permission, callback) => {
    if (permission === 'media' || permission === 'microphone' || permission === 'audioCapture') {
      console.log('[Permissions] Granted:', permission);
      callback(true);
    } else {
      console.log('[Permissions] Denied:', permission);
      callback(false);
    }
  });

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// IPC Handler for opening external URLs
ipcMain.handle('open-external', async (event, url) => {
  const { shell } = require('electron');
  try {
    console.log('[Main] Opening external URL:', url);
    await shell.openExternal(url);
    console.log('[Main] Successfully opened URL');
    return { success: true };
  } catch (error) {
    console.error('[Main] Failed to open URL:', error);
    throw error;
  }
});

// Load .NET ADOMD.NET bridge
let xmlaBridge = null;
let whisperBridge = null;

/**
 * Gets the correct base path for .NET DLL based on whether app is packaged
 * @returns {string} Base path for resources
 */
function getBasePath() {
  if (app.isPackaged) {
    // In production, DLLs are in resources folder (extraResources)
    return process.resourcesPath;
  } else {
    // In development, use __dirname
    return __dirname;
  }
}

/**
 * Loads the .NET bridge for ADOMD.NET queries
 * Uses electron-edge-js to call into .NET assemblies
 * @returns {Function} The loaded bridge function
 * @throws {Error} If the bridge DLL cannot be loaded
 */
function loadXmlaBridge() {
  if (xmlaBridge) return xmlaBridge;

  try {
    const basePath = getBasePath();
    const bridgePath = path.join(basePath, ...CONFIG.BRIDGE.PATH);
    console.log('[.NET Bridge] App packaged:', app.isPackaged);
    console.log('[.NET Bridge] Base path:', basePath);
    console.log('[.NET Bridge] Loading from:', bridgePath);

    // Check if the DLL exists
    if (!fs.existsSync(bridgePath)) {
      const errorMsg = `
===========================================
❌ .NET Bridge DLL Not Found
===========================================

Expected location: ${bridgePath}

This DLL is required to connect to Power BI's XMLA endpoint.

Troubleshooting steps:
1. Build the XmlaBridge project:
   cd XmlaBridge
   dotnet build -c Release

2. Verify the DLL exists at:
   ${bridgePath}

3. Check that .NET Framework 4.8 is installed

4. If using a different target framework, update
   config.js BRIDGE.PATH to match your build output

For more help, see README.md
===========================================
      `.trim();

      throw new Error(errorMsg);
    }

    xmlaBridge = edge.func({
      assemblyFile: bridgePath,
      typeName: CONFIG.BRIDGE.TYPE_NAME,
      methodName: CONFIG.BRIDGE.METHOD_NAME
    });
    console.log('[.NET Bridge] Loaded successfully');

    return xmlaBridge;
  } catch (error) {
    console.error('[.NET Bridge] Failed to load:', error.message);

    // Provide helpful error messages based on error type
    if (error.message.includes('Could not load file or assembly')) {
      const detailedError = `
❌ Failed to load .NET assembly

${error.message}

Possible causes:
- Missing dependencies (Microsoft.AnalysisServices.AdomdClient)
- Wrong .NET Framework version
- Corrupted DLL file

Try rebuilding:
  cd XmlaBridge
  dotnet clean
  dotnet build -c Release
      `.trim();
      throw new Error(detailedError);
    }

    throw error;
  }
}

/**
 * Loads the .NET bridge for Whisper speech-to-text
 * @returns {Function} The loaded bridge function
 * @throws {Error} If the bridge DLL cannot be loaded
 */
function loadWhisperBridge() {
  if (whisperBridge) return whisperBridge;

  try {
    const basePath = getBasePath();
    const bridgePath = path.join(basePath, ...CONFIG.BRIDGE.PATH);
    console.log('[Whisper Bridge] App packaged:', app.isPackaged);
    console.log('[Whisper Bridge] Base path:', basePath);
    console.log('[Whisper Bridge] Loading from:', bridgePath);

    if (!fs.existsSync(bridgePath)) {
      throw new Error('Whisper Bridge DLL not found. Please build the XmlaBridge project first.');
    }

    whisperBridge = edge.func({
      assemblyFile: bridgePath,
      typeName: 'WhisperBridge',
      methodName: 'Invoke'
    });
    console.log('[Whisper Bridge] Loaded successfully');

    return whisperBridge;
  } catch (error) {
    console.error('[Whisper Bridge] Failed to load:', error.message);
    throw error;
  }
}

/**
 * Extracts the query statement from a SOAP XML body
 * @param {string} soapBody - The SOAP XML body containing the query
 * @returns {string} The extracted and unescaped query
 * @throws {Error} If the query cannot be extracted
 */
function extractQueryFromSOAP(soapBody) {
  const match = soapBody.match(/<Statement>([\s\S]*?)<\/Statement>/);
  if (match && match[1]) {
    return unescapeXml(match[1]).trim();
  }
  throw new Error('Could not extract query from SOAP body');
}

// IPC Handler for XMLA requests using ADOMD.NET
ipcMain.handle('xmla-request', async (event, { endpoint, soapBody }) => {
  try {
    // Parse the endpoint URL
    const url = new URL(endpoint);
    const server = `${url.hostname}:${url.port}`;

    // Extract database from the SOAP body if possible
    const catalogMatch = soapBody.match(/<Catalog>(.*?)<\/Catalog>/);
    const database = catalogMatch ? catalogMatch[1] : '';

    // Extract the actual query from SOAP
    const query = extractQueryFromSOAP(soapBody);

    console.log('[ADOMD.NET] Server:', server);
    console.log('[ADOMD.NET] Database:', database);
    console.log('[ADOMD.NET] Query (first 200 chars):', query.substring(0, CONFIG.QUERY.LOG_PREVIEW_LENGTH));

    // Load the .NET bridge
    let bridge;
    try {
      bridge = loadXmlaBridge();
    } catch (bridgeError) {
      // Bridge loading failed - provide user-friendly error
      throw new Error(`Failed to load .NET Bridge: ${bridgeError.message}`);
    }

    // Call the .NET function with timeout
    const result = await new Promise((resolve, reject) => {
      bridge({
        server: server,
        database: database,
        query: query,
        timeout: CONFIG.QUERY.COMMAND_TIMEOUT
      }, (error, result) => {
        if (error) {
          console.error('[ADOMD.NET] Bridge error:', error);
          reject(error);
        } else {
          resolve(result);
        }
      });
    });

    if (result.success) {
      console.log('[ADOMD.NET] Query successful, rows:', result.rowCount);

      // Convert to XML format expected by the parser
      const xmlResponse = convertToXMLResponse(result.data);

      return { success: true, data: xmlResponse };
    } else {
      console.error('[ADOMD.NET] Query failed:', result.error);

      // Provide user-friendly error messages
      let errorMessage = result.error;

      if (result.errorType === 'AdomdConnectionException') {
        errorMessage = `Connection failed: ${result.error}\n\nMake sure:\n• Power BI Desktop is running\n• The report is open\n• External Tools are enabled`;
      } else if (result.errorType === 'AdomdErrorResponseException') {
        errorMessage = `Query error: ${result.error}\n\nCheck your DAX syntax and try again.`;
      }

      throw new Error(errorMessage);
    }

  } catch (error) {
    console.error('[ADOMD.NET] Request error:', error.message);

    // Add context to network errors
    if (error.message.includes('Invalid URL')) {
      throw new Error(`Invalid XMLA endpoint URL. Make sure Power BI Desktop is running and the report is open.`);
    }

    throw error;
  }
});

// IPC Handler for Whisper speech-to-text transcription
ipcMain.handle('whisper-transcribe', async (event, { audioData }) => {
  try {
    console.log('[Whisper] Transcription request received');
    console.log('[Whisper] Audio data length:', audioData?.length || 0);

    // Load the Whisper bridge
    let bridge;
    try {
      bridge = loadWhisperBridge();
    } catch (bridgeError) {
      console.error('[Whisper] Bridge load error:', bridgeError);
      throw new Error(`Failed to load Whisper Bridge: ${bridgeError.message}`);
    }

    // Call the .NET function
    console.log('[Whisper] About to call bridge...');
    const result = await new Promise((resolve, reject) => {
      try {
        bridge({
          audioData: audioData
        }, (error, result) => {
          console.log('[Whisper] Bridge callback received');

          if (error) {
            console.error('[Whisper] Bridge callback ERROR');
            console.error('[Whisper] Bridge error:', error);
            console.error('[Whisper] Bridge error type:', typeof error, error.constructor?.name);
            console.error('[Whisper] Bridge error string:', String(error));
            console.error('[Whisper] Bridge error JSON:', JSON.stringify(error, null, 2));

            // Extract more details from the error
            if (error.InnerException) {
              console.error('[Whisper] Inner exception:', error.InnerException);
            }
            if (error.Message) {
              console.error('[Whisper] Error message:', error.Message);
            }
            if (error.message) {
              console.error('[Whisper] error.message:', error.message);
            }

            reject(error);
          } else {
            console.log('[Whisper] Bridge callback SUCCESS');
            console.log('[Whisper] Result:', JSON.stringify(result, null, 2));
            resolve(result);
          }
        });
      } catch (syncError) {
        console.error('[Whisper] Synchronous error calling bridge:', syncError);
        reject(syncError);
      }
    });

    if (result.success) {
      console.log('[Whisper] Transcription successful:', result.text);
      return { success: true, text: result.text };
    } else {
      console.error('[Whisper] Transcription failed:', result.error);
      console.error('[Whisper] Error type:', result.errorType);
      console.error('[Whisper] Error details:', result.details);

      // Return the detailed error to the renderer
      throw new Error(result.error || 'Unknown transcription error');
    }

  } catch (error) {
    console.error('[Whisper] Request error:', error.message);
    console.error('[Whisper] Full error:', error);
    throw error;
  }
});

/**
 * Converts JSON query results to XML format
 * @param {Array<Object>} data - Array of row objects
 * @returns {string} XML-formatted response
 */
function convertToXMLResponse(data) {
  let xml = '<?xml version="1.0" encoding="utf-8"?>';
  xml += '<root xmlns:xsd="http://www.w3.org/2001/XMLSchema">';

  data.forEach(row => {
    xml += '<row>';
    for (const [key, value] of Object.entries(row)) {
      xml += `<${key}>${escapeXml(value)}</${key}>`;
    }
    xml += '</row>';
  });

  xml += '</root>';
  return xml;
}
