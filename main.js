const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
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

/**
 * Loads the .NET bridge for ADOMD.NET queries
 * Uses electron-edge-js to call into .NET assemblies
 * @returns {Function} The loaded bridge function
 * @throws {Error} If the bridge DLL cannot be loaded
 */
function loadXmlaBridge() {
  if (xmlaBridge) return xmlaBridge;

  try {
    const bridgePath = path.join(__dirname, ...CONFIG.BRIDGE.PATH);
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
