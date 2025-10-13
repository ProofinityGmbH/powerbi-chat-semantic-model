// Global state
let connectionInfo = { server: '', database: '' };
let semanticModel = { tables: [], measures: [], relationships: [], sampleData: {} };
let chatHistory = [];
let xmlaConnection = null;
let connectionMonitorInterval = null;

/**
 * Loads user settings from localStorage
 * Populates API URL, token, and model name fields
 */
function loadSettings() {
    const CONFIG = typeof module !== 'undefined' && module.exports ? require('./config') : window.CONFIG;
    document.getElementById('apiUrl').value = localStorage.getItem('apiUrl') || CONFIG.API.DEFAULT_URL;
    document.getElementById('apiToken').value = localStorage.getItem('apiToken') || '';
    document.getElementById('modelName').value = localStorage.getItem('modelName') || CONFIG.API.DEFAULT_MODEL;
}

/**
 * Saves user settings to localStorage
 * Persists API URL, token, and model name
 */
function saveSettings() {
    localStorage.setItem('apiUrl', document.getElementById('apiUrl').value);
    localStorage.setItem('apiToken', document.getElementById('apiToken').value);
    localStorage.setItem('modelName', document.getElementById('modelName').value);
}

// Listen for connection info from main process
if (window.electronAPI) {
    window.electronAPI.onConnectionInfo((event, info) => {
        connectionInfo = info;
        updateConnectionUI(info);
        connectToSemanticModel(info);
    });

    // Listen for main process logs
    window.electronAPI.onMainLog((event, logData) => {
        if (logData.type === 'error') {
            console.error('[MAIN PROCESS]', logData.message);
        }
    });
}

/**
 * Updates the connection status UI
 * @param {Object} info - Connection information {server, database}
 */
function updateConnectionUI(info) {
    document.getElementById('serverInfo').textContent = info.server || 'Not connected';
    document.getElementById('databaseInfo').textContent = info.database || 'N/A';

    const statusBadge = document.getElementById('statusBadge');
    if (info.server && info.database) {
        statusBadge.textContent = 'Connecting...';
        statusBadge.className = 'status loading';
    } else {
        statusBadge.textContent = 'Standalone Mode';
        statusBadge.className = 'status disconnected';
    }
}

/**
 * Starts monitoring the connection status
 * Periodically checks if the connection is still alive
 */
function startConnectionMonitor() {
    // Clear any existing monitor
    if (connectionMonitorInterval) {
        clearInterval(connectionMonitorInterval);
    }

    // Check connection every 30 seconds
    connectionMonitorInterval = setInterval(async () => {
        if (!xmlaConnection) return;

        try {
            const isAlive = await xmlaConnection.testConnection();
            const statusBadge = document.getElementById('statusBadge');

            if (isAlive) {
                statusBadge.textContent = 'Connected';
                statusBadge.className = 'status connected';
            } else {
                statusBadge.textContent = 'Connection Lost';
                statusBadge.className = 'status disconnected';
                console.warn('[Monitor] Connection lost to Power BI');
            }
        } catch (error) {
            console.error('[Monitor] Failed to check connection:', error);
        }
    }, 30000); // Check every 30 seconds
}

/**
 * Stops monitoring the connection status
 */
function stopConnectionMonitor() {
    if (connectionMonitorInterval) {
        clearInterval(connectionMonitorInterval);
        connectionMonitorInterval = null;
    }
}

/**
 * Connects to the Power BI semantic model via XMLA
 * @param {Object} info - Connection information {server, database}
 */
async function connectToSemanticModel(info) {
    if (!info.server || !info.database) {
        showError('No Power BI connection available. Run this tool from Power BI Desktop.');
        return;
    }

    try {
        // Fetch metadata from the semantic model
        const metadata = await fetchSemanticModelMetadata(info);
        semanticModel = metadata;

        // Update UI
        displayMetadata(metadata);

        const statusBadge = document.getElementById('statusBadge');
        statusBadge.textContent = 'Connected';
        statusBadge.className = 'status connected';

        // Start monitoring the connection
        startConnectionMonitor();
    } catch (error) {
        console.error('[Renderer] Failed to connect:', error);

        // Stop monitoring if connection failed
        stopConnectionMonitor();

        let errorMessage = `Failed to connect to semantic model: ${error.message}`;

        // Add helpful hints based on error type
        if (error.message.includes('ECONNRESET')) {
            errorMessage += '\n\nTip: The connection was reset. This may mean:\n• Power BI Desktop XMLA endpoint is not enabled\n• The port is incorrect\n• A firewall is blocking the connection';
        } else if (error.message.includes('ECONNREFUSED')) {
            errorMessage += '\n\nTip: Connection refused. Make sure:\n• Power BI Desktop is running\n• The report is open\n• External tools are enabled';
        } else if (error.message.includes('timeout')) {
            errorMessage += '\n\nTip: The request timed out. The server may be slow or unreachable.';
        }

        showError(errorMessage);

        const statusBadge = document.getElementById('statusBadge');
        statusBadge.textContent = 'Connection Failed';
        statusBadge.className = 'status disconnected';
    }
}

/**
 * Fetches semantic model metadata using XMLA/SOAP
 * @param {Object} info - Connection information {server, database}
 * @returns {Promise<Object>} Metadata including tables, measures, and relationships
 * @throws {Error} If the connection or metadata fetch fails
 */
async function fetchSemanticModelMetadata(info) {
    try {
        // Create XMLA connection
        xmlaConnection = new XMLAConnection(info.server, info.database);

        // Get config for sample data settings
        const CONFIG = typeof window !== 'undefined' && window.CONFIG ? window.CONFIG : null;
        const maxSampleRows = CONFIG ? CONFIG.QUERY.DEFAULT_SAMPLE_ROWS : 3;

        // Fetch complete semantic model metadata with sample data
        const metadata = await xmlaConnection.getSemanticModelMetadata({
            fetchSampleData: true,
            maxSampleRows: maxSampleRows
        });

        return metadata;

    } catch (error) {
        console.error('XMLA connection failed:', error);
        throw error;
    }
}

/**
 * Displays semantic model metadata in the sidebar
 * @param {Object} metadata - The metadata object containing tables, measures, and relationships
 */
function displayMetadata(metadata) {
    const container = document.getElementById('metadataContainer');
    container.innerHTML = '';

    // Tables section
    if (metadata.tables && metadata.tables.length > 0) {
        const tablesSection = createCollapsibleSection(
            'tables',
            '📋 Tables',
            metadata.tables.length,
            true // expanded by default
        );

        const tablesList = document.createElement('ul');
        tablesList.className = 'metadata-list';

        metadata.tables.forEach(table => {
            const item = document.createElement('li');
            item.className = 'metadata-item';
            const columnCount = table.columns ? table.columns.length : 0;
            item.innerHTML = `
                <span class="metadata-item-name">${table.name}</span>
                <span class="metadata-item-type">${columnCount} columns</span>
            `;
            tablesList.appendChild(item);
        });

        tablesSection.querySelector('.metadata-section-content').appendChild(tablesList);
        container.appendChild(tablesSection);
    }

    // Measures section
    if (metadata.measures && metadata.measures.length > 0) {
        const measuresSection = createCollapsibleSection(
            'measures',
            '📊 Measures',
            metadata.measures.length,
            true // expanded by default
        );

        const measuresList = document.createElement('ul');
        measuresList.className = 'metadata-list';

        metadata.measures.forEach(measure => {
            const item = document.createElement('li');
            item.className = 'metadata-item';
            item.innerHTML = `
                <span class="metadata-item-name">${measure.name}</span>
                <span class="metadata-item-type">${measure.table || ''}</span>
            `;
            measuresList.appendChild(item);
        });

        measuresSection.querySelector('.metadata-section-content').appendChild(measuresList);
        container.appendChild(measuresSection);
    }

    // Relationships section
    if (metadata.relationships && metadata.relationships.length > 0) {
        const relationshipsSection = createCollapsibleSection(
            'relationships',
            '🔗 Relationships',
            metadata.relationships.length,
            false // collapsed by default since it's less commonly used
        );

        const relationshipsList = document.createElement('ul');
        relationshipsList.className = 'metadata-list';

        metadata.relationships.forEach(rel => {
            const item = document.createElement('li');
            item.className = 'metadata-item';
            item.innerHTML = `
                <span class="metadata-item-name">${rel.from}</span>
                <span class="metadata-item-type">→ ${rel.to}</span>
            `;
            relationshipsList.appendChild(item);
        });

        relationshipsSection.querySelector('.metadata-section-content').appendChild(relationshipsList);
        container.appendChild(relationshipsSection);
    }
}

/**
 * Creates a collapsible section for metadata display
 * @param {string} id - Unique identifier for the section
 * @param {string} title - Display title for the section
 * @param {number} count - Number of items in the section
 * @param {boolean} expanded - Whether the section starts expanded
 * @returns {HTMLElement} The collapsible section element
 */
function createCollapsibleSection(id, title, count, expanded = true) {
    const section = document.createElement('div');
    section.className = 'metadata-section';
    section.id = `section-${id}`;

    const header = document.createElement('div');
    header.className = `metadata-section-header${expanded ? '' : ' collapsed'}`;

    const titleDiv = document.createElement('div');
    titleDiv.className = 'metadata-section-title';

    const icon = document.createElement('span');
    icon.className = `collapse-icon${expanded ? '' : ' collapsed'}`;
    icon.textContent = '▼';

    const h4 = document.createElement('h4');
    h4.textContent = `${title} (${count})`;

    titleDiv.appendChild(icon);
    titleDiv.appendChild(h4);
    header.appendChild(titleDiv);

    const content = document.createElement('div');
    content.className = `metadata-section-content${expanded ? '' : ' collapsed'}`;

    // Add click event to toggle collapse
    header.addEventListener('click', () => {
        toggleSection(icon, content, header);
    });

    section.appendChild(header);
    section.appendChild(content);

    return section;
}

/**
 * Toggles the collapse state of a section
 * @param {HTMLElement} icon - The collapse icon element
 * @param {HTMLElement} content - The content element to collapse/expand
 * @param {HTMLElement} header - The header element
 */
function toggleSection(icon, content, header) {
    const isCollapsed = content.classList.contains('collapsed');

    if (isCollapsed) {
        // Expand
        content.classList.remove('collapsed');
        icon.classList.remove('collapsed');
        header.classList.remove('collapsed');
    } else {
        // Collapse
        content.classList.add('collapsed');
        icon.classList.add('collapsed');
        header.classList.add('collapsed');
    }
}

/**
 * Displays an error message in the metadata container
 * @param {string} message - The error message to display
 */
function showError(message) {
    const container = document.getElementById('metadataContainer');
    // Convert newlines to HTML line breaks
    const formattedMessage = message.replace(/\n/g, '<br>');
    container.innerHTML = `
        <div class="empty-state">
            <div class="empty-state-icon">⚠️</div>
            <p style="color: #dc3545; text-align: left; font-size: 12px; line-height: 1.6;">${formattedMessage}</p>
        </div>
    `;
}

// Chat functionality
document.getElementById('sendButton').addEventListener('click', sendMessage);
document.getElementById('chatInput').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
        sendMessage();
    }
});

// Save settings when changed
['apiUrl', 'apiToken', 'modelName'].forEach(id => {
    document.getElementById(id).addEventListener('change', saveSettings);
});

/**
 * Executes a DAX query against the semantic model
 * @param {string} daxQuery - The DAX query to execute
 * @returns {Promise<Array>} Query results
 * @throws {Error} If not connected or query fails
 */
async function executeDAXQuery(daxQuery) {
    if (!xmlaConnection) {
        throw new Error('Not connected to semantic model');
    }

    const results = await xmlaConnection.executeDAX(daxQuery);
    return results;
}

/**
 * Sends a message to the AI assistant or executes a DAX query
 * Handles both chat messages and DAX query execution
 */
async function sendMessage() {
    const input = document.getElementById('chatInput');
    const message = input.value.trim();

    if (!message) return;

    // Clear input
    input.value = '';

    // Add user message to chat
    addMessage('user', message);

    // Check if this is a DAX query command
    if (message.toUpperCase().startsWith('DAX:') || message.toUpperCase().startsWith('EXECUTE:')) {
        const daxQuery = message.substring(message.indexOf(':') + 1).trim();

        // Disable send button
        const sendButton = document.getElementById('sendButton');
        sendButton.disabled = true;
        sendButton.textContent = 'Executing...';

        try {
            const result = await executeDAXQuery(daxQuery);

            // Handle both old format (array) and new format (object)
            const results = result.data || result;
            const warnings = result.warnings || [];
            const rowCount = result.rowCount || results.length;

            // Format and display results
            let resultText = `**Query Results** (${rowCount} rows):\n\n`;

            // Show warnings if any
            if (warnings.length > 0) {
                resultText += '**⚠️ Warnings:**\n';
                warnings.forEach(warning => {
                    resultText += `- ${warning}\n`;
                });
                resultText += '\n';
            }

            if (results.length > 0) {
                // Show first 100 rows in the UI to prevent browser freeze
                const displayRows = results.slice(0, 100);
                resultText += '```json\n' + JSON.stringify(displayRows, null, 2) + '\n```';
                if (results.length > 100) {
                    resultText += `\n_Showing first 100 of ${results.length} rows_`;
                }
            } else {
                resultText += 'No results returned.';
            }

            addMessage('assistant', resultText);
        } catch (error) {
            console.error('DAX execution error:', error);
            addMessage('error', `Failed to execute DAX query: ${error.message}`);
        } finally {
            sendButton.disabled = false;
            sendButton.textContent = 'Send';
        }
        return;
    }

    // Disable send button
    const sendButton = document.getElementById('sendButton');
    sendButton.disabled = true;
    sendButton.textContent = 'Sending...';

    try {
        // Save settings
        saveSettings();

        // Get settings
        const apiUrl = document.getElementById('apiUrl').value;
        const apiToken = document.getElementById('apiToken').value;
        const modelName = document.getElementById('modelName').value;

        if (!apiUrl || !apiToken) {
            throw new Error('Please configure API URL and Token in the settings panel');
        }

        // Build system message with semantic model context
        let systemContent = 'You are a Power BI semantic model assistant with access to both the model structure and actual data. ';

        if (semanticModel.tables.length > 0 || semanticModel.measures.length > 0) {
            systemContent += 'Here is the semantic model structure and sample data:\n\n';

            // Add tables with sample data
            if (semanticModel.tables.length > 0) {
                systemContent += '## Tables:\n';
                semanticModel.tables.forEach(table => {
                    systemContent += `- **${table.name}**`;
                    if (table.description) systemContent += ` - ${table.description}`;
                    systemContent += `\n`;
                    if (table.columns && table.columns.length > 0) {
                        const columnNames = table.columns.map(col => col.name || col.displayName || col).join(', ');
                        systemContent += `  Columns: ${columnNames}\n`;
                    }

                    // Add sample data if available
                    if (semanticModel.sampleData && semanticModel.sampleData[table.name] && semanticModel.sampleData[table.name].length > 0) {
                        systemContent += `  Sample data (top ${semanticModel.sampleData[table.name].length} rows):\n`;
                        systemContent += '  ```\n';
                        systemContent += '  ' + JSON.stringify(semanticModel.sampleData[table.name], null, 2).split('\n').join('\n  ') + '\n';
                        systemContent += '  ```\n';
                    }
                });
                systemContent += '\n';
            }

            // Add measures
            if (semanticModel.measures.length > 0) {
                systemContent += '## Measures:\n';
                semanticModel.measures.forEach(measure => {
                    systemContent += `- **${measure.name}**`;
                    if (measure.table) systemContent += ` (${measure.table})`;
                    if (measure.expression) systemContent += `\n  Expression: ${measure.expression}`;
                    systemContent += '\n';
                });
                systemContent += '\n';
            }

            // Add relationships
            if (semanticModel.relationships.length > 0) {
                systemContent += '## Relationships:\n';
                semanticModel.relationships.forEach(rel => {
                    systemContent += `- ${rel.from} → ${rel.to}\n`;
                });
                systemContent += '\n';
            }

            // Add instructions for querying data
            systemContent += '\n## Instructions:\n';
            systemContent += '- You can see sample data above to help answer questions about the data.\n';
            systemContent += '- If the user asks about specific data values, trends, or analysis, use the sample data and structure to provide informed responses.\n';
            systemContent += '- You can help write DAX queries to analyze the data.\n';
            systemContent += '- When suggesting DAX queries, tell the user they can execute them by typing "DAX: <query>" in the chat.\n';
            systemContent += '- Be specific and reference actual column names and table names from the model.\n';
            systemContent += '- Sample data shown is limited to a few rows - encourage users to run DAX queries for complete analysis.\n';
        } else {
            systemContent += 'The semantic model is not connected or has no metadata available.';
        }

        // Build messages array
        const messages = [
            { role: 'system', content: systemContent },
            ...chatHistory,
            { role: 'user', content: message }
        ];

        // Make API call
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiToken}`
            },
            body: JSON.stringify({
                model: modelName,
                messages: messages
            })
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`HTTP ${response.status}: ${errorText}`);
        }

        const data = await response.json();

        // Extract AI response
        let aiMessage = '';
        if (data.choices && data.choices[0] && data.choices[0].message) {
            aiMessage = data.choices[0].message.content;
        } else {
            aiMessage = 'Received response but could not parse message.';
        }

        // Add AI response to chat
        addMessage('assistant', aiMessage);

        // Update chat history
        chatHistory.push({ role: 'user', content: message });
        chatHistory.push({ role: 'assistant', content: aiMessage });

    } catch (error) {
        console.error('Error sending message:', error);
        addMessage('error', `Error: ${error.message}`);
    } finally {
        // Re-enable send button
        sendButton.disabled = false;
        sendButton.textContent = 'Send';
    }
}

/**
 * Adds a message to the chat UI
 * @param {string} role - The role of the message sender ('user', 'assistant', or 'error')
 * @param {string} content - The message content (supports simple markdown)
 */
function addMessage(role, content) {
    const chatMessages = document.getElementById('chatMessages');

    // Remove empty state if present
    const emptyState = chatMessages.querySelector('.empty-state');
    if (emptyState) {
        emptyState.remove();
    }

    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${role}`;

    const label = role === 'user' ? 'You' : (role === 'assistant' ? 'AI Assistant' : 'Error');

    // Simple markdown-like formatting
    let formattedContent = content
        .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>') // Bold
        .replace(/```(.*?)\n([\s\S]*?)```/g, '<pre><code>$2</code></pre>') // Code blocks
        .replace(/\n/g, '<br>'); // Line breaks

    messageDiv.innerHTML = `
        <div class="message-label">${label}</div>
        <div class="message-content">${formattedContent}</div>
    `;

    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

// Initialize
loadSettings();

// If not connected to Power BI, show message
if (!window.electronAPI) {
    showError('Running in browser mode. This tool requires Electron.');
}

// Modal functionality - wrap in DOMContentLoaded to ensure elements exist
document.addEventListener('DOMContentLoaded', function() {
    const infoButton = document.getElementById('infoButton');
    const infoModal = document.getElementById('infoModal');
    const closeModal = document.getElementById('closeModal');
    const githubLink = document.getElementById('githubLink');
    const donateLink = document.getElementById('donateLink');

    // Open modal
    if (infoButton) {
        infoButton.addEventListener('click', () => {
            infoModal.classList.add('show');
        });
    }

    // Close modal
    if (closeModal) {
        closeModal.addEventListener('click', () => {
            infoModal.classList.remove('show');
        });
    }

    // Close modal when clicking outside
    if (infoModal) {
        infoModal.addEventListener('click', (e) => {
            if (e.target === infoModal) {
                infoModal.classList.remove('show');
            }
        });
    }

    // Handle external links
    if (githubLink) {
        githubLink.addEventListener('click', async (e) => {
            e.preventDefault();
            const url = 'https://github.com/markusbegerow/powerbi-chat-semantic-model';
            console.log('[GitHub Link] Clicked, URL:', url);
            try {
                if (window.electronAPI && window.electronAPI.openExternal) {
                    console.log('[GitHub Link] Using electronAPI.openExternal');
                    await window.electronAPI.openExternal(url);
                    console.log('[GitHub Link] Opened successfully');
                } else {
                    console.log('[GitHub Link] Fallback to window.open');
                    window.open(url, '_blank');
                }
            } catch (error) {
                console.error('[GitHub Link] Error:', error);
                alert('Failed to open link: ' + error.message);
            }
        });
    }

    if (donateLink) {
        donateLink.addEventListener('click', async (e) => {
            e.preventDefault();
            const url = 'https://paypal.me/MarkusBegerow?country.x=DE&locale.x=de_DE';
            console.log('[Donate Link] Clicked, URL:', url);
            try {
                if (window.electronAPI && window.electronAPI.openExternal) {
                    console.log('[Donate Link] Using electronAPI.openExternal');
                    await window.electronAPI.openExternal(url);
                    console.log('[Donate Link] Opened successfully');
                } else {
                    console.log('[Donate Link] Fallback to window.open');
                    window.open(url, '_blank');
                }
            } catch (error) {
                console.error('[Donate Link] Error:', error);
                alert('Failed to open link: ' + error.message);
            }
        });
    }
});
