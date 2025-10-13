// XMLA Connection Module for Power BI Semantic Models
// This module handles SOAP/XMLA requests to Analysis Services

class XMLAConnection {
    constructor(server, database) {
        this.server = server;
        this.database = database;

        // Construct XMLA endpoint
        // If server already includes protocol, use it as-is
        if (server.startsWith('http://') || server.startsWith('https://')) {
            this.xmlaEndpoint = `${server}/xmla`;
        } else {
            // Default to http for localhost connections
            this.xmlaEndpoint = `http://${server}/xmla`;
        }

        // Initialize DAX validator
        const CONFIG = typeof window !== 'undefined' && window.CONFIG ? window.CONFIG : null;
        if (CONFIG && typeof DAXValidator !== 'undefined') {
            this.daxValidator = new DAXValidator(CONFIG.QUERY);
        } else {
            this.daxValidator = null;
        }

        // Connection monitoring
        this.isConnected = false;
        this.lastSuccessfulQuery = null;
    }

    // Test the connection with a simple query
    async testConnection() {
        try {
            // Simple query to check if connection is alive
            const query = `SELECT [CATALOG_NAME] FROM $SYSTEM.DBSCHEMA_CATALOGS WHERE [CATALOG_NAME] = '${this.database}'`;
            await this.executeDMVQuery(query);
            this.isConnected = true;
            this.lastSuccessfulQuery = new Date();
            return true;
        } catch (error) {
            this.isConnected = false;
            console.warn('[XMLA] Connection test failed:', error.message);
            return false;
        }
    }

    // Get connection status
    getConnectionStatus() {
        return {
            isConnected: this.isConnected,
            lastSuccessfulQuery: this.lastSuccessfulQuery,
            server: this.server,
            database: this.database
        };
    }

    // Execute a DMV (Dynamic Management View) query
    async executeDMVQuery(query) {
        const soapRequest = this.buildExecuteRequest(query);
        const result = await this.sendSOAPRequest(soapRequest);

        // Mark connection as successful
        this.isConnected = true;
        this.lastSuccessfulQuery = new Date();

        return result;
    }

    // Build SOAP Execute request for DMV queries
    buildExecuteRequest(query) {
        return `<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <soap:Body>
    <Execute xmlns="urn:schemas-microsoft-com:xml-analysis">
      <Command>
        <Statement>${this.escapeXml(query)}</Statement>
      </Command>
      <Properties>
        <PropertyList>
          <Catalog>${this.escapeXml(this.database)}</Catalog>
          <Format>Tabular</Format>
        </PropertyList>
      </Properties>
    </Execute>
  </soap:Body>
</soap:Envelope>`;
    }

    // Build SOAP Discover request
    buildDiscoverRequest(requestType, restrictions = {}) {
        let restrictionList = '';
        for (const [key, value] of Object.entries(restrictions)) {
            restrictionList += `<${key}>${this.escapeXml(value)}</${key}>`;
        }

        return `<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <soap:Body>
    <Discover xmlns="urn:schemas-microsoft-com:xml-analysis">
      <RequestType>${requestType}</RequestType>
      <Restrictions>
        <RestrictionList>
          ${restrictionList}
        </RestrictionList>
      </Restrictions>
      <Properties>
        <PropertyList>
          <Catalog>${this.escapeXml(this.database)}</Catalog>
        </PropertyList>
      </Properties>
    </Discover>
  </soap:Body>
</soap:Envelope>`;
    }

    // Send SOAP request to XMLA endpoint
    async sendSOAPRequest(soapBody) {
        try {
            // Check if we're in Electron environment with IPC available
            if (window.electronAPI && window.electronAPI.xmlaRequest) {
                // Use IPC to send request via main process (avoids CORS issues)
                const result = await window.electronAPI.xmlaRequest(this.xmlaEndpoint, soapBody);

                if (result.success) {
                    return this.parseXMLResponse(result.data);
                } else {
                    throw new Error('XMLA request failed');
                }
            } else {
                // Fallback to fetch (for browser mode, though it may have CORS issues)
                const response = await fetch(this.xmlaEndpoint, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'text/xml',
                        'SOAPAction': '"urn:schemas-microsoft-com:xml-analysis:Execute"'
                    },
                    body: soapBody
                });

                if (!response.ok) {
                    throw new Error(`XMLA request failed: ${response.status} ${response.statusText}`);
                }

                const xmlText = await response.text();
                return this.parseXMLResponse(xmlText);
            }
        } catch (error) {
            console.error('XMLA request error:', error);
            throw error;
        }
    }

    // Parse XML response to JSON
    parseXMLResponse(xmlText) {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText, 'text/xml');

        // Check for SOAP faults
        const fault = xmlDoc.querySelector('Fault');
        if (fault) {
            const faultString = fault.querySelector('faultstring')?.textContent || 'Unknown SOAP fault';
            throw new Error(`SOAP Fault: ${faultString}`);
        }

        // Parse result rows
        const rows = xmlDoc.querySelectorAll('row');
        const result = [];

        rows.forEach(row => {
            const rowData = {};
            Array.from(row.children).forEach(child => {
                rowData[child.tagName] = child.textContent;
            });
            result.push(rowData);
        });

        return result;
    }

    // Escape XML special characters
    escapeXml(unsafe) {
        // For browser environments, load the utility function
        if (typeof escapeXml === 'undefined') {
            if (!unsafe) return '';
            return unsafe.toString()
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&apos;');
        }
        // Use shared utility if available
        return window.utils ? window.utils.escapeXml(unsafe) : this._escapeXmlFallback(unsafe);
    }

    _escapeXmlFallback(unsafe) {
        if (!unsafe) return '';
        return unsafe.toString()
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    // Get all tables in the model
    async getTables() {
        const query = `
            SELECT *
            FROM $SYSTEM.TMSCHEMA_TABLES
        `;
        return await this.executeDMVQuery(query);
    }

    // Get all columns for all tables
    async getColumns() {
        const query = `
            SELECT *
            FROM $SYSTEM.TMSCHEMA_COLUMNS
        `;
        return await this.executeDMVQuery(query);
    }

    // Get all measures
    async getMeasures() {
        const query = `
            SELECT *
            FROM $SYSTEM.TMSCHEMA_MEASURES
        `;
        return await this.executeDMVQuery(query);
    }

    // Get all relationships
    async getRelationships() {
        try {
            // Try to fetch relationships - column names may vary by Power BI version
            const query = `
                SELECT
                    [Name] as RelationshipName,
                    [FromTableID],
                    [FromColumnID],
                    [ToTableID],
                    [ToColumnID]
                FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS
            `;
            return await this.executeDMVQuery(query);
        } catch (error) {
            console.warn('Failed to fetch relationships (may not be supported in this Power BI version):', error.message);
            return [];
        }
    }

    // Get sample data for a specific table
    async getSampleData(tableName, maxRows = 3) {
        try {
            const daxQuery = `TOPN(${maxRows}, '${tableName}')`;
            const result = await this.executeDAX(daxQuery);
            return result.data || result;
        } catch (error) {
            console.warn(`Failed to fetch sample data for table '${tableName}':`, error.message);
            return [];
        }
    }

    // Execute DAX query and return results
    async executeDAX(daxQuery) {
        let finalQuery = daxQuery;
        let warnings = [];

        // Validate the query if validator is available
        if (this.daxValidator) {
            const validation = this.daxValidator.validate(daxQuery);

            if (!validation.isValid) {
                throw new Error(`Query validation failed: ${validation.errors.join(', ')}`);
            }

            warnings = validation.warnings;
            finalQuery = validation.modified ? validation.modifiedQuery : daxQuery;

            // Sanitize the query
            finalQuery = this.daxValidator.sanitize(finalQuery);
        }

        // Add EVALUATE if not present
        if (!/^\s*EVALUATE/i.test(finalQuery)) {
            finalQuery = `EVALUATE ${finalQuery}`;
        }

        const results = await this.executeDMVQuery(finalQuery);

        // Validate results if validator is available
        if (this.daxValidator) {
            const resultValidation = this.daxValidator.validateResults(results);
            warnings = [...warnings, ...resultValidation.warnings];
        }

        // Return results with any warnings
        return {
            data: results,
            warnings: warnings,
            rowCount: results.length
        };
    }


    // Get complete semantic model metadata
    async getSemanticModelMetadata(options = {}) {
        const fetchSampleData = options.fetchSampleData !== false; // Default to true
        const maxSampleRows = options.maxSampleRows || 3;

        try {
            // Fetch all metadata in parallel
            const [tablesResult, columnsResult, measuresResult, relationshipsResult] = await Promise.all([
                this.getTables().catch(() => []),
                this.getColumns().catch(() => []),
                this.getMeasures().catch(() => []),
                this.getRelationships().catch(() => [])
            ]);

            // First, create maps of IDs to names
            const tableIdToName = {};
            tablesResult.forEach(table => {
                const tableId = table.ID;
                const tableName = table.Name || table.ExplicitName || table.InferredName;
                if (tableId && tableName) {
                    tableIdToName[tableId] = tableName;
                }
            });

            // Create a map of ColumnID -> ColumnName
            const columnIdToName = {};
            columnsResult.forEach(col => {
                const columnId = col.ID;
                const columnName = col.ExplicitName || col.InferredName;
                if (columnId && columnName) {
                    columnIdToName[columnId] = columnName;
                }
            });

            // Group columns by table using TableID
            const columnsByTable = {};

            columnsResult.forEach(col => {
                const tableId = col.TableID;
                const tableName = tableIdToName[tableId];
                const columnName = col.ExplicitName || col.InferredName;

                if (tableName && columnName) {
                    if (!columnsByTable[tableName]) {
                        columnsByTable[tableName] = [];
                    }
                    columnsByTable[tableName].push({
                        name: columnName
                    });
                }
            });

            // Build tables array
            const tables = tablesResult.map(table => {
                const tableName = table.Name || table.ExplicitName || table.InferredName;
                return {
                    name: tableName,
                    columns: columnsByTable[tableName] || []
                };
            });

            // Build measures array
            const measures = measuresResult.map(measure => ({
                name: measure.Name || measure.ExplicitName || measure.InferredName,
                table: tableIdToName[measure.TableID]
            }));

            // Build relationships array
            const relationships = relationshipsResult.map(rel => {
                const fromTable = tableIdToName[rel.FromTableID] || 'Unknown';
                const toTable = tableIdToName[rel.ToTableID] || 'Unknown';
                const fromColumn = columnIdToName[rel.FromColumnID] || 'Unknown';
                const toColumn = columnIdToName[rel.ToColumnID] || 'Unknown';

                return {
                    name: rel.RelationshipName || '',
                    from: `${fromTable}[${fromColumn}]`,
                    to: `${toTable}[${toColumn}]`
                };
            }).filter(rel => !rel.from.includes('Unknown') && !rel.to.includes('Unknown'));

            // Fetch sample data for each table if requested
            const sampleData = {};
            if (fetchSampleData && tables.length > 0) {
                console.log(`[XMLA] Fetching sample data (${maxSampleRows} rows per table)...`);

                // Limit to first 10 tables to avoid overwhelming the system
                // Skip tables with no columns as they cannot be queried
                const tablesToSample = tables
                    .filter(table => table.columns && table.columns.length > 0)
                    .slice(0, 10);

                const sampleDataPromises = tablesToSample.map(async (table) => {
                    try {
                        const samples = await this.getSampleData(table.name, maxSampleRows);
                        return { tableName: table.name, samples };
                    } catch (error) {
                        console.warn(`Failed to fetch sample data for ${table.name}:`, error.message);
                        return { tableName: table.name, samples: [] };
                    }
                });

                const sampleResults = await Promise.all(sampleDataPromises);
                sampleResults.forEach(({ tableName, samples }) => {
                    if (samples.length > 0) {
                        sampleData[tableName] = samples;
                    }
                });

                console.log(`[XMLA] Sample data fetched for ${Object.keys(sampleData).length} tables`);
            }

            return {
                tables,
                measures,
                relationships,
                sampleData,
                summary: {
                    tableCount: tables.length,
                    measureCount: measures.length,
                    relationshipCount: relationships.length,
                    totalColumns: columnsResult.length,
                    tablesWithSampleData: Object.keys(sampleData).length
                }
            };
        } catch (error) {
            console.error('Failed to fetch semantic model metadata:', error);
            throw error;
        }
    }
}

// Export for use in renderer
if (typeof module !== 'undefined' && module.exports) {
    module.exports = XMLAConnection;
}
