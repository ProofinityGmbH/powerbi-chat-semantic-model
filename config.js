/**
 * Application Configuration
 * Central location for all configuration constants
 */

const CONFIG = {
  // Window settings
  WINDOW: {
    WIDTH: 1200,
    HEIGHT: 800,
    TITLE: 'Power BI Chat - Semantic Model Assistant'
  },

  // Sidebar settings
  SIDEBAR: {
    WIDTH: 350
  },

  // Query settings
  QUERY: {
    COMMAND_TIMEOUT: 30, // seconds
    LOG_PREVIEW_LENGTH: 200, // characters
    DEFAULT_SAMPLE_ROWS: 3,
    MAX_SAMPLE_ROWS: 5,
    MAX_DAX_RESULT_ROWS: 10000, // Maximum rows to return from DAX queries
    DAX_WARNING_THRESHOLD: 1000, // Warn if query returns more than this
    BLOCKED_PATTERNS: [
      // Patterns that might cause performance issues
      /CROSSJOIN\s*\(\s*ALL/gi,
      /GENERATE\s*\(\s*GENERATE/gi,
      /ADDCOLUMNS\s*\(\s*ADDCOLUMNS/gi
    ],
    DANGEROUS_FUNCTIONS: [
      'UNION',
      'GENERATE',
      'NATURALLEFTOUTERJOIN',
      'NATURALINNERJOIN'
    ]
  },

  // API settings
  API: {
    DEFAULT_URL: 'https://api.openai.com/v1/chat/completions',
    DEFAULT_MODEL: 'gpt-4'
  },

  // .NET Bridge settings
  BRIDGE: {
    PATH: ['XmlaBridge', 'bin', 'Release', 'net48', 'XmlaBridge.dll'],
    TYPE_NAME: 'Startup',
    METHOD_NAME: 'Invoke'
  },

  // XMLA settings
  XMLA: {
    DEFAULT_PROTOCOL: 'http',
    ENDPOINT_PATH: '/xmla',
    SOAP_ACTION: '"urn:schemas-microsoft-com:xml-analysis:Execute"'
  },

  // Development settings
  DEV: {
    OPEN_DEVTOOLS: typeof process !== 'undefined' && process.env && process.env.NODE_ENV === 'development'
  }
};

// Export for both Node.js and browser environments
if (typeof module !== 'undefined' && module.exports) {
  module.exports = CONFIG;
}

// Also expose to window object for browser environments
if (typeof window !== 'undefined') {
  window.CONFIG = CONFIG;
}
