/**
 * Utility Functions
 * Shared helper functions used across the application
 */

/**
 * Escapes XML special characters to prevent injection and ensure valid XML
 * @param {*} unsafe - The value to escape
 * @returns {string} The escaped XML-safe string
 */
function escapeXml(unsafe) {
    if (unsafe === null || unsafe === undefined || unsafe === '') return '';
    return unsafe.toString()
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * Unescapes XML entities back to their original characters
 * @param {string} escaped - The escaped XML string
 * @returns {string} The unescaped string
 */
function unescapeXml(escaped) {
    if (!escaped) return '';
    return escaped
        .replace(/&apos;/g, "'")
        .replace(/&quot;/g, '"')
        .replace(/&gt;/g, '>')
        .replace(/&lt;/g, '<')
        .replace(/&amp;/g, '&');
}

// Export for Node.js environments
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        escapeXml,
        unescapeXml
    };
}
