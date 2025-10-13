/**
 * DAX Query Validator
 * Provides safety checks for DAX queries before execution
 */

class DAXValidator {
    constructor(config) {
        this.config = config || {};
        this.maxRows = this.config.MAX_DAX_RESULT_ROWS || 10000;
        this.warningThreshold = this.config.DAX_WARNING_THRESHOLD || 1000;
        this.blockedPatterns = this.config.BLOCKED_PATTERNS || [];
        this.dangerousFunctions = this.config.DANGEROUS_FUNCTIONS || [];
    }

    /**
     * Validates a DAX query for safety and performance concerns
     * @param {string} daxQuery - The DAX query to validate
     * @returns {Object} Validation result {isValid, warnings, errors}
     */
    validate(daxQuery) {
        const result = {
            isValid: true,
            warnings: [],
            errors: [],
            modified: false,
            modifiedQuery: daxQuery
        };

        if (!daxQuery || typeof daxQuery !== 'string') {
            result.isValid = false;
            result.errors.push('Query is empty or invalid');
            return result;
        }

        const trimmedQuery = daxQuery.trim();

        // Check for empty query
        if (trimmedQuery.length === 0) {
            result.isValid = false;
            result.errors.push('Query cannot be empty');
            return result;
        }

        // Check for dangerous patterns
        for (const pattern of this.blockedPatterns) {
            if (pattern.test(trimmedQuery)) {
                result.isValid = false;
                result.errors.push(`Query contains blocked pattern: ${pattern.source}. This may cause performance issues.`);
            }
        }

        // Check for nested dangerous functions
        for (const func of this.dangerousFunctions) {
            const regex = new RegExp(`${func}\\s*\\([^)]*${func}`, 'gi');
            if (regex.test(trimmedQuery)) {
                result.warnings.push(`Nested ${func} function detected. This may be slow on large datasets.`);
            }
        }

        // Check if query already has EVALUATE
        const hasEvaluate = /^\s*EVALUATE/i.test(trimmedQuery);

        // Check if query has TOP N clause
        const hasTopN = /TOPN\s*\(/i.test(trimmedQuery) || /TOP\s+\d+/i.test(trimmedQuery);

        // If no TOP N clause and no EVALUATE, suggest adding row limit
        if (!hasTopN && !hasEvaluate) {
            result.warnings.push(`Query does not have a row limit. Consider adding TOPN(${this.warningThreshold}, ...) to limit results.`);
        }

        // Check for SELECT statements (not valid in DAX)
        if (/^\s*SELECT/i.test(trimmedQuery)) {
            result.isValid = false;
            result.errors.push('Invalid syntax: DAX queries use EVALUATE, not SELECT. This tool automatically adds EVALUATE for you.');
        }

        // Check for common SQL keywords that don't belong in DAX
        const sqlKeywords = ['FROM', 'WHERE', 'JOIN', 'GROUP BY', 'HAVING'];
        for (const keyword of sqlKeywords) {
            const regex = new RegExp(`\\b${keyword}\\b`, 'i');
            if (regex.test(trimmedQuery) && !trimmedQuery.toUpperCase().includes('FILTER')) {
                result.warnings.push(`Found SQL keyword '${keyword}'. Make sure you're using DAX syntax, not SQL.`);
            }
        }

        // Auto-add row limit if query seems safe but unlimited
        if (!hasTopN && !hasEvaluate && result.isValid && !trimmedQuery.toUpperCase().includes('$SYSTEM')) {
            // Wrap in TOPN for safety
            result.modified = true;
            result.modifiedQuery = `TOPN(${this.maxRows}, ${trimmedQuery})`;
            result.warnings.push(`Automatically limited to ${this.maxRows} rows for safety.`);
        }

        return result;
    }

    /**
     * Validates DAX query results for safety
     * @param {Array} results - The query results
     * @returns {Object} Validation result {isValid, warnings}
     */
    validateResults(results) {
        const validation = {
            isValid: true,
            warnings: []
        };

        if (!Array.isArray(results)) {
            return validation;
        }

        if (results.length > this.warningThreshold) {
            validation.warnings.push(`Query returned ${results.length} rows. Large result sets may impact performance.`);
        }

        if (results.length >= this.maxRows) {
            validation.warnings.push(`Result set was truncated at ${this.maxRows} rows. Use TOPN() or FILTER to reduce results.`);
        }

        return validation;
    }

    /**
     * Sanitizes a DAX query for safe execution
     * @param {string} daxQuery - The DAX query to sanitize
     * @returns {string} Sanitized query
     */
    sanitize(daxQuery) {
        if (!daxQuery) return '';

        // Remove potential injection attempts
        let sanitized = daxQuery.trim();

        // Remove multiple semicolons (prevents query batching)
        sanitized = sanitized.replace(/;+/g, ';');

        // Remove trailing semicolons
        sanitized = sanitized.replace(/;+\s*$/g, '');

        return sanitized;
    }
}

// Export for both Node.js and browser environments
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DAXValidator;
}
