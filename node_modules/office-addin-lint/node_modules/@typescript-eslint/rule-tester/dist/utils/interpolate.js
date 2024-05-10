"use strict";
// Forked from https://github.com/eslint/eslint/blob/ad9dd6a933fd098a0d99c6a9aa059850535c23ee/lib/linter/interpolate.js
Object.defineProperty(exports, "__esModule", { value: true });
exports.interpolate = void 0;
function interpolate(text, data) {
    if (!data) {
        return text;
    }
    // Substitution content for any {{ }} markers.
    return text.replace(/\{\{([^{}]+?)\}\}/gu, (fullMatch, termWithWhitespace) => {
        const term = termWithWhitespace.trim();
        if (term in data) {
            return String(data[term]);
        }
        // Preserve old behavior: If parameter name not provided, don't replace it.
        return fullMatch;
    });
}
exports.interpolate = interpolate;
//# sourceMappingURL=interpolate.js.map