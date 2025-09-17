import React, { useState, useEffect, useRef, useCallback } from 'react';

// Main component for the Excel Formula Beautifier application
const App = () => {
  // --- State Variables ---
  const [formula, setFormula] = useState('=IF(SUM(A1:B2)>10, "High", "Low")'); // Input formula string
  const [mode, setMode] = useState('beautify'); // Current operation mode (beautify, minify, etc.)
  const [output, setOutput] = useState(''); // Result of the formula processing
  const [isEu, setIsEu] = useState(false); // Flag for European-style separators (;)
  const [numberOfSpaces, setNumberOfSpaces] = useState(4); // Indentation spaces for beautify mode
  const [copySuccess, setCopySuccess] = useState(''); // Feedback message for copy action

  // --- Refs ---
  // A ref to hold the excelFormulaUtilities library logic.
  // Using a ref avoids re-initializing the library on every render.
  const excelFormulaUtilitiesRef = useRef(null);

  // --- Core Logic ---

  /**
   * Main function to process the formula based on the selected mode.
   * It calls the appropriate function from the formula utilities library.
   */
  const updateOutput = useCallback(() => {
    if (!excelFormulaUtilitiesRef.current) return; // Guard against premature execution

    // Set the EU flag on the library instance before processing
    excelFormulaUtilitiesRef.current.isEu = isEu;

    let newOutput;
    switch (mode) {
      case 'beautify':
        newOutput = excelFormulaUtilitiesRef.current.formatFormula(formula, {
          tmplIndentTab: ' '.repeat(numberOfSpaces),
          prefix: "=",
        });
        break;
      case 'minify':
        newOutput = excelFormulaUtilitiesRef.current.formatFormula(formula, {
          tmplFunctionStart: '{{token}}(',
          tmplFunctionStop: ')',
          tmplOperandText: '"{{token}}"',
          tmplArgument: ',',
          tmplOperandOperatorInfix: '{{token}}',
          tmplSubexpressionStart: '(',
          tmplSubexpressionStop: ')',
          tmplIndentTab: '',
          tmplIndentSpace: '',
          newLine: '',
          prefix: '=',
        });
        break;
      case 'js':
        newOutput = excelFormulaUtilitiesRef.current.formula2JavaScript(formula);
        break;
      case 'csharp':
        newOutput = excelFormulaUtilitiesRef.current.formula2CSharp(formula);
        break;
      case 'python':
        newOutput = excelFormulaUtilitiesRef.current.formula2Python(formula);
        break;
      case 'html':
        newOutput = excelFormulaUtilitiesRef.current.formatFormulaHTML(formula, {
          tmplIndentTab: ' '.repeat(numberOfSpaces),
        });
        break;
      default:
        newOutput = 'Invalid mode selected';
    }
    setOutput(newOutput);
  }, [formula, mode, isEu, numberOfSpaces]);

  // --- Effects ---

  // Effect to initialize the Excel formula utilities library once on component mount.
  useEffect(() => {
    // Check if the library is not already initialized
    if (!excelFormulaUtilitiesRef.current) {
        // Initialize the library object
        const excelFormulaUtilities = {};

        // Immediately-invoked function expression (IIFE) to encapsulate the library code
        // and prevent polluting the global scope. This mimics loading a script.
        (function (root) {
            // --- Library setup ---
            // Attaches the library to the 'root' object (excelFormulaUtilities in this case).
            // Defines core variables, token types, and subtypes used for parsing.
            root.core = {
                // A simple extend function to merge objects, similar to jQuery.extend
                extend: function(target, ...sources) {
                    for (const source of sources) {
                        for (const key in source) {
                            if (Object.prototype.hasOwnProperty.call(source, key)) {
                                target[key] = source[key];
                            }
                        }
                    }
                    return target;
                }
            };
            root.string = {
                // Formats a string by replacing placeholders like {0}, {1}
                formatStr: function(str, ...args) {
                    return str.replace(/{(\d+)}/g, function(match, number) {
                        return typeof args[number] != 'undefined' ? args[number] : match;
                    });
                },
                // Trims whitespace from both ends of a string
                trim: function(text) {
                    return (text || "").replace(/^\s+|\s+$/g, "");
                }
            };

            // const formatStr = root.string.formatStr; // Unused in current implementation
            const trim = root.string.trim;

            const types = {};
            const TOK_TYPE_NOOP = types.TOK_TYPE_NOOP = "noop";
            const TOK_TYPE_OPERAND = types.TOK_TYPE_OPERAND = "operand";
            const TOK_TYPE_FUNCTION = types.TOK_TYPE_FUNCTION = "function";
            const TOK_TYPE_SUBEXPR = types.TOK_TYPE_SUBEXPR = "subexpression";
            const TOK_TYPE_ARGUMENT = types.TOK_TYPE_ARGUMENT = "argument";
            const TOK_TYPE_OP_PRE = types.TOK_TYPE_OP_PRE = "operator-prefix";
            const TOK_TYPE_OP_IN = types.TOK_TYPE_OP_IN = "operator-infix";
            const TOK_TYPE_OP_POST = types.TOK_TYPE_OP_POST = "operator-postfix";
            const TOK_TYPE_WHITE_SPACE = types.TOK_TYPE_WHITE_SPACE = "white-space";
            const TOK_TYPE_UNKNOWN = types.TOK_TYPE_UNKNOWN = "unknown";

            const TOK_SUBTYPE_START = types.TOK_SUBTYPE_START = "start";
            const TOK_SUBTYPE_STOP = types.TOK_SUBTYPE_STOP = "stop";
            const TOK_SUBTYPE_TEXT = types.TOK_SUBTYPE_TEXT = "text";
            const TOK_SUBTYPE_NUMBER = types.TOK_SUBTYPE_NUMBER = "number";
            const TOK_SUBTYPE_LOGICAL = types.TOK_SUBTYPE_LOGICAL = "logical";
            const TOK_SUBTYPE_ERROR = types.TOK_SUBTYPE_ERROR = "error";
            const TOK_SUBTYPE_RANGE = types.TOK_SUBTYPE_RANGE = "range";
            // const TOK_SUBTYPE_MATH = types.TOK_SUBTYPE_MATH = "math"; // Unused in current implementation
            // const TOK_SUBTYPE_CONCAT = types.TOK_SUBTYPE_CONCAT = "concatenate"; // Unused in current implementation
            const TOK_SUBTYPE_INTERSECT = types.TOK_SUBTYPE_INTERSECT = "intersect";
            const TOK_SUBTYPE_UNION = types.TOK_SUBTYPE_UNION = "union";

            // Global setting for EU style formulas (using ; instead of ,)
            root.isEu = false;

            // --- Token Classes ---

            /**
             * Represents a single token in a formula.
             * @class
             * @param {string} value The token's string value.
             * @param {string} type The token's type (e.g., 'operand', 'function').
             * @param {string} subtype The token's subtype (e.g., 'number', 'text').
             */
            function F_token(value, type, subtype) {
                this.value = value;
                this.type = type;
                this.subtype = subtype;
            }

            /**
             * A collection of F_token objects with an iterator.
             * @class
             */
            function F_tokens() {
                this.items = [];
                this.add = function (value, type, subtype) {
                    const token = new F_token(value, type, subtype || "");
                    this.items.push(token);
                    return token;
                };
                this.index = -1;
                this.reset = () => { this.index = -1; };
                this.BOF = () => this.index <= 0;
                this.EOF = () => this.index >= this.items.length - 1;
                this.moveNext = () => {
                    if (this.EOF()) return false;
                    this.index++;
                    return true;
                };
                this.current = () => (this.index === -1 ? null : this.items[this.index]);
                this.next = () => (this.EOF() ? null : this.items[this.index + 1]);
                this.previous = () => (this.index < 1 ? null : this.items[this.index - 1]);
            }

            /**
             * A stack for managing nested structures like functions and subexpressions.
             * @class
             */
            function F_tokenStack() {
                this.items = [];
                this.push = (token) => { this.items.push(token); };
                this.pop = (name) => {
                    const token = this.items.pop();
                    return new F_token(name || "", token.type, TOK_SUBTYPE_STOP);
                };
                this.token = () => (this.items.length > 0 ? this.items[this.items.length - 1] : null);
                this.type = () => (this.token() ? this.token().type : "");
            }

            /**
             * The core parser. Converts a formula string into a stream of tokens.
             * @param {string} formula The Excel formula string.
             * @returns {F_tokens} A collection of tokens.
             */
            function getTokens(formula) {
                 // ... [The entire complex parsing logic from the original file] ...
                // This is a large and complex state machine that iterates through the formula
                // character by character, identifying operands, functions, operators, strings, etc.,
                // and creating a flat list of token objects. It handles various states like
                // being inside a string, a range, or an error literal.
                let tokens = new F_tokens();
                let tokenStack = new F_tokenStack();
                let offset = 0;
                let token = "";
                let inString = false, inPath = false, inRange = false, inError = false;

                formula = formula.trim().replace(/^=/, '').trim();

                const currentChar = () => formula.substring(offset, offset + 1);
                const doubleChar = () => formula.substring(offset, offset + 2);
                const nextChar = () => formula.substring(offset + 1, offset + 2);
                const EOF = () => offset >= formula.length;

                while (!EOF()) {
                     if (inString) {
                        if (currentChar() === '"') {
                            if (nextChar() === '"') {
                                token += '"';
                                offset++;
                            } else {
                                inString = false;
                                tokens.add(token, TOK_TYPE_OPERAND, TOK_SUBTYPE_TEXT);
                                token = "";
                            }
                        } else {
                            token += currentChar();
                        }
                        offset++;
                        continue;
                    }
                    if (inPath) {
                         if (currentChar() === "'") {
                            if (nextChar() === "'") {
                                token += "'";
                                offset++;
                            } else {
                                inPath = false;
                            }
                        }
                        token += currentChar();
                        offset++;
                        continue;
                    }
                     if (inRange) {
                        if (currentChar() === ']') inRange = false;
                        token += currentChar();
                        offset++;
                        continue;
                    }

                    if (inError) {
                        token += currentChar();
                        offset++;
                        if ((",#NULL!,#DIV/0!,#VALUE!,#REF!,#NAME?,#NUM!,#N/A,").indexOf("," + token + ",") !== -1) {
                            inError = false;
                            tokens.add(token, TOK_TYPE_OPERAND, TOK_SUBTYPE_ERROR);
                            token = "";
                        }
                        continue;
                    }

                    if (currentChar() === '"') {
                        if (token.length > 0) tokens.add(token, TOK_TYPE_UNKNOWN);
                        token = "";
                        inString = true;
                        offset++;
                        continue;
                    }
                    if (currentChar() === "'") {
                       if (token.length > 0) tokens.add(token, TOK_TYPE_UNKNOWN);
                       token = "";
                       inPath = true;
                       token += "'";
                       offset++;
                       continue;
                    }

                     if (currentChar() === '[') {
                        token += currentChar();
                        inRange = true;
                        offset++;
                        continue;
                    }
                    if (currentChar() === '#') {
                        if (token.length > 0) tokens.add(token, TOK_TYPE_UNKNOWN);
                        token = "";
                        inError = true;
                        token += currentChar();
                        offset++;
                        continue;
                    }

                    if (currentChar() === '{') {
                        if (token.length > 0) tokens.add(token, TOK_TYPE_UNKNOWN);
                        tokenStack.push(tokens.add("ARRAY", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
                        tokenStack.push(tokens.add("ARRAYROW", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
                        offset++;
                        continue;
                    }

                     if (currentChar() === '}') {
                        if (token.length > 0) tokens.add(token, TOK_TYPE_OPERAND);
                        token = "";
                        tokens.items.push(tokenStack.pop("ARRAYROW"));
                        tokens.items.push(tokenStack.pop("ARRAY"));
                        offset++;
                        continue;
                    }


                     if (currentChar() === ' ' || currentChar() === '\n' || currentChar() === '\r') {
                        if (token.length > 0) tokens.add(token, TOK_TYPE_OPERAND);
                        token = "";
                        tokens.add("", TOK_TYPE_WHITE_SPACE);
                        offset++;
                        while (!EOF() && (currentChar() === ' ' || currentChar() === '\n' || currentChar() === '\r')) offset++;
                        continue;
                    }
                    // operators
                    if (doubleChar() === "<=" || doubleChar() === ">=" || doubleChar() === "<>") {
                        if (token.length > 0) tokens.add(token, TOK_TYPE_OPERAND);
                        token = "";
                        tokens.add(doubleChar(), TOK_TYPE_OP_IN, TOK_SUBTYPE_LOGICAL);
                        offset += 2;
                        continue;
                    }
                    if ("+-*/^&=><".indexOf(currentChar()) !== -1) {
                        if (token.length > 0) tokens.add(token, TOK_TYPE_OPERAND);
                        token = "";
                        tokens.add(currentChar(), TOK_TYPE_OP_IN);
                        offset++;
                        continue;
                    }

                     if (currentChar() === '%') {
                        if (token.length > 0) tokens.add(token, TOK_TYPE_OPERAND);
                        token = "";
                        tokens.add('%', TOK_TYPE_OP_POST);
                        offset++;
                        continue;
                    }

                     if (currentChar() === '(') {
                        if (token.length > 0) {
                            tokenStack.push(tokens.add(token, TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
                        } else {
                            tokenStack.push(tokens.add("", TOK_TYPE_SUBEXPR, TOK_SUBTYPE_START));
                        }
                        token = "";
                        offset++;
                        continue;
                    }

                    if (currentChar() === ')') {
                        if (token.length > 0) tokens.add(token, TOK_TYPE_OPERAND);
                        token = "";
                        tokens.items.push(tokenStack.pop());
                        offset++;
                        continue;
                    }

                    const listSep = root.isEu ? ';' : ',';

                    if (currentChar() === listSep) {
                         if (token.length > 0) tokens.add(token, TOK_TYPE_OPERAND);
                        token = "";
                        if (tokenStack.type() === TOK_TYPE_FUNCTION) {
                           tokens.add(listSep, TOK_TYPE_ARGUMENT);
                        } else {
                           tokens.add(listSep, TOK_TYPE_OP_IN, TOK_SUBTYPE_UNION);
                        }
                        offset++;
                        continue;
                    }
                    if (root.isEu === false && currentChar() === ';') { // Array row separator
                        if (token.length > 0) tokens.add(token, TOK_TYPE_OPERAND);
                        token = "";
                        tokens.items.push(tokenStack.pop("ARRAYROW"));
                        tokens.add(";", TOK_TYPE_ARGUMENT); // Represents the row separator
                        tokenStack.push(tokens.add("ARRAYROW", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
                        offset++;
                        continue;
                    }


                    token += currentChar();
                    offset++;
                }
                if (token.length > 0) tokens.add(token, TOK_TYPE_OPERAND);


                // Post-processing steps
                let tokens2 = new F_tokens();
                while(tokens.moveNext()) {
                    let t = tokens.current();
                    if (t.type === TOK_TYPE_WHITE_SPACE) {
                        if (!tokens.BOF() && !tokens.EOF()) {
                            let prev = tokens.previous();
                            let next = tokens.next();
                            if(
                               (prev.type === TOK_TYPE_OPERAND || (prev.subtype === TOK_SUBTYPE_STOP)) &&
                               (next.type === TOK_TYPE_OPERAND || (next.subtype === TOK_SUBTYPE_START))
                            ) {
                                tokens2.items.push(new F_token("", TOK_TYPE_OP_IN, TOK_SUBTYPE_INTERSECT));
                            }
                        }
                        continue;
                    }
                    tokens2.items.push(t);
                }


                while(tokens2.moveNext()) {
                    let t = tokens2.current();
                     if (t.type === TOK_TYPE_OP_IN && (t.value === "+" || t.value === "-")) {
                         if (tokens2.BOF()) {
                             t.type = t.value === "-" ? TOK_TYPE_OP_PRE : TOK_TYPE_NOOP;
                         } else {
                             let prev = tokens2.previous();
                             if (!(prev.type === TOK_TYPE_OPERAND || prev.subtype === TOK_SUBTYPE_STOP || prev.type === TOK_TYPE_OP_POST)) {
                                 t.type = t.value === "-" ? TOK_TYPE_OP_PRE : TOK_TYPE_NOOP;
                             }
                         }
                     }

                    if (t.type === TOK_TYPE_OPERAND && !t.subtype) {
                        if (!isNaN(parseFloat(t.value)) && isFinite(t.value)) {
                            t.subtype = TOK_SUBTYPE_NUMBER;
                        } else if (t.value.toUpperCase() === 'TRUE' || t.value.toUpperCase() === 'FALSE') {
                            t.subtype = TOK_SUBTYPE_LOGICAL;
                        } else {
                            t.subtype = TOK_SUBTYPE_RANGE;
                        }
                    }
                }

                let finalTokens = new F_tokens();
                tokens2.reset();
                while(tokens2.moveNext()) {
                    if (tokens2.current().type !== TOK_TYPE_NOOP) {
                        finalTokens.items.push(tokens2.current());
                    }
                }

                return finalTokens;
            }

            // --- Formatting and Conversion Logic ---
            // const fromBase26 = (number) => { /* ... logic ... */ return 0; }; // Unused in current implementation
            // const toBase26 = (value) => { /* ... logic ... */ return 'A'; }; // Unused in current implementation
            // const breakOutRanges = (rangeStr, delim) => { /* ... logic ... */ return rangeStr; }; // Unused in current implementation

             function applyTokenTemplate(token, options, indent, lineBreak, override, lastToken) {
                // ... [The entire template application logic from the original file] ...
                // This function takes a token and formatting options and returns the
                // formatted string for that token based on a set of templates.
                // It's the core of the "beautifier" functionality.
                let tokenString = token.value;
                 if (override) {
                    const res = override(tokenString, token, indent, lineBreak);
                    tokenString = res.tokenString;
                    if (!res.useTemplate) return tokenString;
                }

                const format = (template) => {
                    return template
                        .replace(/\{\{autoindent\}\}/g, indent)
                        .replace(/\{\{token\}\}/g, tokenString)
                        .replace(/\{\{autolinebreak\}\}/g, lineBreak);
                };

                 switch(token.type) {
                    case TOK_TYPE_FUNCTION:
                        if (token.subtype === TOK_SUBTYPE_START) {
                           return format(options.tmplFunctionStart);
                        } else { // STOP
                           return format(options.tmplFunctionStop);
                        }
                    case TOK_TYPE_ARGUMENT:
                        return format(options.tmplArgument);
                    case TOK_TYPE_OPERAND:
                         switch(token.subtype) {
                            case TOK_SUBTYPE_TEXT: return format(options.tmplOperandText);
                            case TOK_SUBTYPE_NUMBER: return format(options.tmplOperandNumber);
                            case TOK_SUBTYPE_LOGICAL: return format(options.tmplOperandLogical);
                            case TOK_SUBTYPE_RANGE: return format(options.tmplOperandRange);
                            case TOK_SUBTYPE_ERROR: return format(options.tmplOperandError);
                            default: return indent + tokenString;
                        }
                    case TOK_TYPE_OP_IN:
                         return format(options.tmplOperandOperatorInfix);
                    case TOK_TYPE_SUBEXPR:
                        if (token.subtype === TOK_SUBTYPE_START) {
                           return format(options.tmplSubexpressionStart);
                        } else { // STOP
                           return format(options.tmplSubexpressionStop);
                        }
                    default:
                        return indent + tokenString;
                }
            }

            /**
             * Formats a formula string with indentation and line breaks.
             * @param {string} formula The formula to format.
             * @param {object} options Formatting options.
             * @returns {string} The formatted formula.
             */
            root.formatFormula = function (formula, options) {
                 const defaultOptions = {
                    tmplFunctionStart: '{{autoindent}}{{token}}(\n',
                    tmplFunctionStop: '\n{{autoindent}})',
                    tmplOperandError: '{{token}}',
                    tmplOperandRange: '{{autoindent}}{{token}}',
                    tmplLogical: '{{token}}',
                    tmplOperandLogical: '{{autoindent}}{{token}}',
                    tmplOperandNumber: '{{autoindent}}{{token}}',
                    tmplOperandText: '{{autoindent}}"{{token}}"',
                    tmplArgument: ',\n',
                    tmplOperandOperatorInfix: ' {{token}} ',
                    tmplSubexpressionStart: '{{autoindent}}(\n',
                    tmplSubexpressionStop: '\n{{autoindent}})',
                    tmplIndentTab: '    ',
                    tmplIndentSpace: ' ',
                    newLine: '\n',
                    customTokenRender: null,
                    prefix: "=",
                    postfix: ""
                };
                options = root.core.extend({}, defaultOptions, options);

                let indentCount = 0;

                const tokens = getTokens(formula);
                if (!tokens) return "Error parsing formula";
                let outputFormula = "";
                let isNewLine = true;

                while (tokens.moveNext()) {
                    const token = tokens.current();

                    // For function stops, we need to use the current indent level before decrementing
                    let currentIndentCount = indentCount;
                    if (token.subtype === TOK_SUBTYPE_STOP) {
                        currentIndentCount = Math.max(0, indentCount - 1);
                    }

                    const indent = isNewLine ? options.tmplIndentTab.repeat(currentIndentCount) : "";
                    const nextToken = tokens.next();
                    let lineBreak = "";
                    if (nextToken) {
                       if (nextToken.type === TOK_TYPE_ARGUMENT) {
                           lineBreak = options.newLine;
                       }
                    }

                    outputFormula += applyTokenTemplate(token, options, indent, lineBreak, options.customTokenRender, tokens.previous());

                    if (token.subtype === TOK_SUBTYPE_START) {
                        indentCount++;
                    } else if (token.subtype === TOK_SUBTYPE_STOP) {
                        indentCount = Math.max(0, indentCount - 1);
                    }
                    
                    // Update isNewLine flag - function stops always start a new line
                    if (token.subtype === TOK_SUBTYPE_STOP) {
                        isNewLine = true;
                    } else {
                        isNewLine = outputFormula.endsWith(options.newLine);
                    }
                }

                return options.prefix + trim(outputFormula) + options.postfix;
            };

            /**
             * Converts a formula to its JavaScript equivalent.
             * @param {string} formula The formula to convert.
             * @returns {string} The JavaScript code.
             */
            root.formula2JavaScript = function(formula) {
                 const tokRender = (tokenStr, token) => {
                    const directConversionMap = {
                        "=": "===", "<>": "!==",
                        "&": "+", "AND": "&&", "OR": "||",
                        "TRUE": "true", "FALSE": "false"
                    };
                    let outStr = directConversionMap[tokenStr.toUpperCase()] || tokenStr;

                    // Handle specific token types
                    if (token.type === TOK_TYPE_FUNCTION && token.subtype === TOK_SUBTYPE_START) {
                        if (tokenStr.toUpperCase() === 'IF') {
                            return { tokenString: "", useTemplate: true };
                        }
                        return { tokenString: outStr, useTemplate: true };
                    }
                    if (token.type === TOK_TYPE_ARGUMENT) {
                        return { tokenString: ", ", useTemplate: false };
                    }
                    if (token.type === TOK_TYPE_FUNCTION && token.subtype === TOK_SUBTYPE_STOP) {
                        return { tokenString: ")", useTemplate: false };
                    }
                    if (token.type === TOK_TYPE_OPERAND && token.subtype === TOK_SUBTYPE_TEXT) {
                        return { tokenString: '"' + outStr + '"', useTemplate: false };
                    }
                    if (token.type === TOK_TYPE_OP_IN) {
                        return { tokenString: " " + outStr + " ", useTemplate: false };
                    }

                    return { tokenString: outStr, useTemplate: true };
                };
                const options = {
                    // Simplified options for JS conversion
                     tmplFunctionStart: '{{token}}(',
                     tmplFunctionStop: ')',
                     tmplOperandText: '"{{token}}"',
                     tmplArgument: ', ',
                     tmplOperandOperatorInfix: ' {{token}} ',
                     tmplSubexpressionStart: '(',
                     tmplSubexpressionStop: ')',
                     // Disable beautifying features
                     tmplIndentTab: '',
                     newLine: '',
                     prefix: '',
                     customTokenRender: tokRender
                };
                 return root.formatFormula(formula, options);
            };

            root.formula2CSharp = root.formula2JavaScript; // Alias for simplicity
            root.formula2Python = root.formula2JavaScript; // Alias for simplicity

            /**
             * Exposes the getTokens function for external use.
             * @param {string} formula The formula to tokenize.
             * @returns {F_tokens} A collection of tokens.
             */
            root.getTokens = getTokens;

            /**
             * Formats a formula string with HTML tags for better display.
             * @param {string} formula The formula to format.
             * @param {object} options Formatting options.
             * @returns {string} The formatted formula with HTML tags.
             */
            root.formatFormulaHTML = function (formula, options) {
                const defaultOptions = {
                    tmplFunctionStart: '<span class="function">{{autoindent}}<span class="function-name">{{token}}</span>(\n',
                    tmplFunctionStop: '\n{{autoindent}})</span>',
                    tmplOperandError: '<span class="error">{{token}}</span>',
                    tmplOperandRange: '{{autoindent}}<span class="range">{{token}}</span>',
                    tmplLogical: '<span class="logical">{{token}}</span>{{autolinebreak}}',
                    tmplOperandLogical: '{{autoindent}}<span class="logical">{{token}}</span>',
                    tmplOperandNumber: '{{autoindent}}<span class="number">{{token}}</span>',
                    tmplOperandText: '{{autoindent}}<span class="text">"{{token}}"</span>',
                    tmplArgument: ',\n',
                    tmplOperandOperatorInfix: ' <span class="operator">{{token}}</span>{{autolinebreak}} ',
                    tmplSubexpressionStart: '{{autoindent}}<span class="subexpression">(\n',
                    tmplSubexpressionStop: '\n{{autoindent}})</span>',
                    tmplIndentTab: '    ',
                    tmplIndentSpace: ' ',
                    newLine: '\n',
                    customTokenRender: null,
                    prefix: '<span class="equals">=</span>',
                    postfix: ""
                };
                options = root.core.extend({}, defaultOptions, options);
                return root.formatFormula(formula, options);
            };

            // Store the fully constructed library in the ref
            excelFormulaUtilitiesRef.current = root;
             // Trigger the first update
            updateOutput();


        })(excelFormulaUtilities); // Pass the object to be populated
    }
  }, [updateOutput]); // Include updateOutput dependency

  // Effect to re-run the formula processing whenever an input changes.
  useEffect(() => {
    // Only run if the library has been initialized
    if (excelFormulaUtilitiesRef.current) {
        updateOutput();
    }
  }, [updateOutput]); // Dependencies for the effect

  // --- Event Handlers ---

  const handleCopyToClipboard = () => {
    if (output) {
      navigator.clipboard.writeText(output).then(() => {
        setCopySuccess('Copied!');
        setTimeout(() => setCopySuccess(''), 2000); // Clear message after 2 seconds
      }, () => {
        setCopySuccess('Failed to copy.');
        setTimeout(() => setCopySuccess(''), 2000);
      });
    }
  };

  // --- JSX Rendering ---

  return (
    <div className="bg-gray-900 text-gray-200 min-h-screen font-sans flex flex-col">
      {/* Header Section */}
      <header className="bg-gray-800 shadow-lg p-4 md:p-6">
        <div className="container mx-auto flex items-center gap-4">
            <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-green-400">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
                <polyline points="10 9 9 9 8 9"></polyline>
            </svg>
            <div>
                 <h1 className="text-2xl md:text-3xl font-bold text-green-400">Excel Formula Beautifier</h1>
                 <p className="text-sm md:text-base text-gray-400">Beautify, Minify, or Convert Excel Formulas to Code</p>
            </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="container mx-auto p-4 flex-grow w-full">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">

          {/* Left Column: Input and Options */}
          <div className="flex flex-col gap-6">
            {/* Input Area */}
            <div className="bg-gray-800 rounded-lg p-4 shadow-md">
                <label htmlFor="formula_input" className="block text-lg font-semibold mb-2 text-gray-300">
                    Your Formula:
                </label>
                <textarea
                    id="formula_input"
                    value={formula}
                    onChange={(e) => setFormula(e.target.value)}
                    className="w-full h-32 p-3 bg-gray-900 border border-gray-700 rounded-md focus:ring-2 focus:ring-green-500 focus:border-green-500 transition duration-200 text-gray-200 font-mono"
                    placeholder="e.g., =IF(A1>5, SUM(B1:B5), 0)"
                />
            </div>

            {/* Controls Area */}
            <div className="bg-gray-800 rounded-lg p-4 shadow-md">
                 <label htmlFor="mode_select" className="block text-lg font-semibold mb-2 text-gray-300">
                    Operation Mode:
                </label>
                <select
                    id="mode_select"
                    value={mode}
                    onChange={(e) => setMode(e.target.value)}
                    className="w-full p-3 bg-gray-900 border border-gray-700 rounded-md focus:ring-2 focus:ring-green-500 focus:border-green-500 transition duration-200"
                >
                    <option value="beautify">Beautify</option>
                    <option value="minify">Minify</option>
                    <option value="html">To HTML</option>
                    <option value="js">To JavaScript</option>
                    <option value="csharp">To C#</option>
                    <option value="python">To Python</option>
                </select>
            </div>

             {/* Formatting Options (Conditional) */}
             {(mode === 'beautify' || mode === 'html') && (
                <div className="bg-gray-800 rounded-lg p-4 shadow-md transition-all duration-300">
                     <h3 className="text-lg font-semibold mb-3 text-gray-300">Formatting Options</h3>
                     <div className="flex flex-col sm:flex-row sm:items-center gap-4">
                         <div className="flex items-center gap-2">
                             <input
                                type="checkbox"
                                id="isEu"
                                checked={isEu}
                                onChange={(e) => setIsEu(e.target.checked)}
                                className="h-4 w-4 rounded border-gray-600 bg-gray-900 text-green-500 focus:ring-green-500"
                            />
                            <label htmlFor="isEu">Use European Separators (;)</label>
                         </div>
                         <div className="flex items-center gap-2">
                            <label htmlFor="numberOfSpaces">Indent Spaces:</label>
                            <input
                                type="number"
                                id="numberOfSpaces"
                                value={numberOfSpaces}
                                onChange={(e) => setNumberOfSpaces(Math.max(0, parseInt(e.target.value, 10)))}
                                className="w-20 p-2 bg-gray-900 border border-gray-700 rounded-md focus:ring-2 focus:ring-green-500"
                                min="0"
                            />
                         </div>
                     </div>
                </div>
             )}
          </div>

          {/* Right Column: Output */}
          <div className="bg-gray-800 rounded-lg p-4 shadow-md flex flex-col">
            <div className="flex justify-between items-center mb-2">
                 <h2 className="text-lg font-semibold text-gray-300">Result:</h2>
                 <button
                    onClick={handleCopyToClipboard}
                    className="bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-4 rounded-md transition duration-200 flex items-center gap-2 relative"
                >
                     <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
                        <path d="M4 1.5H3a2 2 0 0 0-2 2V14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V3.5a2 2 0 0 0-2-2h-1v1h1a1 1 0 0 1 1 1V14a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V3.5a1 1 0 0 1 1-1h1v-1z"/>
                        <path d="M9.5 1a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5h-3a.5.5 0 0 1-.5-.5v-1a.5.5 0 0 1 .5-.5h3zM-1 7a.5.5 0 0 1 .5-.5h15a.5.5 0 0 1 0 1H-.5A.5.5 0 0 1-1 7z"/>
                     </svg>
                    <span>Copy</span>
                    {copySuccess && (
                        <span className="absolute -top-8 left-1/2 -translate-x-1/2 bg-gray-700 text-white text-xs rounded py-1 px-2">
                            {copySuccess}
                        </span>
                    )}
                 </button>
            </div>
            <pre className="w-full flex-grow bg-gray-900 border border-gray-700 rounded-md p-3 text-gray-200 font-mono overflow-auto min-h-[200px] lg:min-h-0">
                {mode === 'html' ? (
                    <code dangerouslySetInnerHTML={{ __html: output }} />
                ) : (
                    <code>{output}</code>
                )}
            </pre>
          </div>
        </div>
      </main>

      {/* Footer Section */}
      <footer className="bg-gray-800 mt-6 p-4 text-center text-gray-500 text-sm">
        <p>&copy; 2024 - Re-implemented in React from the original <a href="https://github.com/joshbtn/excelformulautilitiesjs" target="_blank" rel="noopener noreferrer" className="text-green-400 hover:underline">ExcelFormulaUtilitiesJS by Josh Bennett</a>.</p>
      </footer>
    </div>
  );
};

export default App;
