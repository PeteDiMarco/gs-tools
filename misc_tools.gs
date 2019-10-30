/**
 * Copyright (c) 2019 Pete DiMarco
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */


//****************************************************************************
// Utility Functions:
// These functions compensate for features missing from Google Apps Script.
// Functions with names that end with "_" cannot be called from Google Sheets.
//****************************************************************************

/**
 * Assert "macro".  Calls Logger.log and then throws error message.  Usage: <br/>
 * <code>
 *     _assert_(function(){return   <i>YOUR_EXPRESSION_HERE</i>   ;});
 * </code>
 * @param {function} fn - Function containing our assertion.  Must return boolean.
 * @throws {!string}
 */
function _assert_(fn) {
  // Call fn without "this" or any parameters.
  if (!fn.apply(null, [])) {
    // If assertion fails, throw an exception in order to generate a stack trace.
    try {
      throw new Error("Get stack");
    } catch (err) {
      var stack_line = err.stack.split('\n')[1];   // Get line #1 from stack trace.
      var expr = /return *([^;]*) *;/.exec(fn.toString())[1];   // Extract the expression that failed from fn.
      var msg = "Assert(" + expr + ") failed " + stack_line.trim();
      Logger.log(msg);
      throw new Error(msg);
    }
  }
}


/**
 * Converts a string used to identify a spreadsheet column into a number,
 * similar to the COLUMN() function.
 * @param {!string} a1 - Just the letter portion of A1 notation.
 * @return {!number}
 * @throws {!string}
 */
function _a1_to_column_(a1) {
  if (typeof a1 != "string" || a1.length < 1) {
    throw new Error("Bad parameter: " + a1);
  }
  a1 = a1.toUpperCase().trim();
  
  var total = 0;
  var power = a1.length - 1;
  var ascii_offset = "A".charCodeAt(0) - 1;	// Subtracting ascii_offset converts ASCII "A" to 1.
  var ascii_max = "Z".charCodeAt(0);
  var ascii_code = 0;

  // Scan every character in a1 from left to right:
  for (var i = 0; i < a1.length; ++i) {
    ascii_code = a1.charCodeAt(i);
    if (ascii_code > ascii_max || ascii_code <= ascii_offset) {
      throw new Error("Illegal column letter: " + a1[i]);
    }
    total += (ascii_code - ascii_offset) * Math.pow(26, power);
    --power;
  }
  return total;
}


/**
 * Returns true if value is: true, 1, "1", or "true" (ignores case and
 * whitespace).  Returns false otherwise.
 * @param {*} value - Any type.
 * @return {!boolean}
 */
function _to_boolean_(value) {
  if (value === false || value === null || value === undefined || value === NaN || value == 0) {
    return false;
  } else if (value === true || value == 1) {
    return true;
  }
  
  var num;
  if (typeof value == "string") {
    value = value.trim();
    num = parseInt(value);
    return value.toUpperCase() == "TRUE" || num == 1;
  }
  
  return false;
}


/**
 * Takes nested arrays and returns a single, flattened array.
 * @param {Array|*} array - Arbitrarily nested array or array element.
 * @param {Array} [result=[]]
 * @return {!Array}
 */
function _flatten_(array, result) {
  if (result === undefined) {
    result = [];
  }
  if (Array.isArray(array)) {
    for (var i = 0; i < array.length; ++i) {
      _flatten_(array[i], result);   // result accumulates all elements in nested array.
    }
  } else {
    result.push(array);
  }
  return result;
}


/**
 * Returns a list of all the unique elements in array.  Ignores null, undefined,
 * and empty strings.
 * @param {Array} array
 * @return {!Array}
 */
function _unique_(array) {
  var map = {};
  var value;
  for (var i = 0; i < array.length; ++i) {
    value = array[i];
    if (value !== null && value !== undefined && (typeof value != "string" || value.trim().length > 0)) {
      map[value] = 1;
    }
  }
  return Object.keys(map);
}


/**
 * Parses a string into a list of substrings denoted by separator.  Trims
 * whitespace and ignores blank strings.
 * @param {string} str - String with 0 or more occurrences of separator.
 * @param {string} separator
 * @return {string[]}
 */
function _parse_list_string_(str, separator) {
  if (typeof str == "string") {
    return str.split(separator).map( function (s) { return s.trim(); } ).filter( function (s) { return s.length > 0; });
  } else {
    return [];
  }
}


/**
 * Finds the next blank row in a matrix by checking column "offset".
 * @param {string[][]} matrix - 2D array of strings.
 * @param {number} offset - Column number (counts from 0).
 * @param {string} [start_at=0] - Start scanning from this row.
 * @throws {string}
 * @return {number} Next row.
 */
function _next_blank_row_(matrix, offset, start_at) {
  _assert_( function(){return matrix instanceof Array ;} );
  _assert_( function(){return typeof offset == "number" ;} );
  if (start_at === undefined) {
    start_at = 0;
  }
  
  var row;
  for (row = start_at; row < matrix.length && !IS_EMPTY(matrix[row][offset]); ++row) {}
  return row;
}


//****************************************************************************
// Formula Functions:
// These functions can be used in Google Sheets formulas... but don't call
// them too often, or Google will rate-limit them!
//****************************************************************************

/**
 * Returns true if value is null, undefined, or an empty string.  Returns false
 * otherwise.  Similar to ISBLANK() function.
 * @param {null|?string} value
 * @return {!boolean}
 */
function IS_EMPTY(value) {
  //Utilities.sleep(500);
  if (value === null || value === undefined || (typeof value == "string" && 0 === value.trim().length)) {
    // Return true if value is null, undefined, "", or /^ *$/:
    return true;
  } else {
    // Return false if value is true, false, 0, [], {}, or anything else:
    return false;
  }
}


/**
 * Returns an array of num elements initialized to value.
 * @param {*} value
 * @param {!number} num
 * @return {!Array}
 */
function ARRAY(value, num) {
  if (typeof num != "number" || num < 0) {
    throw new Error("Bad value for num (" + num + ")");
  }
  var retval = [];
  for (var i = 0; i < num; ++i) {
    retval.push(value);
  }
  return retval;
}


/**
 * Takes a matrix and returns a new matrix with only those rows in which
 * column's contents equal value.
 * If value is empty then return the whole matrix.
 * @param {string[][]} matrix - 2 dimensional array of spreadsheet data.
 * @param {!number} column - Number of column to check against value.  Counts from 1.
 * @param {string} value - String we are looking for.
 * @return {string[][]}
 */
function MAT_FILTER_ROWS(matrix, column, value) {
  var results = [];
  _assert_( function(){return matrix instanceof Array ;} );
  _assert_( function(){return typeof column == "number" && column > 0 ;} );
  _assert_( function(){return typeof value == "string" ;} );
  //Utilities.sleep(700);
  if (IS_EMPTY(value)) {
    return matrix;
  }
  
  for (var row = 0; row < matrix.length; ++row) {
    if (matrix[row][column-1] == value) {         // column counts from 1.
      results.push(matrix[row]);
    }
  }
  return results;
}


/**
 * Takes a matrix and returns a new matrix with only those rows and columns
 * specified in row_list and col_list.
 * @param {string[][]} matrix - 2 dimensional array of spreadsheet data.
 * @param {number|number[][]} row_list - Array of row numbers we want in new matrix.  Counts from 1.
 * @param {number|number[][]} col_list - Array of column numbers we want in new matrix.  Counts from 1.
 * @param {boolean} ignore_undef - If matrix has a "row" that is actually undefined, then create a row of undefined values for the return value.
 * @return {string[][]}
 */
function MAT_SLICE(matrix, row_list, col_list, ignore_undef) {
  var result = [];
  var new_row;
  var rows;
  var columns;

  if (typeof row_list == "number") {
    rows = [ row_list ];            // If the user passes {1} in Google Sheets, this is treated as a number!
  } else {
    rows = row_list[0];             // If the user passes {1,2} in Google Sheets, this is treated as a 2 dimensional array!
  }

  if (typeof col_list == "number") {
    columns = [ col_list ];            // If the user passes {1} in Google Sheets, this is treated as a number!
  } else {
    columns = col_list[0];             // If the user passes {1,2} in Google Sheets, this is treated as a 2 dimensional array!
  }
  
  for (var row = 0; row < rows.length; ++row) {
    if (ignore_undef && matrix[rows[row] - 1] === undefined) {
      new_row = ARRAY(undefined, columns.length);    // Create an array of undefined values to pad the matrix.
    } else {
      new_row = new Array();
      for (var col = 0; col < columns.length; ++col) {
        new_row.push(matrix[rows[row] - 1][columns[col] - 1]);      // columns' contents count from 1.
      }
    }
    result.push(new_row);
  }
  return result;
}


/**
 * Takes a matrix and returns a new matrix with only those columns specified
 * in col_list.
 * @param {string[][]} matrix - 2 dimensional array of spreadsheet data.
 * @param {number|number[][]} col_list - Array of column numbers we want in new matrix.  Counts from 1.
 * @return {string[][]}
 */
function MAT_SLICE_COLS(matrix, col_list) {
  var result = [];
  var new_row;
  var columns;
  
  if (typeof col_list == "number") {
    columns = [ col_list ];            // If the user passes {1} in Google Sheets, this is treated as a number!
  } else {
    columns = col_list[0];             // If the user passes {1,2} in Google Sheets, this is treated as a 2 dimensional array!
  }
  
  for (var row = 0; row < matrix.length; ++row) {
    new_row = new Array();
    for (var col = 0; col < columns.length; ++col) {
      new_row.push(matrix[row][columns[col] - 1]);      // columns' contents count from 1.
    }
    result.push(new_row);
  }
  return result;
}


/**
 * Takes a matrix and returns a new matrix with only those rows specified in
 * row_list.
 * @param {string[][]} matrix - 2 dimensional array of spreadsheet data.
 * @param {number|number[][]} row_list - Array of row numbers we want in new matrix.  Counts from 1.
 * @return {string[][]}
 */
function MAT_SLICE_ROWS(matrix, row_list) {
  var result = [];
  var rows;
  
  if (typeof row_list == "number") {
    rows = [ row_list ];            // If the user passes {1} in Google Sheets, this is treated as a number!
  } else {
    rows = row_list[0];             // If the user passes {1,2} in Google Sheets, this is treated as a 2 dimensional array!
  }
  
  for (var row = 0; row < rows.length; ++row) {
    result.push(matrix[rows[row] - 1]);
  }
  return result;
}


/**
 * Returns the content of the matrix cell at row and column.
 * @param {string[][]} matrix - 2 dimensional array of spreadsheet data.
 * @param {!number} row - Row number counts from 1.
 * @param {!number} column - Column number counts from 1.
 * @return {string}
 */
function MAT_INDEX(matrix, row, column) {
  return matrix[row-1][column-1];
}


/**
 * Scans matrix and returns the number of rows and columns actually used.  The
 * number of rows is determined by finding the last row that doesn't contain
 * undefined or an array of undefined.  The number of columns is determined by
 * finding the right-most column that doesn't contain an undefined value.
 * @param {string[][]} matrix - 2 dimensional array of spreadsheet data.
 * @return {number[]} - Number of rows followed by number of columns.
 */
function MAT_SIZE(matrix) {
  var max_cols = 0;
  var max_rows = 0;
  for (var row = 0; row < matrix.length; ++row) {
    if (matrix[row] !== undefined && matrix[row].length != 0) {
      for (var col = 0; col < matrix[row].length; ++col) {
        if (!IS_EMPTY(matrix[row][col])) {
          max_rows = row;
          if (max_cols < col) {
            max_cols = col;
          }
        }
      }
    }
  }
  return [ max_rows+1, max_cols+1 ];
}


/**
 * Resizes and returns matrix.  If called from another Javascript function,
 * the matrix parameter will be modified in place.  If rows or columns are
 * added, those cells are set to the value of pad.
 * @param {string[][]} matrix - 2 dimensional array of spreadsheet data.  Modified.
 * @param {number} new_row_num
 * @param {number} new_col_num
 * @param {*} pad - Any type.
 * @return {string[][]} - 2 dimensional array of spreadsheet data.
 */
function MAT_RESIZE(matrix, new_row_num, new_col_num, pad) {
  var old_row_num = matrix.length;
  var old_col_num;
  
  // If we need to add rows:
  if (matrix.length < new_row_num) {
    for (var row = matrix.length; row < new_row_num; ++row) {
      matrix[row] = ARRAY(pad, new_col_num);      // New rows have the new column width.
    }
  }
  matrix.length = new_row_num;        // Forces Javascript to redefine the length of matrix.  Excess rows are deleted.
  var min_row_num = Math.min(old_row_num, new_row_num);     // This is the number of rows that may have an incorrect # of columns.

  for (var row = 0; row < min_row_num; ++row) {
    old_col_num = matrix[row].length;
    // If we need additional elements in this row:
    if (old_col_num < new_col_num) {
      matrix[row].concat(ARRAY(pad, new_col_num - old_col_num));
    }
    matrix[row].length = new_col_num;   // Force the row to be the correct length.
  }
  
  return matrix;
}


/**
 * Joins 2 matrices either vertically or horizontally.  As a side effect,
 * matrix1 and matrix2 will be resized.  Returns a new matrix with enough rows
 * and columns to store all the data.
 * @param {string[][]} matrix1 - 2 dimensional array of spreadsheet data.  Modified.
 * @param {string[][]} matrix2 - 2 dimensional array of spreadsheet data.  Modified.
 * @param {boolean} [vertical=true] - If true, join the matrices vertically, else join them horizontally.
 * @return {string[][]} - 2 dimensional array of spreadsheet data.
 */
function MAT_JOIN(matrix1, matrix2, vertical) {
  if (vertical === undefined) {
    vertical = true;
  } else {
    vertical = _to_boolean_(vertical);
  }
  
  var result = [];
  var rows1;
  var cols1;
  var rows2;
  var cols2;
  [ rows1, cols1 ] = MAT_SIZE(matrix1);    // Number of rows and columns actually used in matrix1.
  [ rows2, cols2 ] = MAT_SIZE(matrix2);    // Number of rows and columns actually used in matrix2.
  var max_rows = Math.max(rows1, rows2);
  var max_cols = Math.max(cols1, cols2);
  
  if (vertical) {      // If joining vertically, the matrices must have the same number of columns.
    matrix1 = MAT_RESIZE(matrix1, rows1, max_cols, undefined);
    matrix2 = MAT_RESIZE(matrix2, rows2, max_cols, undefined);
    result = matrix1.concat(matrix2);
  } else {            // If joining horizontally, the matrices must have the same number of rows.
    matrix1 = MAT_RESIZE(matrix1, max_rows, cols1, undefined);
    matrix2 = MAT_RESIZE(matrix2, max_rows, cols2, undefined);
    for (var row = 0; row < max_rows; ++row) {
      result[row] = matrix1[row].concat(matrix2[row]);
    }
  }
  return result;
}


//****************************************************************************
// Class Constructor Functions:
//****************************************************************************

/**
 * Matches a valid JavaScript identifier.
 * @type {regex}
 */
var _id_patt_ = /^[A-Za-z_$][A-Za-z0-9_$]*$/;


/**
 * List of JavaScript reserved words from:
 *     https://docstore.mik.ua/orelly/webprog/jscript/ch02_08.htm
 * @type {string[]}
 */
var _reserved_words_ = [
  "break", "do", "if", "switch", "typeof", "case", "else", "in", "this", "var", 
  "catch", "false", "instanceof", "throw", "void", "continue", "finally", "new",
  "true", "while", "default", "for", "null", "try", "with", "delete", "function",
  "return", "abstract", "double", "goto", "native", "static", "boolean", "enum", 
  "implements", "package", "super", "byte", "export", "import", "private",
  "synchronized", "char", "extends", "int", "protected", "throws", "class",
  "final", "interface", "public", "transient", "const", "float", "long", "short",
  "volatile", "debugger", "arguments", "encodeURI", "Infinity", "Object",
  "String", "Array", "Error", "isFinite", "parseFloat", "SyntaxError", "Boolean",
  "escape", "isNaN", "parseInt", "TypeError", "Date", "eval", "Math",
  "RangeError", "undefined", "decodeURI", "EvalError", "NaN", "ReferenceError",
  "unescape", "decodeURIComponent", "Function", "Number", "RegExp", "URIError",
];


/**
 * Constructor function for Enum class.  Maps both identifiers to numbers and
 * numbers to identifiers (for easy look-up).
 * @param {string[]} list - Array of strings (valid JavaScript identifiers).
 * @param {number} [offset=0] - Starting number for enumeration.
 * @throws {!string}
 */
function Enum(list, offset) {
  if ( !(list instanceof Array) ) {
    throw new Error("Parameter is not an array.");
  }
  if (offset === undefined) {
    offset = 0;
  }
  
  // For every element in list:
  for (var i = 0; i < list.length; ++i) {
    // Verify string is a valid JavaScript identifier:
    if (!_id_patt_.test(list[i])) {
      throw new Error("Enum label \"" + list[i] + "\" is not a valid JavaScript identifier.");
    }
    // Verify string is not a reserved word:
    if (_reserved_words_.some(function (str) { return str == list[i]; })) {
      throw new Error("Enum label \"" + list[i] + "\" is a reserved word in JavaScript.");
    }
    // Check that string isn't supplied more than once:
    if (this[list[i]] !== undefined) {
      throw new Error("Enum label \"" + list[i] + "\" defined more than once.");
    }

    this[list[i]] = i + offset;         // Maps name to number.
    this[i + offset] = list[i];         // Maps number to name (for easy look-up).
  }
}

