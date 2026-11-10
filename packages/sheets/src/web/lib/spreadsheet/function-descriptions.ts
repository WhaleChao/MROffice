/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Spreadsheet function help text — the prose Excel shows around a function in
// the Insert Function dialog (fx / Shift+F3) and in the Function Arguments
// dialog.
//
// Microsoft Support documents both surfaces: "Insert Function in Excel" shows a
// function's syntax plus a one-line description under the function list, and
// "Function Arguments" shows a description of the whole function together with
// a description of whichever argument box currently has focus. This module is
// the data behind exactly those two strings: `summary` is the one-line
// description, and `args` supplies the per-argument line.
//
// PURE data, no imports.
//
// Authored English prose. These strings are NOT translated per locale — they
// are written once in English, exactly like the argument labels in
// function-signatures.ts that this module keys off (every key here must be a
// key there, and every `args` key must be one of that function's labels with
// the optional-argument square brackets stripped).
//
// Every line is a paraphrase of documented, observable behaviour, written from
// scratch. No help text is copied from Microsoft or any other vendor; the
// published documentation is read as a behavioural specification only.
// ---------------------------------------------------------------------------

export interface SpreadsheetFunctionDoc {
  /** One sentence describing what the function returns. Ends with a period. */
  readonly summary: string;
  /** Per-argument help, keyed by the bracket-stripped signature label. */
  readonly args?: Readonly<Record<string, string>>;
}

/**
 * Help text for the functions a normal user meets, keyed by upper-case name.
 * A function with no entry simply shows its syntax with no description, which
 * is what Excel does for a function it has no help page for.
 */
export const FUNCTION_DOCS: Readonly<Record<string, SpreadsheetFunctionDoc>> = {
  ABS: {
    summary: "Returns the absolute value of a number, that is, the number without its sign.",
    args: {
      number: "The real number whose absolute value you want.",
    },
  },
  ACCRINT: {
    summary: "Returns the interest accrued by a security that pays interest periodically.",
  },
  ACOS: {
    summary: "Returns the arccosine of a number, in radians from 0 to pi.",
  },
  ACOSH: {
    summary: "Returns the inverse hyperbolic cosine of a number.",
  },
  ADDRESS: {
    summary: "Builds a cell address as text from a row number and a column number.",
  },
  AGGREGATE: {
    summary:
      "Applies one of a list of aggregate calculations to a range, with the option to ignore hidden rows, errors or nested subtotals.",
    args: {
      function_num: "A number from 1 to 19 naming the calculation to apply, such as 1 for AVERAGE or 9 for SUM.",
      options: "A number from 0 to 7 saying which values to leave out, such as hidden rows, error values or nested subtotals.",
      ref1: "The range or array to aggregate, or the first argument of the chosen calculation.",
      ref2: "Further ranges to aggregate, or the remaining arguments of the chosen calculation.",
    },
  },
  AND: {
    summary: "Returns TRUE when every one of its arguments is TRUE, and FALSE otherwise.",
    args: {
      logical1: "The first condition to test. It may be a value, an expression or a reference that resolves to TRUE or FALSE.",
      logical2: "Further conditions to test, up to 255 in all. Every one must be TRUE for the result to be TRUE.",
    },
  },
  ARABIC: {
    summary: "Converts a Roman numeral written as text to the ordinary number it stands for.",
  },
  AREAS: {
    summary: "Returns how many separate ranges a reference is made of.",
  },
  ARRAYTOTEXT: {
    summary: "Returns an array as a single piece of text, either as a plain list or in a form you could type back into a formula.",
  },
  ASC: {
    summary: "Converts the full-width characters in a piece of text to their half-width equivalents.",
  },
  ASIN: {
    summary: "Returns the arcsine of a number, in radians from -pi/2 to pi/2.",
  },
  ASINH: {
    summary: "Returns the inverse hyperbolic sine of a number.",
  },
  ATAN: {
    summary: "Returns the arctangent of a number, in radians from -pi/2 to pi/2.",
  },
  ATAN2: {
    summary: "Returns the arctangent of the point given by an x and a y coordinate, in radians from -pi to pi.",
  },
  ATANH: {
    summary: "Returns the inverse hyperbolic tangent of a number.",
  },
  AVEDEV: {
    summary: "Returns the average of how far the values are from their own mean.",
  },
  AVERAGE: {
    summary: "Returns the arithmetic mean of its arguments, which may be numbers, ranges or arrays of numbers.",
    args: {
      number1: "The first number, range or reference to average.",
      number2: "Further numbers, ranges or references to average, up to 255 in all.",
    },
  },
  AVERAGEA: {
    summary: "Returns the mean of its arguments, counting text as 0 and logical values as 1 or 0.",
  },
  AVERAGEIF: {
    summary: "Returns the mean of the cells in a range that meet one condition.",
    args: {
      range: "The cells whose values are tested against the condition.",
      criteria: "The condition a cell must meet to be averaged, given as a number, an expression or text, such as 32, \">32\" or \"apples\".",
      average_range: "The cells actually averaged. If omitted, the cells in range are averaged.",
    },
  },
  AVERAGEIFS: {
    summary: "Returns the mean of the cells that meet every one of several conditions.",
    args: {
      average_range: "The cells to average once all the conditions are met.",
      criteria_range1: "The cells tested against the first condition.",
      criteria1: "The condition cells in criteria_range1 must meet, given as a number, an expression or text.",
      criteria_range2: "A further range of cells to test.",
      criteria2: "The condition that further range must meet. Only cells meeting every condition are averaged.",
    },
  },
  BASE: {
    summary: "Converts a number to text in the number base you name, optionally padded to a minimum length.",
  },
  "BETA.DIST": {
    summary: "Returns the beta probability distribution for a value.",
  },
  "BETA.INV": {
    summary: "Returns the value at which the cumulative beta distribution reaches a given probability.",
  },
  BIN2DEC: {
    summary: "Converts a binary number to a decimal number.",
  },
  BIN2HEX: {
    summary: "Converts a binary number to a hexadecimal number.",
  },
  "BINOM.DIST": {
    summary: "Returns the probability of a given number of successes in a fixed series of independent trials.",
  },
  "BINOM.INV": {
    summary: "Returns the smallest number of successes for which the cumulative binomial distribution reaches a given probability.",
  },
  BITAND: {
    summary: "Returns the bitwise AND of two numbers.",
  },
  BITOR: {
    summary: "Returns the bitwise OR of two numbers.",
  },
  BITXOR: {
    summary: "Returns the bitwise exclusive OR of two numbers.",
  },
  BYCOL: {
    summary: "Applies a LAMBDA to each column of an array and returns one result per column.",
  },
  BYROW: {
    summary: "Applies a LAMBDA to each row of an array and returns one result per row.",
  },
  CEILING: {
    summary: "Rounds a number away from zero to the nearest multiple of a step you name.",
  },
  "ECMA.CEILING": {
    summary: "Rounds a number up to the nearest multiple of a step, as the ECMA-376 standard defines CEILING; kept for compatibility.",
  },
  "CEILING.MATH": {
    summary: "Rounds a number up to the nearest multiple of a step, with a choice of how negative numbers are treated.",
  },
  CELL: {
    summary: "Returns one piece of information about a cell, such as its address, its format or its contents.",
  },
  CHAR: {
    summary: "Returns the character with the code number you give, in the character set your computer uses.",
  },
  "CHISQ.TEST": {
    summary: "Returns the probability from a chi-squared test of independence between observed and expected counts.",
  },
  CHOOSE: {
    summary: "Returns one of a list of values, picked by its position in that list.",
    args: {
      index_num: "Which value to return, from 1 for the first value up to 254.",
      value1: "The first value to choose from. It may be a number, text, a formula or a reference.",
      value2: "Further values to choose from, up to 254 in all.",
    },
  },
  CHOOSECOLS: {
    summary: "Returns the columns of an array that you name by position.",
  },
  CHOOSEROWS: {
    summary: "Returns the rows of an array that you name by position.",
  },
  CLEAN: {
    summary: "Removes the non-printing characters from a piece of text.",
  },
  CODE: {
    summary: "Returns the code number of the first character of a piece of text.",
  },
  COLUMN: {
    summary: "Returns the column number of a reference.",
    args: {
      reference: "The cell or range whose column number you want. If omitted, the column of the cell holding the formula is returned.",
    },
  },
  COLUMNS: {
    summary: "Returns how many columns a range or array has.",
    args: {
      array: "The range or array whose number of columns you want.",
    },
  },
  COMBIN: {
    summary: "Returns how many ways a given number of items can be chosen from a set, with order ignored.",
  },
  COMBINA: {
    summary: "Returns how many ways items can be chosen from a set when an item may be chosen more than once.",
  },
  COMPLEX: {
    summary: "Builds a complex number as text from a real part and an imaginary part.",
  },
  CONCAT: {
    summary: "Joins the text from several values or ranges into one piece of text, with no separator between them.",
    args: {
      text1: "The first piece of text, cell or range to join.",
      text2: "Further pieces of text, cells or ranges to join, in the order given.",
    },
  },
  CONCATENATE: {
    summary: "Joins several pieces of text into one. It is kept for compatibility; CONCAT and TEXTJOIN do the same work.",
    args: {
      text1: "The first piece of text to join.",
      text2: "Further pieces of text to join, up to 255 in all.",
    },
  },
  "CONFIDENCE.NORM": {
    summary: "Returns half the width of the confidence interval for a population mean, using the normal distribution.",
  },
  "CONFIDENCE.T": {
    summary: "Returns half the width of the confidence interval for a population mean, using the Student t distribution.",
  },
  CONVERT: {
    summary: "Converts a number from one system of measurement to another.",
  },
  CORREL: {
    summary: "Returns the correlation coefficient between two sets of values.",
  },
  COS: {
    summary: "Returns the cosine of an angle given in radians.",
  },
  COSH: {
    summary: "Returns the hyperbolic cosine of a number.",
  },
  COUNT: {
    summary: "Counts how many of its arguments hold numbers.",
    args: {
      value1: "The first value, range or reference to look at. Only numbers are counted.",
      value2: "Further values, ranges or references to look at, up to 255 in all.",
    },
  },
  COUNTA: {
    summary: "Counts how many of its arguments are not empty, whatever kind of value they hold.",
    args: {
      value1: "The first value, range or reference to look at. Every non-empty cell is counted.",
      value2: "Further values, ranges or references to look at, up to 255 in all.",
    },
  },
  COUNTBLANK: {
    summary: "Counts the empty cells in a range.",
  },
  COUNTIF: {
    summary: "Counts the cells in a range that meet one condition.",
    args: {
      range: "The cells to count.",
      criteria: "The condition a cell must meet to be counted, given as a number, an expression or text, such as 32, \">32\" or \"apples\".",
    },
  },
  COUNTIFS: {
    summary: "Counts the cells that meet every one of several conditions.",
    args: {
      criteria_range1: "The cells tested against the first condition.",
      criteria1: "The condition cells in criteria_range1 must meet, given as a number, an expression or text.",
      criteria_range2: "A further range of cells to test. It must be the same size and shape as criteria_range1.",
      criteria2: "The condition that further range must meet. Only cells meeting every condition are counted.",
    },
  },
  "COVARIANCE.P": {
    summary: "Returns the covariance of two sets of values, treating them as a whole population.",
  },
  "COVARIANCE.S": {
    summary: "Returns the covariance of two sets of values, treating them as a sample.",
  },
  CUMIPMT: {
    summary: "Returns the interest paid on a loan between two periods you name.",
  },
  CUMPRINC: {
    summary: "Returns the principal repaid on a loan between two periods you name.",
  },
  DATE: {
    summary: "Returns the date serial number for a year, a month and a day.",
    args: {
      year: "The year, from 1900 to 9999. A one- or two-digit year is read as a year in the 1900s or 2000s.",
      month: "The month, from 1 for January. A number outside 1 to 12 rolls the year forward or back.",
      day: "The day of the month. A number outside the month's length rolls the month forward or back.",
    },
  },
  DATEDIF: {
    summary: "Returns the whole years, months or days between two dates.",
    args: {
      start_date: "The earlier of the two dates.",
      end_date: "The later of the two dates. It must not fall before start_date.",
      unit: "The unit to count in, as text: \"Y\" for whole years, \"M\" for whole months, \"D\" for days, and \"YM\", \"YD\" or \"MD\" for the remainders.",
    },
  },
  DATEVALUE: {
    summary: "Converts a date written as text to the date serial number it stands for.",
  },
  DAVERAGE: {
    summary: "Returns the mean of the values in one field of the database records that meet your conditions.",
  },
  DAY: {
    summary: "Returns the day of the month of a date, as a number from 1 to 31.",
    args: {
      serial_number: "The date whose day of the month you want.",
    },
  },
  DAYS: {
    summary: "Returns how many days lie between two dates.",
  },
  DAYS360: {
    summary: "Returns how many days lie between two dates on a 360-day year of twelve 30-day months.",
  },
  DB: {
    summary: "Returns the depreciation of an asset for one period by the fixed-declining balance method.",
  },
  DCOUNT: {
    summary: "Counts the cells holding numbers in one field of the database records that meet your conditions.",
  },
  DCOUNTA: {
    summary: "Counts the non-empty cells in one field of the database records that meet your conditions.",
  },
  DDB: {
    summary: "Returns the depreciation of an asset for one period by the double-declining balance method.",
  },
  DEC2BIN: {
    summary: "Converts a decimal number to a binary number.",
  },
  DEC2HEX: {
    summary: "Converts a decimal number to a hexadecimal number.",
  },
  DECIMAL: {
    summary: "Converts a number written as text in a given base to an ordinary decimal number.",
  },
  DEGREES: {
    summary: "Converts an angle in radians to degrees.",
  },
  DELTA: {
    summary: "Returns 1 when two numbers are equal and 0 when they are not.",
  },
  DEVSQ: {
    summary: "Returns the sum of the squared deviations of the values from their own mean.",
  },
  DGET: {
    summary: "Returns the single value in one field of the one database record that meets your conditions.",
  },
  DMAX: {
    summary: "Returns the largest value in one field of the database records that meet your conditions.",
  },
  DMIN: {
    summary: "Returns the smallest value in one field of the database records that meet your conditions.",
  },
  DOLLAR: {
    summary: "Converts a number to text in the currency format, rounded to the number of decimal places you name.",
  },
  DOLLARDE: {
    summary: "Converts a price written as a whole number and a fraction to a decimal number.",
  },
  DOLLARFR: {
    summary: "Converts a decimal price to a whole number and a fraction of the denominator you name.",
  },
  DPRODUCT: {
    summary: "Multiplies the values in one field of the database records that meet your conditions.",
  },
  DROP: {
    summary: "Returns an array with a number of rows or columns removed from its start or its end.",
  },
  DSTDEV: {
    summary: "Returns the sample standard deviation of one field of the database records that meet your conditions.",
  },
  DSUM: {
    summary: "Adds the values in one field of the database records that meet your conditions.",
  },
  DURATION: {
    summary: "Returns the Macaulay duration of a security that pays interest periodically.",
  },
  DVAR: {
    summary: "Returns the sample variance of one field of the database records that meet your conditions.",
  },
  EDATE: {
    summary: "Returns the date a number of whole months before or after a date you name.",
    args: {
      start_date: "The date to count from.",
      months: "How many months to move. A positive number moves forward, a negative one back.",
    },
  },
  EFFECT: {
    summary: "Returns the effective annual interest rate implied by a nominal rate and a compounding frequency.",
  },
  ENCODEURL: {
    summary: "Returns a piece of text encoded so that it can be used inside a URL.",
  },
  FILTERXML: {
    summary: "Returns specific data from XML content by using the XPath you give; several matches spill down.",
    args: {
      xml: "A string in valid XML format.",
      xpath: "A string in standard XPath format.",
    },
  },
  WEBSERVICE: {
    summary: "Returns the text a web service answers at a URL, read by the server for a website you allow.",
    args: {
      url: "The http or https URL of the web service, at most 2,048 characters.",
    },
  },
  IMPORTRANGE: {
    summary: "Imports the values of a range from another spreadsheet file of this workspace that you can read.",
    args: {
      spreadsheet_url: "The address or file id of the spreadsheet file to import from.",
      range_string: "The range to import as text, such as \"Sheet1!A1:C10\"; without a sheet name, the first sheet.",
    },
  },
  EOMONTH: {
    summary: "Returns the last day of the month a number of months before or after a date you name.",
    args: {
      start_date: "The date to count from.",
      months: "How many months to move before taking the last day. A positive number moves forward, a negative one back.",
    },
  },
  "ERROR.TYPE": {
    summary: "Returns a number standing for the kind of error value a cell holds.",
  },
  EVEN: {
    summary: "Rounds a number away from zero to the nearest even whole number.",
  },
  EXACT: {
    summary: "Returns TRUE when two pieces of text are exactly the same, counting a difference in case as a difference.",
  },
  EXP: {
    summary: "Returns e raised to the power of a number.",
  },
  EXPAND: {
    summary: "Returns an array grown to the number of rows and columns you name, padding the new cells.",
  },
  "EXPON.DIST": {
    summary: "Returns the exponential probability distribution for a value.",
  },
  FACT: {
    summary: "Returns the factorial of a number, the product of every whole number from 1 up to it.",
  },
  FALSE: {
    summary: "Returns the logical value FALSE.",
  },
  FILTER: {
    summary: "Returns only the rows or columns of a range that your condition keeps.",
    args: {
      array: "The range or array to filter.",
      include: "An array of TRUE and FALSE values, one per row or column of array. Only the TRUE ones are kept.",
      if_empty: "What to return when no row or column is kept. If omitted, the result is the #CALC! error.",
    },
  },
  FIND: {
    summary: "Returns the position at which one piece of text first appears inside another, counting case as a difference.",
    args: {
      find_text: "The text you are looking for.",
      within_text: "The text to look inside.",
      start_num: "The character to start looking from, counting from 1. If omitted, the search starts at the first character.",
    },
  },
  FIXED: {
    summary: "Converts a number to text with a fixed number of decimal places, with or without thousands separators.",
  },
  FLOOR: {
    summary: "Rounds a number towards zero to the nearest multiple of a step you name.",
  },
  "FLOOR.MATH": {
    summary: "Rounds a number down to the nearest multiple of a step, with a choice of how negative numbers are treated.",
  },
  "FORECAST.LINEAR": {
    summary: "Predicts a y value for a new x by fitting a straight line through the known points.",
  },
  FORMULATEXT: {
    summary: "Returns the formula in a cell as text, exactly as it is written in the formula bar.",
  },
  FREQUENCY: {
    summary: "Counts how many values fall into each of the intervals you name, and returns the counts as a column.",
  },
  FV: {
    summary: "Returns the future value of an investment with constant payments and a constant interest rate.",
    args: {
      rate: "The interest rate for each period.",
      nper: "The total number of payment periods.",
      pmt: "The payment made each period. It cannot change over the life of the investment.",
      pv: "The present value, the lump sum the future payments are worth now. If omitted, 0 is used.",
      type: "When payments fall due: 0 or omitted for the end of the period, 1 for the beginning.",
    },
  },
  FVSCHEDULE: {
    summary: "Returns the future value of a lump sum after applying a series of different interest rates.",
  },
  GCD: {
    summary: "Returns the greatest common divisor of two or more whole numbers.",
  },
  GEOMEAN: {
    summary: "Returns the geometric mean of a set of positive numbers.",
  },
  GESTEP: {
    summary: "Returns 1 when a number reaches a threshold you name, and 0 when it does not.",
  },
  GETPIVOTDATA: {
    summary: "Returns one value held in a PivotTable, found by naming its data field and the items that select it.",
  },
  GROUPBY: {
    summary: "Groups the rows of a table by one or more fields and aggregates the values in each group.",
  },
  GROWTH: {
    summary: "Predicts y values along an exponential curve fitted to the known points.",
  },
  HARMEAN: {
    summary: "Returns the harmonic mean of a set of positive numbers.",
  },
  HEX2DEC: {
    summary: "Converts a hexadecimal number to a decimal number.",
  },
  HLOOKUP: {
    summary: "Looks for a value in the top row of a table and returns a value from the row you name in the same column.",
    args: {
      lookup_value: "The value to look for in the top row of table_array.",
      table_array: "The table of data to search. The values to match must lie in its first row.",
      row_index_num: "Which row of table_array holds the value to return, counting the top row as 1.",
      range_lookup: "TRUE or omitted finds the closest match and needs the top row sorted; FALSE finds only an exact match.",
    },
  },
  HOUR: {
    summary: "Returns the hour of a time, as a number from 0 to 23.",
  },
  HSTACK: {
    summary: "Joins arrays side by side into one wider array.",
  },
  HYPERLINK: {
    summary: "Makes a cell into a link that opens a document, a place in the workbook, or an address on the internet.",
  },
  IF: {
    summary: "Checks whether a condition is met, and returns one value if it is TRUE and another if it is FALSE.",
    args: {
      logical_test: "The value or expression that is evaluated to TRUE or FALSE.",
      value_if_true: "The value returned if logical_test is TRUE. If omitted, TRUE is returned.",
      value_if_false: "The value returned if logical_test is FALSE. If omitted, FALSE is returned.",
    },
  },
  IFERROR: {
    summary: "Returns a value of your own if an expression is an error, and the expression's own result otherwise.",
    args: {
      value: "The value, reference or formula to check for an error.",
      value_if_error: "The value returned when value is any error, including #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME? and #NULL!.",
    },
  },
  IFNA: {
    summary: "Returns a value of your own if an expression is the #N/A error, and the expression's own result otherwise.",
    args: {
      value: "The value, reference or formula to check for #N/A.",
      value_if_na: "The value returned when value is #N/A. Every other error is passed through unchanged.",
    },
  },
  IFS: {
    summary: "Checks several conditions in turn and returns the value that goes with the first one that is TRUE.",
    args: {
      logical_test1: "The first condition, evaluated to TRUE or FALSE.",
      value_if_true1: "The value returned when logical_test1 is TRUE.",
      logical_test2: "A further condition, checked only when every earlier one was FALSE.",
      value_if_true2: "The value returned when that further condition is the first one to be TRUE.",
    },
  },
  IMAGE: {
    summary: "Shows the picture at an https address in the cell, with alternative text and a sizing of your choice.",
    args: {
      source: "The https URL of a BMP, JPG/JPEG, GIF, TIFF, PNG, ICO or WEBP picture.",
      alt_text: "Alternative text that describes the picture for accessibility.",
      sizing: "0 fits the picture in the cell keeping its aspect ratio (the default), 1 fills the cell, 2 keeps the picture's own size, 3 uses height and width.",
      height: "With sizing 3, the height of the picture in pixels.",
      width: "With sizing 3, the width of the picture in pixels.",
    },
  },
  IMABS: {
    summary: "Returns the absolute value, or modulus, of a complex number.",
  },
  IMAGINARY: {
    summary: "Returns the imaginary part of a complex number.",
  },
  IMREAL: {
    summary: "Returns the real part of a complex number.",
  },
  IMSUM: {
    summary: "Returns the sum of two or more complex numbers.",
  },
  INDEX: {
    summary: "Returns the value at a row and column you name inside a range or array.",
    args: {
      array: "The range or array to take the value from.",
      row_num: "Which row of array to use, counting from 1. Use 0 to return the whole column.",
      column_num: "Which column of array to use, counting from 1. Use 0 to return the whole row.",
      area_num: "Which of the areas of a multiple-area reference to use, counting from 1. If omitted, the first area is used.",
    },
  },
  INDIRECT: {
    summary: "Returns the contents of the reference that a piece of text spells out.",
    args: {
      ref_text: "A reference written as text, such as \"A1\" or \"Sheet2!B3\". The text is read as a reference each time the sheet recalculates.",
      a1: "TRUE or omitted reads ref_text in A1 style; FALSE reads it in R1C1 style.",
    },
  },
  INFO: {
    summary: "Returns one piece of information about the current working environment.",
  },
  INT: {
    summary: "Rounds a number down to the nearest whole number.",
    args: {
      number: "The real number to round down. A negative number rounds away from zero, so -4.3 becomes -5.",
    },
  },
  INTERCEPT: {
    summary: "Returns the point at which the least-squares line through the known points crosses the y axis.",
  },
  IPMT: {
    summary: "Returns the interest part of one payment on a loan with constant payments and a constant rate.",
  },
  IRR: {
    summary: "Returns the internal rate of return of a series of cash flows that fall at regular intervals.",
    args: {
      values: "A range or array of cash flows. It must hold at least one negative and one positive value, in the order the flows occur.",
      guess: "A rate you think the answer is near, used to start the search. If omitted, 10 percent is used.",
    },
  },
  ISBLANK: {
    summary: "Returns TRUE when a cell is empty.",
    args: {
      value: "The value or cell to test. A cell holding an empty piece of text is not blank.",
    },
  },
  ISERR: {
    summary: "Returns TRUE when a value is any error except #N/A.",
  },
  ISERROR: {
    summary: "Returns TRUE when a value is any error at all.",
    args: {
      value: "The value or cell to test. Every error counts, including #N/A.",
    },
  },
  ISEVEN: {
    summary: "Returns TRUE when a number is even.",
  },
  ISFORMULA: {
    summary: "Returns TRUE when a cell holds a formula.",
  },
  ISLOGICAL: {
    summary: "Returns TRUE when a value is TRUE or FALSE.",
  },
  ISNA: {
    summary: "Returns TRUE when a value is the #N/A error.",
    args: {
      value: "The value or cell to test. Only #N/A counts; every other error gives FALSE.",
    },
  },
  ISNONTEXT: {
    summary: "Returns TRUE when a value is anything other than text, including an empty cell.",
  },
  ISNUMBER: {
    summary: "Returns TRUE when a value is a number.",
    args: {
      value: "The value or cell to test. A date or a time counts as a number.",
    },
  },
  ISODD: {
    summary: "Returns TRUE when a number is odd.",
  },
  ISOMITTED: {
    summary: "Returns TRUE when an optional LAMBDA argument was left out of the call.",
  },
  ISOWEEKNUM: {
    summary: "Returns the ISO week number of the year a date falls in.",
  },
  ISREF: {
    summary: "Returns TRUE when a value is a reference.",
  },
  ISTEXT: {
    summary: "Returns TRUE when a value is text.",
    args: {
      value: "The value or cell to test. A number stored as text counts as text.",
    },
  },
  KURT: {
    summary: "Returns the kurtosis of a set of values, a measure of how sharply they peak.",
  },
  LAMBDA: {
    summary: "Builds a reusable function from parameters and a calculation, with no code outside the sheet.",
  },
  LARGE: {
    summary: "Returns the value at a given position from the top of a set of values.",
    args: {
      array: "The range or array of values to rank.",
      k: "The position from the largest value, counting 1 as the largest itself.",
    },
  },
  LCM: {
    summary: "Returns the least common multiple of two or more whole numbers.",
  },
  LEFT: {
    summary: "Returns a given number of characters from the start of a piece of text.",
    args: {
      text: "The text to take characters from.",
      num_chars: "How many characters to return. If omitted, 1 is used.",
    },
  },
  LEN: {
    summary: "Returns how many characters a piece of text has, counting spaces.",
    args: {
      text: "The text whose length you want.",
    },
  },
  LET: {
    summary: "Names the results of parts of a calculation so that the rest of the formula can reuse them.",
  },
  LINEST: {
    summary: "Returns the coefficients of the least-squares straight line fitted to a set of points.",
  },
  LN: {
    summary: "Returns the natural logarithm of a number.",
  },
  LOG: {
    summary: "Returns the logarithm of a number to the base you name.",
  },
  LOG10: {
    summary: "Returns the base-10 logarithm of a number.",
  },
  LOOKUP: {
    summary: "Looks for a value in a single row or column and returns the value in the same place in a second one.",
    args: {
      lookup_value: "The value to look for. The vector searched must be sorted in ascending order.",
      lookup_vector: "The single row or column to search.",
      result_vector: "The single row or column to take the answer from. It must be the same size as lookup_vector.",
    },
  },
  LOWER: {
    summary: "Converts a piece of text to lower case.",
    args: {
      text: "The text to convert. Characters that are not letters are left alone.",
    },
  },
  MAKEARRAY: {
    summary: "Builds an array of the size you name by calling a LAMBDA once for each cell.",
  },
  MAP: {
    summary: "Returns an array made by applying a LAMBDA to every value of one or more arrays.",
  },
  MATCH: {
    summary: "Returns the position of a value inside a single row or column, not the value itself.",
    args: {
      lookup_value: "The value to look for in lookup_array.",
      lookup_array: "The single row, column or array to search.",
      match_type: "1 or omitted finds the largest value not above lookup_value in an ascending list; 0 finds an exact match; -1 finds the smallest value not below it in a descending list.",
    },
  },
  MAX: {
    summary: "Returns the largest number among its arguments, ignoring text and logical values.",
    args: {
      number1: "The first number, range or reference to compare.",
      number2: "Further numbers, ranges or references to compare, up to 255 in all.",
    },
  },
  MAXA: {
    summary: "Returns the largest value among its arguments, counting text as 0 and logical values as 1 or 0.",
  },
  MAXIFS: {
    summary: "Returns the largest value among the cells that meet every one of several conditions.",
  },
  MDETERM: {
    summary: "Returns the determinant of a square matrix.",
  },
  MEDIAN: {
    summary: "Returns the median, the value in the middle of a set of numbers once they are ordered.",
    args: {
      number1: "The first number, range or reference to take the median of.",
      number2: "Further numbers, ranges or references, up to 255 in all.",
    },
  },
  MID: {
    summary: "Returns a given number of characters from the middle of a piece of text, starting where you say.",
    args: {
      text: "The text to take characters from.",
      start_num: "The position of the first character to return, counting from 1.",
      num_chars: "How many characters to return. Fewer are returned when the text runs out first.",
    },
  },
  MIN: {
    summary: "Returns the smallest number among its arguments, ignoring text and logical values.",
    args: {
      number1: "The first number, range or reference to compare.",
      number2: "Further numbers, ranges or references to compare, up to 255 in all.",
    },
  },
  MINA: {
    summary: "Returns the smallest value among its arguments, counting text as 0 and logical values as 1 or 0.",
  },
  MINIFS: {
    summary: "Returns the smallest value among the cells that meet every one of several conditions.",
  },
  MINUTE: {
    summary: "Returns the minute of a time, as a number from 0 to 59.",
  },
  MINVERSE: {
    summary: "Returns the inverse of a square matrix.",
  },
  MIRR: {
    summary: "Returns the internal rate of return of a series of cash flows, allowing for different borrowing and reinvestment rates.",
  },
  MMULT: {
    summary: "Returns the matrix product of two arrays.",
  },
  MOD: {
    summary: "Returns the remainder left when one number is divided by another.",
    args: {
      number: "The number to divide.",
      divisor: "The number to divide by. The result takes the sign of the divisor.",
    },
  },
  "MODE.MULT": {
    summary: "Returns every value that occurs most often in a set of numbers.",
  },
  "MODE.SNGL": {
    summary: "Returns the value that occurs most often in a set of numbers.",
  },
  MONTH: {
    summary: "Returns the month of a date, as a number from 1 for January to 12 for December.",
    args: {
      serial_number: "The date whose month you want.",
    },
  },
  MROUND: {
    summary: "Rounds a number to the nearest multiple of a step you name.",
    args: {
      number: "The value to round.",
      multiple: "The step to round to. It must have the same sign as number.",
    },
  },
  MULTINOMIAL: {
    summary: "Returns the multinomial coefficient of a set of numbers.",
  },
  MUNIT: {
    summary: "Returns the identity matrix of the size you name.",
  },
  N: {
    summary: "Converts a value to a number, reading TRUE as 1 and text as 0.",
  },
  NA: {
    summary: "Returns the #N/A error, which marks a value that is not available.",
  },
  NETWORKDAYS: {
    summary: "Returns how many whole working days lie between two dates, leaving out weekends and any holidays you list.",
    args: {
      start_date: "The date the count starts from.",
      end_date: "The date the count runs to. Both end dates are counted when they are working days.",
      holidays: "A range or array of dates to leave out as well, such as public holidays.",
    },
  },
  "NETWORKDAYS.INTL": {
    summary: "Returns how many whole working days lie between two dates, with a weekend of your own choosing.",
  },
  NOMINAL: {
    summary: "Returns the nominal annual interest rate implied by an effective rate and a compounding frequency.",
  },
  "NORM.DIST": {
    summary: "Returns the normal distribution for a value, given a mean and a standard deviation.",
  },
  "NORM.INV": {
    summary: "Returns the value at which the normal distribution reaches a given probability.",
  },
  "NORM.S.DIST": {
    summary: "Returns the standard normal distribution, the one with a mean of 0 and a standard deviation of 1.",
  },
  "NORM.S.INV": {
    summary: "Returns the value at which the standard normal distribution reaches a given probability.",
  },
  NOT: {
    summary: "Reverses a logical value, returning FALSE for TRUE and TRUE for FALSE.",
    args: {
      logical: "The value or expression that is evaluated to TRUE or FALSE.",
    },
  },
  NOW: {
    summary: "Returns the current date and time, and updates whenever the sheet recalculates.",
  },
  NPER: {
    summary: "Returns how many periods an investment takes, given constant payments and a constant interest rate.",
    args: {
      rate: "The interest rate for each period.",
      pmt: "The payment made each period. It cannot change over the life of the loan or investment.",
      pv: "The present value, the lump sum the future payments are worth now.",
      fv: "The cash balance you want left after the last payment. If omitted, 0 is used.",
      type: "When payments fall due: 0 or omitted for the end of the period, 1 for the beginning.",
    },
  },
  NPV: {
    summary: "Returns the net present value of a series of future cash flows at a given discount rate.",
    args: {
      rate: "The discount rate for one period.",
      value1: "The first cash flow. It is treated as falling one period from now, not today.",
      value2: "Further cash flows, one period apart and in the order they occur, up to 254 in all.",
    },
  },
  NUMBERVALUE: {
    summary: "Converts text to a number using the decimal and grouping separators you name.",
  },
  OCT2DEC: {
    summary: "Converts an octal number to a decimal number.",
  },
  ODD: {
    summary: "Rounds a number away from zero to the nearest odd whole number.",
  },
  OFFSET: {
    summary: "Returns a reference shifted from a starting reference by a number of rows and columns, and optionally resized.",
    args: {
      reference: "The reference to measure the offset from.",
      rows: "How many rows to move down. A negative number moves up.",
      cols: "How many columns to move right. A negative number moves left.",
      height: "How many rows tall the returned reference should be. If omitted, the height of reference is kept.",
      width: "How many columns wide the returned reference should be. If omitted, the width of reference is kept.",
    },
  },
  OR: {
    summary: "Returns TRUE when at least one of its arguments is TRUE, and FALSE only when they all are FALSE.",
    args: {
      logical1: "The first condition to test. It may be a value, an expression or a reference that resolves to TRUE or FALSE.",
      logical2: "Further conditions to test, up to 255 in all. One TRUE is enough for the result to be TRUE.",
    },
  },
  PDURATION: {
    summary: "Returns how many periods an investment needs to grow from one value to another at a fixed rate.",
  },
  PEARSON: {
    summary: "Returns the Pearson product-moment correlation coefficient of two sets of values.",
  },
  "PERCENTILE.EXC": {
    summary: "Returns the value at a given percentile of a set, with the percentile strictly between 0 and 1.",
  },
  "PERCENTILE.INC": {
    summary: "Returns the value at a given percentile of a set, with the percentile from 0 to 1 inclusive.",
  },
  PERCENTOF: {
    summary: "Returns what share of a whole set of values a subset of them makes up.",
  },
  "PERCENTRANK.INC": {
    summary: "Returns where a value ranks in a set, as a percentage from 0 to 1 inclusive.",
  },
  PERMUT: {
    summary: "Returns how many ways a given number of items can be chosen from a set when order matters.",
  },
  PI: {
    summary: "Returns the value of pi to fifteen digits.",
  },
  PIVOTBY: {
    summary: "Groups the rows of a table by row and column fields and aggregates the values at each crossing.",
  },
  PMT: {
    summary: "Returns the payment for a loan with constant payments and a constant interest rate.",
    args: {
      rate: "The interest rate for each period. Divide an annual rate by the number of payments a year.",
      nper: "The total number of payments over the life of the loan.",
      pv: "The present value, the amount borrowed today.",
      fv: "The cash balance you want left after the last payment. If omitted, 0 is used.",
      type: "When payments fall due: 0 or omitted for the end of the period, 1 for the beginning.",
    },
  },
  "POISSON.DIST": {
    summary: "Returns the Poisson distribution for a count of events in a fixed interval.",
  },
  POWER: {
    summary: "Returns a number raised to a power.",
    args: {
      number: "The base number, which may be any real number.",
      power: "The exponent to raise the base to.",
    },
  },
  PPMT: {
    summary: "Returns the principal part of one payment on a loan with constant payments and a constant rate.",
  },
  PRICE: {
    summary: "Returns the price per 100 of face value of a security that pays interest periodically.",
  },
  PROB: {
    summary: "Returns the chance that a value falls between two limits, given a list of values and their probabilities.",
  },
  PRODUCT: {
    summary: "Multiplies all the numbers given as arguments and returns the product.",
  },
  PROPER: {
    summary: "Converts a piece of text so that each word starts with a capital letter and the rest is lower case.",
    args: {
      text: "The text to capitalise. A letter after anything other than a letter starts a new word.",
    },
  },
  PV: {
    summary: "Returns the present value of an investment, the lump sum a series of future payments is worth now.",
    args: {
      rate: "The interest rate for each period.",
      nper: "The total number of payment periods.",
      pmt: "The payment made each period. It cannot change over the life of the investment.",
      fv: "The cash balance you want left after the last payment. If omitted, 0 is used.",
      type: "When payments fall due: 0 or omitted for the end of the period, 1 for the beginning.",
    },
  },
  "QUARTILE.EXC": {
    summary: "Returns a quartile of a set of values, with the quartile strictly between 0 and 4.",
  },
  "QUARTILE.INC": {
    summary: "Returns a quartile of a set of values, from the minimum to the maximum inclusive.",
  },
  QUOTIENT: {
    summary: "Returns the whole-number part of a division, dropping any remainder.",
  },
  RADIANS: {
    summary: "Converts an angle in degrees to radians.",
  },
  RAND: {
    summary: "Returns a random number from 0 up to but not including 1, redrawn whenever the sheet recalculates.",
  },
  RANDARRAY: {
    summary: "Returns an array of random numbers of the size, range and kind you name.",
  },
  RANDBETWEEN: {
    summary: "Returns a random whole number between two limits you name, both included.",
  },
  RANK: {
    summary: "Returns where a number ranks among a list of numbers. It is kept for compatibility with RANK.EQ.",
  },
  "RANK.AVG": {
    summary: "Returns where a number ranks among a list of numbers, giving tied values the average of their ranks.",
  },
  "RANK.EQ": {
    summary: "Returns where a number ranks among a list of numbers, giving tied values the highest of their ranks.",
    args: {
      number: "The number whose rank you want.",
      ref: "The list of numbers to rank within. Values that are not numbers are ignored.",
      order: "0 or omitted ranks the list as if sorted from largest to smallest; any other number ranks it the other way round.",
    },
  },
  RATE: {
    summary: "Returns the interest rate per period of a loan or investment with constant payments.",
    args: {
      nper: "The total number of payment periods.",
      pmt: "The payment made each period. It cannot change over the life of the loan.",
      pv: "The present value, the lump sum the future payments are worth now.",
      fv: "The cash balance you want left after the last payment. If omitted, 0 is used.",
      type: "When payments fall due: 0 or omitted for the end of the period, 1 for the beginning.",
      guess: "A rate you think the answer is near, used to start the search. If omitted, 10 percent is used.",
    },
  },
  RECEIVED: {
    summary: "Returns the amount received at maturity for a security that pays no interest before then.",
  },
  REDUCE: {
    summary: "Runs an array through a LAMBDA value by value, carrying a running total, and returns the final result.",
  },
  REGEXEXTRACT: {
    summary: "Returns the parts of a piece of text that match a regular expression.",
  },
  REGEXREPLACE: {
    summary: "Replaces the parts of a piece of text that match a regular expression.",
  },
  REGEXTEST: {
    summary: "Returns TRUE when a piece of text matches a regular expression.",
  },
  REPLACE: {
    summary: "Replaces a stretch of characters inside a piece of text, chosen by position, with new text.",
    args: {
      old_text: "The text in which some characters are to be replaced.",
      start_num: "The position of the first character to replace, counting from 1.",
      num_chars: "How many characters to replace. Use 0 to insert the new text without removing anything.",
      new_text: "The text to put in their place. It need not be the same length.",
    },
  },
  REPT: {
    summary: "Repeats a piece of text a given number of times.",
  },
  RIGHT: {
    summary: "Returns a given number of characters from the end of a piece of text.",
    args: {
      text: "The text to take characters from.",
      num_chars: "How many characters to return. If omitted, 1 is used.",
    },
  },
  ROMAN: {
    summary: "Converts a number to a Roman numeral written as text.",
  },
  ROUND: {
    summary: "Rounds a number to the number of digits you name.",
    args: {
      number: "The number to round.",
      num_digits: "How many decimal places to keep. A negative number rounds to the left of the decimal point, so -2 rounds to hundreds.",
    },
  },
  ROUNDDOWN: {
    summary: "Rounds a number towards zero, to the number of digits you name.",
    args: {
      number: "The number to round down.",
      num_digits: "How many decimal places to keep. A negative number rounds to the left of the decimal point.",
    },
  },
  ROUNDUP: {
    summary: "Rounds a number away from zero, to the number of digits you name.",
    args: {
      number: "The number to round up.",
      num_digits: "How many decimal places to keep. A negative number rounds to the left of the decimal point.",
    },
  },
  ROW: {
    summary: "Returns the row number of a reference.",
    args: {
      reference: "The cell or range whose row number you want. If omitted, the row of the cell holding the formula is returned.",
    },
  },
  ROWS: {
    summary: "Returns how many rows a range or array has.",
    args: {
      array: "The range or array whose number of rows you want.",
    },
  },
  RRI: {
    summary: "Returns the rate of return implied by an investment growing from one value to another over a number of periods.",
  },
  RSQ: {
    summary: "Returns the square of the correlation coefficient between two sets of values.",
  },
  SCAN: {
    summary: "Runs an array through a LAMBDA value by value and returns every intermediate result.",
  },
  SEARCH: {
    summary: "Returns the position at which one piece of text first appears inside another, ignoring case and allowing wildcards.",
    args: {
      find_text: "The text you are looking for. It may use ? for any one character and * for any run of characters.",
      within_text: "The text to look inside.",
      start_num: "The character to start looking from, counting from 1. If omitted, the search starts at the first character.",
    },
  },
  SECOND: {
    summary: "Returns the second of a time, as a number from 0 to 59.",
  },
  SEQUENCE: {
    summary: "Returns an array of numbers counting up in a step you name.",
    args: {
      rows: "How many rows the returned array has.",
      columns: "How many columns the returned array has. If omitted, 1 is used.",
      start: "The first number in the array. If omitted, 1 is used.",
      step: "How much each number rises over the one before. If omitted, 1 is used.",
    },
  },
  SERIESSUM: {
    summary: "Returns the sum of a power series with the coefficients you give.",
  },
  SHEET: {
    summary: "Returns the position of a sheet in the workbook.",
  },
  SHEETS: {
    summary: "Returns how many sheets a reference covers.",
  },
  SIGN: {
    summary: "Returns 1 for a positive number, -1 for a negative one and 0 for zero.",
  },
  SIN: {
    summary: "Returns the sine of an angle given in radians.",
  },
  SINH: {
    summary: "Returns the hyperbolic sine of a number.",
  },
  SKEW: {
    summary: "Returns the skewness of a set of values, a measure of how lopsided they are.",
  },
  SLN: {
    summary: "Returns the depreciation of an asset for one period by the straight-line method.",
  },
  SLOPE: {
    summary: "Returns the slope of the least-squares line through a set of points.",
  },
  SMALL: {
    summary: "Returns the value at a given position from the bottom of a set of values.",
    args: {
      array: "The range or array of values to rank.",
      k: "The position from the smallest value, counting 1 as the smallest itself.",
    },
  },
  SORT: {
    summary: "Returns the rows or columns of a range put in order.",
    args: {
      array: "The range or array to sort.",
      sort_index: "Which row or column to sort on, counting from 1. If omitted, the first is used.",
      sort_order: "1 or omitted sorts in ascending order; -1 sorts in descending order.",
      by_col: "FALSE or omitted sorts the rows; TRUE sorts the columns instead.",
    },
  },
  SORTBY: {
    summary: "Returns a range put in the order of one or more other ranges, which need not be part of the result.",
  },
  SQRT: {
    summary: "Returns the positive square root of a number.",
    args: {
      number: "The number to take the square root of. A negative number gives the #NUM! error.",
    },
  },
  STANDARDIZE: {
    summary: "Returns a value expressed as the number of standard deviations it lies from a mean.",
  },
  "STDEV.P": {
    summary: "Returns the standard deviation of a set of values treated as a whole population.",
  },
  "STDEV.S": {
    summary: "Returns the standard deviation of a set of values treated as a sample of a larger population.",
    args: {
      number1: "The first number, range or reference in the sample.",
      number2: "Further numbers, ranges or references in the sample, up to 255 in all.",
    },
  },
  STDEVA: {
    summary: "Returns the sample standard deviation, counting text as 0 and logical values as 1 or 0.",
  },
  STEYX: {
    summary: "Returns the standard error of the predicted y values in a linear regression.",
  },
  SUBSTITUTE: {
    summary: "Replaces one piece of text with another wherever it appears, or only at the occurrence you name.",
    args: {
      text: "The text in which the replacement is made, or a reference to it.",
      old_text: "The text to look for. It is matched exactly, case included.",
      new_text: "The text to put in its place. An empty piece of text removes old_text altogether.",
      instance_num: "Which occurrence to replace. If omitted, every occurrence is replaced.",
    },
  },
  SUBTOTAL: {
    summary: "Applies one of a list of aggregate calculations to a range, with the option to skip rows hidden by a filter.",
    args: {
      function_num: "A number from 1 to 11 naming the calculation, such as 9 for SUM. Add 100 to skip rows hidden by hand as well.",
      ref1: "The first range to aggregate. Any SUBTOTAL nested inside it is left out of the result.",
      ref2: "Further ranges to aggregate, up to 254 in all.",
    },
  },
  SUM: {
    summary: "Adds all the numbers in a range of cells.",
    args: {
      number1: "The first number, range or reference to add.",
      number2: "Further numbers, ranges or references to add, up to 255 in all.",
    },
  },
  SUMIF: {
    summary: "Adds the cells that meet one condition.",
    args: {
      range: "The cells whose values are tested against the condition.",
      criteria: "The condition a cell must meet to be added, given as a number, an expression or text, such as 32, \">32\" or \"apples\".",
      sum_range: "The cells actually added. If omitted, the cells in range are added.",
    },
  },
  SUMIFS: {
    summary: "Adds the cells that meet every one of several conditions.",
    args: {
      sum_range: "The cells to add once all the conditions are met.",
      criteria_range1: "The cells tested against the first condition.",
      criteria1: "The condition cells in criteria_range1 must meet, given as a number, an expression or text.",
      criteria_range2: "A further range of cells to test. It must be the same size and shape as sum_range.",
      criteria2: "The condition that further range must meet. Only cells meeting every condition are added.",
    },
  },
  SUMPRODUCT: {
    summary: "Multiplies matching values in the arrays you give and returns the sum of those products.",
    args: {
      array1: "The first range or array whose values are to be multiplied and added.",
      array2: "Further ranges or arrays, each the same size and shape as the first, up to 255 in all.",
    },
  },
  SUMSQ: {
    summary: "Returns the sum of the squares of its arguments.",
  },
  SUMX2MY2: {
    summary: "Returns the sum of the difference of the squares of matching values in two arrays.",
  },
  SUMXMY2: {
    summary: "Returns the sum of the squares of the differences between matching values in two arrays.",
  },
  SWITCH: {
    summary: "Compares an expression against a list of values and returns the result that goes with the first match.",
  },
  SYD: {
    summary: "Returns the depreciation of an asset for one period by the sum-of-years-digits method.",
  },
  T: {
    summary: "Returns a value when it is text, and an empty piece of text when it is not.",
  },
  "T.DIST": {
    summary: "Returns the left-tailed Student t distribution for a value.",
  },
  "T.INV": {
    summary: "Returns the value at which the left-tailed Student t distribution reaches a given probability.",
  },
  "T.TEST": {
    summary: "Returns the probability from a Student t test that two samples come from the same population.",
  },
  TAKE: {
    summary: "Returns a given number of rows or columns from the start or the end of an array.",
  },
  TAN: {
    summary: "Returns the tangent of an angle given in radians.",
  },
  TANH: {
    summary: "Returns the hyperbolic tangent of a number.",
  },
  TEXT: {
    summary: "Converts a value to text in the number format you name.",
    args: {
      value: "The number, date or time to convert to text.",
      format_text: "The number format to use, written in quotation marks, such as \"0.00\" or \"dd/mm/yyyy\".",
    },
  },
  TEXTAFTER: {
    summary: "Returns the part of a piece of text that follows a delimiter you name.",
  },
  TEXTBEFORE: {
    summary: "Returns the part of a piece of text that comes before a delimiter you name.",
  },
  TEXTJOIN: {
    summary: "Joins several pieces of text into one, putting a separator of your own between them.",
    args: {
      delimiter: "The text placed between each pair of joined values. An empty piece of text joins them with nothing between.",
      ignore_empty: "TRUE leaves empty cells out, so no separator is doubled; FALSE keeps a place for each one.",
      text1: "The first piece of text, cell or range to join.",
      text2: "Further pieces of text, cells or ranges to join, up to 252 in all.",
    },
  },
  TEXTSPLIT: {
    summary: "Splits a piece of text into columns and rows at the delimiters you name.",
  },
  TIME: {
    summary: "Returns the time serial number for an hour, a minute and a second.",
  },
  TIMEVALUE: {
    summary: "Converts a time written as text to the time serial number it stands for.",
  },
  TOCOL: {
    summary: "Returns the values of an array laid out as a single column.",
  },
  TODAY: {
    summary: "Returns today's date, and updates whenever the sheet recalculates.",
  },
  TOROW: {
    summary: "Returns the values of an array laid out as a single row.",
  },
  TRANSPOSE: {
    summary: "Returns a range turned on its side, so that its rows become columns and its columns become rows.",
    args: {
      array: "The range or array to turn on its side.",
    },
  },
  TREND: {
    summary: "Predicts y values along the straight line that best fits the known points.",
  },
  TRIM: {
    summary: "Removes the leading, trailing and repeated spaces from a piece of text.",
    args: {
      text: "The text to remove the extra spaces from. Single spaces between words are kept.",
    },
  },
  TRIMMEAN: {
    summary: "Returns the mean of a set of values once a share of the highest and lowest have been dropped.",
  },
  TRUE: {
    summary: "Returns the logical value TRUE.",
  },
  TRUNC: {
    summary: "Cuts a number off at the number of digits you name, with no rounding.",
    args: {
      number: "The number to cut off.",
      num_digits: "How many decimal places to keep. If omitted, 0 is used and the whole-number part is returned.",
    },
  },
  TYPE: {
    summary: "Returns a number standing for the kind of value a cell holds, such as a number, text or an error.",
  },
  UNICHAR: {
    summary: "Returns the character with the Unicode code point you give.",
  },
  UNICODE: {
    summary: "Returns the Unicode code point of the first character of a piece of text.",
  },
  UNIQUE: {
    summary: "Returns the distinct values of a range, with the repeats left out.",
    args: {
      array: "The range or array to take the distinct values from.",
      by_col: "FALSE or omitted compares the rows; TRUE compares the columns instead.",
      exactly_once: "TRUE returns only the values that appear exactly once; FALSE or omitted returns one of each.",
    },
  },
  UPPER: {
    summary: "Converts a piece of text to upper case.",
    args: {
      text: "The text to convert. Characters that are not letters are left alone.",
    },
  },
  VALUE: {
    summary: "Converts a number written as text to a number.",
    args: {
      text: "The text to convert. It must be written in a number, date, time or currency format the sheet recognises.",
    },
  },
  VALUETOTEXT: {
    summary: "Returns a value as text, either plainly or in a form you could type back into a formula.",
  },
  "VAR.P": {
    summary: "Returns the variance of a set of values treated as a whole population.",
  },
  "VAR.S": {
    summary: "Returns the variance of a set of values treated as a sample of a larger population.",
  },
  VARA: {
    summary: "Returns the sample variance, counting text as 0 and logical values as 1 or 0.",
  },
  VDB: {
    summary: "Returns the depreciation of an asset over the periods you name by the variable-declining balance method.",
  },
  VLOOKUP: {
    summary: "Looks for a value in the first column of a table and returns a value from the column you name in the same row.",
    args: {
      lookup_value: "The value to look for in the first column of table_array.",
      table_array: "The table of data to search. The values to match must lie in its first column.",
      col_index_num: "Which column of table_array holds the value to return, counting the first column as 1.",
      range_lookup: "TRUE or omitted finds the closest match and needs the first column sorted; FALSE finds only an exact match.",
    },
  },
  VSTACK: {
    summary: "Stacks arrays one under another into one taller array.",
  },
  WEEKDAY: {
    summary: "Returns the day of the week of a date, as a number whose numbering you may choose.",
  },
  WEEKNUM: {
    summary: "Returns the week of the year a date falls in, on the numbering you choose.",
  },
  "WEIBULL.DIST": {
    summary: "Returns the Weibull distribution for a value, often used in reliability work.",
  },
  WORKDAY: {
    summary: "Returns the date a number of working days before or after a date, leaving out weekends and any holidays you list.",
    args: {
      start_date: "The date to count from.",
      days: "How many working days to move. A positive number moves forward, a negative one back.",
      holidays: "A range or array of dates to leave out as well, such as public holidays.",
    },
  },
  "WORKDAY.INTL": {
    summary: "Returns the date a number of working days before or after a date, with a weekend of your own choosing.",
  },
  WRAPCOLS: {
    summary: "Wraps a single row or column into an array of columns of the length you name.",
  },
  WRAPROWS: {
    summary: "Wraps a single row or column into an array of rows of the length you name.",
  },
  XIRR: {
    summary: "Returns the internal rate of return of cash flows that fall on dates you give, at any spacing.",
  },
  XLOOKUP: {
    summary: "Looks for a value in one range and returns the value in the same position in another, in any direction.",
    args: {
      lookup_value: "The value to look for in lookup_array.",
      lookup_array: "The range or array to search.",
      return_array: "The range or array to take the answer from. It must be the same size as lookup_array.",
      if_not_found: "What to return when nothing matches. If omitted, the result is the #N/A error.",
      match_mode: "0 or omitted finds an exact match; -1 or 1 falls back to the next smaller or larger value; 2 matches wildcards.",
      search_mode: "1 or omitted searches from the first value; -1 searches from the last; 2 or -2 runs a binary search on sorted data.",
    },
  },
  XMATCH: {
    summary: "Returns the position of a value inside a range or array, in any direction and with a choice of match.",
    args: {
      lookup_value: "The value to look for in lookup_array.",
      lookup_array: "The range or array to search.",
      match_mode: "0 or omitted finds an exact match; -1 or 1 falls back to the next smaller or larger value; 2 matches wildcards.",
      search_mode: "1 or omitted searches from the first value; -1 searches from the last; 2 or -2 runs a binary search on sorted data.",
    },
  },
  XNPV: {
    summary: "Returns the net present value of cash flows that fall on dates you give, at any spacing.",
  },
  XOR: {
    summary: "Returns TRUE when an odd number of its arguments are TRUE.",
  },
  YEAR: {
    summary: "Returns the year of a date, as a number from 1900 to 9999.",
    args: {
      serial_number: "The date whose year you want.",
    },
  },
  YEARFRAC: {
    summary: "Returns the share of a year that lies between two dates, on the day-count basis you choose.",
  },
  YIELD: {
    summary: "Returns the yield of a security that pays interest periodically.",
  },
  "Z.TEST": {
    summary: "Returns the one-tailed probability from a z test on a sample against a hypothesised mean.",
  },
};

/** The help entry for a (case-insensitive) function name, or undefined. */
export function spreadsheetFunctionDoc(name: string): SpreadsheetFunctionDoc | undefined {
  const upper = name.toUpperCase();
  return Object.prototype.hasOwnProperty.call(FUNCTION_DOCS, upper)
    ? FUNCTION_DOCS[upper]
    : undefined;
}

/**
 * The help line for one argument of one function, or null when there is none.
 *
 * The label is taken as the signature spells it, so the optional-argument
 * brackets are stripped first. A repeating argument is documented once at its
 * base position, so a label that misses is retried once with its trailing
 * digits stripped: the bare base for a function that spells the repeat without
 * a number, and otherwise the second slot, which is the one that carries the
 * "and so on" wording. So number3 reads number2's line, and criteria_range4
 * reads criteria_range2's.
 */
export function spreadsheetArgumentHelp(name: string, argLabel: string): string | null {
  const doc = spreadsheetFunctionDoc(name);
  if (!doc || !doc.args) return null;
  const args = doc.args;
  const label = argLabel.replace(/^\[/, "").replace(/\]$/, "");
  if (Object.prototype.hasOwnProperty.call(args, label)) {
    return args[label];
  }
  const base = label.replace(/\d+$/, "");
  if (base === label) return null;
  if (Object.prototype.hasOwnProperty.call(args, base)) {
    return args[base];
  }
  const repeated = `${base}2`;
  if (repeated !== label && Object.prototype.hasOwnProperty.call(args, repeated)) {
    return args[repeated];
  }
  return null;
}
