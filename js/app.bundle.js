/* Data Practice Lab app bundle. GPL-3.0. Generated from source files; do not edit by hand. */
/* ---------- Excel practice sheet ---------- */
const XL_SHEET = [
  ["Name","Department","City","Sales","Target","Join Date",null,"Department","Manager",null,"Product","Price","Qty"],
  ["Aarav","Sales","Mumbai",52000,50000,"15/03/2021",null,"Sales","Neha",null,"Pen",10,50],
  ["Priya","HR","Pune",31000,35000,"01/07/2020",null,"HR","Rahul",null,"Notebook",40,30],
  ["Rohan","IT","Delhi",47000,45000,"20/11/2019",null,"IT","Sameer",null,"Bag",500,4],
  ["Sneha","Sales","Pune",61000,55000,"05/01/2022",null,null,null,null,"Bottle",150,10],
  ["Vikram","IT","Mumbai",39000,40000,"12/08/2018",null,null,null,null,null,null,null],
  ["Ananya","Sales","Delhi",45000,50000,"23/02/2023",null,null,null,null,null,null,null],
  ["Karan","HR","Mumbai",28000,30000,"30/09/2021",null,null,null,null,"Min Sales","Grade",null],
  ["Meera","IT","Pune",58000,50000,"18/04/2020",null,null,null,null,0,"C",null],
  ["Arjun","Sales","Mumbai",49000,45000,"07/12/2019",null,null,null,null,45000,"B",null],
  ["Divya","HR","Delhi",36000,35000,"26/06/2022",null,null,null,null,55000,"A",null]
];
const XL_COLS = "ABCDEFGHIJKLM".split("");

/* Function catalog used for autocomplete and argument tips */
const XL_FUNCS = {
  SUM:["number1, [number2], …","Adds numbers"], AVERAGE:["number1, [number2], …","Average of numbers"], MIN:["number1, [number2], …","Smallest number"], MAX:["number1, [number2], …","Largest number"],
  COUNT:["value1, [value2], …","Counts cells with numbers"], COUNTA:["value1, [value2], …","Counts non-empty cells"], COUNTBLANK:["range","Counts empty cells"],
  MEDIAN:["number1, [number2], …","Middle value"], MODE:["number1, [number2], …","Most common number"], LARGE:["array, k","k-th largest value"], SMALL:["array, k","k-th smallest value"],
  RANK:["number, ref, [order]","Position of a number in a list"], PRODUCT:["number1, [number2], …","Multiplies numbers"], SUMPRODUCT:["array1, [array2], …","Multiplies ranges row by row, then adds"],
  ROUND:["number, num_digits","Rounds to n decimals"], ROUNDUP:["number, num_digits","Rounds up"], ROUNDDOWN:["number, num_digits","Rounds down"], INT:["number","Drops decimals (rounds down)"], TRUNC:["number, [num_digits]","Cuts off decimals"],
  MOD:["number, divisor","Remainder after division"], ABS:["number","Removes the minus sign"], POWER:["number, power","Number raised to a power"], SQRT:["number","Square root"],
  CEILING:["number, significance","Rounds up to a multiple"], FLOOR:["number, significance","Rounds down to a multiple"], MROUND:["number, multiple","Rounds to nearest multiple"],
  IF:["logical_test, value_if_true, [value_if_false]","Choose a value based on a test"], IFS:["test1, value1, [test2, value2], …","Checks tests in order"], SWITCH:["expression, value1, result1, …, [default]","Match a value to a result"],
  AND:["logical1, [logical2], …","TRUE if all are TRUE"], OR:["logical1, [logical2], …","TRUE if any is TRUE"], NOT:["logical","Flips TRUE/FALSE"], XOR:["logical1, [logical2], …","TRUE if an odd number are TRUE"],
  IFERROR:["value, value_if_error","Replace errors with a value"], IFNA:["value, value_if_na","Replace #N/A with a value"],
  ISBLANK:["value","Is the cell empty?"], ISNUMBER:["value","Is it a number?"], ISTEXT:["value","Is it text?"], ISERROR:["value","Is it an error?"],
  COUNTIF:["range, criteria","Counts cells matching a rule"], COUNTIFS:["criteria_range1, criteria1, …","Counts rows matching several rules"],
  SUMIF:["range, criteria, [sum_range]","Adds cells matching a rule"], SUMIFS:["sum_range, criteria_range1, criteria1, …","Adds with several rules"],
  AVERAGEIF:["range, criteria, [average_range]","Average with a rule"], AVERAGEIFS:["average_range, criteria_range1, criteria1, …","Average with several rules"],
  MAXIFS:["max_range, criteria_range1, criteria1, …","Largest value with rules"], MINIFS:["min_range, criteria_range1, criteria1, …","Smallest value with rules"],
  SUBTOTAL:["function_num, ref1, …","Subtotal (9 = SUM, 1 = AVERAGE)"],
  VLOOKUP:["lookup_value, table_array, col_index_num, [range_lookup]","Look up in the first column, return from another"], HLOOKUP:["lookup_value, table_array, row_index_num, [range_lookup]","Look up across the top row"],
  XLOOKUP:["lookup_value, lookup_array, return_array, [if_not_found]","Modern lookup"], MATCH:["lookup_value, lookup_array, [match_type]","Position of a value"],
  INDEX:["array, row_num, [column_num]","Value at a position"], CHOOSE:["index_num, value1, [value2], …","Pick from a list by number"], OFFSET:["reference, rows, cols, [height], [width]","Cell moved from a starting cell"],
  ROW:["[reference]","Row number"], COLUMN:["[reference]","Column number"], ROWS:["array","Number of rows"], COLUMNS:["array","Number of columns"],
  CONCATENATE:["text1, [text2], …","Joins text"], CONCAT:["text1, [text2], …","Joins text and ranges"], TEXTJOIN:["delimiter, ignore_empty, text1, …","Joins with a separator"],
  LEFT:["text, [num_chars]","First characters"], RIGHT:["text, [num_chars]","Last characters"], MID:["text, start_num, num_chars","Characters from the middle"], LEN:["text","Number of characters"],
  UPPER:["text","CAPITALS"], LOWER:["text","small letters"], PROPER:["text","Title Case"], TRIM:["text","Removes extra spaces"], CLEAN:["text","Removes non-printing characters"],
  FIND:["find_text, within_text, [start_num]","Position of text (case-sensitive)"], SEARCH:["find_text, within_text, [start_num]","Position of text (any case)"],
  SUBSTITUTE:["text, old_text, new_text, [instance_num]","Replace text by matching"], REPLACE:["old_text, start_num, num_chars, new_text","Replace text by position"],
  TEXT:["value, format_text","Number as formatted text"], VALUE:["text","Text to number"], EXACT:["text1, text2","Exact (case-sensitive) match"], REPT:["text, number_times","Repeat text"], CHAR:["number","Character from a code"], CODE:["text","Code of first character"],
  TODAY:["","Today's date"], NOW:["","Date and time now"], DATE:["year, month, day","Builds a date"], YEAR:["serial_number","Year of a date"], MONTH:["serial_number","Month of a date"], DAY:["serial_number","Day of a date"],
  WEEKDAY:["serial_number, [return_type]","Day of week (1 = Sunday)"], EDATE:["start_date, months","Date n months later"], EOMONTH:["start_date, months","Last day of a month"],
  DATEDIF:["start_date, end_date, unit","Difference in Y, M or D"], DAYS:["end_date, start_date","Days between two dates"], NETWORKDAYS:["start_date, end_date, [holidays]","Working days between dates"],
  WORKDAY:["start_date, days, [holidays]","Date after n working days"], YEARFRAC:["start_date, end_date, [basis]","Fraction of a year"],
  PMT:["rate, nper, pv, [fv], [type]","Loan payment per period"], FV:["rate, nper, pmt, [pv], [type]","Future value"], PV:["rate, nper, pmt, [fv], [type]","Present value"], NPER:["rate, pmt, pv, [fv], [type]","Number of periods"],
  STDEV:["number1, [number2], …","Standard deviation (sample)"], VAR:["number1, [number2], …","Variance (sample)"], PI:["","3.14159…"], RAND:["","Random number 0–1"], RANDBETWEEN:["bottom, top","Random whole number"],
  LOG:["number, [base]","Logarithm"], LN:["number","Natural log"], EXP:["number","e to a power"], SIGN:["number","1, 0 or -1"], EVEN:["number","Round up to even"], ODD:["number","Round up to odd"], TRUE:["","Logical TRUE"], FALSE:["","Logical FALSE"]
};

const X = (id, group, title, learn, syntax, ex, exNote, task, answer, must, hint, extra = {}) =>
  ({id, track:"excel", group, title, learn, syntax, example:{code:ex, note:exNote}, task, answer, must, hints:[hint, "Type " + answer], ...extra});

const EXCEL_LESSONS = [
X("e1","Start here","What is a formula?",
`<p>A spreadsheet is a big grid. <b>Columns</b> have letters (A, B, C…) and <b>rows</b> have numbers (1, 2, 3…). Each box is a <b>cell</b>, named by its column letter plus its row number. The cell in column D, row 2 is <code>D2</code>.</p>
<p>A <b>formula</b> tells Excel to calculate something. Every formula starts with <code>=</code>.</p>
<p><b>Tip:</b> you don't have to type cell names. After <code>=</code>, <b>click a cell</b> in the sheet and its name is added for you, just like in Excel.</p>`,
"=cell + cell","=E2+E3","Adds Aarav's target and Priya's target: 85000.",
"Add Aarav's sales (D2) and Priya's sales (D3).","=D2+D3",[],"Type =, click D2, type +, click D3."),

X("e2","Start here","Math: + − * / ^",
`<ul><li><code>+</code> add, <code>-</code> subtract</li><li><code>*</code> multiply, <code>/</code> divide</li><li><code>^</code> power (<code>=2^3</code> gives 8)</li></ul>
<p>Brackets are calculated first, then <code>^</code>, then <code>* /</code>, then <code>+ -</code>. <code>=2+3*4</code> gives 14 but <code>=(2+3)*4</code> gives 20.</p>`,
"=D2-E2","=D3*2","Doubles Priya's sales: 62000.",
"How much did Aarav sell above his target? Subtract E2 from D2.","=D2-E2",[],"Sales are in column D, targets in column E. Aarav is row 2."),

X("e24","Start here","Percentages",
`<p>A percentage is just a division. 52000 out of 50000 is <code>=D2/E2</code> = 1.04. In Excel you would then click the <b>%</b> button to show it as 104%.</p>
<p>To get the plain number 104, multiply by 100.</p>`,
"=part/whole*100","=D3/E3*100","Priya reached 88.57% of her target.",
"What percent of his target did Rohan reach? Use D4 and E4, and multiply by 100.","=D4/E4*100",[],"Divide sales by target, then *100."),

X("e3","Basic functions","SUM",
`<p>A <b>function</b> is a ready-made calculation. Write its name, then put what it needs inside brackets.</p>
<p><code>SUM</code> adds numbers. Give it a <b>range</b>: <i>first cell</i><code>:</code><i>last cell</i>. <code>D2:D11</code> means every cell from D2 down to D11.</p>
<p><b>Like Excel:</b> type <code>=su</code>, press <b>Tab</b> to complete <code>SUM(</code>, then <b>drag</b> over D2 to D11 in the sheet (or click the column letter <b>D</b>). Close with <code>)</code>.</p>`,
"=SUM(first_cell:last_cell)","=SUM(E2:E11)","Adds all ten targets.",
"Find the total sales of all employees (D2 to D11).","=SUM(D2:D11)",["SUM"],"Type =su, press Tab, then drag over D2:D11."),

X("e25","Basic functions","SUM of separate cells",
`<p>SUM can take several things at once, separated by commas: single cells, ranges, or numbers. <code>=SUM(D2,D5,D9)</code> adds three cells.</p>
<p>Hold <b>Ctrl</b> in real Excel to click cells that are not next to each other.</p>`,
"=SUM(cell, cell, range)","=SUM(D2:D4,D9)","Adds D2, D3, D4 and D9.",
"Add the sales of Aarav (D2), Sneha (D5) and Arjun (D10) only.","=SUM(D2,D5,D10)",["SUM"],"Separate the three cells with commas."),

X("e4","Basic functions","AVERAGE",
`<p><code>AVERAGE</code> adds the numbers and divides by how many there are. It answers “what is typical?”.</p>`,
"=AVERAGE(range)","=AVERAGE(E2:E11)","Average target.",
"Find the average sales (D2 to D11).","=AVERAGE(D2:D11)",["AVERAGE"],"Same range as SUM, different function."),

X("e5","Basic functions","MIN and MAX",
`<p><code>MIN</code> gives the smallest number in a range, <code>MAX</code> the largest.</p>`,
"=MIN(range)   =MAX(range)","=MAX(D2:D11)","Highest sales: 61000 (Sneha).",
"Find the lowest sales figure in D2:D11.","=MIN(D2:D11)",["MIN"],"Lowest means smallest."),

X("e6","Basic functions","COUNT and COUNTA",
`<p><code>COUNT</code> counts cells with <b>numbers</b>. <code>COUNTA</code> counts cells that are <b>not empty</b>. Names are text, so <code>=COUNT(A2:A11)</code> gives 0 but <code>=COUNTA(A2:A11)</code> gives 10.</p>`,
"=COUNTA(range)","=COUNT(D2:D11)","10 cells hold numbers.",
"How many employee names are in A2:A11?","=COUNTA(A2:A11)",["COUNTA"],"Names are text. Which count works for text?"),

X("e26","Basic functions","COUNTBLANK",
`<p><code>COUNTBLANK</code> counts empty cells. Useful to find missing data.</p>`,
"=COUNTBLANK(range)","=COUNTBLANK(G1:G11)","Column G is completely empty: 11.",
"How many cells in the Manager column I2:I11 are empty?","=COUNTBLANK(I2:I11)",["COUNTBLANK"],"Use COUNTBLANK on I2:I11."),

X("e27","Basic functions","MEDIAN and MODE",
`<p><code>MEDIAN</code> is the middle value when numbers are sorted. It ignores very high or low outliers better than AVERAGE.</p><p><code>MODE</code> is the number that appears most often.</p>`,
"=MEDIAN(range)","=MODE(E2:E11)","50000 appears most often as a target.",
"Find the median of the sales figures D2:D11.","=MEDIAN(D2:D11)",["MEDIAN"],"MEDIAN takes a range."),

X("e28","Basic functions","LARGE and SMALL",
`<p><code>LARGE(range, 2)</code> gives the 2nd highest value. <code>SMALL(range, 3)</code> gives the 3rd lowest. With k = 1 they work like MAX and MIN.</p>`,
"=LARGE(range, k)","=SMALL(D2:D11,2)","2nd lowest sales: 31000.",
"Find the 2nd highest sales figure.","=LARGE(D2:D11,2)",["LARGE"],"k is 2."),

X("e29","Basic functions","RANK",
`<p><code>RANK(number, list)</code> tells you the position of a number in a list. 1 = the highest. Add a third input <code>1</code> to rank from lowest instead.</p>`,
"=RANK(number, range)","=RANK(D5,D2:D11)","Sneha ranks 1st.",
"What is Meera's (D9) rank in sales?","=RANK(D9,D2:D11)",["RANK"],"The number is D9, the list is D2:D11."),

X("e30","Basic functions","PRODUCT and SUMPRODUCT",
`<p><code>PRODUCT</code> multiplies numbers. <code>SUMPRODUCT</code> multiplies two ranges row by row, then adds the results.</p>
<p>The small table in <b>K1:M5</b> has prices and quantities. Total value = 10×50 + 40×30 + 500×4 + 150×10.</p>`,
"=SUMPRODUCT(range1, range2)","=PRODUCT(L2,M2)","Pen: 10 × 50 = 500.",
"Find the total value of all products: prices L2:L5 times quantities M2:M5, added up.","=SUMPRODUCT(L2:L5,M2:M5)",["SUMPRODUCT"],"Give SUMPRODUCT the price range and the qty range."),

X("e7","Rounding & math","ROUND",
`<p><code>ROUND</code> trims a number to the decimals you choose. You can put a formula inside another: <code>=ROUND(D2/E2, 2)</code> divides first, then rounds.</p><p>Use a negative number to round to tens, hundreds, thousands: <code>=ROUND(47350,-3)</code> gives 47000.</p>`,
"=ROUND(number, decimals)","=ROUND(10/3,1)","3.3",
"Divide Aarav's sales (D2) by his target (E2) and round to 2 decimals.","=ROUND(D2/E2,2)",["ROUND"],"The number to round is D2/E2."),

X("e31","Rounding & math","ROUNDUP, ROUNDDOWN, INT",
`<ul><li><code>ROUNDUP</code> always rounds away from zero</li><li><code>ROUNDDOWN</code> always rounds toward zero</li><li><code>INT</code> drops the decimals (rounds down)</li></ul>`,
"=ROUNDUP(number, decimals)","=ROUNDDOWN(D3/E3,1)","0.88… becomes 0.8",
"Round up the average sales (AVERAGE of D2:D11) to the nearest thousand (use -3).","=ROUNDUP(AVERAGE(D2:D11),-3)",["ROUNDUP"],"Put AVERAGE(D2:D11) inside ROUNDUP with -3."),

X("e32","Rounding & math","MOD and ABS",
`<p><code>MOD(a, b)</code> gives the remainder after dividing a by b. <code>MOD(10,3)</code> = 1. It's handy for “every 2nd row” or “leftover” questions.</p><p><code>ABS</code> removes a minus sign: <code>ABS(-5)</code> = 5.</p>`,
"=MOD(number, divisor)","=ABS(D3-E3)","Priya missed her target by 4000.",
"Vikram (row 6) missed his target. Show the gap as a positive number using ABS.","=ABS(D6-E6)",["ABS"],"D6-E6 is negative. Wrap it in ABS."),

X("e33","Rounding & math","POWER and SQRT",
`<p><code>POWER(5,2)</code> = 25 (same as <code>5^2</code>). <code>SQRT(25)</code> = 5.</p>`,
"=POWER(number, power)","=SQRT(144)","12",
"Square root of Rohan's sales (D4), rounded to 1 decimal.","=ROUND(SQRT(D4),1)",["SQRT","ROUND"],"Put SQRT(D4) inside ROUND(…,1)."),

X("e34","Rounding & math","CEILING, FLOOR, MROUND",
`<p>These round to a <b>multiple</b>:</p><ul><li><code>CEILING(47200, 1000)</code> = 48000 (up)</li><li><code>FLOOR(47200, 1000)</code> = 47000 (down)</li><li><code>MROUND(47600, 5000)</code> = 50000 (nearest)</li></ul>`,
"=CEILING(number, multiple)","=FLOOR(D4,10000)","47000 → 40000",
"Round Priya's sales (D3) to the nearest multiple of 5000.","=MROUND(D3,5000)",["MROUND"],"Nearest multiple means MROUND."),

X("e8","Rounding & math","Locking cells with $",
`<p>When you copy a formula down, Excel shifts the cell names: <code>D2</code> becomes <code>D3</code>. To keep one part fixed, add <code>$</code>: <code>$D$12</code> never moves. Press <b>F4</b> in Excel to add the $ signs.</p>`,
"=D2/SUM($D$2:$D$11)","=E2/SUM($E$2:$E$11)","Aarav's share of total target.",
"What share of total sales did Aarav make? Divide D2 by SUM($D$2:$D$11) and round to 2 decimals.","=ROUND(D2/SUM($D$2:$D$11),2)",["$D$2"],"The locked range is $D$2:$D$11."),

X("e9","Decisions","IF",
`<p><code>IF</code> checks a condition. If it's true it shows one thing, otherwise another.</p>
<p>Comparisons: <code>=</code> <code>&lt;&gt;</code> (not equal) <code>&gt;</code> <code>&lt;</code> <code>&gt;=</code> <code>&lt;=</code>. Text goes in double quotes: <code>"Met"</code>.</p>`,
'=IF(condition, if_true, if_false)','=IF(D2>=E2,"Met","Missed")','Aarav: "Met".',
'Did Priya (row 3) meet her target? Show "Met" or "Missed".','=IF(D3>=E3,"Met","Missed")',["IF"],"Compare D3 with E3."),

X("e10","Decisions","AND and OR",
`<p><code>AND</code>: all conditions true. <code>OR</code>: at least one true. Usually placed inside IF.</p>`,
'=IF(AND(cond1, cond2), "yes", "no")','=IF(OR(C2="Pune",C2="Delhi"),"Yes","No")','"No" for Aarav (Mumbai).',
'Rohan is row 4. Show "Bonus" if he is in IT AND sales D4 > 45000, otherwise "No bonus".','=IF(AND(B4="IT",D4>45000),"Bonus","No bonus")',["AND","IF"],'Conditions: B4="IT" and D4>45000.'),

X("e35","Decisions","Nested IF (grades)",
`<p>Put an IF inside another IF to handle more than two outcomes. Excel checks from left to right and stops at the first true one.</p>`,
'=IF(test1, a, IF(test2, b, c))','=IF(D3>=50000,"A",IF(D3>=40000,"B","C"))','Priya: 31000 → "C".',
'Grade Sneha (D5): "A" if ≥ 60000, "B" if ≥ 45000, otherwise "C".','=IF(D5>=60000,"A",IF(D5>=45000,"B","C"))',["IF"],"Start with the highest limit."),

X("e36","Decisions","IFS",
`<p><code>IFS</code> is a cleaner way to write nested IFs: pairs of <i>test, result</i>. End with <code>TRUE, "something"</code> as the “otherwise”.</p>`,
'=IFS(test1, result1, test2, result2, TRUE, else)','=IFS(D4>=50000,"A",D4>=45000,"B",TRUE,"C")','Rohan → "B".',
'Grade Arjun (D10) with IFS: "A" if ≥ 60000, "B" if ≥ 45000, otherwise "C".','=IFS(D10>=60000,"A",D10>=45000,"B",TRUE,"C")',["IFS"],"Three pairs; the last test is TRUE."),

X("e37","Decisions","SWITCH",
`<p><code>SWITCH</code> compares one value with a list of options and returns the matching result. The last input (optional) is the default.</p>`,
'=SWITCH(value, option1, result1, option2, result2, default)','=SWITCH(C2,"Mumbai","West","Delhi","North","Other")','Aarav → "West".',
'Give Priya (B3) a department code: Sales → "S", HR → "H", IT → "T".','=SWITCH(B3,"Sales","S","HR","H","IT","T")',["SWITCH"],"SWITCH(B3, then pairs)."),

X("e38","Decisions","NOT and IS functions",
`<ul><li><code>NOT(test)</code> flips TRUE to FALSE</li><li><code>ISBLANK(cell)</code>: is it empty?</li><li><code>ISNUMBER</code> / <code>ISTEXT</code>: what type is it?</li></ul>`,
'=ISNUMBER(cell)','=NOT(C2="Pune")','TRUE (Aarav is not in Pune).',
'Is cell H6 empty? Use ISBLANK.','=ISBLANK(H6)',["ISBLANK"],"ISBLANK takes one cell."),

X("e11","Conditional count & sum","COUNTIF",
`<p><code>COUNTIF</code> counts cells that match a rule: first the range to check, then the rule. The rule can be text (<code>"IT"</code>) or a comparison in quotes (<code>"&gt;50000"</code>).</p>`,
'=COUNTIF(range, rule)','=COUNTIF(C2:C11,"Mumbai")','4 people in Mumbai.',
'How many employees are in IT? (B2:B11)','=COUNTIF(B2:B11,"IT")',["COUNTIF"],'Range B2:B11, rule "IT".'),

X("e12","Conditional count & sum","COUNTIF with numbers",
`<p>Number rules go inside quotes: <code>"&gt;50000"</code>, <code>"&lt;=30000"</code>, <code>"&lt;&gt;0"</code>. To use a cell in the rule, join it: <code>"&gt;"&amp;E2</code>.</p>`,
'=COUNTIF(range, ">number")','=COUNTIF(E2:E11,">=50000")','Targets of 50000 or more.',
'How many employees sold more than 50000?','=COUNTIF(D2:D11,">50000")',["COUNTIF"],'The rule is ">50000".'),

X("e40","Conditional count & sum","Wildcards * and ?",
`<p>In COUNTIF, SUMIF and lookups, <code>*</code> means “any characters” and <code>?</code> means “exactly one character”.</p><ul><li><code>"A*"</code>: starts with A</li><li><code>"*a"</code>: ends with a</li><li><code>"?????"</code>: exactly 5 letters</li></ul>`,
'=COUNTIF(range, "A*")','=COUNTIF(A2:A11,"*a")','Names ending in a: 4.',
'How many names start with A?','=COUNTIF(A2:A11,"A*")',["COUNTIF"],'The rule is "A*".'),

X("e13","Conditional count & sum","SUMIF",
`<p><code>SUMIF</code> adds only matching rows. It needs: the range to <b>check</b>, the <b>rule</b>, and the range to <b>add</b>.</p>`,
'=SUMIF(check_range, rule, sum_range)','=SUMIF(B2:B11,"HR",D2:D11)','HR total sales.',
'Total sales from Pune (cities C2:C11, sales D2:D11).','=SUMIF(C2:C11,"Pune",D2:D11)',["SUMIF"],'Check C2:C11 for "Pune", add D2:D11.'),

X("e14","Conditional count & sum","AVERAGEIF",
`<p>Same as SUMIF, but averages the matching rows.</p>`,
'=AVERAGEIF(check_range, rule, average_range)','=AVERAGEIF(C2:C11,"Delhi",D2:D11)','Average sales in Delhi.',
'Average sales of the Sales department.','=AVERAGEIF(B2:B11,"Sales",D2:D11)',["AVERAGEIF"],'Check B2:B11 for "Sales".'),

X("e15","Conditional count & sum","SUMIFS",
`<p>More than one rule? Add an S. In <code>SUMIFS</code> the range to add comes <b>first</b>, then pairs of <i>range, rule</i>.</p>`,
'=SUMIFS(sum_range, range1, rule1, range2, rule2)','=SUMIFS(D2:D11,B2:B11,"IT",D2:D11,">40000")','IT sales above 40000.',
'Total sales of the Sales department in Mumbai.','=SUMIFS(D2:D11,B2:B11,"Sales",C2:C11,"Mumbai")',["SUMIFS"],'D2:D11 first, then B2:B11,"Sales", then C2:C11,"Mumbai".'),

X("e41","Conditional count & sum","COUNTIFS",
`<p><code>COUNTIFS</code> counts rows where <b>all</b> rules are true. Pairs of <i>range, rule</i>.</p>`,
'=COUNTIFS(range1, rule1, range2, rule2)','=COUNTIFS(B2:B11,"IT",C2:C11,"Pune")','IT people in Pune: 1.',
'How many Sales department people sold 50000 or more?','=COUNTIFS(B2:B11,"Sales",D2:D11,">=50000")',["COUNTIFS"],'Rules: B2:B11 = "Sales", D2:D11 ">=50000".'),

X("e42","Conditional count & sum","AVERAGEIFS",
`<p>Average with several rules. The range to average comes first.</p>`,
'=AVERAGEIFS(avg_range, range1, rule1, …)','=AVERAGEIFS(D2:D11,B2:B11,"HR",C2:C11,"Mumbai")','HR in Mumbai: 28000.',
'Average sales of IT people who are NOT in Delhi.','=AVERAGEIFS(D2:D11,B2:B11,"IT",C2:C11,"<>Delhi")',["AVERAGEIFS"],'"<>Delhi" means not Delhi.'),

X("e43","Conditional count & sum","MAXIFS and MINIFS",
`<p>The largest or smallest value among matching rows. Same pattern as SUMIFS.</p>`,
'=MAXIFS(max_range, range1, rule1)','=MINIFS(D2:D11,C2:C11,"Pune")','Lowest Pune sales: 31000.',
'Highest sales in the HR department.','=MAXIFS(D2:D11,B2:B11,"HR")',["MAXIFS"],'Max range D2:D11, rule B2:B11 = "HR".'),

X("e44","Conditional count & sum","SUBTOTAL",
`<p><code>SUBTOTAL</code> does a calculation chosen by a number: 9 = SUM, 1 = AVERAGE, 2 = COUNT, 4 = MAX, 5 = MIN. Its special power in Excel: it ignores rows hidden by a filter.</p>`,
'=SUBTOTAL(9, range)','=SUBTOTAL(1,D2:D11)','Average sales.',
'Use SUBTOTAL to find the total of the targets E2:E11.','=SUBTOTAL(9,E2:E11)',["SUBTOTAL"],"9 means SUM."),

X("e16","Lookup","VLOOKUP",
`<p><code>VLOOKUP</code> finds a value in the <b>first column</b> of a table and returns something from the same row. The table <b>H1:I4</b> lists department managers.</p>
<ol><li>What to find: "IT"</li><li>Where: H2:I4</li><li>Which column to return: 2</li><li>FALSE = exact match. Always use FALSE as a beginner.</li></ol>`,
'=VLOOKUP(find, table, column_number, FALSE)','=VLOOKUP("HR",H2:I4,2,FALSE)','"Rahul".',
'Who is the manager of IT?','=VLOOKUP("IT",H2:I4,2,FALSE)',["VLOOKUP"],'Find "IT" in H2:I4, return column 2.'),

X("e45","Lookup","XLOOKUP",
`<p><code>XLOOKUP</code> is the modern lookup in Excel 365/2021. Give it: what to find, the column to search, the column to return. It can look left or right, and has a built-in “not found” value.</p>`,
'=XLOOKUP(find, search_range, return_range, [if_not_found])','=XLOOKUP("Rohan",A2:A11,D2:D11)','47000',
"Which city does Karan live in? Use XLOOKUP.",'=XLOOKUP("Karan",A2:A11,C2:C11)',["XLOOKUP"],'Search A2:A11, return C2:C11.'),

X("e46","Lookup","HLOOKUP",
`<p><code>HLOOKUP</code> is VLOOKUP turned sideways: it searches the <b>top row</b> and returns from a row below.</p>`,
'=HLOOKUP(find, table, row_number, FALSE)','=HLOOKUP("City",A1:F11,2,FALSE)','Row 2 of the City column: "Mumbai".',
'Use HLOOKUP on A1:F11 to find the Target in the 5th row of the table (Sneha).','=HLOOKUP("Target",A1:F11,5,FALSE)',["HLOOKUP"],'Find "Target" in the top row, row number 5.'),

X("e17","Lookup","IFERROR",
`<p>If a lookup fails you get <code>#N/A</code>. <code>IFERROR</code> catches any error and shows your own message.</p>`,
'=IFERROR(formula, value_if_error)','=IFERROR(D2/0,"Check data")','"Check data".',
'Look up the manager of "Finance" (not in the table) and show "Not found" instead of an error.','=IFERROR(VLOOKUP("Finance",H2:I4,2,FALSE),"Not found")',["IFERROR","VLOOKUP"],'Wrap the VLOOKUP inside IFERROR(…, "Not found").'),

X("e47","Lookup","MATCH",
`<p><code>MATCH</code> tells you <b>which position</b> a value is at in a list. Use 0 at the end for exact match.</p>`,
'=MATCH(find, range, 0)','=MATCH("Pune",C2:C11,0)','First Pune is position 2.',
'At which position is "Vikram" in A2:A11?','=MATCH("Vikram",A2:A11,0)',["MATCH"],"Search A2:A11, exact match 0."),

X("e48","Lookup","INDEX",
`<p><code>INDEX(range, row, column)</code> returns the value at that position inside the range.</p>`,
'=INDEX(range, row_num, col_num)','=INDEX(A2:F11,1,4)','Row 1, column 4 of the table: 52000.',
'Return the value in row 3, column 3 of A2:F11.','=INDEX(A2:F11,3,3)',["INDEX"],"Range, then 3, then 3."),

X("e18","Lookup","INDEX + MATCH",
`<p>MATCH finds the position, INDEX returns the value at that position. Together they look up in any direction.</p>`,
'=INDEX(return_range, MATCH(find, search_range, 0))','=INDEX(C2:C11,MATCH("Karan",A2:A11,0))','"Mumbai".',
"What are Meera's sales? Use INDEX and MATCH.",'=INDEX(D2:D11,MATCH("Meera",A2:A11,0))',["INDEX","MATCH"],'Return from D2:D11, search "Meera" in A2:A11.'),

X("e49","Lookup","CHOOSE",
`<p><code>CHOOSE(n, a, b, c)</code> returns the n-th item from the list you give.</p>`,
'=CHOOSE(number, item1, item2, …)','=CHOOSE(3,"Mon","Tue","Wed")','"Wed".',
'Return the 2nd quarter name from "Q1","Q2","Q3","Q4".','=CHOOSE(2,"Q1","Q2","Q3","Q4")',["CHOOSE"],"Number 2, then the four names."),

X("e19","Text","Joining text with &",
`<p><code>&amp;</code> joins text. Put spaces or symbols in quotes: <code>=A2&amp;" "&amp;B2</code> gives "Aarav Sales".</p>`,
'=A2&" - "&C2','=A3&" works in "&B3','"Priya works in HR".',
'Join A2 and C2 with " - " between to get "Aarav - Mumbai".','=A2&" - "&C2',[],"A2, then \" - \", then C2, joined with &."),

X("e50","Text","CONCAT and TEXTJOIN",
`<p><code>CONCAT</code> joins several pieces. <code>TEXTJOIN</code> joins a whole range with a separator, and can skip empty cells (TRUE).</p>`,
'=TEXTJOIN(separator, TRUE, range)','=CONCAT(A4," (",B4,")")','"Rohan (IT)".',
'Join the first three names A2:A4 separated by ", ".','=TEXTJOIN(", ",TRUE,A2:A4)',["TEXTJOIN"],'Separator ", ", TRUE, then A2:A4.'),

X("e20","Text","LEFT and RIGHT",
`<p><code>LEFT(text, n)</code> gives the first n characters. <code>RIGHT(text, n)</code> the last n.</p>`,
'=LEFT(text, how_many)','=RIGHT(A2,3)','"rav".',
'Make a 3-letter code from Aarav\'s city (C2): "Mum".','=LEFT(C2,3)',["LEFT"],"First 3 characters of C2."),

X("e51","Text","MID",
`<p><code>MID(text, start, n)</code> takes n characters starting at position start.</p>`,
'=MID(text, start, how_many)','=MID(C3,2,3)','"Pune" → "une".',
'Take 3 characters from Karan (A8), starting at position 2.','=MID(A8,2,3)',["MID"],"MID(A8, 2, 3)."),

X("e21","Text","UPPER, LOWER, PROPER",
`<ul><li><code>UPPER</code>: ALL CAPS</li><li><code>LOWER</code>: all small</li><li><code>PROPER</code>: First Letter Capital</li></ul>`,
'=UPPER(text)','=LOWER(B2)','"sales".',
"Show Sneha's name (A5) in capitals.","=UPPER(A5)",["UPPER"],"Which one makes ALL CAPS?"),

X("e52","Text","LEN and TRIM",
`<p><code>LEN</code> counts characters (spaces count too). <code>TRIM</code> removes extra spaces, which is very common in messy data.</p>`,
'=LEN(text)','=TRIM("  Aarav   Shah ")','"Aarav Shah".',
'How many characters are in Ananya\'s name (A7)?','=LEN(A7)',["LEN"],"LEN takes one cell."),

X("e53","Text","FIND and SEARCH",
`<p>Both give the <b>position</b> of some text. <code>FIND</code> is case-sensitive, <code>SEARCH</code> is not and allows wildcards.</p>`,
'=FIND(what, where)','=SEARCH("A",A2)','1 (SEARCH ignores case).',
'At which position is the letter "e" in Meera (A9)? Use FIND.','=FIND("e",A9)',["FIND"],'FIND("e", A9).'),

X("e54","Text","SUBSTITUTE and REPLACE",
`<p><code>SUBSTITUTE</code> swaps text you name. <code>REPLACE</code> swaps text at a position.</p>`,
'=SUBSTITUTE(text, old, new)','=REPLACE(A2,1,1,"E")','"Earav".',
'Change "Mumbai" to "Bombay" in C2.','=SUBSTITUTE(C2,"Mumbai","Bombay")',["SUBSTITUTE"],'SUBSTITUTE(C2, "Mumbai", "Bombay").'),

X("e55","Text","TEXT and VALUE",
`<p><code>TEXT(number, format)</code> turns a number or date into formatted text, for example <code>"0.00"</code> or <code>"dd/mm/yyyy"</code>. Real Excel also understands <code>"#,##0"</code>, <code>"0%"</code> and <code>"mmm"</code>.</p><p><code>VALUE</code> does the opposite: text that looks like a number becomes a real number.</p>`,
'=TEXT(value, "format")','=TEXT(D2/E2,"0.00")','"1.04" (as text).',
'Show Aarav\'s join date F2 as text in the format "dd/mm/yyyy".','=TEXT(F2,"dd/mm/yyyy")',["TEXT"],'Format text is "dd/mm/yyyy".'),

X("e22","Dates","YEAR, MONTH, DAY",
`<p>Excel stores dates as numbers (days since 1900), so it can do math with them. These pull out one part:</p><ul><li><code>YEAR</code>, <code>MONTH</code>, <code>DAY</code></li></ul>`,
'=YEAR(date)','=MONTH(F3)','7 (July).',
'In which year did Aarav join? (F2)','=YEAR(F2)',["YEAR"],"YEAR on F2."),

X("e56","Dates","TODAY and NOW",
`<p><code>TODAY()</code> gives today's date, <code>NOW()</code> the date and time. They need no inputs, but keep the brackets. They update every time the sheet recalculates.</p>`,
'=TODAY()','=TODAY()-F2','Days since Aarav joined.',
'Show the current year using YEAR and TODAY.','=YEAR(TODAY())',["YEAR","TODAY"],"Put TODAY() inside YEAR."),

X("e57","Dates","DATE",
`<p><code>DATE(year, month, day)</code> builds a date from three numbers. Here you'll see the date's serial number with the date beside it. In Excel you'd format the cell as a date.</p>`,
'=DATE(year, month, day)','=DATE(2026,12,25)','Christmas 2026.',
'Build the date 1 April 2026.','=DATE(2026,4,1)',["DATE"],"Year, month, day.",{dateResult:true}),

X("e23","Dates","Days between dates",
`<p>Subtract dates to get the number of days between them.</p>`,
'=later - earlier','=F5-F2','Days between Aarav and Sneha joining.',
'How many days after Priya (F3) did Aarav (F2) join?','=F2-F3',[],"Later date first."),

X("e58","Dates","DATEDIF",
`<p><code>DATEDIF(start, end, unit)</code> counts complete years ("Y"), months ("M") or days ("D") between two dates. Great for age or experience. The start must be the earlier date.</p>`,
'=DATEDIF(start, end, "Y")','=DATEDIF(F3,F2,"M")','8 full months.',
'Full years between Vikram joining (F6) and Ananya joining (F7).','=DATEDIF(F6,F7,"Y")',["DATEDIF"],'Start F6, end F7, unit "Y".'),

X("e59","Dates","EDATE and EOMONTH",
`<p><code>EDATE(date, n)</code>: the same day n months later. <code>EOMONTH(date, n)</code>: the last day of the month n months away (0 = same month).</p>`,
'=EOMONTH(date, months)','=EDATE(F2,6)','6 months after Aarav joined.',
'Last day of the month in which Aarav joined (F2).','=EOMONTH(F2,0)',["EOMONTH"],"Months = 0.",{dateResult:true}),

X("e60","Dates","WEEKDAY",
`<p><code>WEEKDAY(date)</code> returns 1 (Sunday) to 7 (Saturday).</p>`,
'=WEEKDAY(date)','=WEEKDAY(F3)','1 July 2020 was a Wednesday: 4.',
'Which weekday number did Aarav join on? (F2)','=WEEKDAY(F2)',["WEEKDAY"],"WEEKDAY of F2."),

X("e61","Dates","NETWORKDAYS",
`<p><code>NETWORKDAYS(start, end)</code> counts working days (Mon–Fri) between two dates, including both.</p>`,
'=NETWORKDAYS(start, end)','=NETWORKDAYS(F2,F5)','Working days between two joins.',
'Working days from Priya joining (F3) to Aarav joining (F2).','=NETWORKDAYS(F3,F2)',["NETWORKDAYS"],"Start F3, end F2."),

X("e62","Finance","PMT (loan EMI)",
`<p><code>PMT</code> gives the regular payment for a loan. Use the <b>monthly</b> rate (yearly ÷ 12) and number of <b>months</b>. Put the loan as a negative number to get a positive payment.</p>`,
'=PMT(rate/12, months, -loan)','=PMT(0.1/12,12,-100000)','EMI for 1 lakh at 10% for 1 year.',
'EMI for a 5,00,000 loan at 9% per year for 60 months. Round to 0 decimals.','=ROUND(PMT(0.09/12,60,-500000),0)',["PMT"],"rate 0.09/12, 60 months, -500000."),

X("e63","Finance","FV (savings)",
`<p><code>FV</code> gives the future value of regular savings. Payments you make are negative.</p>`,
'=FV(rate/12, months, -monthly_saving)','=FV(0.06/12,12,-1000)','Saving 1000/month for a year at 6%.',
'Future value of saving 5000 per month for 24 months at 8% per year. Round to 0 decimals.','=ROUND(FV(0.08/12,24,-5000),0)',["FV"],"rate 0.08/12, 24, -5000.")
];
/* ---------- SQL practice database ---------- */
const SQL_SCHEMA = `
CREATE TABLE employees (id INTEGER, name TEXT, department TEXT, city TEXT, salary INTEGER, join_year INTEGER, bonus INTEGER, hire_date TEXT, mentor_id INTEGER);
INSERT INTO employees VALUES
(1,'Aarav','Sales','Mumbai',52000,2021,5000,'2021-03-15',NULL),
(2,'Priya','HR','Pune',31000,2020,NULL,'2020-07-01',NULL),
(3,'Rohan','IT','Delhi',47000,2019,3000,'2019-11-20',8),
(4,'Sneha','Sales','Pune',61000,2022,6000,'2022-01-05',NULL),
(5,'Vikram','IT','Mumbai',39000,2018,NULL,'2018-08-12',8),
(6,'Ananya','Sales','Delhi',45000,2023,2000,'2023-02-23',1),
(7,'Karan','HR','Mumbai',28000,2021,NULL,'2021-09-30',2),
(8,'Meera','IT','Pune',58000,2020,4000,'2020-04-18',NULL),
(9,'Arjun','Sales','Mumbai',49000,2019,3500,'2019-12-07',4),
(10,'Divya','HR','Delhi',36000,2022,1500,'2022-06-26',2);
CREATE TABLE departments (department TEXT, manager TEXT, floor INTEGER);
INSERT INTO departments VALUES ('Sales','Neha',2),('HR','Rahul',1),('IT','Sameer',3),('Finance','Kavya',4);
CREATE TABLE projects (project_id INTEGER, project_name TEXT, employee_id INTEGER);
INSERT INTO projects VALUES (101,'Website Revamp',3),(102,'Payroll Setup',2),(103,'Mobile App',8),(104,'Client Pitch',1),(105,'Data Cleanup',8),(106,'Market Study',NULL);
CREATE TABLE sales (sale_id INTEGER, employee_id INTEGER, sale_date TEXT, amount INTEGER);
INSERT INTO sales VALUES
(1,1,'2026-01-05',12000),(2,4,'2026-01-09',18000),(3,6,'2026-01-15',9000),(4,9,'2026-01-22',15000),
(5,1,'2026-02-03',8000),(6,4,'2026-02-11',22000),(7,9,'2026-02-18',11000),(8,6,'2026-02-25',7000),
(9,1,'2026-03-04',16000),(10,4,'2026-03-12',14000),(11,9,'2026-03-19',13000),(12,4,'2026-03-28',9500);
`;
const SQL_TABLES = ["employees","departments","projects","sales"];
const SQL_WORDS = ["SELECT","FROM","WHERE","AND","OR","NOT","IN","BETWEEN","LIKE","IS","NULL","ORDER BY","GROUP BY","HAVING","LIMIT","OFFSET","DISTINCT","AS","ASC","DESC",
 "JOIN","INNER JOIN","LEFT JOIN","RIGHT JOIN","FULL OUTER JOIN","CROSS JOIN","ON","UNION","UNION ALL","INTERSECT","EXCEPT","WITH","CASE","WHEN","THEN","ELSE","END","EXISTS",
 "INSERT INTO","VALUES","UPDATE","SET","DELETE FROM","CREATE TABLE","OVER","PARTITION BY"];
const SQL_FUNCS = {COUNT:"expression",SUM:"expression",AVG:"expression",MIN:"expression",MAX:"expression",ROUND:"number, decimals",UPPER:"text",LOWER:"text",LENGTH:"text",
 SUBSTR:"text, start, length",REPLACE:"text, old, new",TRIM:"text",COALESCE:"value1, value2, …",IFNULL:"value, fallback",ABS:"number",CAST:"value AS type",
 STRFTIME:"format, date",DATE:"date, modifier",JULIANDAY:"date",ROW_NUMBER:"",RANK:"",DENSE_RANK:"",LAG:"column, [offset]",LEAD:"column, [offset]",NTILE:"n"};

const Q = (id, group, title, learn, syntax, ex, exNote, task, answer, hint, extra = {}) =>
  ({id, track:"sql", group, title, learn, syntax, example:{code:ex, note:exNote}, task, answer, hints:[hint, "Type " + answer], ...extra});

const SQL_LESSONS = [
Q("s1","Start here","Your first query",
`<p>A <b>database</b> stores data in <b>tables</b>, like sheets in Excel. <b>SQL</b> is the language for asking it questions.</p>
<ul><li><code>SELECT</code> means “show me”</li><li><code>*</code> means “all columns”</li><li><code>FROM employees</code> says which table</li></ul>
<p>Keywords work in any case. <b>Tip:</b> type a few letters and press <b>Tab</b> to complete a word. Click a column name in the tables to add it to your query.</p>`,
"SELECT * FROM table_name;","SELECT * FROM departments;","All 4 departments.",
"Show every column and every row of the employees table.","SELECT * FROM employees;","SELECT, then *, then FROM and the table."),

Q("s2","Start here","Choose columns",`<p>List only the columns you want, separated by commas.</p>`,
"SELECT col1, col2 FROM table;","SELECT name, salary FROM employees;","Names and salaries.",
"Show only the name and city of every employee.","SELECT name, city FROM employees;","List name, city after SELECT."),

Q("s3","Filter rows","WHERE with text",`<p><code>WHERE</code> keeps only matching rows. Text goes in <b>single quotes</b>: <code>'Pune'</code>.</p>`,
"WHERE column = 'value'","SELECT * FROM employees WHERE department = 'HR';","Only HR.",
"Show all columns for employees who live in Pune.","SELECT * FROM employees WHERE city = 'Pune';","Add WHERE city = 'Pune'."),

Q("s4","Filter rows","WHERE with numbers",`<p>Numbers need no quotes. Compare with <code>=</code> <code>&lt;&gt;</code> <code>&gt;</code> <code>&lt;</code> <code>&gt;=</code> <code>&lt;=</code>.</p>`,
"WHERE salary > 50000","SELECT name FROM employees WHERE join_year < 2020;","Joined before 2020.",
"Show name and salary of employees earning more than 50000.","SELECT name, salary FROM employees WHERE salary > 50000;","WHERE salary > 50000."),

Q("s5","Filter rows","AND / OR",`<p><code>AND</code>: both true. <code>OR</code>: either true. Use brackets when mixing them.</p>`,
"WHERE cond1 AND cond2","SELECT name FROM employees WHERE city = 'Delhi' OR city = 'Pune';","Delhi or Pune.",
"Names of employees in IT AND living in Pune.","SELECT name FROM employees WHERE department = 'IT' AND city = 'Pune';","department = 'IT' AND city = 'Pune'."),

Q("s21","Filter rows","NOT and <>",`<p><code>&lt;&gt;</code> (or <code>!=</code>) means “not equal”. <code>NOT</code> flips any condition: <code>NOT city = 'Pune'</code>.</p>`,
"WHERE column <> 'value'","SELECT name FROM employees WHERE NOT department = 'Sales';","Everyone outside Sales.",
"Show name and city of employees who do NOT live in Mumbai.","SELECT name, city FROM employees WHERE city <> 'Mumbai';","city <> 'Mumbai'."),

Q("s6","Filter rows","IN and BETWEEN",`<ul><li><code>city IN ('Pune','Delhi')</code>: short for many ORs</li><li><code>salary BETWEEN 40000 AND 50000</code>: a range, both ends included</li></ul>`,
"WHERE col BETWEEN low AND high","SELECT name FROM employees WHERE department IN ('HR','IT');","HR or IT.",
"Name and salary where salary is between 40000 and 50000.","SELECT name, salary FROM employees WHERE salary BETWEEN 40000 AND 50000;","BETWEEN 40000 AND 50000."),

Q("s22","Filter rows","NOT IN",`<p><code>NOT IN</code> keeps rows whose value is <b>not</b> in the list.</p>`,
"WHERE col NOT IN ('a','b')","SELECT name FROM employees WHERE city NOT IN ('Pune');","Not in Pune.",
"Names of employees not in HR or IT.","SELECT name FROM employees WHERE department NOT IN ('HR','IT');","department NOT IN ('HR','IT')."),

Q("s7","Filter rows","LIKE with %",`<p><code>LIKE</code> matches patterns. <code>%</code> = any characters. <code>'A%'</code> starts with A, <code>'%a'</code> ends with a, <code>'%ee%'</code> contains ee.</p>`,
"WHERE col LIKE 'A%'","SELECT name FROM employees WHERE name LIKE '%a';","Ends in a.",
"Names that start with A.","SELECT name FROM employees WHERE name LIKE 'A%';","Pattern 'A%'."),

Q("s23","Filter rows","LIKE with _",`<p><code>_</code> matches exactly <b>one</b> character. <code>'_r%'</code> means: any first letter, then r, then anything.</p>`,
"WHERE col LIKE '_r%'","SELECT name FROM employees WHERE name LIKE '_____';","Names with exactly 5 letters.",
"Names whose second letter is r.","SELECT name FROM employees WHERE name LIKE '_r%';","Pattern '_r%'."),

Q("s16","Filter rows","IS NULL",`<p><code>NULL</code> means no value. You can't use <code>= NULL</code>; use <code>IS NULL</code>.</p>`,
"WHERE col IS NULL","SELECT name FROM employees WHERE mentor_id IS NULL;","People without a mentor.",
"Names of employees with no bonus.","SELECT name FROM employees WHERE bonus IS NULL;","bonus IS NULL."),

Q("s24","Filter rows","IS NOT NULL",`<p>The opposite: rows that <b>do</b> have a value.</p>`,
"WHERE col IS NOT NULL","SELECT name FROM employees WHERE mentor_id IS NOT NULL;","People with a mentor.",
"Name and bonus of employees who got a bonus.","SELECT name, bonus FROM employees WHERE bonus IS NOT NULL;","bonus IS NOT NULL."),

Q("s8","Sort & limit","ORDER BY",`<p><code>ORDER BY</code> sorts. <code>ASC</code> = smallest first (default), <code>DESC</code> = largest first.</p>`,
"ORDER BY col DESC","SELECT name FROM employees ORDER BY name;","A to Z.",
"Name and salary, highest salary first.","SELECT name, salary FROM employees ORDER BY salary DESC;","ORDER BY salary DESC.",{ordered:true}),

Q("s25","Sort & limit","Sort by two columns",`<p>Sort by a first column, and use a second column to break ties: <code>ORDER BY department, salary DESC</code>.</p>`,
"ORDER BY col1, col2 DESC","SELECT name, city FROM employees ORDER BY city, name;","By city, then name.",
"Show department, name and salary sorted by department (A–Z), then salary (highest first).","SELECT department, name, salary FROM employees ORDER BY department, salary DESC;","ORDER BY department, salary DESC.",{ordered:true}),

Q("s9","Sort & limit","LIMIT (top N)",`<p><code>LIMIT 3</code> keeps the first 3 rows. With ORDER BY it answers “top 3” questions. SQL Server writes <code>SELECT TOP 3</code>.</p>`,
"ORDER BY col DESC LIMIT 3","SELECT name FROM employees ORDER BY join_year LIMIT 2;","Two longest-serving.",
"Name and salary of the top 3 earners.","SELECT name, salary FROM employees ORDER BY salary DESC LIMIT 3;","ORDER BY salary DESC LIMIT 3.",{ordered:true}),

Q("s26","Sort & limit","OFFSET (skip rows)",`<p><code>OFFSET 3</code> skips the first 3 rows. <code>LIMIT 3 OFFSET 3</code> gives rows 4–6. This is how websites show “page 2”.</p>`,
"LIMIT n OFFSET skip","SELECT name FROM employees ORDER BY name LIMIT 5 OFFSET 5;","Names 6–10.",
"Name and salary of the 4th, 5th and 6th highest earners.","SELECT name, salary FROM employees ORDER BY salary DESC LIMIT 3 OFFSET 3;","LIMIT 3 OFFSET 3.",{ordered:true}),

Q("s10","Sort & limit","DISTINCT",`<p><code>DISTINCT</code> removes duplicate rows.</p>`,
"SELECT DISTINCT col FROM table;","SELECT DISTINCT department FROM employees;","Sales, HR, IT.",
"Show the unique cities.","SELECT DISTINCT city FROM employees;","DISTINCT after SELECT."),

Q("s11","Summaries","COUNT(*)",`<p><code>COUNT(*)</code> counts rows. With WHERE it counts matching rows.</p>`,
"SELECT COUNT(*) FROM table WHERE …","SELECT COUNT(*) FROM employees;","10.",
"How many employees are in Sales?","SELECT COUNT(*) FROM employees WHERE department = 'Sales';","COUNT(*) with WHERE."),

Q("s27","Summaries","COUNT(column) and COUNT(DISTINCT)",`<p><code>COUNT(column)</code> skips NULLs. <code>COUNT(DISTINCT column)</code> counts unique values.</p>`,
"SELECT COUNT(DISTINCT col) FROM table;","SELECT COUNT(bonus) FROM employees;","7 people have a bonus.",
"How many different cities are there?","SELECT COUNT(DISTINCT city) FROM employees;","COUNT(DISTINCT city)."),

Q("s12","Summaries","SUM, AVG, MIN, MAX",`<p>Totals, averages and extremes of a column.</p>`,
"SELECT AVG(col) FROM table;","SELECT SUM(salary) FROM employees;","Total salary bill.",
"Average salary of all employees.","SELECT AVG(salary) FROM employees;","AVG(salary)."),

Q("s28","Summaries","ROUND",`<p><code>ROUND(value, 2)</code> keeps 2 decimals. Averages often need it.</p>`,
"ROUND(expression, decimals)","SELECT ROUND(AVG(bonus), 0) FROM employees;","Average bonus, no decimals.",
"Average salary of HR, rounded to 2 decimals.","SELECT ROUND(AVG(salary), 2) FROM employees WHERE department = 'HR';","ROUND(AVG(salary), 2) with WHERE."),

Q("s13","Summaries","Math and AS",`<p>Calculate inside SELECT, and name the result with <code>AS</code>.</p>`,
"SELECT col * 12 AS new_name","SELECT name, salary + 1000 AS new_salary FROM employees;","Everyone +1000.",
"Name and yearly salary (salary × 12) as yearly_salary.","SELECT name, salary * 12 AS yearly_salary FROM employees;","salary * 12 AS yearly_salary."),

Q("s14","Group & summarise","GROUP BY",`<p><code>GROUP BY</code> makes one row per group with a summary, like a Pivot Table. Every SELECT column must be grouped or inside COUNT/SUM/AVG…</p>`,
"SELECT col, COUNT(*) FROM t GROUP BY col;","SELECT city, SUM(salary) FROM employees GROUP BY city;","Salary per city.",
"Each department and its number of employees.","SELECT department, COUNT(*) FROM employees GROUP BY department;","GROUP BY department."),

Q("s29","Group & summarise","GROUP BY two columns",`<p>Group by more than one column to get every combination.</p>`,
"GROUP BY col1, col2","SELECT department, city, SUM(salary) FROM employees GROUP BY department, city;","Salary per dept and city.",
"Count employees for each department and city combination.","SELECT department, city, COUNT(*) FROM employees GROUP BY department, city;","GROUP BY department, city."),

Q("s15","Group & summarise","HAVING",`<p><code>WHERE</code> filters rows before grouping. <code>HAVING</code> filters groups after, so it can use COUNT, SUM, AVG.</p>`,
"GROUP BY col HAVING AVG(x) > 100","SELECT department FROM employees GROUP BY department HAVING COUNT(*) > 3;","Sales only.",
"Each city and its average salary, only where the average is above 45000.","SELECT city, AVG(salary) FROM employees GROUP BY city HAVING AVG(salary) > 45000;","HAVING AVG(salary) > 45000."),

Q("s30","Group & summarise","Sort groups",`<p>Sort grouped results by a summary: <code>ORDER BY SUM(salary) DESC</code>, or by its alias.</p>`,
"GROUP BY col ORDER BY total DESC","SELECT city, COUNT(*) AS people FROM employees GROUP BY city ORDER BY people DESC;","Biggest city first.",
"Each department with its total salary as total, biggest total first.","SELECT department, SUM(salary) AS total FROM employees GROUP BY department ORDER BY total DESC;","SUM(salary) AS total … ORDER BY total DESC.",{ordered:true}),

Q("s31","Functions","UPPER and LOWER",`<p>Change text case: <code>UPPER(name)</code>, <code>LOWER(city)</code>.</p>`,
"SELECT UPPER(col) FROM t;","SELECT LOWER(department) FROM employees;","sales, hr, it…",
"Show every name in capital letters.","SELECT UPPER(name) FROM employees;","UPPER(name)."),

Q("s32","Functions","LENGTH and SUBSTR",`<p><code>LENGTH(text)</code> counts characters. <code>SUBSTR(text, start, n)</code> cuts out n characters from start (1 = first).</p>`,
"SUBSTR(col, 1, 3)","SELECT name, LENGTH(name) FROM employees;","Name lengths.",
"Show each name and a 3-letter city code (first 3 letters of city) named code.","SELECT name, SUBSTR(city, 1, 3) AS code FROM employees;","SUBSTR(city, 1, 3) AS code."),

Q("s33","Functions","Joining text with ||",`<p>In SQLite and Oracle, <code>||</code> joins text: <code>name || ' - ' || city</code>. In SQL Server use <code>+</code> or <code>CONCAT()</code>.</p>`,
"col1 || ' ' || col2","SELECT name || ' (' || department || ')' FROM employees;","Aarav (Sales)…",
"Show “name - city” for everyone, named label.","SELECT name || ' - ' || city AS label FROM employees;","name || ' - ' || city AS label."),

Q("s34","Functions","REPLACE and TRIM",`<p><code>REPLACE(text, old, new)</code> swaps text. <code>TRIM</code> removes spaces at the ends.</p>`,
"REPLACE(col, 'old', 'new')","SELECT TRIM('  Pune  ');","'Pune'",
"Show each city with 'Mumbai' replaced by 'Bombay' (other cities unchanged).","SELECT REPLACE(city, 'Mumbai', 'Bombay') FROM employees;","REPLACE(city, 'Mumbai', 'Bombay')."),

Q("s35","Functions","COALESCE (fill NULLs)",`<p><code>COALESCE(a, b)</code> returns the first value that isn't NULL. Use it to show 0 instead of NULL. (<code>IFNULL</code> and SQL Server's <code>ISNULL</code> are similar.)</p>`,
"COALESCE(col, 0)","SELECT name, COALESCE(mentor_id, 0) FROM employees;","0 where no mentor.",
"Show name and bonus, with 0 instead of NULL.","SELECT name, COALESCE(bonus, 0) FROM employees;","COALESCE(bonus, 0)."),

Q("s36","Functions","Dates: strftime",`<p>Dates here are text like <code>'2026-01-05'</code>. <code>strftime('%Y', date)</code> gives the year, <code>'%m'</code> the month, <code>'%d'</code> the day. (SQL Server uses <code>YEAR()</code>, <code>MONTH()</code>.)</p>`,
"strftime('%m', date_col)","SELECT name, strftime('%Y', hire_date) FROM employees;","Hire years.",
"From the sales table, show the month ('%m') and total amount for each month.","SELECT strftime('%m', sale_date) AS month, SUM(amount) FROM sales GROUP BY month;","GROUP BY strftime('%m', sale_date)."),

Q("s37","Functions","Filtering dates",`<p>Dates in YYYY-MM-DD form sort correctly as text, so you can use <code>BETWEEN</code> or <code>&gt;=</code> directly.</p>`,
"WHERE date_col BETWEEN '2026-01-01' AND '2026-01-31'","SELECT * FROM sales WHERE sale_date >= '2026-03-01';","March sales.",
"Show all sales in February 2026.","SELECT * FROM sales WHERE sale_date BETWEEN '2026-02-01' AND '2026-02-28';","BETWEEN '2026-02-01' AND '2026-02-28'."),

Q("s17","Functions","CASE (if-then)",`<p><code>CASE</code> is SQL's IF. It makes a new column from conditions.</p>`,
"CASE WHEN cond THEN 'a' ELSE 'b' END AS col","SELECT name, CASE WHEN join_year < 2020 THEN 'Senior' ELSE 'New' END AS level FROM employees;","Labels.",
"Show name and band: 'High' if salary ≥ 50000, else 'Normal'.","SELECT name, CASE WHEN salary >= 50000 THEN 'High' ELSE 'Normal' END AS band FROM employees;","CASE WHEN salary >= 50000 THEN 'High' ELSE 'Normal' END AS band"),

Q("s38","Functions","CASE inside SUM",`<p>Count or add only some rows inside a group: <code>SUM(CASE WHEN … THEN 1 ELSE 0 END)</code>. This is a classic interview trick.</p>`,
"SUM(CASE WHEN cond THEN 1 ELSE 0 END)","SELECT city, SUM(CASE WHEN department = 'IT' THEN 1 ELSE 0 END) FROM employees GROUP BY city;","IT people per city.",
"For each department, count people with salary ≥ 50000, named high_earners.","SELECT department, SUM(CASE WHEN salary >= 50000 THEN 1 ELSE 0 END) AS high_earners FROM employees GROUP BY department;","SUM(CASE WHEN salary >= 50000 THEN 1 ELSE 0 END)."),

Q("s18","Joins","INNER JOIN",`<p>A <b>JOIN</b> puts two tables side by side by matching a shared column. <code>employees e</code> gives the table a short alias. An INNER JOIN keeps only rows that match in both tables, like VLOOKUP.</p>`,
"FROM a JOIN b ON a.col = b.col","SELECT e.name, d.floor FROM employees e JOIN departments d ON e.department = d.department;","Each person's floor.",
"Each employee's name and their manager.","SELECT e.name, d.manager FROM employees e JOIN departments d ON e.department = d.department;","JOIN departments d ON e.department = d.department."),

Q("s39","Joins","JOIN + GROUP BY",`<p>Join first, then group. Here: total sales amount per employee name.</p>`,
"FROM a JOIN b … GROUP BY …","SELECT e.name, COUNT(*) FROM sales s JOIN employees e ON s.employee_id = e.id GROUP BY e.name;","Number of sales per person.",
"Total sales amount per employee name (from the sales table).","SELECT e.name, SUM(s.amount) FROM sales s JOIN employees e ON s.employee_id = e.id GROUP BY e.name;","Join sales to employees ON s.employee_id = e.id."),

Q("s19","Joins","LEFT JOIN",`<p>A <b>LEFT JOIN</b> keeps every row from the first table, even without a match (NULL fills the gap).</p>`,
"FROM a LEFT JOIN b ON …","SELECT d.department, e.name FROM departments d LEFT JOIN employees e ON d.department = e.department;","Finance appears with NULL.",
"Every employee's name with their project_name (people without a project still appear).","SELECT e.name, p.project_name FROM employees e LEFT JOIN projects p ON e.id = p.employee_id;","employees e LEFT JOIN projects p ON e.id = p.employee_id."),

Q("s40","Joins","Find missing matches",`<p>LEFT JOIN + <code>WHERE right_table.column IS NULL</code> finds rows with <b>no</b> match. Very common in real work.</p>`,
"LEFT JOIN b ON … WHERE b.id IS NULL","SELECT d.department FROM departments d LEFT JOIN employees e ON d.department = e.department WHERE e.id IS NULL;","Finance.",
"Names of employees who have no project.","SELECT e.name FROM employees e LEFT JOIN projects p ON e.id = p.employee_id WHERE p.project_id IS NULL;","WHERE p.project_id IS NULL."),

Q("s41","Joins","RIGHT and FULL JOIN",`<p><code>RIGHT JOIN</code> keeps every row of the second table. <code>FULL OUTER JOIN</code> keeps all rows from both.</p>`,
"FROM a FULL OUTER JOIN b ON …","SELECT e.name, p.project_name FROM employees e RIGHT JOIN projects p ON e.id = p.employee_id;","Market Study has no owner.",
"Show employee name and project_name keeping ALL employees and ALL projects.","SELECT e.name, p.project_name FROM employees e FULL OUTER JOIN projects p ON e.id = p.employee_id;","FULL OUTER JOIN projects p ON e.id = p.employee_id."),

Q("s42","Joins","Self join",`<p>A table can be joined to <b>itself</b>. <code>mentor_id</code> points to another employee's <code>id</code>. Use two aliases.</p>`,
"FROM employees e JOIN employees m ON e.mentor_id = m.id","SELECT e.name, m.city FROM employees e JOIN employees m ON e.mentor_id = m.id;","Mentor's city.",
"Show each employee's name and their mentor's name, as mentor (only people with a mentor).","SELECT e.name, m.name AS mentor FROM employees e JOIN employees m ON e.mentor_id = m.id;","JOIN employees m ON e.mentor_id = m.id."),

Q("s43","Joins","Three-table join",`<p>Chain joins one after another.</p>`,
"FROM a JOIN b ON … JOIN c ON …","SELECT e.name, d.floor, p.project_name FROM employees e JOIN departments d ON e.department = d.department JOIN projects p ON p.employee_id = e.id;","",
"Show project_name, employee name and their manager for every project that has an employee.","SELECT p.project_name, e.name, d.manager FROM projects p JOIN employees e ON p.employee_id = e.id JOIN departments d ON e.department = d.department;","projects → employees → departments."),

Q("s20","Subqueries & sets","Subquery",`<p>A <b>subquery</b> is a query in brackets inside another. The inner one runs first.</p>`,
"WHERE col > (SELECT AVG(col) FROM t)","SELECT name FROM employees WHERE salary = (SELECT MAX(salary) FROM employees);","Top earner.",
"Name and salary of people earning more than the average.","SELECT name, salary FROM employees WHERE salary > (SELECT AVG(salary) FROM employees);","Inner: (SELECT AVG(salary) FROM employees)."),

Q("s44","Subqueries & sets","IN (subquery)",`<p>Use a subquery to build the list for IN.</p>`,
"WHERE id IN (SELECT … FROM …)","SELECT name FROM employees WHERE id NOT IN (SELECT employee_id FROM sales);","People with no sales.",
"Names of employees who have at least one project.","SELECT name FROM employees WHERE id IN (SELECT employee_id FROM projects);","id IN (SELECT employee_id FROM projects)."),

Q("s45","Subqueries & sets","EXISTS",`<p><code>EXISTS (subquery)</code> is true if the subquery returns any row. The subquery can refer to the outer row.</p>`,
"WHERE EXISTS (SELECT 1 FROM t WHERE t.x = outer.x)","SELECT d.department FROM departments d WHERE EXISTS (SELECT 1 FROM employees e WHERE e.department = d.department);","Departments with staff.",
"Names of employees who made a sale over 15000.","SELECT name FROM employees e WHERE EXISTS (SELECT 1 FROM sales s WHERE s.employee_id = e.id AND s.amount > 15000);","EXISTS (SELECT 1 FROM sales s WHERE s.employee_id = e.id AND s.amount > 15000)."),

Q("s46","Subqueries & sets","UNION",`<p><code>UNION</code> stacks results of two queries (same number of columns) and removes duplicates. <code>UNION ALL</code> keeps duplicates.</p>`,
"query1 UNION query2","SELECT city FROM employees UNION SELECT 'Chennai';","Adds a city.",
"Names of people in Pune, combined with names of people earning over 55000 (no duplicates).","SELECT name FROM employees WHERE city = 'Pune' UNION SELECT name FROM employees WHERE salary > 55000;","Two SELECTs joined with UNION."),

Q("s47","Subqueries & sets","INTERSECT and EXCEPT",`<p><code>INTERSECT</code>: rows in both results. <code>EXCEPT</code> (MINUS in Oracle): rows in the first but not the second.</p>`,
"query1 EXCEPT query2","SELECT name FROM employees WHERE city = 'Pune' INTERSECT SELECT name FROM employees WHERE department = 'IT';","Meera.",
"Names of Sales people EXCEPT those in Mumbai.","SELECT name FROM employees WHERE department = 'Sales' EXCEPT SELECT name FROM employees WHERE city = 'Mumbai';","Two SELECTs joined with EXCEPT."),

Q("s48","Subqueries & sets","WITH (CTE)",`<p>A <b>CTE</b> gives a subquery a name so the main query reads cleanly. Start with <code>WITH name AS (…)</code>.</p>`,
"WITH t AS (SELECT …) SELECT … FROM t;","WITH it AS (SELECT * FROM employees WHERE department = 'IT') SELECT name FROM it;","IT names.",
"Using a CTE named totals (employee_id, total amount from sales), show employee_id and total where total > 35000.","WITH totals AS (SELECT employee_id, SUM(amount) AS total FROM sales GROUP BY employee_id) SELECT employee_id, total FROM totals WHERE total > 35000;","WITH totals AS (SELECT employee_id, SUM(amount) AS total FROM sales GROUP BY employee_id)."),

Q("s49","Window functions","ROW_NUMBER",`<p><b>Window functions</b> calculate across rows without grouping them away. <code>ROW_NUMBER() OVER (ORDER BY salary DESC)</code> numbers rows 1, 2, 3…</p>`,
"ROW_NUMBER() OVER (ORDER BY col)","SELECT name, ROW_NUMBER() OVER (ORDER BY name) FROM employees;","Alphabetical numbering.",
"Show name, salary and a row number by salary (highest = 1), named rn.","SELECT name, salary, ROW_NUMBER() OVER (ORDER BY salary DESC) AS rn FROM employees;","ROW_NUMBER() OVER (ORDER BY salary DESC) AS rn."),

Q("s50","Window functions","RANK and DENSE_RANK",`<p><code>RANK</code> gives ties the same number and skips the next (1,2,2,4). <code>DENSE_RANK</code> doesn't skip (1,2,2,3).</p>`,
"RANK() OVER (ORDER BY col DESC)","SELECT name, bonus, DENSE_RANK() OVER (ORDER BY bonus DESC) FROM employees;","Bonus ranks.",
"Rank employees by join_year (earliest = 1) with RANK, named r. Show name, join_year, r.","SELECT name, join_year, RANK() OVER (ORDER BY join_year) AS r FROM employees;","RANK() OVER (ORDER BY join_year) AS r."),

Q("s51","Window functions","PARTITION BY",`<p><code>PARTITION BY</code> restarts the calculation for each group. Example: rank within each department.</p>`,
"RANK() OVER (PARTITION BY grp ORDER BY col DESC)","SELECT name, department, AVG(salary) OVER (PARTITION BY department) FROM employees;","Dept average on every row.",
"Show name, department, salary and rank within department by salary (highest = 1), named dept_rank.","SELECT name, department, salary, RANK() OVER (PARTITION BY department ORDER BY salary DESC) AS dept_rank FROM employees;","RANK() OVER (PARTITION BY department ORDER BY salary DESC)."),

Q("s52","Window functions","Running total",`<p><code>SUM(amount) OVER (ORDER BY sale_date)</code> adds up as it goes: a running total.</p>`,
"SUM(col) OVER (ORDER BY date)","SELECT sale_date, amount, AVG(amount) OVER (ORDER BY sale_date) FROM sales;","Running average.",
"Show sale_date, amount and a running total named running_total, ordered by sale_date.","SELECT sale_date, amount, SUM(amount) OVER (ORDER BY sale_date) AS running_total FROM sales ORDER BY sale_date;","SUM(amount) OVER (ORDER BY sale_date).",{ordered:true}),

Q("s53","Window functions","LAG and LEAD",`<p><code>LAG(col)</code> reads the value from the previous row; <code>LEAD(col)</code> from the next row. Great for “change vs last time”.</p>`,
"LAG(col) OVER (ORDER BY date)","SELECT sale_date, amount, LEAD(amount) OVER (ORDER BY sale_date) FROM sales;","Next sale's amount.",
"Show sale_date, amount and the previous sale's amount named prev_amount, ordered by sale_date.","SELECT sale_date, amount, LAG(amount) OVER (ORDER BY sale_date) AS prev_amount FROM sales ORDER BY sale_date;","LAG(amount) OVER (ORDER BY sale_date).",{ordered:true}),

Q("s54","Change data","INSERT",`<p><code>INSERT INTO</code> adds a new row. List the columns, then the values in the same order. (Don't worry, this practice database resets every time you run.)</p>`,
"INSERT INTO t (col1, col2) VALUES (v1, v2);","INSERT INTO departments (department, manager, floor) VALUES ('Legal', 'Isha', 5);","Adds a department.",
"Add a new department 'Admin' with manager 'Farhan' on floor 1.","INSERT INTO departments (department, manager, floor) VALUES ('Admin', 'Farhan', 1);","INSERT INTO departments (department, manager, floor) VALUES (…).",{check:"SELECT * FROM departments"}),

Q("s55","Change data","UPDATE",`<p><code>UPDATE … SET … WHERE …</code> changes existing rows. <b>Always add WHERE</b>, or every row changes.</p>`,
"UPDATE t SET col = value WHERE …;","UPDATE employees SET city = 'Pune' WHERE name = 'Karan';","Karan moves.",
"Give everyone in IT a 10% raise (salary × 1.1).","UPDATE employees SET salary = salary * 1.1 WHERE department = 'IT';","SET salary = salary * 1.1 WHERE department = 'IT'.",{check:"SELECT name, salary FROM employees"}),

Q("s56","Change data","DELETE",`<p><code>DELETE FROM … WHERE …</code> removes rows. Without WHERE it empties the table.</p>`,
"DELETE FROM t WHERE …;","DELETE FROM projects WHERE employee_id IS NULL;","Removes unassigned projects.",
"Delete all sales with amount below 10000.","DELETE FROM sales WHERE amount < 10000;","WHERE amount < 10000.",{check:"SELECT * FROM sales"}),

Q("s57","Change data","CREATE TABLE",`<p><code>CREATE TABLE</code> makes a new table with column names and types (INTEGER, TEXT, REAL, DATE…).</p>`,
"CREATE TABLE t (col1 TYPE, col2 TYPE);","CREATE TABLE cities (name TEXT, state TEXT);","",
"Create a table named training with columns course (TEXT) and hours (INTEGER), then insert one row: 'Excel', 6.","CREATE TABLE training (course TEXT, hours INTEGER); INSERT INTO training VALUES ('Excel', 6);","Two statements separated by ;",{check:"SELECT * FROM training"})
];
/* ---------- Power BI practice model ---------- */
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const PBI_COLS = ["OrderID","Date","Month","MonthNum","Region","Product","Category","Quantity","Price","Discount","Amount"];
const PBI_SALES = [
 [1001,"2026-01-05","North","Laptop","Electronics",1,45000,2000],[1002,"2026-01-07","South","Phone","Electronics",2,15000,0],
 [1003,"2026-01-09","East","Chair","Furniture",4,3000,500],[1004,"2026-01-12","North","Desk","Furniture",1,8000,0],
 [1005,"2026-01-15","West","Phone","Electronics",1,15000,1000],[1006,"2026-01-18","South","Laptop","Electronics",1,45000,3000],
 [1007,"2026-01-20","East","Pen","Stationery",20,50,0],[1008,"2026-01-22","West","Notebook","Stationery",15,80,100],
 [1009,"2026-02-02","North","Chair","Furniture",6,3000,0],[1010,"2026-02-05","South","Pen","Stationery",30,50,200],
 [1011,"2026-02-08","East","Laptop","Electronics",2,45000,5000],[1012,"2026-02-11","West","Desk","Furniture",2,8000,500],
 [1013,"2026-02-14","North","Notebook","Stationery",25,80,0],[1014,"2026-02-17","South","Chair","Furniture",3,3000,0],
 [1015,"2026-02-20","East","Phone","Electronics",3,15000,1500],[1016,"2026-02-24","North","Phone","Electronics",1,15000,0],
 [1017,"2026-03-02","West","Laptop","Electronics",1,45000,2500],[1018,"2026-03-05","North","Desk","Furniture",2,8000,1000],
 [1019,"2026-03-09","South","Notebook","Stationery",40,80,200],[1020,"2026-03-12","East","Chair","Furniture",5,3000,0],
 [1021,"2026-03-16","North","Laptop","Electronics",1,45000,0],[1022,"2026-03-20","South","Phone","Electronics",1,15000,500],
 [1023,"2026-03-24","East","Pen","Stationery",50,50,100],[1024,"2026-03-27","West","Chair","Furniture",2,3000,0]
].map(([OrderID,Date,Region,Product,Category,Quantity,Price,Discount]) => {
  const MonthNum = +Date.slice(5,7);
  return {OrderID,Date,Month:MONTHS[MonthNum-1],MonthNum,Region,Product,Category,Quantity,Price,Discount,Amount:Quantity*Price-Discount};
});
const REG_COLS = ["Region","Manager","Target"];
const PBI_REGIONS_T = [{Region:"North",Manager:"Amit",Target:120000},{Region:"South",Manager:"Kavitha",Target:110000},{Region:"East",Manager:"Rahul",Target:150000},{Region:"West",Manager:"Fatima",Target:60000}];
const PBI_TABLES = {sales:{name:"Sales", cols:PBI_COLS, rows:PBI_SALES}, regions:{name:"Regions", cols:REG_COLS, rows:PBI_REGIONS_T}};
const PBI_MEASURES = {"Total Sales":"SUM(Sales[Amount])", "Orders":"COUNTROWS(Sales)", "Total Qty":"SUM(Sales[Quantity])"};
const MONTH_COLS = ["Date","Month","MonthNum"];

class DaxError extends Error {}
const DAX_FN = {
 SUM:"column",AVERAGE:"column",MIN:"column",MAX:"column",COUNT:"column",COUNTA:"column",COUNTBLANK:"column",COUNTROWS:"table",DISTINCTCOUNT:"column",
 SUMX:"table, expression",AVERAGEX:"table, expression",MINX:"table, expression",MAXX:"table, expression",COUNTX:"table, expression",
 FILTER:"table, condition",ALL:"table or column",REMOVEFILTERS:"table or column",ALLEXCEPT:"table, column",VALUES:"column",DISTINCT:"column",
 CALCULATE:"expression, filter1, …",KEEPFILTERS:"condition",DIVIDE:"numerator, denominator, [alternate]",IF:"condition, if_true, [if_false]",SWITCH:"expression, value1, result1, …, [else]",
 AND:"cond1, cond2",OR:"cond1, cond2",NOT:"condition",ROUND:"number, decimals",ABS:"number",INT:"number",BLANK:"",TRUE:"",FALSE:"",ISBLANK:"value",COALESCE:"value1, value2, …",
 SELECTEDVALUE:"column, [alternate]",HASONEVALUE:"column",RANKX:"table, expression, [value], [ASC/DESC]",CONCATENATEX:"table, expression, [separator]",
 FORMAT:"value, format",RELATED:"column",YEAR:"date",MONTH:"date",DAY:"date",UPPER:"text",LOWER:"text",LEFT:"text, n",RIGHT:"text, n",LEN:"text",
 TOTALYTD:"expression, dates",DATESYTD:"dates",PREVIOUSMONTH:"dates",DATEADD:"dates, number, MONTH"};
const DAX_KW = ["VAR","RETURN","ASC","DESC","MONTH","YEAR","DAY","QUARTER","IN"];

function findTable(name) { const k = String(name).toLowerCase(); const t = PBI_TABLES[k]; if (!t) throw new DaxError(`There is no table called ${name}. Tables here: Sales and Regions.`); return t; }
function findCol(t, c) { const hit = t.cols.find(x => x.toLowerCase() === c.trim().toLowerCase()); if (!hit) throw new DaxError(`${t.name} has no column [${c}]. Its columns are: ${t.cols.join(", ")}.`); return hit; }

function daxTokenize(src) {
  const t = []; let i = 0;
  while (i < src.length) {
    const c = src[i];
    if (/\s/.test(c)) { i++; continue; }
    if (c === "/" && src[i+1] === "/") { while (i < src.length && src[i] !== "\n") i++; continue; }
    if (/[0-9]/.test(c) || (c === "." && /[0-9]/.test(src[i+1]||""))) { const m = src.slice(i).match(/^\d*\.?\d+/); t.push({k:"num", v:parseFloat(m[0])}); i += m[0].length; continue; }
    if (c === '"') { let j = i+1, s = ""; while (j < src.length) { if (src[j] === '"') { if (src[j+1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; } if (j >= src.length) throw new DaxError('A text value is missing its closing " quote.'); t.push({k:"str", v:s}); i = j+1; continue; }
    if (c === "'" || /[A-Za-z_]/.test(c)) {
      let name;
      if (c === "'") { const j = src.indexOf("'", i+1); if (j < 0) throw new DaxError("A table name is missing its closing ' quote."); name = src.slice(i+1, j); i = j+1; }
      else { const m = src.slice(i).match(/^[A-Za-z_][A-Za-z0-9_.]*/); name = m[0]; i += name.length; }
      let k = i; while (/\s/.test(src[k]||"")) k++;
      if (src[k] === "[") { const j = src.indexOf("]", k); if (j < 0) throw new DaxError("A column name is missing its closing ] bracket."); t.push({k:"col", table:name, col:src.slice(k+1, j)}); i = j+1; continue; }
      if (src[k] === "(") { t.push({k:"fn", v:name.toUpperCase()}); i = k; continue; }
      t.push({k:"id", v:name}); continue;
    }
    if (c === "[") { const j = src.indexOf("]", i); if (j < 0) throw new DaxError("A name is missing its closing ] bracket."); t.push({k:"col", table:null, col:src.slice(i+1, j)}); i = j+1; continue; }
    const two = src.slice(i, i+2);
    if (["&&","||","<=",">=","<>","=="].includes(two)) { t.push({k:"op", v: two === "==" ? "=" : two}); i += 2; continue; }
    if ("+-*/&=<>^(),{}".includes(c)) { t.push({k:"op", v:c}); i++; continue; }
    throw new DaxError(`I don't understand the character "${c}".`);
  }
  return t;
}

function daxParse(src) {
  const toks = daxTokenize(src); let p = 0;
  const peek = () => toks[p], isOp = v => toks[p] && toks[p].k === "op" && toks[p].v === v, isId = v => toks[p] && toks[p].k === "id" && toks[p].v.toUpperCase() === v;
  const expect = v => { if (!isOp(v)) throw new DaxError(v === ")" ? "A closing bracket ) seems to be missing." : `Expected "${v}" here.`); p++; };
  const bin = (next, ops) => () => { let a = next(); while (toks[p] && toks[p].k === "op" && ops.includes(toks[p].v)) { const op = toks[p++].v; a = {t:"bin", op, a, b:next()}; } return a; };
  let expr;
  const primary = () => {
    const tk = peek(); if (!tk) throw new DaxError("The formula ends too early. Is something missing at the end?");
    if (tk.k === "num") { p++; return {t:"num", v:tk.v}; }
    if (tk.k === "str") { p++; return {t:"str", v:tk.v}; }
    if (tk.k === "col") { p++;
      if (tk.table) { const T = findTable(tk.table); return {t:"col", tbl:T.name, col:findCol(T, tk.col)}; }
      const mk = Object.keys(PBI_MEASURES).find(m => m.toLowerCase() === tk.col.trim().toLowerCase());
      if (mk) return {t:"measure", name:mk};
      const sc = PBI_COLS.find(x => x.toLowerCase() === tk.col.trim().toLowerCase());
      if (sc) return {t:"col", tbl:"Sales", col:sc};
      throw new DaxError(`[${tk.col}] isn't a measure or column here. Ready-made measures: ${Object.keys(PBI_MEASURES).map(m => "[" + m + "]").join(", ")}.`); }
    if (tk.k === "fn") { p++; expect("("); const args = []; const arg = () => (isOp(",") || isOp(")")) ? {t:"empty"} : expr();
      if (!isOp(")")) { args.push(arg()); while (isOp(",")) { p++; args.push(arg()); } } expect(")");
      if (!(tk.v in DAX_FN)) throw new DaxError(`${tk.v} isn't available in this practice lab. You can use: ${Object.keys(DAX_FN).join(", ")}.`);
      return {t:"fn", name:tk.v, args}; }
    if (tk.k === "id") { p++; const u = tk.v.toUpperCase();
      if (u === "VAR") { const vars = []; let name = tk;
        for (;;) { const nm = peek(); if (!nm || nm.k !== "id") throw new DaxError("After VAR write a name, like VAR total = ..."); p++; expect("="); vars.push([nm.v.toLowerCase(), expr()]);
          if (isId("VAR")) { p++; continue; } if (isId("RETURN")) { p++; break; } throw new DaxError("Every VAR block must end with RETURN and the final expression."); }
        return {t:"let", vars, body:expr()}; }
      if (u === "TRUE" || u === "FALSE") return {t:"bool", v:u === "TRUE"};
      if (["ASC","DESC","MONTH","YEAR","DAY","QUARTER"].includes(u)) return {t:"kw", v:u};
      if (PBI_TABLES[tk.v.toLowerCase()]) return {t:"tbl", name:PBI_TABLES[tk.v.toLowerCase()].name};
      return {t:"var", name:tk.v.toLowerCase(), raw:tk.v}; }
    if (isOp("(")) { p++; const e = expr(); expect(")"); return e; }
    if (isOp("-")) { p++; return {t:"neg", a:primary()}; }
    if (isOp("{")) { p++; const items = []; if (!isOp("}")) { items.push(expr()); while (isOp(",")) { p++; items.push(expr()); } } expect("}"); return {t:"list", items}; }
    throw new DaxError(`Unexpected "${tk.v ?? tk.col}" in the formula.`);
  };
  const pow = bin(primary, ["^"]), mul = bin(pow, ["*","/"]), add = bin(mul, ["+","-"]), cat = bin(add, ["&"]);
  const cmp = () => { let a = bin(cat, ["=","<>","<",">","<=",">="])(); while (isId("IN") || (isId("NOT") && toks[p+1] && toks[p+1].k === "id" && toks[p+1].v.toUpperCase() === "IN")) { let neg = false; if (isId("NOT")) { neg = true; p++; } p++; const list = primary(); a = {t:"in", a, list, neg}; } return a; };
  const and = bin(cmp, ["&&"]); expr = bin(and, ["||"]);
  const tree = expr();
  if (p < toks.length) throw new DaxError(`Something extra is left at the end: "${toks[p].v ?? toks[p].col}". Check your brackets and commas.`);
  return tree;
}

function stripMeasureName(src) {
  src = src.replace(/^\s*=(?!=)/, "");
  const m = src.match(/^\s*([A-Za-z_][\w %]*?)\s*:?=(?!=)/);
  if (m && !/[\[\("]/.test(m[1]) && !/^VAR\b/i.test(m[1].trim())) return {name:m[1].trim(), body:src.slice(m[0].length)};
  return {name:null, body:src};
}

const blank = v => v === null || v === undefined;
const truthy = v => !blank(v) && v !== false && v !== 0 && v !== "";
const sub = (ctx, o) => Object.assign({}, ctx, o);
function daxRows(ctx) { return PBI_SALES.filter(r => ctx.filters.every(f => f.test(r))); }
function colsIn(n, set = new Set()) { if (!n || typeof n !== "object") return set; if (n.t === "col" && n.tbl === "Sales") set.add(n.col); for (const k of ["a","b","list","body"]) colsIn(n[k], set); (n.args||[]).forEach(a => colsIn(a, set)); (n.items||[]).forEach(a => colsIn(a, set)); return set; }
const TABLE_FNS = ["FILTER","ALL","REMOVEFILTERS","VALUES","DISTINCT","ALLEXCEPT"];
const isTableNode = (n, ctx) => n.t === "tbl" || (n.t === "fn" && TABLE_FNS.includes(n.name)) || (n.t === "var" && ctx.env && ctx.env[n.name] && ctx.env[n.name].__table);
function needCol(n, fname) { if (!n || n.t !== "col") throw new DaxError(`${fname} needs a column inside, like ${fname}(Sales[Amount]).`); return n.col; }
function pseudo(col, vals) { return [...new Set(vals)].sort((a,b) => a < b ? -1 : a > b ? 1 : 0).map(v => ({__pseudo:true, [col]:v})); }

function daxTable(n, ctx) {
  if (n.t === "tbl") { if (n.name !== "Sales") return PBI_TABLES.regions.rows.slice(); return daxRows(ctx); }
  if (n.t === "var") { const v = ctx.env && ctx.env[n.name]; if (v && v.__table) return v.__table; }
  if (n.t === "fn") {
    const A = n.args;
    if (n.name === "FILTER") { if (A.length !== 2) throw new DaxError("FILTER needs 2 parts: FILTER(table, condition)."); return daxTable(A[0], ctx).filter(r => truthy(daxEval(A[1], sub(ctx, {row:r})))); }
    if (n.name === "ALL" || n.name === "REMOVEFILTERS") {
      if (!A.length || A[0].t === "tbl") return PBI_SALES.slice();
      const col = needCol(A[0], n.name); return pseudo(col, PBI_SALES.map(r => r[col]));
    }
    if (n.name === "VALUES" || n.name === "DISTINCT") { const col = needCol(A[0], n.name); return pseudo(col, daxRows(ctx).map(r => r[col])); }
  }
  throw new DaxError("This part needs a table, like Sales, FILTER(Sales, …), VALUES(Sales[Column]) or ALL(Sales).");
}

function visibleMonths(ctx) { return [...new Set(daxRows(ctx).map(r => r.MonthNum))]; }
function applyModifier(m, filters, ctx) {
  const dropMonth = fs => fs.filter(f => !MONTH_COLS.includes(f.col));
  if (m.t === "fn" && (m.name === "ALL" || m.name === "REMOVEFILTERS")) {
    if (!m.args.length || m.args[0].t === "tbl") return [];
    m.args.forEach(c => { const col = needCol(c, m.name); filters = filters.filter(f => f.col !== col); }); return filters;
  }
  if (m.t === "fn" && m.name === "ALLEXCEPT") { const keep = m.args.slice(1).map(c => needCol(c, "ALLEXCEPT")); return filters.filter(f => f.col && keep.includes(f.col)); }
  if (m.t === "fn" && m.name === "KEEPFILTERS") { const inner = m.args[0]; return filters.concat({col:"__keep", test:r => truthy(daxEval(inner, sub(ctx, {filters:[], row:r})))}); }
  if (m.t === "fn" && m.name === "DATESYTD") { const ms = visibleMonths(ctx); if (!ms.length) return filters; const mx = Math.max(...ms); return dropMonth(filters).concat({col:"MonthNum", test:r => r.MonthNum <= mx}); }
  if (m.t === "fn" && m.name === "PREVIOUSMONTH") { const ms = visibleMonths(ctx); const mn = ms.length ? Math.min(...ms) - 1 : -99; return dropMonth(filters).concat({col:"MonthNum", test:r => r.MonthNum === mn}); }
  if (m.t === "fn" && m.name === "DATEADD") {
    const n = daxEval(m.args[1], ctx), unit = m.args[2] && m.args[2].t === "kw" ? m.args[2].v : null;
    if (unit !== "MONTH") throw new DaxError("In this lab, DATEADD works with MONTH, like DATEADD(Sales[Date], -1, MONTH).");
    const set = new Set(visibleMonths(ctx).map(x => x + n)); return dropMonth(filters).concat({col:"MonthNum", test:r => set.has(r.MonthNum)});
  }
  if (isTableNode(m, ctx)) {
    const rows = daxTable(m, ctx);
    if (rows.length && rows[0].__pseudo) {
      const cs = Object.keys(rows[0]).filter(k => k !== "__pseudo"), key = r => JSON.stringify(cs.map(c => r[c])), vals = new Set(rows.map(key));
      const monthy = cs.some(c => MONTH_COLS.includes(c));
      filters = filters.filter(f => !cs.includes(f.col) && !(monthy && MONTH_COLS.includes(f.col)));
      return filters.concat({col:cs[0], test:r => vals.has(key(r))});
    }
    if (!rows.length && m.t === "fn" && m.name === "FILTER" && m.args[0].t === "fn" && m.args[0].args[0] && m.args[0].args[0].t === "col") { const c = m.args[0].args[0].col; return filters.filter(f => f.col !== c).concat({col:c, test:() => false}); }
    const keep = new Set(rows); return filters.concat({col:null, test:r => keep.has(r)});
  }
  const cols = [...colsIn(m)];
  if (!cols.length) throw new DaxError('A CALCULATE filter must test a column, like Sales[Region] = "North".');
  filters = filters.filter(f => !cols.includes(f.col));
  return filters.concat({col:cols.length === 1 ? cols[0] : null, test:r => truthy(daxEval(m, sub(ctx, {filters:[], row:r})))});
}

function rowFilters(ctx) {
  const r = ctx.row; if (!r) return ctx.filters;
  if (r.__pseudo) { const cs = Object.keys(r).filter(k => k !== "__pseudo"); return ctx.filters.filter(f => !cs.includes(f.col)).concat(cs.map(c => ({col:c, test:x => x[c] === r[c]}))); }
  if (PBI_SALES.includes(r)) return ctx.filters.concat({col:null, test:x => x === r});
  return ctx.filters;
}

function cmpVals(op, a, b) {
  if (typeof a === "string" || typeof b === "string") { a = String(blank(a) ? "" : a).toLowerCase(); b = String(blank(b) ? "" : b).toLowerCase(); }
  else { a = blank(a) ? 0 : a; b = blank(b) ? 0 : b; }
  return {"=":a===b,"<>":a!==b,"<":a<b,">":a>b,"<=":a<=b,">=":a>=b}[op];
}

function daxEval(n, ctx) {
  switch (n.t) {
    case "num": case "str": case "bool": return n.v;
    case "empty": return null;
    case "kw": throw new DaxError(`${n.v} can't be used on its own here.`);
    case "list": throw new DaxError("A list in { } is used with IN, like Sales[Region] IN {\"North\", \"South\"}.");
    case "tbl": throw new DaxError("A whole table can't be one value. Wrap it, for example COUNTROWS(Sales).");
    case "neg": { const v = daxEval(n.a, ctx); return blank(v) ? null : -v; }
    case "var": { const env = ctx.env || {}; if (!(n.name in env)) throw new DaxError(`I don't know "${n.raw}". Columns need Table[Column], measures need [Name], and text needs "double quotes".`);
      if (env[n.name] && env[n.name].__table) throw new DaxError(`${n.raw} holds a table. Use it inside COUNTROWS, SUMX or similar.`); return env[n.name]; }
    case "let": { const env = Object.assign({}, ctx.env || {}); let c = sub(ctx, {env});
      for (const [name, e] of n.vars) { env[name] = isTableNode(e, c) ? {__table:daxTable(e, c)} : daxEval(e, c); c = sub(ctx, {env}); }
      return daxEval(n.body, c); }
    case "measure": { const tree = daxParse(PBI_MEASURES[n.name]); return daxEval(tree, {filters:rowFilters(ctx), row:null, env:ctx.env}); }
    case "col":
      if (!ctx.row) throw new DaxError(`${n.tbl}[${n.col}] is a whole column, and a measure needs one value. Wrap it in a function like SUM(${n.tbl}[${n.col}]), or use it inside SUMX or FILTER.`);
      if (n.tbl === "Regions" && !("Manager" in ctx.row)) throw new DaxError(`To read Regions[${n.col}] from a Sales row, use RELATED(Regions[${n.col}]).`);
      if (!(n.col in ctx.row)) throw new DaxError(`${n.tbl}[${n.col}] isn't available in this row.`);
      return ctx.row[n.col];
    case "in": { const v = daxEval(n.a, ctx); if (n.list.t !== "list") throw new DaxError('After IN write a list in curly brackets: {"North", "South"}.');
      const hit = n.list.items.some(it => cmpVals("=", v, daxEval(it, ctx))); return n.neg ? !hit : hit; }
    case "bin": {
      if (n.op === "&&") return truthy(daxEval(n.a, ctx)) && truthy(daxEval(n.b, ctx));
      if (n.op === "||") return truthy(daxEval(n.a, ctx)) || truthy(daxEval(n.b, ctx));
      let a = daxEval(n.a, ctx), b = daxEval(n.b, ctx);
      if (n.op === "&") return String(blank(a) ? "" : a) + String(blank(b) ? "" : b);
      if (["=","<>","<",">","<=",">="].includes(n.op)) return cmpVals(n.op, a, b);
      if (blank(a) && blank(b)) return null;
      a = blank(a) ? 0 : a; b = blank(b) ? 0 : b;
      if (typeof a !== "number" || typeof b !== "number") throw new DaxError(`Math (${n.op}) only works on numbers. To join text, use &.`);
      if (n.op === "/" && b === 0) throw new DaxError("Division by zero. Use DIVIDE(a, b) instead of a / b to avoid this.");
      return {"+":a+b,"-":a-b,"*":a*b,"/":a/b,"^":Math.pow(a,b)}[n.op];
    }
    case "fn": return daxFn(n, ctx);
  }
}
const nums = a => a.filter(v => typeof v === "number");
const sumArr = a => a.reduce((s,v) => s+v, 0);
function fmtDax(v, f) {
  if (blank(v)) return "";
  const F = String(f);
  if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}$/.test(v)) { const [y,m,d] = v.split("-");
    return F.replace(/MMMM/i, ["January","February","March","April","May","June","July","August","September","October","November","December"][+m-1]).replace(/MMM/i, MONTHS[+m-1]).replace(/MM/, m).replace(/YYYY/i, y).replace(/DD/i, d); }
  if (typeof v !== "number") return String(v);
  const pm = F.match(/^0(?:\.(0+))?%$/); if (pm) return (v*100).toFixed(pm[1] ? pm[1].length : 0) + "%";
  const dm = F.match(/^#,(?:##)?0(?:\.(0+))?$/); if (dm) return v.toLocaleString("en-US", {minimumFractionDigits:dm[1]?dm[1].length:0, maximumFractionDigits:dm[1]?dm[1].length:0});
  const zm = F.match(/^0(?:\.(0+))?$/); if (zm) return v.toFixed(zm[1] ? zm[1].length : 0);
  throw new DaxError('FORMAT here supports "0", "0.00", "#,0", "#,0.00", "0%", "0.0%" and dates like "MMM YYYY".');
}

function daxFn(n, ctx) {
  const A = n.args, need = k => { if (A.length < k) throw new DaxError(`${n.name} needs more inputs: ${n.name}(${DAX_FN[n.name]})`); };
  const ev = (i, c = ctx) => daxEval(A[i], c);
  switch (n.name) {
    case "SUM": case "AVERAGE": case "MIN": case "MAX": case "COUNT": case "COUNTA": case "COUNTBLANK": case "DISTINCTCOUNT": {
      need(1);
      if ((n.name === "MIN" || n.name === "MAX") && A.length === 2) { const a = ev(0), b = ev(1); return n.name === "MIN" ? Math.min(a, b) : Math.max(a, b); }
      const col = needCol(A[0], n.name); if (A[0].tbl !== "Sales") throw new DaxError("In this lab, aggregate Sales columns. For Regions use RELATED in a column.");
      const vals = daxRows(ctx).map(r => r[col]), nv = nums(vals);
      if (n.name === "COUNTA") return vals.length || null;
      if (n.name === "COUNTBLANK") return vals.filter(blank).length;
      if (n.name === "DISTINCTCOUNT") return new Set(vals).size || null;
      if (n.name === "COUNT") return nv.length || null;
      if (!nv.length) { if (vals.length) throw new DaxError(`${n.name} only works on number columns. Sales[${col}] contains text.`); return null; }
      if (n.name === "SUM") return sumArr(nv);
      if (n.name === "AVERAGE") return sumArr(nv) / nv.length;
      return n.name === "MIN" ? Math.min(...nv) : Math.max(...nv);
    }
    case "COUNTROWS": { need(1); return daxTable(A[0], ctx).length || null; }
    case "SUMX": case "AVERAGEX": case "MINX": case "MAXX": case "COUNTX": {
      need(2); const vals = nums(daxTable(A[0], ctx).map(r => daxEval(A[1], sub(ctx, {row:r}))));
      if (!vals.length) return null;
      if (n.name === "SUMX") return sumArr(vals); if (n.name === "AVERAGEX") return sumArr(vals) / vals.length; if (n.name === "COUNTX") return vals.length;
      return n.name === "MINX" ? Math.min(...vals) : Math.max(...vals);
    }
    case "FILTER": case "VALUES": case "DISTINCT": case "ALLEXCEPT": throw new DaxError(`${n.name} gives back a table. Wrap it in something like COUNTROWS(…), SUMX(…) or use it inside CALCULATE.`);
    case "ALL": case "REMOVEFILTERS": throw new DaxError(`${n.name} is used inside CALCULATE, like CALCULATE([Total Sales], ALL(Sales)).`);
    case "KEEPFILTERS": case "DATESYTD": case "PREVIOUSMONTH": case "DATEADD": throw new DaxError(`${n.name} is used inside CALCULATE, like CALCULATE([Total Sales], ${n.name}(…)).`);
    case "CALCULATE": { need(1); let filters = rowFilters(ctx); const c0 = sub(ctx, {row:null, filters});
      for (const m of A.slice(1)) filters = applyModifier(m, filters, sub(c0, {filters}));
      return daxEval(A[0], sub(ctx, {filters, row:null})); }
    case "TOTALYTD": { need(2); const filters = applyModifier({t:"fn", name:"DATESYTD", args:[A[1]]}, rowFilters(ctx), sub(ctx, {row:null})); return daxEval(A[0], sub(ctx, {filters, row:null})); }
    case "DIVIDE": { need(2); const a = ev(0), b = ev(1); if (blank(b) || b === 0) return A[2] ? ev(2) : null; if (blank(a)) return null; return a / b; }
    case "IF": need(2); return truthy(ev(0)) ? ev(1) : (A[2] ? ev(2) : null);
    case "SWITCH": { need(3); const v = ev(0); let i = 1; for (; i + 1 < A.length; i += 2) if (cmpVals("=", v, ev(i))) return ev(i + 1); return i < A.length ? ev(i) : null; }
    case "AND": need(2); return truthy(ev(0)) && truthy(ev(1));
    case "OR": need(2); return truthy(ev(0)) || truthy(ev(1));
    case "NOT": need(1); return !truthy(ev(0));
    case "ROUND": { need(2); const v = ev(0), d = ev(1); if (blank(v)) return null; const f = Math.pow(10, d); return Math.round(v * f) / f; }
    case "ABS": { need(1); const v = ev(0); return blank(v) ? null : Math.abs(v); }
    case "INT": { need(1); const v = ev(0); return blank(v) ? null : Math.floor(v); }
    case "BLANK": return null;
    case "TRUE": return true;
    case "FALSE": return false;
    case "ISBLANK": need(1); return blank(ev(0));
    case "COALESCE": { need(2); for (let i = 0; i < A.length; i++) { const v = ev(i); if (!blank(v)) return v; } return null; }
    case "SELECTEDVALUE": case "HASONEVALUE": { need(1); const col = needCol(A[0], n.name); const vs = [...new Set(daxRows(ctx).map(r => r[col]))];
      if (n.name === "HASONEVALUE") return vs.length === 1; return vs.length === 1 ? vs[0] : (A[1] ? ev(1) : null); }
    case "RANKX": { need(2); const tbl = daxTable(A[0], sub(ctx, {row:null}));
      const vals = tbl.map(r => daxEval(A[1], sub(ctx, {row:r, filters:rowFilters(sub(ctx, {row:r}))}))).filter(v => !blank(v));
      const cur = A[2] && A[2].t !== "kw" && A[2].t !== "empty" ? ev(2) : daxEval(A[1], sub(ctx, {row:null}));
      const order = A.find(a => a.t === "kw" && (a.v === "ASC" || a.v === "DESC")); const asc = order && order.v === "ASC";
      if (blank(cur)) return null; return 1 + vals.filter(v => asc ? v < cur : v > cur).length; }
    case "CONCATENATEX": { need(2); const sep = A[2] ? ev(2) : ""; return daxTable(A[0], ctx).map(r => daxEval(A[1], sub(ctx, {row:r}))).filter(v => !blank(v)).join(sep); }
    case "FORMAT": need(2); return fmtDax(ev(0), ev(1));
    case "RELATED": { need(1); if (A[0].t !== "col" || A[0].tbl !== "Regions") throw new DaxError("RELATED reads from the Regions table, like RELATED(Regions[Manager]).");
      if (!ctx.row || !("Region" in ctx.row)) throw new DaxError("RELATED needs a Sales row. Use it in a calculated column or inside SUMX/FILTER.");
      const hit = PBI_REGIONS_T.find(x => x.Region === ctx.row.Region); return hit ? hit[A[0].col] : null; }
    case "YEAR": case "MONTH": case "DAY": { need(1); const v = ev(0); if (blank(v)) return null; const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/); if (!m) throw new DaxError(`${n.name} needs a date, like ${n.name}(Sales[Date]).`); return +m[{YEAR:1,MONTH:2,DAY:3}[n.name]]; }
    case "UPPER": need(1); return String(ev(0) ?? "").toUpperCase();
    case "LOWER": need(1); return String(ev(0) ?? "").toLowerCase();
    case "LEFT": need(2); return String(ev(0) ?? "").slice(0, ev(1));
    case "RIGHT": { need(2); const s = String(ev(0) ?? ""); return s.slice(Math.max(0, s.length - ev(1))); }
    case "LEN": need(1); return String(ev(0) ?? "").length;
  }
}

/** Evaluate a measure: grand total + one value per item of `by`. */
function daxGroups(by) {
  const vals = [...new Set(PBI_SALES.map(r => r[by]))];
  if (by === "Month") vals.sort((a,b) => MONTHS.indexOf(a) - MONTHS.indexOf(b)); else if (by !== "Region") vals.sort();
  return vals;
}
function daxMeasure(src, by = "Region") {
  const {name, body} = stripMeasureName(src);
  if (!body.trim()) throw new DaxError("Write a formula after the = sign.");
  const tree = daxParse(body);
  const total = daxEval(tree, {filters:[], row:null, env:{}});
  const groups = daxGroups(by).map(g => ({key:g, value:daxEval(tree, {filters:[{col:by, test:r => r[by] === g}], row:null, env:{}})}));
  return {name, total, groups};
}
/** Evaluate a calculated column on every Sales row. */
function daxColumn(src) {
  const {name, body} = stripMeasureName(src);
  if (!body.trim()) throw new DaxError("Write a formula after the = sign.");
  const tree = daxParse(body);
  return {name, values:PBI_SALES.map(r => daxEval(tree, {filters:[], row:r, env:{}}))};
}
/* ---------- Power BI lessons ---------- */
const PQ = (id, group, title, learn, questions) => ({id, track:"pbi", group, title, kind:"quiz", learn, questions:questions.map(([q, opts, a, why]) => ({q, opts, a, why}))});
const PM = (id, group, title, learn, syntax, ex, exNote, task, answer, must, hint, extra = {}) =>
  ({id, track:"pbi", group, title, learn, syntax, example:{code:ex, note:exNote}, task, answer, must, hints:[hint, "Type " + answer], ...extra});

const PBI_LESSONS = [
PQ("p1","Power BI basics","Power BI in 2 minutes",
`<p><b>Power BI</b> turns data into interactive reports and dashboards.</p>
<ul><li><b>Power BI Desktop</b>: free Windows app where you build reports</li><li><b>Power BI Service</b> (app.powerbi.com): where you publish and share</li><li><b>Power BI Mobile</b>: view reports on a phone</li></ul>
<p>The workflow: <b>Get data</b> → <b>Clean</b> it in Power Query → <b>Model</b> it (relationships) → <b>Calculate</b> with DAX → <b>Visualize</b> → <b>Publish</b>.</p>
<p>Desktop has three views on the left: <b>Report</b> (build visuals), <b>Table</b> (see rows), <b>Model</b> (see relationships).</p>`,
[["Where do you build a report from scratch?",["Power BI Desktop","Power BI Mobile","Microsoft Paint"],0,"Reports are built in Desktop, then published to the Service."],
 ["Where do you share a report with your team online?",["Power Query","Power BI Service","Model view"],1,"The Service (app.powerbi.com) is the online home for sharing."],
 ["Which view shows how tables are connected?",["Report view","Model view","Table view"],1,"Model view shows tables as boxes with relationship lines."]]),

PQ("p2","Power BI basics","Power Query: cleaning data",
`<p><b>Power Query</b> is the cleaning tool. Open it with <b>Home → Transform data</b>.</p>
<ul><li><b>Remove columns</b>, <b>Change type</b> (numbers as numbers, dates as dates)</li><li><b>Remove duplicates</b>, <b>Remove blank rows</b>, <b>Trim</b>, <b>Replace values</b>, <b>Fill down</b></li><li><b>Split column</b> by a delimiter, e.g. "Mumbai-West" by "-"</li></ul>
<p>Every step is saved in <b>Applied Steps</b>. Click <b>Refresh</b> and all steps run again on the new data.</p>`,
[["Sales was imported as text. What do you do?",["Delete the column","Change type to a number","Rename it"],1,"Numbers stored as text can't be summed."],
 ["A Region column is filled only on the first row of each group. Which step fills the gaps?",["Fill down","Split column","Pivot"],0,"Fill down copies the value above into empty cells."],
 ["What happens to your steps when you Refresh?",["They are deleted","They run again on the new data","You redo them by hand"],1,"Applied Steps replay on every refresh."]]),

PQ("p16","Power BI basics","Power Query: combine & reshape",
`<ul><li><b>Append queries</b>: stack tables with the same columns (January + February)</li><li><b>Merge queries</b>: add columns from another table by matching a key, like VLOOKUP / SQL JOIN. Join kinds: Left outer, Inner, Full outer, Left anti (rows with no match)</li>
<li><b>Unpivot columns</b>: turn month columns (Jan, Feb, Mar) into two columns: Month and Value. Most visuals need this “long” shape</li><li><b>Group By</b>: summarise rows like a pivot</li><li>Behind every step is the <b>M language</b>. You rarely need to write it yourself</li></ul>`,
[["You have separate files for each month with the same columns. How do you combine them?",["Merge","Append","Unpivot"],1,"Append stacks rows."],
 ["You need to bring the Manager name from a Regions table into Sales.",["Merge queries","Append queries","Fill down"],0,"Merge matches a key column and adds columns."],
 ["A sheet has columns Jan, Feb, Mar with sales under each. What shape step helps charting?",["Unpivot columns","Remove columns","Split column"],0,"Unpivot turns the month headers into a Month column."],
 ["Which merge kind shows only rows with NO match in the other table?",["Inner","Left anti","Full outer"],1,"Left anti keeps rows from the first table that didn't match."]]),

PQ("p3","Visuals","Choosing the right visual",
`<ul><li><b>Card</b>: one big number</li><li><b>Bar / Column</b>: compare categories</li><li><b>Line</b>: change over time</li><li><b>Pie / Donut</b>: share of a whole (only a few slices)</li><li><b>Table / Matrix</b>: detailed numbers; a matrix is a pivot table</li><li><b>Slicer</b>: a clickable filter</li><li><b>Map</b>: data by location</li><li><b>KPI / Gauge</b>: progress toward a target</li></ul>
<p>Click a visual in the <b>Visualizations</b> pane, then drag fields into its wells (X-axis, Y-axis, Legend, Values).</p>`,
[["Show sales month by month.",["Pie chart","Line chart","Card"],1,"Line charts show trends over time."],
 ["One big Total Revenue number.",["Card","Matrix","Map"],0,"A Card shows a single value."],
 ["Users should pick a Region to filter the page.",["Slicer","Line chart","Tooltip"],0,"A Slicer is an on-page filter."],
 ["Sales by Region and Category in a grid with subtotals.",["Matrix","Card","Gauge"],0,"A Matrix works like a pivot table."]]),

PQ("p17","Visuals","Formatting and interactions",
`<ul><li><b>Format pane</b> (paint brush): titles, colors, data labels, fonts</li><li><b>Conditional formatting</b>: color cells or bars by value (e.g. red below target)</li><li><b>Cross-filtering</b>: clicking a bar filters the other visuals. Change it with <b>Format → Edit interactions</b></li><li><b>Drill down</b>: put Year → Month in the axis and use the arrows to go deeper</li><li><b>Tooltips</b>: extra info when hovering</li><li><b>Bookmarks</b>: save a view (filters, visible visuals) and jump to it with a button</li></ul>`,
[["Make cells red when sales are below target.",["Conditional formatting","Bookmarks","Append"],0,"Conditional formatting colors by rules or values."],
 ["Stop a chart from filtering a card when clicked.",["Edit interactions","Drill down","Unpivot"],0,"Edit interactions controls how visuals affect each other."],
 ["Axis has Year then Month. How do you go from years to months?",["Drill down","Fill down","Merge"],0,"Drill down moves to the next level of a hierarchy."]]),

PQ("p4","Data model","Fact and dimension tables",
`<ul><li><b>Fact table</b>: events with numbers, many rows (each sale)</li><li><b>Dimension table</b>: descriptions, fewer rows (Products, Regions, Dates)</li></ul>
<p>Connect them in <b>Model view</b> by dragging a key column onto its match. Usually <b>one-to-many</b>: one region, many sales. A fact table surrounded by dimensions is a <b>star schema</b>.</p>
<p>In this lab, <b>Sales</b> is the fact table and <b>Regions</b> is a dimension, related on Region.</p>`,
[["Which is a fact table?",["Products (ID, Name)","Sales (Date, ProductID, Qty, Amount)","Regions (Region, Manager)"],1,"Fact tables hold events and numbers."],
 ["One region appears in many sales. The relationship is…",["One-to-many","Many-to-many","One-to-one"],0,"One region (one side), many sales (many side)."],
 ["Fact in the middle, dimensions around it is called…",["Snowball","Star schema","Pivot"],1,"It looks like a star."]]),

PQ("p18","Data model","Relationships, date table, filter direction",
`<ul><li><b>Cardinality</b>: one-to-many (best), one-to-one, many-to-many (avoid when you can)</li><li><b>Cross-filter direction</b>: by default filters flow from the one side (dimension) to the many side (fact). <b>Both</b> directions can cause confusing results; use it rarely</li>
<li><b>Date table</b>: a separate table with one row per day (Date, Year, Month, Quarter). Mark it with <b>Mark as date table</b>. Time intelligence like YTD works best with it</li><li>Only one relationship between two tables can be <b>active</b> (solid line); others are dotted</li></ul>`,
[["Filters normally flow from…",["Fact to dimension","Dimension to fact","Nowhere"],1,"From the one side to the many side."],
 ["What is the recommended way to handle dates?",["Use text dates in each table","A separate date table marked as date table","Avoid dates"],1,"A proper date table makes time intelligence reliable."],
 ["A dotted line in Model view means…",["An inactive relationship","A broken table","A measure"],0,"Only one relationship between two tables can be active."]]),

PQ("p5","DAX basics","Measures vs calculated columns",
`<p><b>DAX</b> is Power BI's formula language. Columns are written <code>Table[Column]</code>.</p>
<ul><li><b>Measure</b>: one answer that changes with filters: <code>Total Sales = SUM(Sales[Amount])</code>. Use for totals, averages, %. Created with <b>New measure</b>. You'll use these most.</li><li><b>Calculated column</b>: computed row by row and stored: <code>Gross = Sales[Quantity] * Sales[Price]</code>. Created with <b>New column</b>.</li></ul>
<p>In the next lessons you'll write real DAX. Results show as a <b>card</b> and a <b>matrix</b>. Three ready-made measures exist: <code>[Total Sales]</code>, <code>[Orders]</code>, <code>[Total Qty]</code>.</p>`,
[["Total Sales that updates with a Region slicer.",["Measure","Calculated column","Power Query step"],0,"Measures recalculate for the current filters."],
 ["How is the Amount column of Sales written?",["Sales.Amount","Sales[Amount]","Amount(Sales)"],1,"DAX uses Table[Column]."],
 ["Which is calculated row by row and stored?",["Measure","Calculated column","Slicer"],1,"Calculated columns are stored per row."]]),

PQ("p19","DAX basics","Row context vs filter context",
`<p>The two most important DAX ideas:</p>
<ul><li><b>Filter context</b>: the filters active where a measure is shown. In a matrix row “North”, the filter is Region = North, so <code>SUM(Sales[Amount])</code> adds only North rows</li><li><b>Row context</b>: “the current row”. It exists in calculated columns and inside X-functions like SUMX and FILTER, so you can read <code>Sales[Quantity]</code> directly</li></ul>
<p><code>CALCULATE</code> changes the filter context. That's why it's so powerful.</p>`,
[["In a matrix row 'East', SUM(Sales[Amount]) adds…",["All rows","Only East rows","Only the first row"],1,"The matrix row creates a filter context."],
 ["Where can you write Sales[Quantity] * Sales[Price] directly?",["In a calculated column or inside SUMX","Only in a card","Nowhere"],0,"These have a row context."],
 ["Which function changes the filter context?",["CALCULATE","SUM","LEN"],0,"CALCULATE adds, removes or replaces filters."]]),

PM("p6","Measures: basics","SUM",
`<p>Your first measure. <code>SUM</code> adds up a column. In the matrix, each Region row shows only that region's total.</p><p>Write a name before the formula (<code>Total Sales = …</code>) or just the formula. <b>Tip:</b> type <code>Sa</code> and press <b>Tab</b>, then <code>[</code> to pick a column. Or click a column name in the table.</p>`,
"Total Sales = SUM(Sales[Amount])","Total Qty = SUM(Sales[Quantity])","Adds all quantities.",
"Create a measure that adds up the Amount column.","Total Sales = SUM(Sales[Amount])",["SUM"],"The column is Sales[Amount]."),

PM("p7","Measures: basics","AVERAGE, MIN, MAX",
`<p>Same as Excel: <code>AVERAGE(Sales[Amount])</code>, <code>MIN(…)</code>, <code>MAX(…)</code>.</p>`,
"Biggest Order = MAX(Sales[Amount])","Avg Order = AVERAGE(Sales[Amount])","Average order amount.",
"Create a measure for the biggest single order Amount.","Biggest Order = MAX(Sales[Amount])",["MAX"],"Biggest means MAX."),

PM("p8","Measures: basics","COUNTROWS",
`<p><code>COUNTROWS</code> counts rows in a table. Each Sales row is one order. Give it the <b>table</b>, not a column.</p>`,
"Orders = COUNTROWS(Sales)","Regions = COUNTROWS(Regions)","4 regions.",
"Create a measure that counts orders.","Orders = COUNTROWS(Sales)",["COUNTROWS"],"COUNTROWS(Sales)."),

PM("p9","Measures: basics","DISTINCTCOUNT",
`<p><code>DISTINCTCOUNT</code> counts unique values. Product repeats a lot, but there are only 6 different products.</p>`,
"Products Sold = DISTINCTCOUNT(Sales[Product])","Categories = DISTINCTCOUNT(Sales[Category])","3 categories.",
"Create a measure counting how many different products were sold.","Products Sold = DISTINCTCOUNT(Sales[Product])",["DISTINCTCOUNT"],"DISTINCTCOUNT on Sales[Product]."),

PM("p10","Measures: basics","DIVIDE",
`<p><code>/</code> fails when the bottom is 0. <code>DIVIDE(a, b)</code> returns blank instead. Functions can go inside functions.</p>`,
"DIVIDE(numerator, denominator)","Avg Qty = DIVIDE(SUM(Sales[Quantity]), COUNTROWS(Sales))","Quantity per order.",
"Create an Average Order Value: total Amount divided by number of orders.","AOV = DIVIDE(SUM(Sales[Amount]), COUNTROWS(Sales))",["DIVIDE"],"Top: SUM(Sales[Amount]). Bottom: COUNTROWS(Sales)."),

PM("p20","Measures: basics","Reusing measures",
`<p>Measures can use other measures. Write a measure name in square brackets: <code>[Total Sales]</code>. This keeps formulas short and consistent. Ready-made here: <code>[Total Sales]</code>, <code>[Orders]</code>, <code>[Total Qty]</code>.</p>`,
"New = DIVIDE([Measure A], [Measure B])","AOV = DIVIDE([Total Sales], [Orders])","Same as the last lesson, shorter.",
"Create a measure for average quantity per order using [Total Qty] and [Orders].","Qty per Order = DIVIDE([Total Qty], [Orders])",["DIVIDE"],"DIVIDE([Total Qty], [Orders])."),

PM("p11","CALCULATE","CALCULATE",
`<p><code>CALCULATE</code> runs a calculation <b>with filters you choose</b>. It's the most important DAX function.</p><p>Notice the matrix: every Region row shows the same number, because your Region filter replaces the row's filter.</p>`,
'CALCULATE(expression, Table[Column] = "value")','North Sales = CALCULATE([Total Sales], Sales[Region] = "North")',"Always North's total.",
'Create a measure for the total Amount of the "Electronics" category only.','Electronics Sales = CALCULATE(SUM(Sales[Amount]), Sales[Category] = "Electronics")',["CALCULATE"],'Filter: Sales[Category] = "Electronics".'),

PM("p21","CALCULATE","CALCULATE with two filters",
`<p>Add more filters separated by commas. All must be true (AND).</p>`,
'CALCULATE(expr, filter1, filter2)','CALCULATE([Orders], Sales[Region] = "South", Sales[Category] = "Furniture")',"South furniture orders.",
'Total Amount for Phones in the East region.','East Phones = CALCULATE([Total Sales], Sales[Product] = "Phone", Sales[Region] = "East")',["CALCULATE"],"Two filters, separated by a comma."),

PM("p22","CALCULATE","IN { list }",
`<p><code>IN</code> checks against a list in curly brackets. Shorter than many ORs.</p>`,
'Table[Column] IN {"a", "b"}','CALCULATE([Total Sales], Sales[Region] IN {"North", "South"})',"North + South.",
'Total Amount for Pen and Notebook together.','Stationery Mix = CALCULATE([Total Sales], Sales[Product] IN {"Pen", "Notebook"})',["CALCULATE"],'Sales[Product] IN {"Pen", "Notebook"}.'),

PM("p23","CALCULATE","Number filters",
`<p>Filters can compare numbers: <code>Sales[Quantity] &gt;= 5</code>.</p>`,
'CALCULATE(expr, Table[Column] > number)','CALCULATE([Orders], Sales[Discount] > 0)',"Orders with a discount.",
'Total Amount for orders with Quantity of 5 or more.','Bulk Sales = CALCULATE([Total Sales], Sales[Quantity] >= 5)',["CALCULATE"],"Sales[Quantity] >= 5."),

PM("p14","CALCULATE","ALL and % of total",
`<p><code>ALL(Sales)</code> removes all filters. Inside CALCULATE it gives the grand total even in a Region row. Divide by it to get a <b>share</b>.</p>`,
"DIVIDE([Total Sales], CALCULATE([Total Sales], ALL(Sales)))","Grand Total = CALCULATE([Total Sales], ALL(Sales))","Same in every row.",
"Create a measure showing each region's share of total Amount.","Share = DIVIDE(SUM(Sales[Amount]), CALCULATE(SUM(Sales[Amount]), ALL(Sales)))",["ALL","CALCULATE","DIVIDE"],"Bottom: CALCULATE(SUM(Sales[Amount]), ALL(Sales))."),

PM("p24","CALCULATE","ALL on one column",
`<p><code>ALL(Sales[Category])</code> removes only the Category filter and keeps others. The matrix here is <b>by Category</b>.</p>`,
"CALCULATE(expr, ALL(Table[Column]))","CALCULATE([Orders], ALL(Sales[Category]))","All orders in every row.",
"Each category's share of all categories' Amount, using ALL(Sales[Category]).","Category Share = DIVIDE([Total Sales], CALCULATE([Total Sales], ALL(Sales[Category])))",["ALL"],"Bottom: CALCULATE([Total Sales], ALL(Sales[Category])).",{by:"Category"}),

PM("p25","CALCULATE","KEEPFILTERS",
`<p>Normally a CALCULATE filter <b>replaces</b> the existing filter on that column. <code>KEEPFILTERS</code> <b>adds</b> it instead, so rows outside both become blank.</p>`,
"CALCULATE(expr, KEEPFILTERS(condition))",'CALCULATE([Total Sales], Sales[Region] = "North")',"North in every row.",
'Total Amount for North, but only in the North row (others blank). Use KEEPFILTERS.','North Only = CALCULATE([Total Sales], KEEPFILTERS(Sales[Region] = "North"))',["KEEPFILTERS"],'KEEPFILTERS(Sales[Region] = "North").'),

PM("p12","Iterators","SUMX (row by row)",
`<p>X-functions go through a table <b>row by row</b>, calculate on each row, then combine. Amount already has the discount removed; to get the amount <b>before discount</b>, multiply Quantity × Price per row.</p>`,
"SUMX(Table, expression)","Total Discount = SUMX(Sales, Sales[Discount])","Adds discounts.",
"Create a measure for gross sales: Quantity × Price per row, added up.","Gross Sales = SUMX(Sales, Sales[Quantity] * Sales[Price])",["SUMX"],"SUMX(Sales, Sales[Quantity] * Sales[Price])."),

PM("p26","Iterators","AVERAGEX and MAXX",
`<p>Same idea: <code>AVERAGEX</code> averages, <code>MAXX</code> takes the biggest, <code>MINX</code> the smallest value of the row calculation.</p>`,
"MAXX(Table, expression)","Avg Gross = AVERAGEX(Sales, Sales[Quantity] * Sales[Price])","Average gross per order.",
"Biggest single order before discount (Quantity × Price).","Max Gross = MAXX(Sales, Sales[Quantity] * Sales[Price])",["MAXX"],"MAXX(Sales, Sales[Quantity] * Sales[Price])."),

PM("p13","Iterators","FILTER",
`<p><code>FILTER(table, condition)</code> returns matching rows as a <b>table</b>. Wrap it in something that uses a table, like COUNTROWS or SUMX.</p>`,
"COUNTROWS(FILTER(Sales, condition))","Small Orders = COUNTROWS(FILTER(Sales, Sales[Amount] < 2000))","Orders under 2000.",
"Count orders with Amount greater than 20000.","Big Orders = COUNTROWS(FILTER(Sales, Sales[Amount] > 20000))",["FILTER","COUNTROWS"],"COUNTROWS(FILTER(Sales, Sales[Amount] > 20000))."),

PM("p27","Iterators","FILTER inside CALCULATE",
`<p>A FILTER table can be a CALCULATE filter. Useful when the condition compares two columns.</p>`,
"CALCULATE(expr, FILTER(Sales, condition))","CALCULATE([Orders], FILTER(Sales, Sales[Discount] > 1000))","Big-discount orders.",
"Total Amount where Discount is more than 5% of Quantity × Price.","Heavy Discount Sales = CALCULATE([Total Sales], FILTER(Sales, Sales[Discount] > 0.05 * Sales[Quantity] * Sales[Price]))",["CALCULATE","FILTER"],"FILTER(Sales, Sales[Discount] > 0.05 * Sales[Quantity] * Sales[Price])."),

PM("p28","Iterators","CONCATENATEX",
`<p><code>CONCATENATEX(table, expression, separator)</code> joins text from many rows. <code>VALUES(Sales[Product])</code> gives the unique products in the current filter.</p>`,
'CONCATENATEX(VALUES(col), col, ", ")','CONCATENATEX(VALUES(Sales[Category]), Sales[Category], " / ")',"Categories per region.",
'List the unique products for each region, separated by ", ".','Products = CONCATENATEX(VALUES(Sales[Product]), Sales[Product], ", ")',["CONCATENATEX","VALUES"],'CONCATENATEX(VALUES(Sales[Product]), Sales[Product], ", ").'),

PM("p15","Logic","IF in a measure",
`<p><code>IF</code> works like Excel's IF. In a measure, the test usually uses another measure.</p>`,
'IF(condition, "yes", "no")','Size = IF([Orders] > 5, "Busy", "Quiet")',"Labels regions.",
'Show "Top" if total Amount is more than 150000, otherwise "Growing".','Status = IF([Total Sales] > 150000, "Top", "Growing")',["IF"],"Test: [Total Sales] > 150000."),

PM("p29","Logic","SWITCH(TRUE())",
`<p><code>SWITCH(TRUE(), test1, result1, test2, result2, else)</code> is the clean way to write many IFs. It returns the result of the first true test.</p>`,
'SWITCH(TRUE(), test1, r1, test2, r2, else)','SWITCH([Orders], 6, "Six", 7, "Seven", "Other")',"SWITCH on a value.",
'Band regions: "High" if [Total Sales] ≥ 150000, "Medium" if ≥ 100000, else "Low".','Band = SWITCH(TRUE(), [Total Sales] >= 150000, "High", [Total Sales] >= 100000, "Medium", "Low")',["SWITCH"],"SWITCH(TRUE(), …) with the higher test first."),

PM("p30","Logic","VAR and RETURN",
`<p>Variables make long formulas readable. Write <code>VAR name = …</code> for each piece, then <code>RETURN</code> the final result.</p>`,
"VAR a = … VAR b = … RETURN a / b","VAR s = [Total Sales] RETURN s * 1.18","Sales with 18% GST.",
"Discount % of gross: VAR gross = SUMX(Sales, Sales[Quantity] * Sales[Price]), VAR disc = SUM(Sales[Discount]), RETURN DIVIDE(disc, gross).","Discount % = VAR gross = SUMX(Sales, Sales[Quantity] * Sales[Price]) VAR disc = SUM(Sales[Discount]) RETURN DIVIDE(disc, gross)",["SUMX","DIVIDE"],"Two VAR lines, then RETURN."),

PM("p31","Logic","SELECTEDVALUE",
`<p><code>SELECTEDVALUE(column, alternate)</code> returns the single value in the current filter, or the alternate if there are many (like in the Total row). Great for dynamic titles.</p>`,
'SELECTEDVALUE(Table[Column], "All")','SELECTEDVALUE(Sales[Region])',"Blank in the total.",
'Show "Region: " followed by the selected region, or "Region: All" in the total.','Region Label = "Region: " & SELECTEDVALUE(Sales[Region], "All")',["SELECTEDVALUE"],'"Region: " & SELECTEDVALUE(Sales[Region], "All").'),

PM("p32","Logic","FORMAT",
`<p><code>FORMAT(value, "format")</code> turns a number into formatted text: <code>"0.0%"</code>, <code>"#,0"</code>, <code>"0.00"</code>. The result is text, so use it for labels, not for further math.</p>`,
'FORMAT(value, "0.0%")','FORMAT([Total Sales], "#,0")',"With thousands separators.",
'Show each region\'s share of total Amount as text like "23.4%".','Share Text = FORMAT(DIVIDE([Total Sales], CALCULATE([Total Sales], ALL(Sales))), "0.0%")',["FORMAT"],'Wrap the share calculation in FORMAT(…, "0.0%").'),

PM("p33","Ranking","RANKX",
`<p><code>RANKX(ALL(Sales[Product]), [Total Sales])</code> ranks each product by sales (1 = highest). <code>ALL</code> makes it compare against every product. The matrix here is <b>by Product</b>.</p>`,
"RANKX(ALL(Table[Column]), expression)","RANKX(ALL(Sales[Product]), [Orders])","Rank by number of orders.",
"Rank products by [Total Sales].","Product Rank = RANKX(ALL(Sales[Product]), [Total Sales])",["RANKX","ALL"],"RANKX(ALL(Sales[Product]), [Total Sales]).",{by:"Product"}),

PM("p34","Ranking","RANKX ascending",
`<p>Add <code>, , ASC</code> to rank from lowest. The empty middle input means “use the current value”.</p>`,
"RANKX(ALL(col), expr, , ASC)","RANKX(ALL(Sales[Region]), [Total Sales])","Highest region = 1.",
"Rank regions by [Orders] where the fewest orders = 1.","Low Rank = RANKX(ALL(Sales[Region]), [Orders], , ASC)",["RANKX"],"End with , , ASC"),

PM("p35","Time intelligence","TOTALYTD",
`<p><b>Year-to-date</b> adds everything from 1 January up to the current month. The matrix here is <b>by Month</b>.</p>`,
"TOTALYTD(expression, Dates)","CALCULATE([Total Sales], DATESYTD(Sales[Date]))","Same thing, with CALCULATE.",
"Create a year-to-date sales measure.","YTD Sales = TOTALYTD([Total Sales], Sales[Date])",["TOTALYTD"],"TOTALYTD([Total Sales], Sales[Date]).",{by:"Month"}),

PM("p36","Time intelligence","PREVIOUSMONTH",
`<p><code>PREVIOUSMONTH(Sales[Date])</code> inside CALCULATE shifts to the month before. January shows blank because there is no December data.</p>`,
"CALCULATE(expr, PREVIOUSMONTH(Dates))","CALCULATE([Orders], PREVIOUSMONTH(Sales[Date]))","Last month's orders.",
"Show last month's [Total Sales] on each month.","Prev Month Sales = CALCULATE([Total Sales], PREVIOUSMONTH(Sales[Date]))",["CALCULATE","PREVIOUSMONTH"],"CALCULATE([Total Sales], PREVIOUSMONTH(Sales[Date])).",{by:"Month"}),

PM("p37","Time intelligence","Month-over-month growth",
`<p>Combine VAR, DATEADD and DIVIDE: growth = (this month − last month) ÷ last month. <code>DATEADD(Sales[Date], -1, MONTH)</code> moves back one month.</p>`,
"VAR cur = … VAR prev = CALCULATE(…, DATEADD(Dates, -1, MONTH)) RETURN DIVIDE(cur - prev, prev)","CALCULATE([Total Sales], DATEADD(Sales[Date], -1, MONTH))","Last month's sales.",
"Create a MoM growth % measure (as a decimal) using VAR, DATEADD and DIVIDE.","MoM % = VAR cur = [Total Sales] VAR prev = CALCULATE([Total Sales], DATEADD(Sales[Date], -1, MONTH)) RETURN DIVIDE(cur - prev, prev)",["DATEADD","DIVIDE"],"prev uses DATEADD(Sales[Date], -1, MONTH).",{by:"Month"}),

PM("p38","Calculated columns","Your first calculated column",
`<p>A calculated column runs once for <b>every row</b> and is stored in the table. Inside it you can use the row's own values directly. The result below shows the new column next to each order.</p>`,
"Column Name = Sales[A] * Sales[B]","Net Price = Sales[Price] - Sales[Discount]","Per row.",
"Add a Gross column: Quantity × Price.","Gross = Sales[Quantity] * Sales[Price]",[],"Sales[Quantity] * Sales[Price].",{kind:"column"}),

PM("p39","Calculated columns","IF in a column",
`<p>Use IF on each row to create categories you can slice by.</p>`,
'Size = IF(condition, "a", "b")','Has Discount = IF(Sales[Discount] > 0, "Yes", "No")',"Per row.",
'Add an Order Size column: "Bulk" if Quantity ≥ 10, otherwise "Single".','Order Size = IF(Sales[Quantity] >= 10, "Bulk", "Single")',["IF"],"Test Sales[Quantity] >= 10.",{kind:"column"}),

PM("p40","Calculated columns","RELATED",
`<p><code>RELATED</code> brings a value from a related (dimension) table into the current row, like VLOOKUP. Sales and Regions are related on Region.</p>`,
"RELATED(OtherTable[Column])","Region Target = RELATED(Regions[Target])","Target for each row's region.",
"Add the region's Manager name to each sale.","Manager = RELATED(Regions[Manager])",["RELATED"],"RELATED(Regions[Manager]).",{kind:"column"}),

PM("p41","Calculated columns","Date parts",
`<p><code>YEAR</code>, <code>MONTH</code>, <code>DAY</code> and <code>FORMAT(date, "MMM")</code> split dates into useful pieces for slicers and axes.</p>`,
"MONTH(Sales[Date])",'Month Name = FORMAT(Sales[Date], "MMM")',"Jan, Feb, Mar.",
"Add a Day column with the day of the month from Date.","Day = DAY(Sales[Date])",["DAY"],"DAY(Sales[Date]).",{kind:"column"}),

PM("p42","Calculated columns","Joining text",
`<p>Use <code>&amp;</code> to build labels from columns.</p>`,
'Label = col1 & " - " & col2','Code = LEFT(Sales[Region], 1) & Sales[OrderID]',"N1001, S1002…",
'Add a Label column like "Laptop - North".','Label = Sales[Product] & " - " & Sales[Region]',[],"Sales[Product] & \" - \" & Sales[Region].",{kind:"column"})
];
/* ---------- Theory chapters ---------- */
const T = (id, track, group, title, body, terms, questions, practice = []) =>
  ({id, track, mode:"theory", group, title, body, terms, practice, questions:questions.map(([q, opts, a, why]) => ({q, opts, a, why}))});

const VENN = (l, r, both, label) => `<figure class="venn"><svg viewBox="0 0 120 70" role="img" aria-label="${label}"><circle cx="45" cy="35" r="26" class="${l ? "on" : "off"}"/><circle cx="75" cy="35" r="26" class="${r ? "on" : "off"}"/>${both ? `<path class="on" d="M60 13.7 A26 26 0 0 1 60 56.3 A26 26 0 0 1 60 13.7Z"/>` : `<path class="mid" d="M60 13.7 A26 26 0 0 1 60 56.3 A26 26 0 0 1 60 13.7Z"/>`}<circle cx="45" cy="35" r="26" class="ring"/><circle cx="75" cy="35" r="26" class="ring"/></svg><figcaption>${label}</figcaption></figure>`;

const EXCEL_THEORY = [
T("te1","excel","Getting started","What is Excel?",
`<p><b>Microsoft Excel</b> is a spreadsheet program. You use it to store data in a grid, calculate with it, and turn it into tables, summaries and charts. Google Sheets and LibreOffice Calc work almost the same way.</p>
<h4>The parts of the screen</h4>
<ul><li><b>Workbook</b>: the whole Excel file (.xlsx)</li><li><b>Worksheet</b> (sheet): one grid inside the workbook. Tabs at the bottom switch sheets</li><li><b>Ribbon</b>: the toolbar at the top (Home, Insert, Formulas, Data…)</li><li><b>Name Box</b>: top-left box showing the selected cell, like <code>D2</code></li><li><b>Formula Bar</b>: shows what's really inside the cell, such as <code>=SUM(D2:D11)</code>, while the cell shows the result</li><li><b>Status bar</b>: bottom bar. Select numbers and it shows their Sum, Average and Count</li></ul>
<h4>Cells, rows and columns</h4>
<p>Columns are letters (A to XFD), rows are numbers (1 to 1,048,576). A cell's address is column + row: <code>B3</code>. A group of cells is a <b>range</b>: <code>B3:D10</code>.</p>`,
[["Workbook","The Excel file"],["Worksheet","One grid/tab inside it"],["Cell","One box, e.g. D2"],["Range","A block of cells, e.g. D2:D11"]],
[["What is a workbook?",["One cell","The whole Excel file","One column"],1,"A workbook is the file; it can hold many worksheets."],
 ["Where do you see the real formula inside a cell?",["Status bar","Formula Bar","Name Box"],1,"The cell shows the result; the Formula Bar shows the formula."],
 ["What is B3:D10?",["A range","A function","A sheet name"],0,"Two cells joined by a colon describe a range."]],["e1"]),

T("te2","excel","Getting started","Types of data",
`<p>Every cell holds one of these:</p>
<ul><li><b>Numbers</b>: 52000, 3.5, -10. Right-aligned by default</li><li><b>Text</b>: names, codes. Left-aligned by default. A number typed with a leading apostrophe ('0123) is stored as text</li><li><b>Dates and times</b>: stored as numbers. 1 = 1 Jan 1900, so 15/03/2021 is really 44270. That's why you can subtract dates</li><li><b>Logical</b>: TRUE or FALSE</li><li><b>Errors</b>: #N/A, #DIV/0! and friends</li></ul>
<h4>Value vs format</h4>
<p><b>Formatting</b> changes how a value <i>looks</i>, not what it <i>is</i>. 0.25 shown as 25% is still 0.25. 1234.567 shown as 1,235 is still 1234.567 in calculations.</p>
<p><b>Watch out:</b> numbers stored as text (often left-aligned with a small green triangle) won't add up. Fix with <b>Data → Text to Columns</b> or VALUE().</p>`,
[["Text","Words or codes"],["Serial date","Date stored as a day count"],["Format","How a value looks"]],
[["A date in Excel is really…",["Text","A number","A formula"],1,"Dates are day counts, which is why date math works."],
 ["You format 0.5 as 50%. What value does Excel use in formulas?",["50","0.5","50%"],1,"Formatting changes only the display."],
 ["Numbers are left-aligned with a green triangle. Most likely…",["They are stored as text","They are dates","They are negative"],0,"Text-numbers don't calculate. Convert them."]],["e22","e55"]),

T("te3","excel","Formulas","Formulas, functions and operators",
`<p>A <b>formula</b> is any calculation that starts with <code>=</code>. A <b>function</b> is a named, ready-made calculation used inside a formula: <code>=SUM(D2:D11)*2</code> is a formula that uses the SUM function.</p>
<h4>Parts of a function</h4>
<p><code>=ROUND(D2/E2, 2)</code>: the name is ROUND, the brackets hold the <b>arguments</b> (inputs), separated by commas. Some arguments are optional and shown in [square brackets] in help tips.</p>
<h4>Operators</h4>
<table class="mini"><tr><th>Math</th><td><code>+ - * / ^ %</code></td></tr><tr><th>Compare</th><td><code>= &lt;&gt; &gt; &lt; &gt;= &lt;=</code> (give TRUE/FALSE)</td></tr><tr><th>Join text</th><td><code>&amp;</code></td></tr><tr><th>Range</th><td><code>:</code> (A1:A10) and <code>,</code> (separate items)</td></tr></table>
<h4>Order of operations</h4>
<p>Like school maths (BODMAS): <b>Brackets</b> → <code>^</code> → <code>* /</code> → <code>+ -</code> → <code>&amp;</code> → comparisons. When unsure, add brackets.</p>
<h4>Nesting</h4><p>Functions can go inside other functions: <code>=ROUND(AVERAGE(D2:D11), 0)</code>. Excel works from the inside out.</p>`,
[["Formula","Anything starting with ="],["Function","Named calculation like SUM"],["Argument","An input inside the brackets"],["Nesting","A function inside another"]],
[["What does =2+3*4 give?",["20","14","24"],1,"Multiply first: 3*4 = 12, then +2."],
 ["In =ROUND(D2/E2, 2), what are D2/E2 and 2?",["Arguments","Operators","Sheets"],0,"Values inside the brackets are arguments."],
 ["Which operator joins text?",["+","&",":"],1,"& joins text: =A2&\" \"&B2."]],["e2","e3","e7"]),

T("te4","excel","Formulas","Cell references: relative, absolute, mixed",
`<p>When you copy a formula, Excel adjusts its references. This is the single most important idea for working fast.</p>
<table class="mini"><tr><th>Type</th><th>Written</th><th>When copied down/right</th></tr>
<tr><td>Relative</td><td><code>D2</code></td><td>Both column and row move</td></tr>
<tr><td>Absolute</td><td><code>$D$2</code></td><td>Never moves</td></tr>
<tr><td>Mixed</td><td><code>$D2</code> / <code>D$2</code></td><td>Column fixed / row fixed</td></tr></table>
<p>Press <kbd>F4</kbd> while editing a reference to cycle through these.</p>
<h4>Other references</h4>
<ul><li>Another sheet: <code>=Sheet2!A1</code> or <code>='Sales 2026'!A1</code> (quotes when the name has spaces)</li><li>Whole column: <code>D:D</code>. Whole row: <code>2:2</code></li><li><b>Named range</b>: give a range a name (Formulas → Define Name), then write <code>=SUM(Sales)</code></li></ul>`,
[["Relative","Moves when copied"],["Absolute","$ locks it"],["F4","Toggles the $ signs"]],
[["You copy =D2*$H$1 from row 2 to row 3. It becomes…",["=D3*$H$2","=D3*$H$1","=D2*$H$1"],1,"D2 moves to D3; $H$1 is locked."],
 ["Which keeps the column fixed but lets the row change?",["$D2","D$2","D2"],0,"$ before the letter locks the column."],
 ["How do you refer to A1 on a sheet called Sales 2026?",["Sales 2026!A1","'Sales 2026'!A1","A1@Sales 2026"],1,"Sheet names with spaces need single quotes."]],["e8"]),

T("te5","excel","Formulas","Errors and how to fix them",
`<table class="mini"><tr><th>Error</th><th>Meaning</th><th>Usual fix</th></tr>
<tr><td><code>#NAME?</code></td><td>Unknown name</td><td>Check spelling; put text in "quotes"</td></tr>
<tr><td><code>#DIV/0!</code></td><td>Divided by zero or blank</td><td>Check the divisor; use IFERROR</td></tr>
<tr><td><code>#N/A</code></td><td>Lookup value not found</td><td>Check spelling/extra spaces (TRIM); use IFNA</td></tr>
<tr><td><code>#VALUE!</code></td><td>Wrong type, e.g. text in math</td><td>Convert text to numbers</td></tr>
<tr><td><code>#REF!</code></td><td>Cell reference deleted</td><td>Undo or rewrite the reference</td></tr>
<tr><td><code>#NUM!</code></td><td>Invalid number</td><td>Check inputs (e.g. SQRT of negative)</td></tr>
<tr><td><code>#SPILL!</code></td><td>Result can't spill into cells</td><td>Clear cells below/right</td></tr>
<tr><td><code>#####</code></td><td>Column too narrow (not an error)</td><td>Widen the column</td></tr></table>
<p><b>Debug tools:</b> select a part of a formula and press <kbd>F9</kbd> to see its value; use <b>Formulas → Evaluate Formula</b> to step through; <b>Trace Precedents</b> draws arrows to the cells a formula uses.</p>`,
[["#N/A","Not found"],["#REF!","Broken reference"],["IFERROR","Catches any error"]],
[["A VLOOKUP shows #N/A. Most likely…",["Value not found","Column too narrow","Deleted sheet"],0,"#N/A means the lookup failed."],
 ["A cell shows ##### . What do you do?",["Delete the formula","Widen the column","Use IFERROR"],1,"##### just means the column is too narrow."],
 ["You deleted a column a formula used. You'll see…",["#REF!","#NAME?","#NUM!"],0,"#REF! means the reference no longer exists."]],["e17"]),

T("te6","excel","Working with data","Tables, sort, filter and validation",
`<h4>Excel Tables</h4><p>Select your data and press <kbd>Ctrl</kbd>+<kbd>T</kbd>. A Table auto-expands when you add rows, adds filter buttons, keeps formatting, and lets formulas use names like <code>=SUM(Table1[Sales])</code>. Always use Tables for lists of records.</p>
<h4>Sort and filter</h4><ul><li><b>Sort</b> (Data → Sort): order rows by one or more columns</li><li><b>Filter</b> (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>L</kbd>): hide rows that don't match. SUBTOTAL ignores hidden rows; SUM doesn't</li></ul>
<h4>Clean-up tools</h4><ul><li><b>Remove Duplicates</b> (Data tab)</li><li><b>Text to Columns</b>: split "Mumbai-West" by "-"</li><li><b>Flash Fill</b> (<kbd>Ctrl</kbd>+<kbd>E</kbd>): type one example, Excel copies the pattern</li><li><b>Find &amp; Replace</b> (<kbd>Ctrl</kbd>+<kbd>H</kbd>)</li></ul>
<h4>Data Validation</h4><p>Data → Data Validation limits what people can type: a dropdown list, whole numbers between 1–100, dates only. It prevents bad data at the source.</p>
<h4>Freeze Panes</h4><p>View → Freeze Panes keeps the header row visible while you scroll.</p>`,
[["Ctrl+T","Create a Table"],["Filter","Hide non-matching rows"],["Data Validation","Control what can be typed"]],
[["Best way to make a list that auto-expands for formulas?",["Merge cells","Format as Table (Ctrl+T)","Freeze panes"],1,"Tables grow automatically."],
 ["You want a dropdown of departments in a cell.",["Data Validation → List","Conditional Formatting","PivotTable"],0,"Validation with a List gives a dropdown."],
 ["After filtering, which function totals only visible rows?",["SUM","SUBTOTAL","COUNT"],1,"SUBTOTAL ignores hidden rows."]],["e44"]),

T("te7","excel","Working with data","Lookups explained",
`<p>A lookup answers: “find this value in a list, and give me something from the same row”.</p>
<table class="mini"><tr><th>Function</th><th>Strengths</th><th>Limits</th></tr>
<tr><td>VLOOKUP</td><td>Everywhere, simple</td><td>Searches only the first column, returns only to the right, breaks if columns are inserted</td></tr>
<tr><td>HLOOKUP</td><td>For data laid out across</td><td>Same limits, sideways</td></tr>
<tr><td>INDEX + MATCH</td><td>Any direction, robust</td><td>Two functions to learn</td></tr>
<tr><td>XLOOKUP</td><td>Any direction, built-in "not found", simplest</td><td>Excel 2021/365 only</td></tr></table>
<h4>Exact vs approximate</h4><p>VLOOKUP's last argument: <b>FALSE</b> = exact match (use this for names, IDs). <b>TRUE</b> = approximate, for grade or tax slabs, and the first column must be sorted ascending.</p>
<p><b>Common cause of #N/A:</b> extra spaces or numbers stored as text in one list and as numbers in the other.</p>`,
[["Lookup value","What you search for"],["Exact match","FALSE / 0"],["XLOOKUP","Modern replacement"]],
[["Which lookup can return a column to the LEFT of the search column?",["VLOOKUP","INDEX + MATCH","HLOOKUP"],1,"INDEX+MATCH (and XLOOKUP) work in any direction."],
 ["Looking up employee IDs, the last VLOOKUP argument should be…",["TRUE","FALSE","1"],1,"IDs need exact match."],
 ["Approximate match needs the first column to be…",["Sorted ascending","Text","Hidden"],0,"Otherwise results are wrong."]],["e16","e45","e18"]),

T("te8","excel","Working with data","PivotTables",
`<p>A <b>PivotTable</b> summarises thousands of rows in seconds, with no formulas. Select your data → <b>Insert → PivotTable</b>.</p>
<p>Drag fields into four areas:</p>
<ul><li><b>Rows</b>: categories down the side (Department)</li><li><b>Columns</b>: categories across the top (City)</li><li><b>Values</b>: numbers to summarise (Sum of Sales). Click it → Value Field Settings to switch to Count, Average, Max, % of total</li><li><b>Filters</b>: filter the whole pivot (Year = 2026)</li></ul>
<p>Extras: <b>Slicers</b> are clickable filter buttons; <b>Group</b> dates by month/quarter; <b>PivotCharts</b> chart the pivot. When the source data changes, click <b>Refresh</b>.</p>
<p>A PivotTable does the same job as SUMIFS/COUNTIFS for every combination at once. SQL's GROUP BY and Power BI's matrix are the same idea.</p>`,
[["Rows/Columns","Categories"],["Values","Numbers summarised"],["Refresh","Update after data changes"]],
[["Where do you put Sales to see total sales?",["Rows","Values","Filters"],1,"Values holds what gets summarised."],
 ["Source data changed. The pivot shows old numbers. You…",["Rebuild it","Click Refresh","Use F4"],1,"Pivots update on Refresh."],
 ["Which SQL idea is closest to a PivotTable?",["GROUP BY","DELETE","LIKE"],0,"Both group rows and summarise."]],["e15"]),

T("te9","excel","Working with data","Charts and conditional formatting",
`<h4>Choosing a chart</h4><ul><li><b>Column/Bar</b>: compare categories</li><li><b>Line</b>: trend over time</li><li><b>Pie</b>: parts of a whole, max 5–6 slices</li><li><b>Scatter</b>: relationship between two numbers</li><li><b>Combo</b>: two measures with different scales (sales and %)</li></ul>
<p>Select data → <b>Insert → Recommended Charts</b>. Always give the chart a clear title and label the axes.</p>
<h4>Conditional formatting</h4><p>Home → Conditional Formatting colours cells by rules:</p><ul><li><b>Highlight rules</b>: greater than, text contains, duplicates</li><li><b>Top/Bottom</b>: top 10, above average</li><li><b>Data bars, colour scales, icon sets</b>: visual heat maps</li><li><b>Use a formula</b>: e.g. <code>=$D2&lt;$E2</code> to colour the whole row when sales are below target</li></ul>`,
[["Line chart","Trends over time"],["Conditional formatting","Colour by rules"]],
[["Show monthly sales over a year.",["Pie","Line","Scatter"],1,"Lines show trends."],
 ["Colour duplicate names red.",["Conditional formatting → Duplicate Values","Data Validation","Sort"],0,"It's a built-in highlight rule."],
 ["A formula rule =$D2<$E2 applied to A2:F11 colours…",["Only column D","Whole rows where sales < target","Nothing"],1,"$D/$E lock the columns so each row checks its own values."]]),

T("te10","excel","Working smarter","Keyboard shortcuts",
`<table class="mini"><tr><th>Shortcut</th><th>Does</th></tr>
<tr><td><kbd>Ctrl</kbd>+<kbd>C</kbd> / <kbd>V</kbd> / <kbd>Z</kbd></td><td>Copy / paste / undo</td></tr>
<tr><td><kbd>Ctrl</kbd>+<kbd>Arrow</kbd></td><td>Jump to the edge of data</td></tr>
<tr><td><kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>Arrow</kbd></td><td>Select to the edge of data</td></tr>
<tr><td><kbd>Ctrl</kbd>+<kbd>T</kbd></td><td>Make a Table</td></tr>
<tr><td><kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>L</kbd></td><td>Filter on/off</td></tr>
<tr><td><kbd>Alt</kbd>+<kbd>=</kbd></td><td>AutoSum</td></tr>
<tr><td><kbd>F2</kbd></td><td>Edit the active cell</td></tr>
<tr><td><kbd>F4</kbd></td><td>Toggle $ in a reference / repeat last action</td></tr>
<tr><td><kbd>Ctrl</kbd>+<kbd>D</kbd></td><td>Fill down</td></tr>
<tr><td><kbd>Ctrl</kbd>+<kbd>;</kbd></td><td>Insert today's date</td></tr>
<tr><td><kbd>Ctrl</kbd>+<kbd>1</kbd></td><td>Format Cells</td></tr>
<tr><td><kbd>Ctrl</kbd>+<kbd>E</kbd></td><td>Flash Fill</td></tr>
<tr><td><kbd>Tab</kbd> while typing a function</td><td>Accept the suggestion</td></tr>
<tr><td>Double-click the fill handle</td><td>Copy a formula down the whole column</td></tr></table>`,
[["Alt+=","AutoSum"],["F2","Edit cell"],["Ctrl+D","Fill down"]],
[["Quickest way to add a SUM below a column?",["Alt+=","Ctrl+T","F9"],0,"Alt+= inserts AutoSum."],
 ["Select from the current cell to the last filled cell below?",["Ctrl+Shift+Down","Ctrl+Down","Shift+F4"],0,"Adding Shift selects while jumping."],
 ["Open the Format Cells window?",["Ctrl+1","Ctrl+F","Ctrl+P"],0,"Ctrl+1 opens Format Cells."]]),

T("te11","excel","Working smarter","Good spreadsheet habits",
`<ul><li><b>One header row</b>, one record per row, one kind of data per column</li><li><b>No merged cells</b> in data; they break sorting, filtering and pivots. Use “Center Across Selection” instead</li><li><b>No blank rows or columns</b> inside a data block</li><li><b>Keep inputs separate</b> from calculations. Put assumptions (tax %, rates) in their own labelled cells and reference them with $</li><li><b>Don't hard-code numbers</b> inside formulas: <code>=D2*$H$1</code>, not <code>=D2*0.18</code></li><li>Use <b>consistent formats</b>: same date format, same spelling (“Mumbai” not “mumbai ”)</li><li><b>Label everything</b> and add a notes sheet explaining what the file does</li><li>Keep raw data untouched; do the analysis on another sheet</li></ul>
<p>Data laid out this way also loads cleanly into SQL and Power BI.</p>`,
[["Tidy data","One row per record, one column per field"],["Hard-coding","Typing fixed numbers into formulas"]],
[["Why avoid merged cells in data?",["They use more memory","They break sort, filter and pivots","They change fonts"],1,"Merged cells disrupt the grid structure."],
 ["The GST rate may change. Best approach?",["Type 0.18 in every formula","Put it in one cell and reference it with $","Use a chart"],1,"One input cell is easy to update."],
 ["Tidy data means…",["Colourful formatting","One record per row, one field per column","Many blank rows"],1,"That's the shape all tools expect."]])
];

const SQL_THEORY = [
T("ts1","sql","Database basics","What is a database?",
`<p>A <b>database</b> is an organised store of data that many people and apps can use safely at the same time. Unlike a spreadsheet, it enforces rules, handles millions of rows, and many users can read and write at once.</p>
<p>A <b>DBMS</b> (Database Management System) is the software that runs it. Most business databases are <b>relational</b> (RDBMS): data lives in <b>tables</b> that are linked to each other.</p>
<ul><li><b>Table</b>: a set of records about one thing (employees)</li><li><b>Row</b> (record): one item (one employee)</li><li><b>Column</b> (field): one attribute (salary)</li></ul>
<p>Popular RDBMS: <b>MySQL</b>, <b>PostgreSQL</b>, <b>Microsoft SQL Server</b>, <b>Oracle</b>, <b>SQLite</b> (used in this lab). All speak <b>SQL</b> (Structured Query Language), with small differences called <i>dialects</i>.</p>`,
[["RDBMS","Relational database software"],["Row","One record"],["Column","One field"],["SQL","Language to talk to it"]],
[["A row in the employees table represents…",["One attribute","One employee","The whole company"],1,"Each row is one record."],
 ["Which is NOT a relational database?",["PostgreSQL","SQL Server","Microsoft Paint"],2,"The others are RDBMS products."],
 ["SQL is…",["A database product","A language to query databases","A file type"],1,"SQL is the language; MySQL etc. are products."]],["s1"]),

T("ts2","sql","Database basics","Keys and relationships",
`<ul><li><b>Primary key (PK)</b>: a column that uniquely identifies each row, never NULL. Here: <code>employees.id</code></li><li><b>Foreign key (FK)</b>: a column that points to a primary key in another table. Here: <code>projects.employee_id</code> → <code>employees.id</code></li></ul>
<h4>Relationship types</h4>
<ul><li><b>One-to-many</b> (most common): one employee, many sales</li><li><b>One-to-one</b>: one employee, one locker</li><li><b>Many-to-many</b>: students and courses. Solved with a <b>junction table</b> (enrolments: student_id, course_id)</li></ul>
<p>Keys are what JOINs use to bring related data together.</p>`,
[["Primary key","Unique ID of a row"],["Foreign key","Points to another table's PK"],["Junction table","Resolves many-to-many"]],
[["Which column is the best primary key for employees?",["name","id","city"],1,"IDs are unique and stable; names can repeat."],
 ["sales.employee_id refers to employees.id. It is a…",["Foreign key","Primary key","Index"],0,"It points to another table's key."],
 ["Students ↔ Courses is many-to-many. You need…",["A junction table","Two primary keys in one column","A view"],0,"A middle table links them."]],["s18"]),

T("ts3","sql","Database basics","Data types and NULL",
`<table class="mini"><tr><th>Type</th><th>Stores</th><th>Example</th></tr>
<tr><td>INT / INTEGER</td><td>Whole numbers</td><td>52000</td></tr><tr><td>DECIMAL(10,2) / NUMERIC</td><td>Exact decimals (money)</td><td>1234.50</td></tr>
<tr><td>FLOAT / REAL</td><td>Approximate decimals</td><td>3.14159</td></tr><tr><td>VARCHAR(n) / TEXT</td><td>Text</td><td>'Mumbai'</td></tr>
<tr><td>DATE / DATETIME / TIMESTAMP</td><td>Dates and times</td><td>'2026-03-15'</td></tr><tr><td>BOOLEAN / BIT</td><td>True/false</td><td>1</td></tr></table>
<h4>NULL</h4><p><b>NULL</b> means “unknown / no value”. It is not 0 and not an empty string. Any comparison with NULL is unknown, so <code>bonus = NULL</code> never matches. Use <code>IS NULL</code>. Aggregates like AVG and COUNT(column) skip NULLs; COUNT(*) counts every row.</p>`,
[["VARCHAR","Variable-length text"],["DECIMAL","Exact numbers, good for money"],["NULL","Missing value"]],
[["Best type for prices?",["FLOAT","DECIMAL","VARCHAR"],1,"DECIMAL is exact; FLOAT can have rounding errors."],
 ["WHERE bonus = NULL returns…",["Rows with no bonus","No rows","An error in every database"],1,"NULL never equals anything; use IS NULL."],
 ["COUNT(bonus) vs COUNT(*): COUNT(bonus)…",["Skips NULLs","Counts all rows","Adds bonuses"],0,"COUNT(column) ignores NULLs."]],["s16","s27"]),

T("ts4","sql","SQL language","SQL command families",
`<table class="mini"><tr><th>Family</th><th>Purpose</th><th>Commands</th></tr>
<tr><td><b>DQL</b> Query</td><td>Read data</td><td>SELECT</td></tr>
<tr><td><b>DML</b> Manipulation</td><td>Change rows</td><td>INSERT, UPDATE, DELETE, MERGE</td></tr>
<tr><td><b>DDL</b> Definition</td><td>Change structure</td><td>CREATE, ALTER, DROP, TRUNCATE</td></tr>
<tr><td><b>DCL</b> Control</td><td>Permissions</td><td>GRANT, REVOKE</td></tr>
<tr><td><b>TCL</b> Transactions</td><td>Save or undo work</td><td>COMMIT, ROLLBACK, SAVEPOINT</td></tr></table>
<p><b>DELETE vs TRUNCATE vs DROP</b>: DELETE removes chosen rows (can use WHERE, can roll back); TRUNCATE empties the whole table quickly; DROP removes the table itself.</p>
<p>As an analyst you'll mostly write DQL (SELECT). Be very careful with UPDATE/DELETE: always write the WHERE first, and test it with a SELECT.</p>`,
[["DQL","SELECT"],["DML","INSERT/UPDATE/DELETE"],["DDL","CREATE/ALTER/DROP"],["TCL","COMMIT/ROLLBACK"]],
[["CREATE TABLE belongs to…",["DML","DDL","DCL"],1,"It defines structure."],
 ["Remove only rows where amount < 100.",["DELETE … WHERE","TRUNCATE","DROP"],0,"Only DELETE accepts a WHERE."],
 ["GRANT SELECT ON sales TO priya is…",["DCL","TCL","DQL"],0,"Permissions are DCL."]],["s54","s55","s56","s57"]),

T("ts5","sql","SQL language","How a SELECT runs",
`<p>You <i>write</i> a query in this order:</p>
<p><code>SELECT … FROM … WHERE … GROUP BY … HAVING … ORDER BY … LIMIT</code></p>
<p>But the database <i>runs</i> it in this logical order:</p>
<ol class="steps"><li><b>FROM / JOIN</b>: get the tables</li><li><b>WHERE</b>: filter rows</li><li><b>GROUP BY</b>: make groups</li><li><b>HAVING</b>: filter groups</li><li><b>SELECT</b>: pick and calculate columns (aliases are created here)</li><li><b>DISTINCT</b></li><li><b>ORDER BY</b>: sort (aliases can be used here)</li><li><b>LIMIT / OFFSET</b></li></ol>
<p>This explains common errors: you <b>can't use a SELECT alias in WHERE</b> (it doesn't exist yet), and you <b>can't use SUM() in WHERE</b> (groups don't exist yet), so use HAVING.</p>`,
[["Logical order","FROM → WHERE → GROUP BY → HAVING → SELECT → ORDER BY"]],
[["Which runs first?",["SELECT","WHERE","ORDER BY"],1,"WHERE runs before SELECT."],
 ["Why does WHERE yearly > 500000 fail when yearly is a SELECT alias?",["Aliases are created after WHERE runs","WHERE can't use numbers","Aliases need quotes"],0,"SELECT runs after WHERE."],
 ["To filter on COUNT(*) > 3 you use…",["WHERE","HAVING","LIMIT"],1,"HAVING filters groups."]],["s15"]),

T("ts6","sql","SQL language","Joins explained",
`<p>A join combines rows from two tables where a condition matches (usually key = key).</p>
<div class="venns">${VENN(false,false,true,"INNER JOIN: only matches")}${VENN(true,false,true,"LEFT JOIN: all left + matches")}${VENN(false,true,true,"RIGHT JOIN: all right + matches")}${VENN(true,true,true,"FULL OUTER: everything")}</div>
<ul><li><b>INNER JOIN</b>: only rows that match in both</li><li><b>LEFT JOIN</b>: every row of the left table; NULLs where no match</li><li><b>RIGHT JOIN</b>: every row of the right table</li><li><b>FULL OUTER JOIN</b>: all rows from both</li><li><b>CROSS JOIN</b>: every combination (rows × rows)</li><li><b>SELF JOIN</b>: a table joined to itself (employee → mentor)</li></ul>
<p><b>Row count tip:</b> if one employee has 2 projects, a join gives that employee 2 rows. Unexpected row growth usually means a one-to-many join.</p>`,
[["INNER","Matches only"],["LEFT","Keep all left rows"],["Anti join","LEFT JOIN … WHERE right.key IS NULL"]],
[["List all departments, even with no employees.",["INNER JOIN employees","departments LEFT JOIN employees","CROSS JOIN"],1,"LEFT keeps every department."],
 ["3 rows × 4 rows with CROSS JOIN gives…",["7","12","4"],1,"Every combination: 3 × 4."],
 ["Find customers with no orders.",["LEFT JOIN + WHERE order_id IS NULL","INNER JOIN","UNION"],0,"The classic anti-join."]],["s18","s19","s40","s41","s42"]),

T("ts7","sql","SQL language","Grouping and aggregation",
`<p><b>Aggregate functions</b> turn many rows into one value: COUNT, SUM, AVG, MIN, MAX.</p>
<p><b>GROUP BY</b> splits rows into groups first, then aggregates each group. Rule: every column in SELECT must be either in GROUP BY or inside an aggregate.</p>
<table class="mini"><tr><th>WHERE</th><th>HAVING</th></tr><tr><td>Filters rows</td><td>Filters groups</td></tr><tr><td>Before grouping</td><td>After grouping</td></tr><tr><td>Can't use SUM/COUNT</td><td>Uses SUM/COUNT</td></tr></table>
<p><b>Window functions</b> (… OVER (…)) also aggregate, but keep every row, e.g. show each salary next to its department average.</p>`,
[["Aggregate","Many rows → one value"],["GROUP BY","One row per group"],["Window function","Aggregate without collapsing rows"]],
[["SELECT city, name, COUNT(*) … GROUP BY city fails because…",["name isn't grouped or aggregated","COUNT needs a column","city is text"],0,"Every column must be grouped or aggregated."],
 ["Show each row with its department average.",["GROUP BY","A window function: AVG(salary) OVER (PARTITION BY department)","HAVING"],1,"Window functions keep rows."],
 ["Keep only departments with total salary over 1,00,000.",["WHERE SUM(salary) > 100000","HAVING SUM(salary) > 100000","LIMIT 100000"],1,"Group filters go in HAVING."]],["s14","s15","s51"]),

T("ts8","sql","SQL language","Subqueries, CTEs and set operations",
`<ul><li><b>Subquery</b>: a query inside brackets. Can return one value (<code>WHERE salary &gt; (SELECT AVG…)</code>), a list (<code>IN (SELECT …)</code>) or a table (<code>FROM (SELECT …) t</code>)</li><li><b>Correlated subquery</b>: refers to the outer row, runs once per row (EXISTS examples)</li><li><b>CTE</b> (<code>WITH name AS (…)</code>): a named subquery at the top. Easier to read, can be reused in the query</li></ul>
<h4>Set operations</h4><ul><li><b>UNION</b>: rows from both, duplicates removed. <b>UNION ALL</b> keeps duplicates (faster)</li><li><b>INTERSECT</b>: rows in both</li><li><b>EXCEPT</b> / MINUS: rows in the first only</li></ul><p>Set operations need the same number of columns with compatible types.</p>`,
[["CTE","WITH … AS (…)"],["UNION ALL","Stack, keep duplicates"],["Correlated","Subquery uses outer row"]],
[["Which is easiest to read for a multi-step query?",["Nested subqueries","CTEs","Cross joins"],1,"CTEs name each step."],
 ["Combine two lists and keep duplicates.",["UNION","UNION ALL","INTERSECT"],1,"UNION ALL doesn't remove duplicates."],
 ["Names in list A but not in list B.",["EXCEPT","INTERSECT","JOIN"],0,"EXCEPT returns the difference."]],["s20","s46","s48"]),

T("ts9","sql","Database design","Constraints",
`<p><b>Constraints</b> are rules the database enforces so bad data can't get in.</p>
<table class="mini"><tr><th>Constraint</th><th>Rule</th></tr>
<tr><td>NOT NULL</td><td>Value required</td></tr><tr><td>UNIQUE</td><td>No duplicates (e.g. email)</td></tr><tr><td>PRIMARY KEY</td><td>NOT NULL + UNIQUE row ID</td></tr>
<tr><td>FOREIGN KEY</td><td>Value must exist in the parent table</td></tr><tr><td>CHECK</td><td>Custom rule, e.g. <code>CHECK (salary &gt; 0)</code></td></tr><tr><td>DEFAULT</td><td>Value used when none is given</td></tr></table>
<pre class="code-block">CREATE TABLE employees (
  id INT PRIMARY KEY,
  email VARCHAR(100) UNIQUE NOT NULL,
  salary DECIMAL(10,2) CHECK (salary > 0),
  dept_id INT REFERENCES departments(id),
  created DATE DEFAULT CURRENT_DATE
);</pre>`,
[["CHECK","Custom rule"],["UNIQUE","No duplicates"],["DEFAULT","Fallback value"]],
[["Stop two employees having the same email.",["UNIQUE","DEFAULT","CHECK (email > 0)"],0,"UNIQUE blocks duplicates."],
 ["Stop sales.employee_id pointing to a non-existent employee.",["FOREIGN KEY","NOT NULL","INDEX"],0,"A foreign key checks the parent table."],
 ["PRIMARY KEY is the same as…",["NOT NULL + UNIQUE","CHECK + DEFAULT","FOREIGN KEY"],0,"It requires a unique, non-null value."]]),

T("ts10","sql","Database design","Normalisation",
`<p><b>Normalisation</b> means organising tables to avoid repeating data, so each fact is stored once.</p>
<ul><li><b>1NF</b>: one value per cell, no repeating groups. Not “Pen, Bag” in one cell</li><li><b>2NF</b>: every column depends on the whole key, not part of it</li><li><b>3NF</b>: columns depend only on the key, not on other columns. Don't store the manager's name in every employee row; store it once in departments</li></ul>
<p><b>Why it matters:</b> if the HR manager changes and her name is in 50 rows, you must update 50 rows and might miss one. In a normalised design you update one row.</p>
<p>Reporting models (data warehouses, Power BI) are sometimes deliberately <b>denormalised</b> into star schemas for speed and simplicity.</p>`,
[["1NF","One value per cell"],["3NF","No column depends on another non-key column"],["Denormalise","Duplicate on purpose for reporting"]],
[["A cell holds 'Pen, Bag, Bottle'. Which form is broken?",["1NF","3NF","None"],0,"1NF requires one value per cell."],
 ["Manager name repeated in every employee row. Better design?",["Keep it","Move it to a departments table","Add another column"],1,"Store each fact once."],
 ["Main benefit of normalisation?",["Prettier output","Less duplication and fewer update mistakes","More columns"],1,"One fact, one place."]]),

T("ts11","sql","Database design","Indexes, views and transactions",
`<h4>Indexes</h4><p>An <b>index</b> is like a book's index: it helps the database find rows without scanning the whole table. Speeds up WHERE and JOIN on that column, but slows inserts a bit. Primary keys are indexed automatically.</p>
<h4>Views</h4><p>A <b>view</b> is a saved SELECT that behaves like a table: <code>CREATE VIEW it_staff AS SELECT * FROM employees WHERE department='IT';</code>. Great for reuse and for hiding complexity.</p>
<h4>Transactions and ACID</h4><p>A <b>transaction</b> groups changes so they all succeed or all fail (transfer money: debit and credit together). <code>BEGIN; … COMMIT;</code> or <code>ROLLBACK;</code> to undo.</p>
<ul><li><b>A</b>tomicity: all or nothing</li><li><b>C</b>onsistency: rules always hold</li><li><b>I</b>solation: transactions don't see each other's half-done work</li><li><b>D</b>urability: committed data survives a crash</li></ul>`,
[["Index","Faster lookups"],["View","Saved query"],["ACID","Reliable transactions"]],
[["Queries filtering by email are slow. Add…",["An index on email","A view","A CHECK constraint"],0,"Indexes speed up searches."],
 ["Undo changes in the current transaction.",["COMMIT","ROLLBACK","GRANT"],1,"ROLLBACK cancels uncommitted work."],
 ["'All or nothing' is which ACID property?",["Atomicity","Durability","Isolation"],0,"Atomicity = indivisible."]]),

T("ts12","sql","Database design","SQL dialect differences",
`<p>The core of SQL is the same everywhere, but details differ:</p>
<table class="mini"><tr><th>Task</th><th>MySQL / PostgreSQL / SQLite</th><th>SQL Server</th><th>Oracle</th></tr>
<tr><td>Top 5 rows</td><td><code>LIMIT 5</code></td><td><code>SELECT TOP 5</code></td><td><code>FETCH FIRST 5 ROWS ONLY</code></td></tr>
<tr><td>Join text</td><td><code>CONCAT()</code> / <code>||</code></td><td><code>+</code> / <code>CONCAT()</code></td><td><code>||</code></td></tr>
<tr><td>Replace NULL</td><td><code>COALESCE</code> / <code>IFNULL</code></td><td><code>ISNULL</code> / <code>COALESCE</code></td><td><code>NVL</code> / <code>COALESCE</code></td></tr>
<tr><td>Current date</td><td><code>CURRENT_DATE</code></td><td><code>GETDATE()</code></td><td><code>SYSDATE</code></td></tr>
<tr><td>Year of a date</td><td><code>YEAR(d)</code> / <code>EXTRACT</code> / <code>strftime</code></td><td><code>YEAR(d)</code></td><td><code>EXTRACT(YEAR FROM d)</code></td></tr></table>
<p>Tip: <code>COALESCE</code>, <code>CASE</code>, and standard joins work almost everywhere, so prefer them.</p>`,
[["Dialect","A database's flavour of SQL"]],
[["Top 3 rows in SQL Server?",["LIMIT 3","SELECT TOP 3","ROWNUM 3"],1,"SQL Server uses TOP."],
 ["Which NULL function works in nearly every database?",["NVL","ISNULL","COALESCE"],2,"COALESCE is standard SQL."],
 ["Today's date in SQL Server?",["GETDATE()","SYSDATE","NOW"],0,"GETDATE() is SQL Server's function."]],["s9","s35"])
];

/* Power BI theory: new chapters plus the concept lessons moved from practice */
const PBI_THEORY_NEW = {
 tp0:T("tp0","pbi","Power BI basics","What is Business Intelligence?",
`<p><b>Business Intelligence (BI)</b> turns raw data into information people can act on. The chain is: <b>data</b> (rows in systems) → <b>information</b> (totals, trends) → <b>insight</b> (why sales dropped in the South) → <b>decision</b>.</p>
<ul><li><b>Report</b>: one or more pages of visuals exploring a topic</li><li><b>Dashboard</b>: a single page of key tiles for monitoring at a glance</li><li><b>KPI</b> (Key Performance Indicator): a number tracked against a target (sales vs target, ticket SLA %)</li></ul>
<p><b>Power BI</b> is Microsoft's BI tool. Similar tools include Tableau, Qlik and Looker. Excel, SQL and Power BI work together: SQL pulls data from databases, Power BI models and visualises it, and Excel is used for quick analysis and sharing.</p>`,
[["BI","Data → decisions"],["KPI","Measured against a target"],["Dashboard","One-page overview"]],
[["A number tracked against a target is a…",["KPI","Slicer","Column"],0,"Key Performance Indicator."],
 ["A one-page, at-a-glance overview is usually called a…",["Dashboard","Query","Table"],0,"Dashboards are for monitoring."],
 ["Correct order?",["Insight → data → decision","Data → information → insight → decision","Decision → data"],1,"Raw data becomes information, then insight, then action."]]),
 tp1:T("tp1","pbi","Power BI basics","Data sources and connection modes",
`<p><b>Get data</b> connects to hundreds of sources: Excel, CSV, SQL Server, SharePoint lists, web pages, Salesforce, and more.</p>
<table class="mini"><tr><th>Mode</th><th>How it works</th><th>Good for</th></tr>
<tr><td><b>Import</b></td><td>Copies data into Power BI; refreshed on a schedule</td><td>Fastest visuals, most features. The default</td></tr>
<tr><td><b>DirectQuery</b></td><td>Leaves data in the source; asks it on every click</td><td>Very large or near-real-time data</td></tr>
<tr><td><b>Live connection</b></td><td>Connects to an existing published model</td><td>Reusing a company “golden” dataset</td></tr>
<tr><td><b>Composite</b></td><td>Mix of Import and DirectQuery</td><td>Big fact table + small imported dimensions</td></tr></table>
<p>With Import, numbers are only as fresh as the last <b>refresh</b>.</p>`,
[["Import","Copy data in, refresh later"],["DirectQuery","Query the source live"],["Refresh","Reload imported data"]],
[["Default and fastest mode?",["Import","DirectQuery","Live connection"],0,"Import holds data in memory."],
 ["Data must be near real time and is huge.",["Import","DirectQuery","Copy-paste"],1,"DirectQuery reads the source each time."],
 ["An imported report shows yesterday's numbers. Why?",["It hasn't been refreshed yet","DAX is wrong","Visuals are cached forever"],0,"Imported data updates on refresh."]]),
 tp2:T("tp2","pbi","Sharing","Power BI Service and sharing",
`<ul><li><b>Publish</b>: from Desktop, Home → Publish sends the report and its <b>semantic model</b> (dataset) to a <b>workspace</b></li><li><b>Workspace</b>: a shared folder for a team. Roles: Admin, Member, Contributor, Viewer</li><li><b>App</b>: a polished package of reports from a workspace for a wide audience</li><li><b>Dashboard</b> (in the Service): pin tiles from different reports onto one page</li><li><b>Scheduled refresh</b>: refresh imported data automatically, e.g. daily at 8 AM</li><li><b>On-premises data gateway</b>: lets the cloud Service refresh data that lives on your company's own servers</li><li>Other ways to share: share link, Teams tab, embed in SharePoint, subscribe by email, export to PDF/PowerPoint, Analyze in Excel</li></ul>`,
[["Workspace","Team folder in the Service"],["App","Packaged reports for consumers"],["Gateway","Bridge to on-premises data"]],
[["Refreshing a report whose data is on a company server needs…",["A gateway","A slicer","DAX"],0,"The gateway bridges the cloud and local data."],
 ["Best way to distribute finished reports to 200 viewers?",["Email the .pbix","Publish an App","Screenshot it"],1,"Apps are made for broad distribution."],
 ["Where do team members collaborate on reports in the Service?",["Workspace","Power Query","Model view"],0,"Workspaces are shared team areas."]]),
 tp3:T("tp3","pbi","Sharing","Security and licences",
`<h4>Row-Level Security (RLS)</h4><p>RLS shows each user only their rows. In Desktop: <b>Modeling → Manage roles</b>, add a role with a DAX filter like <code>[Region] = "North"</code>, or a dynamic one: <code>[Email] = USERPRINCIPALNAME()</code>. Test with <b>View as</b>. Assign people to roles in the Service.</p>
<h4>Licence types (concepts)</h4><ul><li><b>Free</b>: build in Desktop, personal use</li><li><b>Pro</b>: publish and share with other Pro users</li><li><b>Premium Per User</b>: Pro plus advanced features</li><li><b>Capacity</b> (Premium/Fabric): the organisation buys capacity so many viewers can consume content</li></ul>
<p>Licence names and prices change over time. Check with your IT team for what your company uses.</p>`,
[["RLS","Row-level security"],["USERPRINCIPALNAME()","Signed-in user's email"]],
[["Each regional manager should see only their region.",["Row-Level Security","Bookmarks","Conditional formatting"],0,"RLS filters rows per user."],
 ["How do you test a role in Desktop?",["View as","Publish","Refresh"],0,"View as previews a role."],
 ["Dynamic RLS usually compares a column with…",["TODAY()","USERPRINCIPALNAME()","RAND()"],1,"It returns the signed-in user."]]),
 tp4:T("tp4","pbi","Sharing","Best practices and performance",
`<ul><li><b>Star schema</b>: one fact table, clean dimension tables, one-to-many relationships</li><li>Prefer <b>measures</b> over calculated columns for numbers you aggregate</li><li><b>Remove</b> unused columns and rows in Power Query; fewer columns = smaller, faster model</li><li>Use a proper <b>date table</b></li><li>Avoid <b>bi-directional</b> relationships unless really needed</li><li><b>Name things clearly</b> (“Total Sales”, not “Measure 3”) and put measures in a dedicated table</li><li>Limit visuals per page (around 8) and use <b>Performance Analyzer</b> (Optimize tab) to find slow ones</li><li>Design for the reader: most important number top-left, consistent colours, clear titles</li></ul>`,
[["Performance Analyzer","Times each visual"],["Star schema","Fact + dimensions"]],
[["A report is slow. Which tool shows which visual is slow?",["Performance Analyzer","Bookmarks","Gateway"],0,"It times each visual's query."],
 ["Which helps model size most?",["Removing unused columns","Adding more pages","Using pie charts"],0,"Fewer columns compress better."],
 ["For totals that respond to slicers, prefer…",["Measures","Calculated columns","Hard-coded numbers"],0,"Measures recalculate with filters."]])
};

const PBI_QUIZ = Object.fromEntries(PBI_LESSONS.filter(l => l.kind === "quiz").map(l => [l.id, l]));
const conv = (id, group, practice = []) => Object.assign({}, PBI_QUIZ[id], {mode:"theory", group, body:PBI_QUIZ[id].learn, terms:[], practice});
const N = PBI_THEORY_NEW;
const PBI_THEORY = [N.tp0, conv("p1","Power BI basics"), N.tp1, conv("p2","Power Query"), conv("p16","Power Query"), conv("p4","Data model"), conv("p18","Data model"),
  conv("p5","DAX concepts",["p6","p38"]), conv("p19","DAX concepts",["p11","p12"]), conv("p3","Visuals"), conv("p17","Visuals"), N.tp2, N.tp3, N.tp4];
const THEORY = [...EXCEL_THEORY, ...SQL_THEORY, ...PBI_THEORY];
/* ---------- Advanced practice lessons ---------- */
const XA = (id, group, title, learn, syntax, ex, exNote, task, answer, must, hint, extra = {}) => X(id, group, title, learn, syntax, ex, exNote, task, answer, must, hint, Object.assign({level:"advanced"}, extra));
const QA = (id, group, title, learn, syntax, ex, exNote, task, answer, hint, extra = {}) => Q(id, group, title, learn, syntax, ex, exNote, task, answer, hint, Object.assign({level:"advanced"}, extra));
const PA = (id, group, title, learn, syntax, ex, exNote, task, answer, must, hint, extra = {}) => PM(id, group, title, learn, syntax, ex, exNote, task, answer, must, hint, Object.assign({level:"advanced"}, extra));

const EXCEL_ADV = [
XA("ea1","Expert lookups","Two-way lookup",
`<p>Combine one INDEX with <b>two</b> MATCHes: one finds the row, the other finds the column. This answers “value at row X and column Y” and keeps working even if columns move.</p>`,
'=INDEX(table, MATCH(row_item, first_col, 0), MATCH(col_item, header_row, 0))','=INDEX(A1:F11,MATCH("Rohan",A1:A11,0),MATCH("City",A1:F1,0))','"Delhi".',
"Find Meera's Target using a two-way INDEX/MATCH on A1:F11.",'=INDEX(A1:F11,MATCH("Meera",A1:A11,0),MATCH("Target",A1:F1,0))',["INDEX","MATCH"],'Row MATCH on A1:A11, column MATCH on A1:F1.'),
XA("ea2","Expert lookups","Approximate match for slabs",
`<p>For grade bands, tax slabs or commission tiers, use VLOOKUP with <b>TRUE</b>. It finds the largest value that is <b>less than or equal</b> to what you look up. The slab table (K7:L10) must be sorted ascending.</p>`,
'=VLOOKUP(value, slab_table, 2, TRUE)','=VLOOKUP(D5,K8:L10,2,TRUE)','Sneha 61000 → "A".',
"What grade does Aarav (D2) get from the slab table K8:L10?",'=VLOOKUP(D2,K8:L10,2,TRUE)',["VLOOKUP"],"Last argument TRUE for approximate match."),
XA("ea3","Expert lookups","Wildcard lookup",
`<p>Lookups accept wildcards with exact match: <code>"Ar*"</code> finds the first value starting with Ar.</p>`,
'=VLOOKUP("text*", table, col, FALSE)','=VLOOKUP("*ya",A2:C11,3,FALSE)','First name ending in "ya": Priya → "Pune".',
'Sales (column 4) of the first person whose name starts with "Ar" (search A2:D11).','=VLOOKUP("Ar*",A2:D11,4,FALSE)',["VLOOKUP"],'Lookup value "Ar*".'),
XA("ea4","Expert lookups","Safe lookup with a default",
`<p>Wrap INDEX/MATCH in IFERROR so missing items return a clean default (0 or "Not found") instead of #N/A breaking your totals.</p>`,
'=IFERROR(INDEX(…, MATCH(…)), 0)','=IFERROR(INDEX(C2:C11,MATCH("Zoya",A2:A11,0)),"Not found")','"Not found".',
'Return Zoya\'s sales from D2:D11, or 0 if she is not in the list.','=IFERROR(INDEX(D2:D11,MATCH("Zoya",A2:A11,0)),0)',["IFERROR","INDEX","MATCH"],"IFERROR(INDEX(D2:D11, MATCH(\"Zoya\", A2:A11, 0)), 0)"),
XA("ea5","Conditional logic","Criteria from a cell",
`<p>Build a rule from a cell by joining an operator with it: <code>"&gt;"&amp;E2</code>. Change E2 and the result updates. This is how you build interactive reports.</p>`,
'=COUNTIF(range, ">"&cell)','=SUMIF(D2:D11,">="&E5,D2:D11)','Sales at or above Sneha\'s target.',
'How many people sold more than Aarav\'s target (E2)? Use ">"&E2.','=COUNTIF(D2:D11,">"&E2)',["COUNTIF"],'Rule: ">"&E2'),
XA("ea6","Conditional logic","OR logic in counts",
`<p>COUNTIFS works as AND. For OR (Pune <b>or</b> Delhi), add two COUNTIFs together.</p>`,
'=COUNTIF(r,"a") + COUNTIF(r,"b")','=SUMIF(B2:B11,"HR",D2:D11)+SUMIF(B2:B11,"IT",D2:D11)','HR or IT sales.',
'How many employees are in Pune OR Delhi?','=COUNTIF(C2:C11,"Pune")+COUNTIF(C2:C11,"Delhi")',["COUNTIF"],"Two COUNTIFs added."),
XA("ea7","Conditional logic","Conditions on dates",
`<p>Dates are numbers, so compare them like numbers. Build the date with DATE and join it to the operator.</p>`,
'=SUMIFS(sum_range, date_range, ">="&DATE(y,m,d))','=COUNTIFS(F2:F11,"<"&DATE(2020,1,1))','Joined before 2020: 3.',
'Total sales of people who joined on or after 1 Jan 2021.','=SUMIFS(D2:D11,F2:F11,">="&DATE(2021,1,1))',["SUMIFS","DATE"],'">="&DATE(2021,1,1)'),
XA("ea8","Conditional logic","Text contains",
`<p><code>SEARCH</code> returns a position if text is found, or an error if not. <code>ISNUMBER(SEARCH(…))</code> turns that into TRUE/FALSE: a “contains” test.</p>`,
'=IF(ISNUMBER(SEARCH("text", cell)), "Yes", "No")','=ISNUMBER(SEARCH("an",A8))','"Karan" contains "an": TRUE.',
'Does Meera\'s name (A9) contain "ee"? Show "Yes" or "No".','=IF(ISNUMBER(SEARCH("ee",A9)),"Yes","No")',["ISNUMBER","SEARCH"],"Wrap SEARCH in ISNUMBER inside IF."),
XA("ea9","Calculations","Weighted average",
`<p>A simple AVERAGE treats every row equally. A <b>weighted</b> average gives more weight to bigger quantities: SUMPRODUCT(values, weights) ÷ SUM(weights).</p>`,
'=SUMPRODUCT(values, weights) / SUM(weights)','=AVERAGE(L2:L5)','Simple average price: 175.',
'Average price weighted by quantity (prices L2:L5, qty M2:M5), rounded to 2 decimals.','=ROUND(SUMPRODUCT(L2:L5,M2:M5)/SUM(M2:M5),2)',["SUMPRODUCT"],"SUMPRODUCT ÷ SUM of quantities."),
XA("ea10","Calculations","Dynamic ranges with OFFSET",
`<p><code>OFFSET(start, rows, cols, height, width)</code> builds a range on the fly. <code>SUM(OFFSET(D2,0,0,5,1))</code> adds the first 5 cells from D2 down. Used for “last N” totals and dynamic charts.</p>`,
'=SUM(OFFSET(start, 0, 0, n, 1))','=AVERAGE(OFFSET(D2,5,0,5,1))','Average of the last 5 sales.',
'Add the first 5 sales figures starting at D2 using OFFSET.','=SUM(OFFSET(D2,0,0,5,1))',["OFFSET"],"OFFSET(D2,0,0,5,1) inside SUM."),
XA("ea11","Calculations","Split full names",
`<p>FIND locates the space, then LEFT/MID cut around it. Same idea splits emails, codes and addresses.</p>`,
'=MID(text, FIND(" ", text)+1, 50)','=LEFT("Priya Sharma",FIND(" ","Priya Sharma")-1)','"Priya".',
'Get the surname from the text "Priya Sharma".','=MID("Priya Sharma",FIND(" ","Priya Sharma")+1,50)',["MID","FIND"],'Start one character after the space.'),
XA("ea12","Dates & finance","Indian financial-year quarter",
`<p>India's financial year starts in April, so April–June is Q1 and January–March is Q4. CHOOSE with MONTH maps each month to its quarter.</p>`,
'=CHOOSE(MONTH(date), 4,4,4,1,1,1,2,2,2,3,3,3)','=ROUNDUP(MONTH(F3)/3,0)','Calendar quarter of July: 3.',
'Which Indian FY quarter did Aarav (F2) join in?','=CHOOSE(MONTH(F2),4,4,4,1,1,1,2,2,2,3,3,3)',["CHOOSE","MONTH"],"Jan–Mar map to 4."),
XA("ea13","Dates & finance","Years of service",
`<p>DATEDIF with a fixed “as of” date gives completed years. Use TODAY() in real files so it updates daily.</p>`,
'=DATEDIF(join_date, as_of_date, "Y")','=DATEDIF(F3,DATE(2026,3,31),"M")','Priya\'s months of service.',
'Completed years of service for Aarav (F2) as of 31 March 2026.','=DATEDIF(F2,DATE(2026,3,31),"Y")',["DATEDIF","DATE"],'End date DATE(2026,3,31), unit "Y".'),
XA("ea14","Dates & finance","WORKDAY deadlines",
`<p><code>WORKDAY(start, n)</code> returns the date n working days later, skipping weekends. Add a holiday range as a third input.</p>`,
'=WORKDAY(start, days)','=WORKDAY(F3,5)','5 working days after Priya joined.',
'Date 10 working days after Aarav joined (F2).','=WORKDAY(F2,10)',["WORKDAY"],"WORKDAY(F2,10).",{dateResult:true}),
XA("ea15","Dates & finance","Loan tenure with NPER",
`<p><code>NPER(rate, payment, loan)</code> gives how many periods to repay. Payment is negative (money going out).</p>`,
'=NPER(rate/12, -payment, loan)','=NPER(0.09/12,-5000,200000)','Months to repay 2 lakh at 9% paying 5000.',
'Months to repay 5,00,000 at 10% a year paying 10,000 per month. Round up to whole months.','=ROUNDUP(NPER(0.1/12,-10000,500000),0)',["NPER","ROUNDUP"],"ROUNDUP(NPER(0.1/12,-10000,500000),0)"),
XA("ea16","Dates & finance","Readable labels with dates",
`<p>Joining a date directly shows its serial number. Wrap it in TEXT to keep the date format.</p>`,
'=A2&" … "&TEXT(date, "dd/mm/yyyy")','=A2&" joined on "&F2','Shows a number like 44270.',
'Build the text "Aarav joined on 15/03/2021" from A2 and F2.','=A2&" joined on "&TEXT(F2,"dd/mm/yyyy")',["TEXT"],'A2&" joined on "&TEXT(F2,"dd/mm/yyyy")')
,
XA("ea17","Ageing & SLA","Ageing in business days",
`<p>Ageing reports count how long an item has been open. <code>NETWORKDAYS(start, end)</code> counts Mon–Fri including both days, so subtract 1 for completed business days. At work you'd use <code>NOW()</code> or <code>TODAY()</code> as the end; here we use a fixed date so the answer can be checked.</p>`,
'=NETWORKDAYS(created, TODAY())-1','=INT(DATE(2026,3,31)-F2)','Calendar days (24×7) since Aarav joined.',
"Business days from Aarav's join date (F2) to 31 March 2026, excluding the start day.",'=NETWORKDAYS(F2,DATE(2026,3,31))-1',["NETWORKDAYS"],"NETWORKDAYS(F2, DATE(2026,3,31)) then −1."),
XA("ea18","Ageing & SLA","Age bucket",
`<p>Buckets group ages into ranges for dashboards. IFS checks from the smallest range up; <code>TRUE</code> catches the rest.</p>`,
'=IFS(age<=a,"…",age<=b,"…",TRUE,"…")','=IFS(D3<=30000,"Low",D3<=50000,"Mid",TRUE,"High")','Priya → "Mid".',
`Bucket Ananya's days of service to 31/03/2026 (DATE(2026,3,31)-F7): "0-365", "366-1095", otherwise "1095+".`,'=IFS(DATE(2026,3,31)-F7<=365,"0-365",DATE(2026,3,31)-F7<=1095,"366-1095",TRUE,"1095+")',["IFS"],"Use DATE(2026,3,31)-F7 in each test."),
XA("ea19","Ageing & SLA","Blank-safe logic",
`<p>Open tickets have no resolved date. Check for blank <b>first</b> (<code>cell=""</code>), otherwise the rest of the formula calculates with an empty cell and gives wrong results.</p>`,
'=IF(cell="","No value",…)','=IF(H2="","No Manager",H2)','H2 has a value: "Sales".',
'H6 is empty. Show "No Manager" if H6 is blank, otherwise show H6.','=IF(H6="","No Manager",H6)',["IF"],'IF(H6="", "No Manager", H6)'),
XA("ea20","Ageing & SLA","SLA due date by category",
`<p>Different categories often have different SLAs. SWITCH converts the category to days, and WORKDAY adds those working days.</p>`,
'=WORKDAY(created, SWITCH(category, "A",1, "B",2, 3))','=SWITCH(B3,"Sales",2,"HR",5,3)','HR → 5 days.',
'Due date for Aarav (F2): Sales = 2 days, HR = 5 days, anything else 3 (use B2).','=WORKDAY(F2,SWITCH(B2,"Sales",2,"HR",5,3))',["WORKDAY","SWITCH"],"Put SWITCH(B2, …) inside WORKDAY(F2, …).",{dateResult:true}),
XA("ea21","Ageing & SLA","Week number for weekly reports",
`<p><code>ISOWEEKNUM</code> gives the international week number (weeks start Monday). <code>WEEKNUM</code> counts from 1 January.</p>`,
'=ISOWEEKNUM(date)','=WEEKNUM(F2)','Week 12 by WEEKNUM.',
"ISO week number of Aarav's join date (F2).",'=ISOWEEKNUM(F2)',["ISOWEEKNUM"],"ISOWEEKNUM(F2)")
];

const SQL_ADV = [
QA("sa1","Interview patterns","Second highest salary",
`<p>A classic interview question. Find the maximum salary that is <b>below</b> the overall maximum.</p>`,
"SELECT MAX(col) FROM t WHERE col < (SELECT MAX(col) FROM t)","SELECT MIN(salary) FROM employees WHERE salary > (SELECT MIN(salary) FROM employees);","Second lowest.",
"Return the second highest salary.","SELECT MAX(salary) FROM employees WHERE salary < (SELECT MAX(salary) FROM employees);","Use a subquery for the top salary."),
QA("sa2","Interview patterns","Top earner per department",
`<p><b>Top-N per group</b>: number rows inside each group with ROW_NUMBER, then keep rn = 1 (or ≤ N).</p>`,
"WITH r AS (SELECT …, ROW_NUMBER() OVER (PARTITION BY g ORDER BY x DESC) AS rn FROM t) SELECT … FROM r WHERE rn = 1","WITH r AS (SELECT name, city, salary, ROW_NUMBER() OVER (PARTITION BY city ORDER BY salary DESC) AS rn FROM employees) SELECT city, name FROM r WHERE rn = 1;","Top earner per city.",
"Show department, name and salary of the highest paid person in each department.","WITH r AS (SELECT name, department, salary, ROW_NUMBER() OVER (PARTITION BY department ORDER BY salary DESC) AS rn FROM employees) SELECT department, name, salary FROM r WHERE rn = 1;","ROW_NUMBER() OVER (PARTITION BY department ORDER BY salary DESC)."),
QA("sa3","Interview patterns","Groups with many rows",
`<p>Finding duplicates or frequent items = GROUP BY + HAVING COUNT(*) &gt; 1.</p>`,
"GROUP BY col HAVING COUNT(*) > n","SELECT city, COUNT(*) FROM employees GROUP BY city HAVING COUNT(*) > 3;","Mumbai.",
"Show employee_id and number of sales for employees with more than 2 sales.","SELECT employee_id, COUNT(*) FROM sales GROUP BY employee_id HAVING COUNT(*) > 2;","GROUP BY employee_id HAVING COUNT(*) > 2."),
QA("sa4","Interview patterns","Mentor counts (self join + group)",
`<p>Combine a self join with GROUP BY to count relationships inside one table.</p>`,
"FROM t e JOIN t m ON e.mentor_id = m.id GROUP BY m.name","SELECT m.name, e.name FROM employees e JOIN employees m ON e.mentor_id = m.id;","Pairs.",
"Show each mentor's name and how many people they mentor.","SELECT m.name, COUNT(*) FROM employees e JOIN employees m ON e.mentor_id = m.id GROUP BY m.name;","Join employees to itself, group by m.name."),
QA("sa5","Advanced windows","Percent of total",
`<p><code>SUM(salary) OVER ()</code> with empty brackets is the grand total on every row. Divide by it for a share. Use <code>100.0</code> so SQLite does decimal division.</p>`,
"100.0 * col / SUM(col) OVER ()","SELECT name, salary, SUM(salary) OVER () AS total FROM employees;","Total on each row.",
"Show name, salary and pct = salary as % of total salary, rounded to 2 decimals.","SELECT name, salary, ROUND(100.0 * salary / SUM(salary) OVER (), 2) AS pct FROM employees;","ROUND(100.0 * salary / SUM(salary) OVER (), 2)."),
QA("sa6","Advanced windows","Moving average (window frame)",
`<p>A <b>frame</b> limits which rows the window sees. <code>ROWS BETWEEN 2 PRECEDING AND CURRENT ROW</code> = this row and the two before: a 3-row moving average.</p>`,
"AVG(x) OVER (ORDER BY d ROWS BETWEEN 2 PRECEDING AND CURRENT ROW)","SELECT sale_date, SUM(amount) OVER (ORDER BY sale_date ROWS BETWEEN 1 PRECEDING AND CURRENT ROW) FROM sales;","Sum of each sale and the previous one.",
"Show sale_date, amount and moving_avg (current + 2 previous sales), ordered by sale_date.","SELECT sale_date, amount, AVG(amount) OVER (ORDER BY sale_date ROWS BETWEEN 2 PRECEDING AND CURRENT ROW) AS moving_avg FROM sales ORDER BY sale_date;","ROWS BETWEEN 2 PRECEDING AND CURRENT ROW.",{ordered:true}),
QA("sa7","Advanced windows","Running total per person",
`<p>PARTITION BY + ORDER BY gives a running total that restarts for each group.</p>`,
"SUM(x) OVER (PARTITION BY g ORDER BY d)","SELECT employee_id, sale_date, COUNT(*) OVER (PARTITION BY employee_id ORDER BY sale_date) FROM sales;","Sale number per person.",
"Show employee_id, sale_date, amount and running (running total per employee), ordered by employee_id, sale_date.","SELECT employee_id, sale_date, amount, SUM(amount) OVER (PARTITION BY employee_id ORDER BY sale_date) AS running FROM sales ORDER BY employee_id, sale_date;","PARTITION BY employee_id ORDER BY sale_date.",{ordered:true}),
QA("sa8","Advanced windows","NTILE buckets",
`<p><code>NTILE(4)</code> splits ordered rows into 4 nearly equal buckets: quartiles. Good for “top 25%” segments.</p>`,
"NTILE(n) OVER (ORDER BY col DESC)","SELECT name, NTILE(2) OVER (ORDER BY salary DESC) FROM employees;","Top half / bottom half.",
"Show name, salary and quartile (NTILE 4 by salary, highest first).","SELECT name, salary, NTILE(4) OVER (ORDER BY salary DESC) AS quartile FROM employees;","NTILE(4) OVER (ORDER BY salary DESC) AS quartile."),
QA("sa9","Advanced windows","Gap to department max",
`<p>Window aggregates compare each row with its group without collapsing rows.</p>`,
"MAX(x) OVER (PARTITION BY g) - x","SELECT name, salary - AVG(salary) OVER () FROM employees;","Difference from overall average.",
"Show name, department, salary and gap = department's max salary minus the person's salary.","SELECT name, department, salary, MAX(salary) OVER (PARTITION BY department) - salary AS gap FROM employees;","MAX(salary) OVER (PARTITION BY department) - salary."),
QA("sa10","Advanced windows","Month-over-month change",
`<p>First summarise by month in a CTE, then use LAG on the summary.</p>`,
"WITH m AS (…GROUP BY month) SELECT month, total, total - LAG(total) OVER (ORDER BY month) FROM m","WITH m AS (SELECT strftime('%m', sale_date) AS month, COUNT(*) AS n FROM sales GROUP BY month) SELECT * FROM m;","Sales count per month.",
"Show month, total and change vs previous month (change), from the sales table.","WITH m AS (SELECT strftime('%m', sale_date) AS month, SUM(amount) AS total FROM sales GROUP BY month) SELECT month, total, total - LAG(total) OVER (ORDER BY month) AS change FROM m;","CTE by month, then LAG(total)."),
QA("sa11","Reshaping","Pivot with CASE",
`<p>SQL has no simple PIVOT in every database. Use conditional aggregation: one SUM(CASE…) per output column.</p>`,
"SUM(CASE WHEN col = 'x' THEN 1 ELSE 0 END) AS x","SELECT city, SUM(CASE WHEN department = 'HR' THEN 1 ELSE 0 END) AS hr FROM employees GROUP BY city;","HR count per city.",
"For each department, count people in mumbai, pune and delhi as three columns named mumbai, pune, delhi.","SELECT department, SUM(CASE WHEN city = 'Mumbai' THEN 1 ELSE 0 END) AS mumbai, SUM(CASE WHEN city = 'Pune' THEN 1 ELSE 0 END) AS pune, SUM(CASE WHEN city = 'Delhi' THEN 1 ELSE 0 END) AS delhi FROM employees GROUP BY department;","Three SUM(CASE …) columns."),
QA("sa12","Reshaping","String aggregation",
`<p>Join many rows into one text: <code>GROUP_CONCAT</code> (SQLite/MySQL), <code>STRING_AGG</code> (SQL Server, PostgreSQL), <code>LISTAGG</code> (Oracle).</p>`,
"GROUP_CONCAT(col, ', ')","SELECT GROUP_CONCAT(DISTINCT city) FROM employees;","Mumbai,Pune,Delhi",
"Show each department and its employee names joined with ', '.","SELECT department, GROUP_CONCAT(name, ', ') FROM employees GROUP BY department;","GROUP_CONCAT(name, ', ') with GROUP BY."),
QA("sa13","Reshaping","Date differences",
`<p>In SQLite, <code>julianday(date)</code> turns a date into a day number; subtract two to get days. (SQL Server: <code>DATEDIFF(day, a, b)</code>; PostgreSQL: <code>b - a</code>.)</p>`,
"julianday(end) - julianday(start)","SELECT name, julianday('2026-01-01') - julianday(hire_date) FROM employees;","Days employed by 2026.",
"Show name and days (whole number) from hire_date to 2026-01-01, named days.","SELECT name, CAST(julianday('2026-01-01') - julianday(hire_date) AS INTEGER) AS days FROM employees;","CAST(julianday('2026-01-01') - julianday(hire_date) AS INTEGER)."),
QA("sa14","Reshaping","Recursive CTE",
`<p>A <b>recursive CTE</b> refers to itself. It starts with an anchor row, then keeps adding rows until a condition stops it. Used for number/date series and hierarchies (org charts).</p>`,
"WITH RECURSIVE n(x) AS (SELECT 1 UNION ALL SELECT x + 1 FROM n WHERE x < 5) SELECT x FROM n;","WITH RECURSIVE d(x) AS (SELECT '2026-01-01' UNION ALL SELECT date(x, '+1 day') FROM d WHERE x < '2026-01-07') SELECT x FROM d;","A week of dates.",
"Generate the numbers 1 to 12 in a column named x.","WITH RECURSIVE n(x) AS (SELECT 1 UNION ALL SELECT x + 1 FROM n WHERE x < 12) SELECT x FROM n;","Anchor SELECT 1, then SELECT x + 1 … WHERE x < 12."),
QA("sa15","Database objects","Create a view",
`<p>A <b>view</b> saves a query under a name. Others can query it like a table, without knowing the logic behind it.</p>`,
"CREATE VIEW name AS SELECT …;","CREATE VIEW pune_staff AS SELECT * FROM employees WHERE city = 'Pune';","",
"Create a view high_earners with name and salary of people earning 50000 or more.","CREATE VIEW high_earners AS SELECT name, salary FROM employees WHERE salary >= 50000;","CREATE VIEW high_earners AS SELECT …",{check:"SELECT * FROM high_earners"}),
QA("sa16","Database objects","Transactions and ROLLBACK",
`<p>Inside <code>BEGIN … ROLLBACK</code>, changes are undone. With <code>COMMIT</code> they are saved. Always wrap risky changes in a transaction at work.</p>`,
"BEGIN; …changes…; ROLLBACK;","BEGIN; UPDATE employees SET salary = 0; COMMIT;","Saved for real (don't do this!).",
"Start a transaction, delete everything from sales, then undo it with ROLLBACK.","BEGIN; DELETE FROM sales; ROLLBACK;","BEGIN; DELETE FROM sales; ROLLBACK;",{check:"SELECT COUNT(*) FROM sales", mustKw:["BEGIN","DELETE","ROLLBACK"]})
];

const PBI_ADV = [
PA("pa1","Advanced patterns","Running total",
`<p>A running total keeps every month up to the current one. FILTER over <code>ALL</code> of the month column keeps months ≤ the latest visible month (<code>MAX</code>).</p><p>In a real report you would use a Date table: <code>FILTER(ALL('Date'), 'Date'[Date] &lt;= MAX('Date'[Date]))</code>. This lab uses Sales[MonthNum].</p>`,
"CALCULATE([m], FILTER(ALL(col), col <= MAX(col)))","CALCULATE([Orders], FILTER(ALL(Sales[MonthNum]), Sales[MonthNum] <= MAX(Sales[MonthNum])))","Running order count.",
"Create a running total of [Total Sales] over MonthNum.","Running Sales = CALCULATE([Total Sales], FILTER(ALL(Sales[MonthNum]), Sales[MonthNum] <= MAX(Sales[MonthNum])))",["CALCULATE","FILTER","ALL"],"FILTER(ALL(Sales[MonthNum]), Sales[MonthNum] <= MAX(Sales[MonthNum]))",{by:"Month"}),
PA("pa2","Advanced patterns","Average per month",
`<p>Iterate over <code>VALUES(Sales[Month])</code> and evaluate a measure for each month (context transition), then average: “average monthly sales”.</p>`,
"AVERAGEX(VALUES(col), [measure])","MAXX(VALUES(Sales[Month]), [Total Sales])","Best month's sales.",
"Average monthly [Total Sales] for each region.","Avg Monthly Sales = AVERAGEX(VALUES(Sales[Month]), [Total Sales])",["AVERAGEX","VALUES"],"AVERAGEX(VALUES(Sales[Month]), [Total Sales])"),
PA("pa3","Advanced patterns","Count items meeting a condition",
`<p>Filter a list of items by a measure, then count them: “how many products sold over 50,000?”</p>`,
"COUNTROWS(FILTER(VALUES(col), [measure] > n))","COUNTROWS(FILTER(VALUES(Sales[Product]), [Orders] >= 3))","Products with 3+ orders.",
"Count products whose [Total Sales] is above 50000.","Big Products = COUNTROWS(FILTER(VALUES(Sales[Product]), [Total Sales] > 50000))",["COUNTROWS","FILTER","VALUES"],"FILTER(VALUES(Sales[Product]), [Total Sales] > 50000)"),
PA("pa4","Advanced patterns","Target vs actual",
`<p>Targets live in the Regions table. Iterate over the visible regions and bring each target with RELATED, then compare with actual sales.</p>`,
"DIVIDE([Total Sales], SUMX(VALUES(Sales[Region]), RELATED(Regions[Target])))","SUMX(VALUES(Sales[Region]), RELATED(Regions[Target]))","Target per region; total = 440,000.",
"Create Achievement %: [Total Sales] divided by the regions' total Target.","Achievement % = DIVIDE([Total Sales], SUMX(VALUES(Sales[Region]), RELATED(Regions[Target])))",["DIVIDE","SUMX","RELATED"],"Bottom: SUMX(VALUES(Sales[Region]), RELATED(Regions[Target]))"),
PA("pa5","Advanced patterns","Hide the total row",
`<p><code>HASONEVALUE</code> is TRUE when exactly one item is in context (a region row) and FALSE in the total. Use it to blank out totals that don't make sense.</p>`,
"IF(HASONEVALUE(col), [measure])","IF(HASONEVALUE(Sales[Region]), [Orders], \"All\")","",
"Show [Total Sales] only on single-region rows, blank in the total.","Region Sales = IF(HASONEVALUE(Sales[Region]), [Total Sales])",["HASONEVALUE"],"IF(HASONEVALUE(Sales[Region]), [Total Sales])"),
PA("pa6","Advanced patterns","Compare with the average region",
`<p><code>AVERAGEX(ALL(Sales[Region]), [Total Sales])</code> is the average of all regions, even inside one region's row. Subtract to see who is above or below average.</p>`,
"[m] - AVERAGEX(ALL(col), [m])","AVERAGEX(ALL(Sales[Region]), [Total Sales])","Average region sales in every row.",
"Create vs Avg Region: [Total Sales] minus the average of all regions' [Total Sales].","vs Avg Region = [Total Sales] - AVERAGEX(ALL(Sales[Region]), [Total Sales])",["AVERAGEX","ALL"],"[Total Sales] - AVERAGEX(ALL(Sales[Region]), [Total Sales])"),
PA("pa7","Advanced patterns","Best product's sales",
`<p>MAXX over products with a measure finds the highest product total inside each region.</p>`,
"MAXX(VALUES(col), [measure])","MINX(VALUES(Sales[Product]), [Total Sales])","Weakest product's sales.",
"For each region, the sales of its best-selling product.","Best Product Sales = MAXX(VALUES(Sales[Product]), [Total Sales])",["MAXX","VALUES"],"MAXX(VALUES(Sales[Product]), [Total Sales])"),
PA("pa8","Advanced patterns","Rate of discounted orders",
`<p>A ratio of two counts: orders with a discount ÷ all orders, by Category.</p>`,
"DIVIDE(CALCULATE([Orders], condition), [Orders])","CALCULATE([Orders], Sales[Discount] > 0)","Discounted orders.",
"Share of orders that had a discount, by category.","Discount Rate = DIVIDE(CALCULATE([Orders], Sales[Discount] > 0), [Orders])",["DIVIDE","CALCULATE"],"Top: CALCULATE([Orders], Sales[Discount] > 0)",{by:"Category"})
];
/* ---------- Advanced theory ---------- */
const TA = (...a) => Object.assign(T(...a), {level:"advanced"});

const EXCEL_THEORY_ADV = [
TA("tea1","excel","Modern Excel","Dynamic arrays and spilling",
`<p>In Excel 365/2021, one formula can return <b>many</b> results that “spill” into the cells below and to the right.</p>
<table class="mini"><tr><th>Function</th><th>Does</th><th>Example</th></tr>
<tr><td>FILTER</td><td>Rows that match</td><td><code>=FILTER(A2:D11, C2:C11="Pune")</code></td></tr>
<tr><td>SORT / SORTBY</td><td>Sorted copy</td><td><code>=SORT(A2:D11, 4, -1)</code></td></tr>
<tr><td>UNIQUE</td><td>Distinct values</td><td><code>=UNIQUE(B2:B11)</code></td></tr>
<tr><td>SEQUENCE</td><td>Number series</td><td><code>=SEQUENCE(12)</code></td></tr>
<tr><td>XLOOKUP</td><td>Can return whole rows</td><td><code>=XLOOKUP("Meera", A2:A11, B2:D11)</code></td></tr></table>
<p>Refer to a whole spilled result with <code>#</code>: if UNIQUE is in H2, <code>=COUNTA(H2#)</code> counts it. If cells below are not empty you get <code>#SPILL!</code>.</p>
<p>Old “array formulas” needed <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>Enter</kbd> and showed { } braces; dynamic arrays replace them.</p>`,
[["Spill","Results flow into neighbouring cells"],["H2#","The whole spilled range"],["#SPILL!","Blocked by non-empty cells"]],
[["=UNIQUE(B2:B11) returns…",["One value","A list of distinct values that spills","An error always"],1,"It spills all distinct values."],
 ["You see #SPILL!. Why?",["Cells in the spill area aren't empty","Wrong function name","Division by zero"],0,"Clear the cells the result needs."],
 ["How do you reference the whole spilled result starting in H2?",["H2:H","H2#","$H$2"],1,"The # operator refers to the spill range."]]),
TA("tea2","excel","Modern Excel","LET and LAMBDA",
`<p><b>LET</b> names parts of a formula, so long formulas become readable and faster (each part is calculated once):</p>
<pre class="code-block">=LET(total, SUM(D2:D11),
     target, SUM(E2:E11),
     total / target)</pre>
<p><b>LAMBDA</b> lets you create your <b>own</b> function. Define it in Name Manager, e.g. name <code>GST</code> = <code>=LAMBDA(amount, amount*1.18)</code>, then use <code>=GST(D2)</code> anywhere in the workbook.</p>
<p>Helper functions like <code>MAP</code>, <code>REDUCE</code>, <code>BYROW</code> apply a LAMBDA to every item of a range.</p>`,
[["LET","Name values inside a formula"],["LAMBDA","Build a custom function"]],
[["Main benefit of LET?",["Pretty colours","Readable formulas that calculate parts once","Protects sheets"],1,"Names make formulas clearer and avoid repeated work."],
 ["Where do you save a LAMBDA so it works like a function?",["Name Manager","Conditional formatting","Power Query"],0,"Give it a name in Name Manager."],
 ["=LET(x, 5, x*2) returns…",["5","10","x*2"],1,"x is 5, so x*2 = 10."]]),
TA("tea3","excel","Modern Excel","Power Query in Excel",
`<p>Excel has the same <b>Power Query</b> engine as Power BI: <b>Data → Get Data</b>. Use it instead of repeating manual clean-up every month.</p>
<ul><li>Import from files, folders (combine all CSVs in a folder automatically), databases, web</li><li>Clean: remove columns, change types, split, trim, fill down, unpivot</li><li>Combine: Merge (lookup/join) and Append (stack)</li><li><b>Close &amp; Load</b> to a table or to the Data Model</li><li>Next month: drop the new file in the folder and click <b>Refresh All</b></li></ul>
<p>Rule of thumb: if you do the same cleaning steps twice, record them in Power Query.</p>`,
[["Refresh All","Reruns every query"],["From Folder","Combine many files at once"]],
[["Monthly CSVs arrive in a folder. Best approach?",["Copy-paste each month","Power Query From Folder + Refresh","Retype the data"],1,"It combines all files and repeats on refresh."],
 ["Merge in Power Query is like…",["VLOOKUP / JOIN","SUM","Freeze panes"],0,"Merge matches keys and brings columns."],
 ["Steps recorded in Power Query run again when you…",["Refresh","Save","Print"],0,"Refresh replays the steps."]]),
TA("tea4","excel","Analysis tools","What-If analysis",
`<ul><li><b>Goal Seek</b> (Data → What-If): “what input gives this result?” e.g. what sales are needed for a 1,00,000 bonus pool</li><li><b>Scenario Manager</b>: save sets of inputs (Best / Expected / Worst case) and switch or summarise them</li><li><b>Data Table</b>: see a result for many input values at once, e.g. EMI for interest rates 8%–12% and tenures 3–7 years</li><li><b>Solver</b> (add-in): find the best mix of inputs under constraints (maximise profit with limited budget)</li></ul>`,
[["Goal Seek","Find one input for a target result"],["Data Table","Result grid for many inputs"],["Solver","Optimise with constraints"]],
[["Which input gives exactly 10,000 EMI?",["Goal Seek","Sort","Freeze panes"],0,"Goal Seek changes one input to hit a target."],
 ["EMIs for 5 rates × 5 tenures in one grid.",["Data Table","Scenario Manager","Flash Fill"],0,"Two-variable data table."],
 ["Maximise profit with budget limits.",["Solver","COUNTIF","Text to Columns"],0,"Solver handles constraints."]]),
TA("tea5","excel","Analysis tools","Advanced PivotTables and Power Pivot",
`<ul><li><b>Show Values As</b>: % of grand total, % of row, running total, difference from previous</li><li><b>Calculated fields</b>: new fields from others inside the pivot</li><li><b>Group</b> dates by months/quarters/years, or numbers into bands</li><li><b>Slicers and Timelines</b> filter several pivots at once (Report Connections)</li><li><b>GETPIVOTDATA</b>: pull a number from a pivot into a report cell</li><li><b>Data Model / Power Pivot</b>: load several tables, create relationships, and write <b>DAX measures</b> in Excel, the same language as Power BI. Handles millions of rows</li></ul>`,
[["Show Values As","Percentages and running totals"],["Data Model","Related tables inside Excel"],["Power Pivot","DAX in Excel"]],
[["Show each region's share of total in a pivot.",["Show Values As → % of Grand Total","Conditional formatting","SORT"],0,"Show Values As does it without formulas."],
 ["One slicer filtering three pivots uses…",["Report Connections","Goal Seek","Freeze panes"],0,"Connect the slicer to each pivot."],
 ["Which Excel feature lets you write DAX measures?",["Power Pivot / Data Model","Flash Fill","Solver"],0,"Power Pivot uses DAX like Power BI."]]),
TA("tea6","excel","Analysis tools","Building Excel dashboards",
`<ol class="steps"><li>Keep <b>raw data</b> in a Table on its own sheet</li><li>Build <b>calculations</b> (pivots or formulas) on a second sheet</li><li>Put <b>KPIs and charts</b> on a clean dashboard sheet</li><li>Add <b>slicers</b> or dropdowns (Data Validation) to make it interactive</li><li>Use consistent colours, align everything to a grid, hide gridlines</li><li>Protect the sheet so users can't break formulas</li></ol>
<p>KPI cards are just cells with big fonts linked to formulas like <code>=SUM(Table1[Sales])</code>. Use <b>sparklines</b> (Insert → Sparklines) for tiny trend charts inside cells.</p>`,
[["Sparkline","Tiny in-cell chart"],["KPI card","Big number linked to a formula"]],
[["Where should raw data live?",["On the dashboard","On its own sheet as a Table","Inside charts"],1,"Separate data, calculation and presentation."],
 ["Tiny trend lines inside cells are…",["Sparklines","Slicers","Pivots"],0,"Insert → Sparklines."],
 ["Make a dashboard interactive with…",["Slicers or dropdowns","Merged cells","Hard-coded numbers"],0,"They drive the formulas and pivots."]]),
TA("tea7","excel","Automation","Macros and VBA basics",
`<p>A <b>macro</b> records your clicks and keystrokes so you can replay them. <b>VBA</b> (Visual Basic for Applications) is the code behind it.</p>
<ul><li>Turn on the <b>Developer</b> tab (File → Options → Customize Ribbon)</li><li><b>Record Macro</b>, do the steps, Stop. Run it again with one click or a shortcut</li><li>Edit the code with <kbd>Alt</kbd>+<kbd>F11</kbd></li><li>Save as <b>.xlsm</b> (macro-enabled). .xlsx files can't keep macros</li><li>Only enable macros from trusted files: they can run any code</li></ul>
<pre class="code-block">Sub HighlightLowSales()
  Dim c As Range
  For Each c In Range("D2:D11")
    If c.Value &lt; 40000 Then c.Interior.Color = vbYellow
  Next c
End Sub</pre>
<p>Newer options: <b>Office Scripts</b> (Excel on the web) and <b>Power Automate</b> for cloud flows.</p>`,
[["Macro","Recorded steps"],["VBA","Macro code language"],[".xlsm","Macro-enabled workbook"]],
[["Which file type keeps macros?",[".xlsx",".xlsm",".csv"],1,".xlsm is macro-enabled."],
 ["Open the VBA editor with…",["Alt+F11","Ctrl+T","F4"],0,"Alt+F11 opens the VBA editor."],
 ["A file from an unknown sender asks to enable macros. You…",["Enable immediately","Only enable if you trust the source","Rename it"],1,"Macros can run harmful code."]]),
TA("tea8","excel","Automation","Protection and collaboration",
`<ul><li><b>Protect Sheet</b>: lock formulas, leave input cells unlocked (Format Cells → Protection → uncheck Locked first)</li><li><b>Protect Workbook</b>: stop sheets being added, deleted or renamed</li><li><b>Encrypt with password</b>: File → Info → Protect Workbook</li><li><b>Co-authoring</b>: save to OneDrive/SharePoint and edit together live</li><li><b>Version History</b>: restore an older copy</li><li><b>Comments / @mentions</b> for review, <b>Track changes</b> via Show Changes</li></ul>`,
[["Locked cell","Protected when sheet protection is on"],["Co-authoring","Many people editing at once"]],
[["Let users type in inputs but not edit formulas.",["Unlock input cells, then Protect Sheet","Hide the sheet","Merge cells"],0,"All cells are locked by default; unlock inputs first."],
 ["Recover yesterday's version of a OneDrive file.",["Version History","Undo","Goal Seek"],0,"Version History keeps earlier copies."],
 ["Several people edit the same file at once using…",["Co-authoring on OneDrive/SharePoint","Email attachments","Macros"],0,"Cloud storage enables live co-authoring."]])
];

const SQL_THEORY_ADV = [
TA("tsa1","sql","Advanced querying","Window functions in depth",
`<p>A window function: <code>FUNCTION() OVER (PARTITION BY … ORDER BY … frame)</code></p>
<ul><li><b>PARTITION BY</b>: restart for each group</li><li><b>ORDER BY</b>: the order inside the window (needed for running totals, LAG, ranking)</li><li><b>Frame</b>: which rows count, e.g. <code>ROWS BETWEEN 2 PRECEDING AND CURRENT ROW</code></li></ul>
<table class="mini"><tr><th>Function</th><th>Use</th></tr><tr><td>ROW_NUMBER</td><td>Unique 1,2,3 (top-N per group)</td></tr><tr><td>RANK / DENSE_RANK</td><td>Ranking with ties</td></tr><tr><td>NTILE(n)</td><td>Buckets</td></tr><tr><td>LAG / LEAD</td><td>Previous / next row</td></tr><tr><td>FIRST_VALUE / LAST_VALUE</td><td>First/last in window</td></tr><tr><td>SUM/AVG … OVER</td><td>Running totals, moving averages, % of total</td></tr></table>
<p><b>ROWS vs RANGE</b>: ROWS counts physical rows; RANGE groups rows with the same ORDER BY value. With ORDER BY and no frame, the default is RANGE from the start to the current row, which is why ties get the same running total.</p>`,
[["PARTITION BY","Group for the window"],["Frame","Rows the window sees"]],
[["Top 3 salaries per department needs…",["ROW_NUMBER with PARTITION BY department","GROUP BY only","UNION"],0,"Number within each department, keep ≤ 3."],
 ["Previous month's value on the same row?",["LAG","LEAD","NTILE"],0,"LAG looks back."],
 ["3-row moving average frame:",["ROWS BETWEEN 2 PRECEDING AND CURRENT ROW","PARTITION BY 3","LIMIT 3"],0,"Current row plus two before."]],["sa5","sa6","sa7"]),
TA("tsa2","sql","Advanced querying","Recursive CTEs and hierarchies",
`<p>A recursive CTE has two parts joined by UNION ALL: an <b>anchor</b> (starting rows) and a <b>recursive</b> part that references the CTE itself. It repeats until no new rows appear.</p>
<pre class="code-block">WITH RECURSIVE chain AS (
  SELECT id, name, mentor_id, 1 AS level
  FROM employees WHERE mentor_id IS NULL      -- anchor: top people
  UNION ALL
  SELECT e.id, e.name, e.mentor_id, c.level + 1
  FROM employees e JOIN chain c ON e.mentor_id = c.id   -- recursive step
)
SELECT * FROM chain;</pre>
<p>Uses: org charts, folder trees, bill of materials, generating date or number series. Always make sure the recursion stops (a WHERE condition).</p>`,
[["Anchor","Starting rows"],["Recursive member","Joins back to the CTE"]],
[["Two parts of a recursive CTE are joined by…",["UNION ALL","JOIN","HAVING"],0,"Anchor UNION ALL recursive step."],
 ["Good use case?",["Org chart levels","Simple COUNT","Renaming a column"],0,"Hierarchies are the classic case."],
 ["Without a stop condition the query…",["May loop forever","Returns nothing","Is faster"],0,"Recursion needs a limit."]],["sa14"]),
TA("tsa3","sql","Performance","Query performance and indexes",
`<ul><li><b>Indexes</b> speed up WHERE, JOIN and ORDER BY on the indexed columns. A <b>composite index</b> (city, salary) helps filters on city, or city + salary, but not salary alone</li><li>Read the plan: <code>EXPLAIN</code> / <code>EXPLAIN QUERY PLAN</code> (SQLite), “Display Execution Plan” (SQL Server). Look for full <b>table scans</b> on big tables</li><li><b>Sargable</b> filters can use indexes: <code>WHERE hire_date &gt;= '2026-01-01'</code> is good; <code>WHERE YEAR(hire_date) = 2026</code> wraps the column in a function and usually can't</li><li>Select only the columns you need (avoid <code>SELECT *</code> in production)</li><li>Filter early; avoid functions on join keys; prefer EXISTS over IN for large correlated lists</li><li>Too many indexes slow down INSERT/UPDATE</li></ul>`,
[["EXPLAIN","Shows how a query will run"],["Sargable","Filter that can use an index"],["Composite index","Index on several columns"]],
[["Which filter can use an index on hire_date?",["WHERE YEAR(hire_date) = 2026","WHERE hire_date >= '2026-01-01'","WHERE hire_date LIKE '%26%'"],1,"Don't wrap the indexed column in a function."],
 ["Index on (city, salary) helps a filter on…",["salary only","city (and city + salary)","name"],1,"The leftmost column must be used."],
 ["Downside of many indexes?",["Slower inserts and updates","Slower SELECTs","Data loss"],0,"Every write must update each index."]]),
TA("tsa4","sql","Performance","Transactions, isolation and locking",
`<p>When many users work at once, problems can occur:</p>
<ul><li><b>Dirty read</b>: reading data another transaction hasn't committed</li><li><b>Non-repeatable read</b>: the same row changes between two reads</li><li><b>Phantom read</b>: new rows appear in a repeated query</li></ul>
<table class="mini"><tr><th>Isolation level</th><th>Prevents</th></tr><tr><td>READ UNCOMMITTED</td><td>Nothing</td></tr><tr><td>READ COMMITTED (common default)</td><td>Dirty reads</td></tr><tr><td>REPEATABLE READ</td><td>+ non-repeatable reads</td></tr><tr><td>SERIALIZABLE</td><td>+ phantoms (safest, slowest)</td></tr></table>
<p>Higher isolation = more locking = less concurrency. A <b>deadlock</b> is when two transactions wait for each other; the database cancels one.</p>`,
[["Isolation level","How much transactions see of each other"],["Deadlock","Two transactions waiting on each other"]],
[["Reading uncommitted data from another transaction is a…",["Dirty read","Phantom read","Deadlock"],0,"Dirty reads see uncommitted changes."],
 ["Safest isolation level?",["SERIALIZABLE","READ UNCOMMITTED","READ COMMITTED"],0,"It prevents all three problems."],
 ["Two transactions each wait for the other's lock. This is…",["A deadlock","A view","Normalisation"],0,"The DB resolves it by cancelling one."]],["sa16"]),
TA("tsa5","sql","Beyond queries","Stored procedures, functions and triggers",
`<ul><li><b>Stored procedure</b>: saved SQL program with parameters, run with <code>EXEC</code>/<code>CALL</code>. Used for repeatable jobs (monthly close, data loads)</li><li><b>User-defined function</b>: returns a value or table; usable inside SELECT</li><li><b>Trigger</b>: code that runs automatically on INSERT/UPDATE/DELETE, e.g. write to an audit log</li></ul>
<pre class="code-block">-- SQL Server style
CREATE PROCEDURE GetDeptStaff @dept VARCHAR(20) AS
BEGIN
  SELECT name, salary FROM employees WHERE department = @dept;
END;
EXEC GetDeptStaff @dept = 'IT';</pre>
<p>Syntax varies a lot between databases; the ideas are the same.</p>`,
[["Procedure","Saved program, called explicitly"],["Trigger","Runs automatically on a change"]],
[["Automatically log every salary change.",["Trigger","View","Index"],0,"Triggers fire on data changes."],
 ["Reusable logic you call with parameters.",["Stored procedure","CTE","Alias"],0,"Procedures take parameters."],
 ["Which can be used inside a SELECT?",["Function","Procedure","Trigger"],0,"Functions return values usable in queries."]]),
TA("tsa6","sql","Beyond queries","Data warehousing basics",
`<table class="mini"><tr><th></th><th>OLTP</th><th>OLAP</th></tr><tr><td>Purpose</td><td>Run the business (orders, payments)</td><td>Analyse the business</td></tr><tr><td>Design</td><td>Normalised, many small tables</td><td>Star schema, denormalised</td></tr><tr><td>Queries</td><td>Many small reads/writes</td><td>Big aggregations</td></tr></table>
<ul><li><b>Fact tables</b> (sales, tickets) + <b>dimension tables</b> (date, product, customer) = star schema. Normalised dimensions = snowflake</li><li><b>ETL / ELT</b>: Extract, Transform, Load data from source systems into the warehouse</li><li><b>Slowly Changing Dimensions</b>: Type 1 overwrite; Type 2 keep history with valid_from/valid_to rows</li><li>Modern platforms: Snowflake, BigQuery, Azure Synapse, Microsoft Fabric, Databricks</li></ul>`,
[["OLTP","Transactional systems"],["OLAP","Analytical systems"],["SCD Type 2","Keeps history rows"]],
[["Reporting warehouse usually uses…",["Star schema","Only one table","Spreadsheets"],0,"Facts plus dimensions."],
 ["Keep a customer's old city when it changes.",["SCD Type 2","SCD Type 1","DELETE"],0,"Type 2 adds a new row with dates."],
 ["Order-entry system is…",["OLTP","OLAP","ETL"],0,"It runs day-to-day transactions."]]),
TA("tsa7","sql","Beyond queries","SQL security",
`<ul><li><b>SQL injection</b>: an attacker types SQL into an input box, e.g. <code>' OR '1'='1</code>, and changes your query. Prevent it with <b>parameterised queries</b> (prepared statements); never build SQL by joining user text</li><li><b>Least privilege</b>: give each user/app only the permissions it needs (GRANT SELECT on reporting views, not DELETE)</li><li>Use <b>views</b> to expose only safe columns (no salaries or personal data)</li><li><b>Row-level security</b> and <b>masking</b> for sensitive data</li><li>Audit logs and backups</li></ul>`,
[["SQL injection","Attack via crafted input"],["Parameterised query","Safe way to pass user values"],["Least privilege","Minimum permissions"]],
[["Best defence against SQL injection?",["Parameterised queries","Longer table names","More indexes"],0,"Parameters keep data separate from code."],
 ["A report user should only read data. Grant…",["SELECT","DELETE","DROP"],0,"Least privilege."],
 ["Hide the salary column from analysts.",["Give them a view without it","Rename the column","Use LIMIT"],0,"Views expose only chosen columns."]]),
TA("tsa8","sql","Beyond queries","Interview problem patterns",
`<table class="mini"><tr><th>Question</th><th>Pattern</th></tr>
<tr><td>Nth highest salary</td><td>Subquery with MAX &lt;, or DENSE_RANK() = N</td></tr>
<tr><td>Top N per group</td><td>ROW_NUMBER() OVER (PARTITION BY … ORDER BY …) ≤ N</td></tr>
<tr><td>Find duplicates</td><td>GROUP BY … HAVING COUNT(*) &gt; 1</td></tr>
<tr><td>Delete duplicates keep one</td><td>ROW_NUMBER over the duplicate columns, delete rn &gt; 1</td></tr>
<tr><td>Rows with no match</td><td>LEFT JOIN … WHERE right.id IS NULL, or NOT EXISTS</td></tr>
<tr><td>Running total / MoM change</td><td>SUM() OVER (ORDER BY …), LAG()</td></tr>
<tr><td>Pivot</td><td>SUM(CASE WHEN … THEN … END)</td></tr>
<tr><td>Employees earning more than their manager</td><td>Self join</td></tr>
<tr><td>Consecutive days (gaps &amp; islands)</td><td>date − ROW_NUMBER() groups consecutive rows</td></tr></table>
<p>In interviews, talk through: what's the grain (one row = ?), which join, which filter goes in WHERE vs HAVING, and how NULLs behave.</p>`,
[["Grain","What one row represents"],["Gaps & islands","Finding consecutive runs"]],
[["Pattern for 'top 2 products per region'?",["ROW_NUMBER with PARTITION BY region","GROUP BY product","UNION"],0,"Rank within each region."],
 ["Pattern for 'customers who never ordered'?",["LEFT JOIN … IS NULL","INNER JOIN","CROSS JOIN"],0,"Anti-join."],
 ["Pattern for 'month-over-month change'?",["LAG","NTILE","DISTINCT"],0,"Compare with the previous row."]],["sa1","sa2","sa3"])
];

const PBI_THEORY_ADV = [
TA("tpa1","pbi","Advanced DAX","Context transition",
`<p>Context transition is the step from intermediate to expert DAX.</p>
<ul><li>In a <b>row context</b> (calculated column, SUMX), a plain <code>SUM(Sales[Amount])</code> ignores the current row and returns the grand total</li><li><b>CALCULATE</b> turns the current row into a <b>filter</b>. So <code>CALCULATE(SUM(Sales[Amount]))</code> inside a row gives that row's slice</li><li>Every <b>measure reference</b> like <code>[Total Sales]</code> has an invisible CALCULATE around it. That's why <code>AVERAGEX(VALUES(Sales[Month]), [Total Sales])</code> works per month</li></ul>
<p>Rule: inside iterators, use measures (or CALCULATE) when you want the value for each item.</p>`,
[["Context transition","Row context becomes filter context"],["Implicit CALCULATE","Wrapped around every measure reference"]],
[["In a calculated column, SUM(Sales[Amount]) returns…",["The row's amount","The grand total","Blank"],1,"SUM ignores row context."],
 ["Which function performs context transition?",["CALCULATE","SUM","FORMAT"],0,"CALCULATE converts row to filter context."],
 ["Why does AVERAGEX(VALUES(Sales[Month]), [Total Sales]) work per month?",["Measure references include an implicit CALCULATE","VALUES sums data","AVERAGEX ignores filters"],0,"Each month row becomes a filter."]],["pa2","pa3"]),
TA("tpa2","pbi","Advanced DAX","Date tables and time intelligence",
`<p>Professional models have a dedicated <b>Date</b> table: one row per day, no gaps, related to the fact table's date column, and <b>marked as date table</b>.</p>
<pre class="code-block">Date = ADDCOLUMNS(
  CALENDAR(DATE(2024,1,1), DATE(2026,12,31)),
  "Year", YEAR([Date]),
  "Month", FORMAT([Date], "MMM"),
  "MonthNo", MONTH([Date]),
  "Quarter", "Q" & QUARTER([Date]))</pre>
<table class="mini"><tr><th>Need</th><th>DAX</th></tr><tr><td>Year to date</td><td>TOTALYTD / DATESYTD (fiscal: add "03-31")</td></tr><tr><td>Month/quarter to date</td><td>DATESMTD / DATESQTD</td></tr><tr><td>Same period last year</td><td>SAMEPERIODLASTYEAR</td></tr><tr><td>Shift periods</td><td>DATEADD('Date'[Date], -1, MONTH)</td></tr><tr><td>Previous month</td><td>PREVIOUSMONTH</td></tr><tr><td>Rolling 3 months</td><td>DATESINPERIOD('Date'[Date], MAX('Date'[Date]), -3, MONTH)</td></tr></table>
<p>Indian FY YTD: <code>TOTALYTD([Total Sales], 'Date'[Date], "03-31")</code>.</p>`,
[["CALENDAR","Creates a date range table"],["Mark as date table","Enables reliable time intelligence"]],
[["Last year's sales for the same period?",["SAMEPERIODLASTYEAR","PREVIOUSMONTH","RANKX"],0,"It shifts the period back one year."],
 ["Indian financial-year YTD ends on…",["\"03-31\"","\"12-31\"","\"06-30\""],0,"Pass the year-end date as the third argument."],
 ["A proper date table must…",["Have one row per day with no gaps","Have only month names","Be hidden"],0,"Continuous dates make time intelligence work."]],["p35","p36","p37","pa1"]),
TA("tpa3","pbi","Advanced modelling","Advanced relationships",
`<ul><li><b>Role-playing dimensions</b>: Order Date and Ship Date both relate to one Date table. Only one relationship is active; activate the other in a measure with <code>USERELATIONSHIP</code>:<br><code>CALCULATE([Total Sales], USERELATIONSHIP(Sales[ShipDate], 'Date'[Date]))</code></li><li><b>Many-to-many</b>: use a <b>bridge table</b> between the two sides, or the many-to-many cardinality setting with care</li><li><b>TREATAS</b>: apply a filter from a table that has no relationship</li><li><b>Snowflake vs star</b>: flatten snowflaked dimensions into one table where you can</li><li><b>Aggregation tables</b> and <b>composite models</b> keep huge facts in DirectQuery while summary tables are imported for speed</li></ul>`,
[["USERELATIONSHIP","Use an inactive relationship"],["Bridge table","Resolves many-to-many"],["TREATAS","Virtual relationship"]],
[["Use the inactive Ship Date relationship in one measure.",["USERELATIONSHIP","ALL","RELATED"],0,"It activates that relationship inside CALCULATE."],
 ["Customers ↔ Accounts (many-to-many) is usually solved with…",["A bridge table","A slicer","A bookmark"],0,"The bridge holds the pairs."],
 ["Filter a table using values from an unrelated table.",["TREATAS","FORMAT","DIVIDE"],0,"TREATAS applies values as a filter."]]),
TA("tpa4","pbi","Advanced modelling","Calculation groups and field parameters",
`<ul><li><b>Field parameters</b> (Modeling → New parameter → Fields): a slicer that lets users switch which measure or column a visual shows (Sales / Orders / Profit)</li><li><b>Calculation groups</b> (Model view or Tabular Editor): write time logic once (YTD, PY, YoY %) and apply it to every measure, instead of creating 3 versions of every measure. Uses <code>SELECTEDMEASURE()</code></li><li><b>Dynamic format strings</b>: change number format per measure or per calculation item</li></ul>
<pre class="code-block">-- calculation item "YTD"
CALCULATE(SELECTEDMEASURE(), DATESYTD('Date'[Date]))</pre>`,
[["Field parameter","User picks the measure/axis"],["Calculation group","Reusable logic for any measure"],["SELECTEDMEASURE()","The measure being modified"]],
[["Let users pick Sales or Orders for a chart.",["Field parameter","RLS","Gateway"],0,"Field parameters switch fields."],
 ["Avoid writing YTD for 20 measures separately.",["Calculation group","Bookmarks","Merge queries"],0,"One calculation item covers all measures."],
 ["Inside a calculation item, the current measure is…",["SELECTEDMEASURE()","[Total Sales]","VALUES()"],0,"SELECTEDMEASURE refers to whichever measure is used."]]),
TA("tpa5","pbi","Advanced Power Query","M language and query folding",
`<p>Every Power Query step writes <b>M</b> code. Open <b>Advanced Editor</b> to see it:</p>
<pre class="code-block">let
  Source = Excel.Workbook(File.Contents("C:\\data\\sales.xlsx")),
  Sales  = Source{[Item="Sales"]}[Data],
  Typed  = Table.TransformColumnTypes(Sales, {{"Amount", type number}}),
  Big    = Table.SelectRows(Typed, each [Amount] > 20000)
in
  Big</pre>
<ul><li><b>Parameters</b>: e.g. a file path or date, reused across queries</li><li><b>Custom functions</b>: turn a query into a function and apply it to many files</li><li><b>Query folding</b>: Power Query translates steps into SQL that runs on the source database. Folded steps are much faster. Right-click a step → <b>View Native Query</b> to check. Some steps (custom columns with complex M, index columns) break folding; do them last</li></ul>`,
[["M","Power Query's language"],["Query folding","Steps pushed to the source as SQL"],["Parameter","Reusable value"]],
[["How do you see the M code of a query?",["Advanced Editor","Performance Analyzer","DAX Studio"],0,"Home → Advanced Editor."],
 ["Query folding means…",["Steps run in the source database","Queries are hidden","Data is compressed"],0,"The source does the work."],
 ["Check whether a step folds with…",["View Native Query","Mark as date table","Bookmarks"],0,"If it's enabled, the step folds."]]),
TA("tpa6","pbi","Advanced reports","Advanced visuals and report design",
`<ul><li><b>Drillthrough</b>: right-click a bar to jump to a detail page filtered to that item</li><li><b>Report page tooltips</b>: a mini page shown on hover</li><li><b>Decomposition tree</b>: break a number down by any dimension step by step</li><li><b>Key influencers</b>: AI visual showing what drives a result</li><li><b>Q&amp;A</b>: ask questions in plain English</li><li><b>Smart narrative</b>: automatic text summary</li><li><b>Custom visuals</b> from AppSource</li><li><b>Paginated reports</b> (Report Builder): pixel-perfect, printable reports like invoices</li><li>Design: a clear visual hierarchy, consistent colour meaning, and a theme JSON for branding</li></ul>`,
[["Drillthrough","Jump to a filtered detail page"],["Decomposition tree","Interactive breakdown"],["Paginated report","Print-ready layout"]],
[["Right-click a region to open a details page for it.",["Drillthrough","Bookmark","Slicer"],0,"Drillthrough passes the filter."],
 ["Printable invoice-style report with many pages.",["Paginated report","Card","Smart narrative"],0,"Built with Power BI Report Builder."],
 ["Find what drives low customer ratings.",["Key influencers","Pie chart","Gauge"],0,"An AI visual for drivers."]]),
TA("tpa7","pbi","Advanced reports","Deployment, refresh and governance",
`<ul><li><b>Deployment pipelines</b>: Development → Test → Production workspaces</li><li><b>Incremental refresh</b>: only refresh recent data (e.g. last 3 days) instead of years of history. Uses RangeStart/RangeEnd parameters</li><li><b>Endorsement</b>: mark semantic models as Promoted or Certified so people trust them</li><li><b>Sensitivity labels</b> (Microsoft Purview) protect confidential data</li><li><b>Shared semantic models</b>: one golden model used by many reports (thin reports)</li><li><b>Microsoft Fabric</b>: one platform with Lakehouse, Data Factory, Warehouse and Power BI; <b>Direct Lake</b> mode reads lake files directly</li><li>Version control: save as <b>.pbip</b> (project format) to use Git</li></ul>`,
[["Incremental refresh","Refresh only recent data"],["Deployment pipeline","Dev → Test → Prod"],["Certified","Trusted official dataset"]],
[["Model takes hours to refresh 5 years of data.",["Incremental refresh","More visuals","Pie charts"],0,"Refresh only new partitions."],
 ["Move a report from Test to Production safely.",["Deployment pipeline","Email the .pbix","Screenshot"],0,"Pipelines promote content between stages."],
 ["Mark the official sales model as trusted.",["Certify/endorse it","Hide it","Rename it"],0,"Endorsement signals quality."]]),
TA("tpa8","pbi","Advanced reports","Performance tuning",
`<ul><li>Power BI's engine (<b>VertiPaq</b>) compresses columns. <b>High-cardinality</b> columns (unique IDs, timestamps with seconds) compress badly: remove or split them</li><li>Prefer integers for keys; reduce decimal precision; disable Auto date/time</li><li><b>Performance Analyzer</b> shows DAX query time per visual; copy the query to <b>DAX Studio</b> for deeper analysis (server timings, VertiPaq Analyzer)</li><li>DAX tips: use variables; filter columns, not whole tables (<code>CALCULATE([m], Sales[Region]="North")</code> rather than <code>FILTER(Sales, …)</code>); avoid iterating huge tables with nested iterators</li><li><b>Tabular Editor</b> for bulk edits and best-practice checks</li></ul>`,
[["VertiPaq","In-memory columnar engine"],["Cardinality","Number of distinct values"],["DAX Studio","Query analysis tool"]],
[["Which column hurts compression most?",["Transaction ID with millions of unique values","Region with 4 values","Year"],0,"High cardinality compresses poorly."],
 ["Better filter in CALCULATE?",["Sales[Region] = \"North\"","FILTER(Sales, Sales[Region] = \"North\")","Both identical always"],0,"Column filters are lighter than table filters."],
 ["Free tool to analyse DAX query timings?",["DAX Studio","Paint","Notepad"],0,"DAX Studio shows server timings."]])
];
/* ---------- Rich theory: Excel ---------- */
/* Tutorial chapter */
const RT = (id, track, level, group, title, o) => Object.assign({id, track, mode:"theory", level, group, title, rich:true, terms:[], questions:[], practice:[]}, o);
/* Excel function reference page: examples run live on the practice sheet */
const XF = (id, level, group, name, intro, syntax, args, ex, o = {}) => RT(id, "excel", level, group, name, Object.assign({
  intro, syntax, args, ex: ex.map(([desc, code]) => ({desc, code})), sheet:true,
  steps: o.steps || [
    "Click the cell where you want the result.",
    `Type <code>=${name.split(/[ /+(]/)[0].toUpperCase()}(</code>. Excel shows the function in a list; press <kbd>Tab</kbd> to accept it.`,
    ...args.map(([a]) => `Enter <b>${a}</b>: click or drag over the cells, or type a value. ${""}`).slice(0, 4),
    "Separate each input with a comma (,). Some regional settings use a semicolon (;) instead.",
    "Type <code>)</code> and press <kbd>Enter</kbd>."
  ]}, o));

const EXCEL_RICH = [
/* ===== Getting started ===== */
RT("xg1","excel","basic","Getting started","Get started: workbooks and saving",{
 intro:`<p>When you open Excel you see the start screen. Choose <b>Blank workbook</b> to start fresh, or pick a template. Each new workbook starts with one worksheet called <b>Sheet1</b>.</p>`,
 steps:["Open Excel and click <b>Blank workbook</b>.","Click a cell and start typing. Press <kbd>Enter</kbd> to move down or <kbd>Tab</kbd> to move right.","Press <kbd>Ctrl</kbd>+<kbd>S</kbd> to save. Choose a folder (or OneDrive) and give the file a clear name, such as <i>Sales_Report_Jan2026</i>.","To add another sheet, click the <b>+</b> next to the sheet tabs at the bottom. Double-click a tab to rename it."],
 table:{cols:["File type","Use it for"],rows:[[".xlsx","Normal Excel workbook (default)"],[".xlsm","Workbook with macros"],[".csv","Plain data shared with other systems; keeps only one sheet and no formatting"],[".xlsb","Very large files (binary, faster)"],[".pdf","Sharing a read-only copy"]]},
 notes:["AutoSave is available when the file is saved to OneDrive or SharePoint.","Saving as .csv removes formulas, formatting and extra sheets. Keep an .xlsx copy."],
 mistake:{wrong:"Saving the only copy of a report as .csv",why:"CSV keeps only the values of the active sheet, so formulas and other sheets are lost.",fix:"Save as .xlsx, and export a CSV copy only when another system needs it."},
 terms:[["Workbook","The file"],["Worksheet","A tab inside it"],["AutoSave","Automatic saving to the cloud"]],
 questions:[["Which format keeps formulas and several sheets?",[".csv",".xlsx",".txt"],1,".xlsx is the normal workbook format."],["Shortcut to save?",["Ctrl+S","Ctrl+P","Ctrl+Z"],0,"Ctrl+S saves the workbook."]]}),
RT("xg2","excel","basic","Getting started","Entering and editing data",{
 intro:`<p>Click a cell and type. What you type appears in the cell and in the <b>Formula Bar</b>. Press <kbd>Enter</kbd> to confirm.</p>`,
 steps:["Click a cell, type the value, press <kbd>Enter</kbd> (moves down) or <kbd>Tab</kbd> (moves right).","To change a cell completely, click it and type over it.","To edit part of a cell, double-click it or press <kbd>F2</kbd>, then use the arrow keys.","Press <kbd>Esc</kbd> to cancel an edit before pressing Enter.","Press <kbd>Delete</kbd> to clear the content of selected cells.","Double-click the line between two column letters to <b>AutoFit</b> the column width."],
 table:{cols:["You type","Excel stores"],rows:[["52000","Number (right-aligned)"],["Mumbai","Text (left-aligned)"],["15/03/2021","Date (a number, shown as a date)"],["14:30","Time (a fraction of a day)"],["'00123","Text that keeps the leading zeros"],["=D2*2","A formula; the cell shows the result"]]},
 notes:["Use <kbd>Alt</kbd>+<kbd>Enter</kbd> to start a new line inside the same cell.","Dates are read using your computer's regional format (dd/mm/yyyy in India)."],
 mistake:{wrong:"Typing 0123456789 for a phone number",why:"Excel treats it as a number and removes the leading zero.",fix:"Format the column as Text first, or type an apostrophe: '0123456789."},
 questions:[["How do you edit part of a cell's content?",["Press F2","Press F4","Press Ctrl+T"],0,"F2 enters edit mode."],["Start a new line inside a cell:",["Alt+Enter","Shift+Enter","Ctrl+Enter"],0,"Alt+Enter adds a line break."]]}),
RT("xg3","excel","basic","Getting started","Ranges and selecting cells",{
 intro:`<p>A <b>range</b> is a group of cells. You write it as <i>top-left cell</i> <code>:</code> <i>bottom-right cell</i>. <code>A1:C3</code> is 9 cells (3 columns × 3 rows).</p>`,
 steps:["Click a cell to select it.","Drag with the mouse to select a block of cells.","Click the first cell, hold <kbd>Shift</kbd> and click the last cell to select everything between.","Hold <kbd>Ctrl</kbd> and click to add separate cells or ranges.","Click a column letter (or row number) to select the whole column (or row).","Press <kbd>Ctrl</kbd>+<kbd>A</kbd> to select the whole data block; press it again for the whole sheet.","Type a range like <code>D2:D11</code> in the <b>Name Box</b> (left of the Formula Bar) and press Enter to jump to it."],
 table:{cols:["Reference","Means"],rows:[["D2","One cell"],["D2:D11","Cells D2 down to D11"],["A1:F11","A block of 6 columns × 11 rows"],["D:D","The whole column D"],["2:2","The whole row 2"],["D2:D11,F2:F11","Two separate ranges"]]},
 ex:[{desc:"Sum a range of sales",code:"=SUM(D2:D11)"},{desc:"Count cells in a block (6 columns × 10 rows)",code:"=COUNTA(A2:F11)"}], sheet:true,
 notes:["The status bar at the bottom shows Sum, Average and Count of the selected cells: a quick check without formulas."],
 questions:[["How many cells are in B2:C5?",["8","6","10"],0,"2 columns × 4 rows = 8."],["Select non-adjacent cells by holding…",["Ctrl","Shift","Alt"],0,"Ctrl adds separate selections."]], practice:["e3"]}),
RT("xg4","excel","basic","Getting started","Fill, AutoFill and Flash Fill",{
 intro:`<p>The small square at the bottom-right of a selected cell is the <b>fill handle</b>. Drag it to copy a value or formula, or to continue a series.</p>`,
 steps:["Type <b>Jan</b> in a cell, then drag the fill handle down: Excel fills Feb, Mar, Apr…","Type <b>1</b> and <b>2</b> in two cells, select both, and drag: Excel continues 3, 4, 5…","Write a formula in the first row, then <b>double-click</b> the fill handle: Excel copies it down to the last row of your data.","After filling, click the <b>Auto Fill Options</b> icon to choose Copy Cells, Fill Series or Fill Formatting Only.","<b>Flash Fill</b>: type the result you want for the first row (e.g. the first name from a full name), then press <kbd>Ctrl</kbd>+<kbd>E</kbd>. Excel copies the pattern."],
 table:{cols:["Start with","Drag gives"],rows:[["Mon","Tue, Wed, Thu…"],["Q1","Q2, Q3, Q4, Q1…"],["01/01/2026","02/01/2026, 03/01/2026…"],["5, 10","15, 20, 25…"],["Item 1","Item 2, Item 3…"]]},
 notes:["Double-click fill stops at the first blank cell in the neighbouring column.","Formulas adjust their references as they fill. See the chapter on relative and absolute references."],
 mistake:{wrong:"Dragging one number expecting 1, 2, 3",why:"With only one number selected, Excel copies it (1, 1, 1).",fix:"Enter the first two values, or hold Ctrl while dragging, or choose Fill Series from Auto Fill Options."},
 questions:[["Copy a formula to the last data row fastest:",["Double-click the fill handle","Press F4","Ctrl+T"],0,"Double-click fills to the end of the data."],["Shortcut for Flash Fill?",["Ctrl+E","Ctrl+F","Ctrl+L"],0,"Ctrl+E runs Flash Fill."]]}),
RT("xg5","excel","basic","Getting started","Move, copy, insert and delete",{
 intro:`<p>Rearranging data is part of every report. Excel can move cells, copy them, insert space, and remove rows or columns.</p>`,
 steps:["<b>Copy</b>: select cells, <kbd>Ctrl</kbd>+<kbd>C</kbd>, click the destination, <kbd>Ctrl</kbd>+<kbd>V</kbd>.","<b>Move</b>: select cells, <kbd>Ctrl</kbd>+<kbd>X</kbd>, click the destination, <kbd>Ctrl</kbd>+<kbd>V</kbd>. Or drag the selection border.","<b>Insert</b> a row/column: right-click a row number or column letter → <b>Insert</b>. Shortcut <kbd>Ctrl</kbd>+<kbd>+</kbd>.","<b>Delete</b> a row/column: right-click → <b>Delete</b>. Shortcut <kbd>Ctrl</kbd>+<kbd>-</kbd>.","<b>Hide</b>: right-click → Hide. To unhide, select the rows/columns around it → right-click → Unhide."],
 notes:["Formulas pointing to moved cells update automatically.","Deleting a row that a formula uses gives #REF!. Check before deleting."],
 mistake:{wrong:"Pressing Delete to remove a whole row",why:"Delete only clears the contents; the empty row stays.",fix:"Right-click the row number and choose Delete."},
 questions:[["Insert a row shortcut:",["Ctrl and +","Ctrl and -","Ctrl+R"],0,"Ctrl + plus inserts."],["You deleted a column a formula used. You'll see:",["#REF!","#N/A","#####"],0,"The reference no longer exists."]]}),
RT("xg6","excel","basic","Getting started","Undo, redo and Paste Special",{
 intro:`<p><kbd>Ctrl</kbd>+<kbd>Z</kbd> undoes the last action, <kbd>Ctrl</kbd>+<kbd>Y</kbd> redoes it. <b>Paste Special</b> (<kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>V</kbd>) pastes only part of what you copied.</p>`,
 table:{cols:["Paste option","Result"],rows:[["Values","Only the results, no formulas (freeze numbers)"],["Formats","Only the formatting"],["Formulas","Only formulas, no formatting"],["Transpose","Rows become columns and columns become rows"],["Column widths","Copy the width of columns"],["Multiply / Add","Multiply or add the copied number into the selected cells"]]},
 steps:["Copy the cells (<kbd>Ctrl</kbd>+<kbd>C</kbd>).","Right-click the destination → <b>Paste Special</b>, or press <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>V</kbd>.","Choose an option such as <b>Values</b> and click OK."],
 notes:["Paste Values is the standard way to send a report without formulas or links."],
 questions:[["Keep only the results of formulas:",["Paste Values","Paste Formats","Transpose"],0,"Values removes formulas."],["Turn a vertical list into a horizontal one:",["Transpose","Undo","Format Painter"],0,"Transpose swaps rows and columns."]]}),

/* ===== Formulas ===== */
RT("xf1","excel","basic","Formulas","Arithmetic operators",{
 intro:`<p>Excel uses the same maths you learned in school. Every formula starts with <code>=</code>.</p>`,
 table:{cols:["Operator","Meaning","Example","Result"],rows:[["+","Add","=D2+D3","83000"],["-","Subtract","=D2-E2","2000"],["*","Multiply","=D3*2","62000"],["/","Divide","=D2/E2","1.04"],["^","Power","=2^3","8"],["%","Percent","=D2*10%","5200"],["&","Join text","=A2&\" \"&C2","Aarav Mumbai"]]},
 ex:[{desc:"Add two cells",code:"=D2+D3"},{desc:"Sales above target",code:"=D2-E2"},{desc:"10% bonus on sales",code:"=D2*10%"}], sheet:true,
 steps:["Click an empty cell and type <code>=</code>.","Click the first cell (Excel writes its address).","Type an operator such as <code>+</code>.","Click the second cell and press <kbd>Enter</kbd>."],
 mistake:{wrong:"D2+D3 (without =)",why:"Without the equals sign Excel stores it as text.",fix:"Always start with =."},
 questions:[["What does =2^3 give?",["6","8","9"],1,"2 × 2 × 2 = 8."],["Which operator joins text?",["&","+","*"],0,"& joins text."]], practice:["e1","e2"]}),
RT("xf2","excel","basic","Formulas","Parentheses and order of calculation",{
 intro:`<p>Excel calculates in a fixed order, like BODMAS: <b>brackets</b> → <b>^</b> → <b>* and /</b> → <b>+ and -</b>. Use parentheses to control the order.</p>`,
 ex:[{desc:"Multiply first",code:"=2+3*4"},{desc:"Brackets first",code:"=(2+3)*4"},{desc:"Average of two cells",code:"=(D2+D3)/2"}], sheet:true,
 mistake:{wrong:"=D2+D3/2",why:"Only D3 is divided by 2, then D2 is added.",fix:"=(D2+D3)/2"},
 notes:["Every opening bracket needs a closing one. Excel colours matching pairs while you type."],
 questions:[["=10-4/2 gives…",["3","8","7"],1,"4/2 = 2, then 10 − 2 = 8."]]}),
RT("xf3","excel","basic","Formulas","Using functions",{
 intro:`<p>A <b>function</b> is a ready-made formula, such as SUM or AVERAGE. Excel has about 500 of them, grouped on the <b>Formulas</b> tab.</p>`,
 steps:["<b>AutoSum</b>: select the cell below a column of numbers and press <kbd>Alt</kbd>+<kbd>=</kbd>, then Enter.","<b>Typing</b>: type <code>=su</code>, choose SUM from the list with the arrow keys, press <kbd>Tab</kbd>, select the range, type <code>)</code>, Enter.","<b>Insert Function</b>: click the <b>fx</b> button next to the Formula Bar, search for a function, and fill in the boxes in the dialog.","<b>Formula tips</b>: while typing, a tooltip shows the arguments. The bold one is the argument you are typing now. Optional arguments are in [square brackets]."],
 ex:[{desc:"Total",code:"=SUM(D2:D11)"},{desc:"Average",code:"=AVERAGE(D2:D11)"},{desc:"Highest",code:"=MAX(D2:D11)"}], sheet:true,
 notes:["Function names are not case-sensitive: =sum(d2:d11) works and Excel changes it to capitals."],
 questions:[["Shortcut for AutoSum:",["Alt+=","Ctrl+S","F4"],0,"Alt+= inserts SUM."],["In the tooltip SUM(number1, [number2], …), the brackets mean…",["Optional argument","Required argument","Error"],0,"Square brackets mark optional inputs."]], practice:["e3"]}),

/* ===== Formatting ===== */
RT("xm1","excel","basic","Formatting","Formatting cells",{
 intro:`<p>Formatting changes how cells <b>look</b>, never what they contain. The main tools are on the <b>Home</b> tab.</p>`,
 table:{cols:["Tool","Where","Use"],rows:[["Font, size, bold","Home → Font","Headers and emphasis"],["Fill colour","Home → Font → paint bucket","Highlight header rows"],["Borders","Home → Font → Borders","Grid lines that print"],["Alignment","Home → Alignment","Left/centre/right, top/middle/bottom"],["Wrap Text","Home → Alignment","Show long text on several lines"],["Merge & Center","Home → Alignment","Titles only; avoid inside data"],["Cell Styles","Home → Styles","Ready-made formats (Good, Bad, Heading)"]]},
 steps:["Select the cells.","Choose a format on the Home tab, or press <kbd>Ctrl</kbd>+<kbd>1</kbd> for the full <b>Format Cells</b> window.","<kbd>Ctrl</kbd>+<kbd>B</kbd> bold, <kbd>Ctrl</kbd>+<kbd>I</kbd> italic, <kbd>Ctrl</kbd>+<kbd>U</kbd> underline."],
 mistake:{wrong:"Merging cells inside a data table",why:"Merged cells break sorting, filtering, copying and PivotTables.",fix:"Use Format Cells → Alignment → Horizontal: Center Across Selection."},
 questions:[["Open Format Cells:",["Ctrl+1","Ctrl+F","Ctrl+9"],0,"Ctrl+1 opens Format Cells."],["Best alternative to merging cells:",["Center Across Selection","Wrap Text","Delete"],0,"It looks the same without merging."]]}),
RT("xm2","excel","basic","Formatting","Format Painter",{
 intro:`<p><b>Format Painter</b> copies formatting (font, colour, border, number format) from one cell to others.</p>`,
 steps:["Click a cell that has the formatting you want.","Click <b>Format Painter</b> (brush icon) on the Home tab.","Click or drag over the cells you want to format.","Double-click the brush to keep it active for several areas; press <kbd>Esc</kbd> to stop."],
 questions:[["Keep Format Painter on for many areas:",["Double-click the brush","Press F2","Ctrl+Z"],0,"Double-click locks it on."]]}),
RT("xm3","excel","basic","Formatting","Number formats",{
 intro:`<p>Number formats decide how values display: currency, percentage, date, decimals. The value itself does not change.</p>`,
 table:{cols:["Format","52000.5 / 0.25 / a date shows as","How"],rows:[["General","52000.5","Default"],["Number (2 dp)","52,000.50","Home → Number"],["Currency ₹","₹52,000.50","Home → Number → Currency"],["Percentage","25%","Ctrl+Shift+%"],["Short date","15-03-2021","Home → Number → Short Date"],["Custom: #,##0","52,001","Ctrl+1 → Custom"],["Custom: [h]:mm","27:30 (durations over 24 h)","Ctrl+1 → Custom"],["Custom: 00000","00123 (leading zeros)","Ctrl+1 → Custom"]]},
 steps:["Select the cells.","Pick a format from the dropdown in Home → Number, or press <kbd>Ctrl</kbd>+<kbd>1</kbd> → Number tab.","For Indian grouping (1,00,000) use Custom: <code>[&gt;=10000000]##\\,##\\,##\\,##0;[&gt;=100000]##\\,##\\,##0;##,##0</code>."],
 notes:["To use a formatted number inside text, use TEXT(): =\"Total: \"&TEXT(D2,\"#,##0\")."],
 mistake:{wrong:"Typing 25 and applying % expecting 25%",why:"Percentage format multiplies by 100 for display, so 25 shows as 2500%.",fix:"Type 0.25 (or 25%) and then apply the format."},
 questions:[["Format to show durations over 24 hours:",["[h]:mm","hh:mm","dd"],0,"Square brackets keep counting hours."]]}),
RT("xm4","excel","basic","Formatting","View, freeze panes and printing",{
 intro:`<p>The <b>View</b> and <b>Page Layout</b> tabs control what you see on screen and on paper.</p>`,
 steps:["<b>Freeze the header row</b>: View → Freeze Panes → Freeze Top Row.","<b>Freeze rows and columns</b>: click the cell below and to the right of what should stay, then View → Freeze Panes.","<b>Gridlines</b>: View → untick Gridlines for a clean dashboard.","<b>Zoom</b>: bottom-right slider, or <kbd>Ctrl</kbd> + mouse wheel.","<b>Print</b>: Page Layout → Orientation (Landscape), Size, Print Area, and <b>Print Titles</b> to repeat headers on each page. Preview with <kbd>Ctrl</kbd>+<kbd>P</kbd>.","<b>Fit to one page</b>: in Print settings choose \"Fit All Columns on One Page\"."],
 questions:[["Keep the header row visible while scrolling:",["Freeze Top Row","Wrap Text","Merge"],0,"View → Freeze Panes → Freeze Top Row."],["Repeat the header on every printed page:",["Print Titles","Gridlines","Zoom"],0,"Page Layout → Print Titles."]]}),

/* ===== Data analysis ===== */
RT("xd1","excel","basic","Working with data","Sort",{
 intro:`<p>Sorting reorders rows by one or more columns: A→Z, smallest→largest, oldest→newest, or by colour.</p>`,
 steps:["Click any cell in the column you want to sort by.","Data → <b>Sort A to Z</b> or <b>Sort Z to A</b>.","For several levels: Data → <b>Sort</b> → Add Level (e.g. Department, then Sales largest to smallest).","Tick <b>My data has headers</b> so the header row stays on top."],
 notes:["Select one cell, not one column: selecting only a column and sorting it scrambles your rows."],
 mistake:{wrong:"Selecting only column D and sorting",why:"Only that column moves; names no longer match their sales.",fix:"Click one cell in the data, or choose \"Expand the selection\" when Excel asks."},
 questions:[["Sort by Department, then by Sales:",["Data → Sort → Add Level","Filter","Freeze Panes"],0,"Multi-level sort uses Add Level."]]}),
RT("xd2","excel","basic","Working with data","Filter",{
 intro:`<p>A filter hides rows that don't match, so you can focus on part of the data. Hidden rows are not deleted.</p>`,
 steps:["Click inside your data and press <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>L</kbd> (or Data → Filter).","Click a header's dropdown arrow.","Untick <b>Select All</b>, tick the values you want, click OK.","Use <b>Number Filters</b> (greater than, top 10) or <b>Date Filters</b> (this month, last quarter) for conditions.","Clear: Data → <b>Clear</b>, or press <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>L</kbd> again to remove filters."],
 ex:[{desc:"SUM counts hidden rows too",code:"=SUM(D2:D11)"},{desc:"SUBTOTAL ignores rows hidden by a filter",code:"=SUBTOTAL(9,D2:D11)"}], sheet:true,
 notes:["The row numbers turn blue when a filter is active."],
 questions:[["Which function totals only visible filtered rows?",["SUM","SUBTOTAL","COUNTA"],1,"SUBTOTAL skips filtered-out rows."]], practice:["e44"]}),
RT("xd3","excel","basic","Working with data","Excel Tables",{
 intro:`<p>An <b>Excel Table</b> turns a plain range into a smart list: it grows automatically, keeps formatting, adds filter buttons and lets formulas use column names.</p>`,
 steps:["Click inside your data and press <kbd>Ctrl</kbd>+<kbd>T</kbd>. Confirm <b>My table has headers</b>.","Rename it: Table Design → Table Name (e.g. <i>tblSales</i>).","Add a row by typing right below the table: it expands, and formulas fill down automatically.","Tick <b>Total Row</b> in Table Design to add totals with a dropdown (Sum, Average, Count…).","Remove duplicates: Table Design → <b>Remove Duplicates</b>.","Back to a normal range: Table Design → <b>Convert to Range</b>."],
 table:{cols:["Structured reference","Means"],rows:[["tblSales[Sales]","The whole Sales column"],["[@Sales]","Sales in the same row"],["tblSales[#Totals]","The total row"],["=SUM(tblSales[Sales])","Total that grows with the table"]]},
 notes:["Use Tables as the source for PivotTables and charts so new rows are included on refresh."],
 questions:[["Create a Table:",["Ctrl+T","Ctrl+L","Alt+T"],0,"Ctrl+T formats as a table."],["[@Sales] refers to…",["Sales in the same row","The whole column","The total"],0,"@ means this row."]]}),
RT("xd4","excel","basic","Working with data","Cleaning data: duplicates, split and trim",{
 intro:`<p>Real data is messy: extra spaces, duplicates, combined fields. Excel has quick tools for each.</p>`,
 table:{cols:["Problem","Tool"],rows:[["Duplicate rows","Data → Remove Duplicates"],["\"Mumbai-West\" in one cell","Data → Text to Columns (delimiter -) or Flash Fill"],["Extra spaces","=TRIM(A2), then Paste Values"],["Numbers stored as text","Text to Columns → Finish, or =VALUE(A2)"],["Wrong capitals","=PROPER(A2)"],["Find & replace words","Ctrl+H"],["Blank cells to fill","Select → F5 → Special → Blanks → type = ↑ → Ctrl+Enter"]]},
 ex:[{desc:"Remove extra spaces",code:'=TRIM("  Aarav   Shah ")'},{desc:"Fix capitals",code:'=PROPER("aARAV shah")'}], sheet:false,
 mistake:{wrong:"Removing duplicates on the only copy of data",why:"Removed rows can't be recovered after saving.",fix:"Copy the sheet first, or use a formula flag (COUNTIF) to review duplicates."},
 questions:[["Split \"Mumbai-West\" into two columns:",["Text to Columns","Freeze Panes","Paste Special"],0,"Use - as the delimiter."]], practice:["e52","e54"]}),
RT("xd5","excel","basic","Working with data","Data Validation (dropdown lists)",{
 intro:`<p><b>Data Validation</b> controls what people can type: a dropdown list, numbers in a range, or dates only. It stops bad data at the source.</p>`,
 steps:["Select the input cells (e.g. the Status column).","Data → <b>Data Validation</b>.","Allow: <b>List</b>. Source: <code>Open,In Progress,Closed</code> (or select a range of options).","On the <b>Error Alert</b> tab, write a friendly message.","Click OK. Each cell now shows a dropdown arrow."],
 table:{cols:["Allow","Example rule"],rows:[["Whole number","Between 1 and 100"],["Decimal","Greater than 0"],["List","Open, Closed"],["Date","Between 01/04/2026 and 31/03/2027"],["Text length","Exactly 10 (mobile number)"],["Custom","=COUNTIF($A:$A,A2)=1 (no duplicates)"]]},
 questions:[["Create a dropdown in a cell:",["Data Validation → List","Conditional Formatting","Sort"],0,"List validation shows a dropdown."]]}),

/* ===== Conditional formatting ===== */
RT("xc1","excel","basic","Conditional formatting","Highlight Cell Rules",{
 intro:`<p>Conditional formatting changes a cell's colour automatically based on its value. <b>Highlight Cell Rules</b> cover the common cases.</p>`,
 table:{cols:["Rule","Example"],rows:[["Greater Than","Sales &gt; 50000 in green"],["Less Than","Sales &lt; 35000 in red"],["Between","Targets between 40000 and 50000"],["Equal To","Department = \"IT\""],["Text that Contains","City contains \"Pu\""],["A Date Occurring","Due dates Yesterday / This week / Last month"],["Duplicate Values","Repeated ticket IDs"]]},
 steps:["Select the cells (e.g. D2:D11).","Home → <b>Conditional Formatting</b> → <b>Highlight Cells Rules</b> → choose a rule.","Type the value (or click a cell) and pick a colour. Click OK."],
 notes:["Rules update automatically when values change."],
 questions:[["Highlight repeated IDs:",["Duplicate Values","Data Bars","Freeze Panes"],0,"It's a built-in highlight rule."]]}),
RT("xc2","excel","basic","Conditional formatting","Top/Bottom rules, data bars, colour scales, icon sets",{
 intro:`<p>Besides simple highlights, Excel can show rankings and visual scales inside cells.</p>`,
 table:{cols:["Type","What it shows"],rows:[["Top 10 Items / Top 10%","Best performers"],["Bottom 10","Lowest values"],["Above/Below Average","Compared with the average"],["Data Bars","A bar inside each cell, length = value"],["Color Scales","Heat map: green high, red low"],["Icon Sets","Arrows, traffic lights, ratings"]]},
 steps:["Select the numbers.","Home → Conditional Formatting → Top/Bottom Rules, Data Bars, Color Scales or Icon Sets.","Fine-tune: Conditional Formatting → <b>Manage Rules</b> → Edit Rule (e.g. icon thresholds)."],
 questions:[["Show a heat map:",["Color Scales","Icon Sets","Top 10"],0,"Color scales shade by value."]]}),
RT("xc3","excel","basic","Conditional formatting","Formula rules and Manage Rules",{
 intro:`<p>For anything the built-in rules can't do, use <b>Use a formula to determine which cells to format</b>. The formula must return TRUE for cells to colour. Write it for the <b>first row</b> of your selection.</p>`,
 steps:["Select the whole data area, e.g. A2:F11 (start from row 2).","Home → Conditional Formatting → <b>New Rule</b> → <b>Use a formula…</b>","Type the formula for row 2, e.g. <code>=$D2&lt;$E2</code> (sales below target). The $ before the letter keeps the column fixed; the row stays relative.","Click Format → Fill → choose a colour → OK.","Manage everything in Conditional Formatting → <b>Manage Rules</b>: edit, delete, change order, or tick <b>Stop If True</b>."],
 table:{cols:["Rule formula","Colours"],rows:[["=$D2&lt;$E2","Rows where sales are below target"],["=$O4&gt;=$P4","Uptime meeting the SLA target (green)"],["=$O4&lt;$P4","Uptime below target (red)"],["=AND($E2&lt;&gt;\"Closed\",$D2&lt;TODAY())","Open items past due"],["=MOD(ROW(),2)=0","Zebra stripes"]]},
 mistake:{wrong:"=D2<E2 applied to A2:F11",why:"Without $, each column compares its own neighbours, so colours land in the wrong places.",fix:"Lock the column: =$D2<$E2."},
 questions:[["In =$D2<$E2, the $ keeps…",["The column fixed","The row fixed","Both fixed"],0,"$ before the letter locks the column."]]}),

/* ===== Charts ===== */
RT("xh1","excel","basic","Charts","Create a chart",{
 intro:`<p>Charts turn numbers into pictures. Excel recommends charts based on your data.</p>`,
 steps:["Select the data including headers, e.g. Name and Sales (A1:A11 and D1:D11 with <kbd>Ctrl</kbd>).","Insert → <b>Recommended Charts</b>, or pick a type in the Charts group.","Click the chart, then use the <b>+</b> button for Chart Elements (title, labels, legend), the brush for Styles, and the funnel for Filters.","Move the chart to its own sheet: Chart Design → Move Chart.","Shortcut: select data and press <kbd>Alt</kbd>+<kbd>F1</kbd> for an instant chart."],
 table:{cols:["Chart","Best for"],rows:[["Column / Bar","Comparing categories (sales by person)"],["Stacked column","Parts of a total per category"],["Line","Trends over time"],["Pie / Doughnut","Share of a whole, 5 slices or fewer"],["Combo","Two measures, e.g. sales (columns) and % achieved (line)"],["Scatter","Relationship between two numbers"]]},
 notes:["Charts update automatically when data changes. Base them on a Table so new rows are included."],
 questions:[["Monthly trend over a year:",["Line","Pie","Scatter"],0,"Lines show change over time."],["Instant chart shortcut:",["Alt+F1","Ctrl+1","F4"],0,"Alt+F1 inserts a default chart."]]}),
RT("xh2","excel","basic","Charts","Customise charts",{
 intro:`<p>A clear chart has a title that says the point, labelled values, and no clutter.</p>`,
 steps:["<b>Title</b>: click it and type a message, e.g. \"Sneha leads sales in Q1\".","<b>Data labels</b>: + → Data Labels.","<b>Axis</b>: right-click the axis → Format Axis to set minimum, maximum and number format.","<b>Colours</b>: click one bar twice to colour just that bar (highlight the key item).","<b>Secondary axis</b>: right-click a series → Format Data Series → Secondary Axis (for combo charts).","<b>Remove clutter</b>: delete gridlines and legends you don't need."],
 questions:[["Colour a single bar:",["Click it twice, then format","Change the theme","Sort the data"],0,"The second click selects one data point."]]}),

/* ===== PivotTables ===== */
RT("xp1","excel","basic","PivotTables","Create a PivotTable",{
 intro:`<p>A PivotTable summarises a whole table in seconds by dragging fields: no formulas needed.</p>`,
 steps:["Click inside your data (ideally an Excel Table).","Insert → <b>PivotTable</b> → New Worksheet → OK.","In the <b>PivotTable Fields</b> pane, drag <b>Department</b> to <b>Rows</b>.","Drag <b>Sales</b> to <b>Values</b>. Excel shows Sum of Sales per department.","Drag <b>City</b> to <b>Columns</b> to get a grid.","Drag a field to <b>Filters</b> to filter the whole pivot."],
 table:{cols:["Area","Holds","Example"],rows:[["Rows","Categories down the side","Department"],["Columns","Categories across the top","City"],["Values","Numbers to summarise","Sum of Sales"],["Filters","Report-level filter","Year"]]},
 ex:[{desc:"What the pivot calculates for Sales department (same as SUMIF)",code:'=SUMIF(B2:B11,"Sales",D2:D11)'}], sheet:true,
 questions:[["Where do you drop Sales to total it?",["Values","Rows","Filters"],0,"Values holds the numbers."]], practice:["e13"]}),
RT("xp2","excel","basic","PivotTables","Pivot values, filters and slicers",{
 intro:`<p>Values can be summed, counted, averaged or shown as percentages. Filters and slicers narrow down what the pivot shows.</p>`,
 steps:["Click a value → right-click → <b>Summarize Values By</b> → Count / Average / Max.","Right-click → <b>Show Values As</b> → % of Grand Total, % of Row Total, Running Total In, Difference From.","Right-click → <b>Value Field Settings</b> → Number Format to format all values at once.","Filter a row field with its dropdown (e.g. Top 10).","Insert a <b>Slicer</b>: PivotTable Analyze → Insert Slicer → tick Department. Click buttons to filter.","Insert a <b>Timeline</b> for date fields: PivotTable Analyze → Insert Timeline."],
 questions:[["Show each department's share of total:",["Show Values As → % of Grand Total","Sort","Group"],0,"Show Values As converts to %."],["Clickable filter buttons:",["Slicer","Chart","Data bar"],0,"Slicers are visual filters."]]}),
RT("xp3","excel","basic","PivotTables","Group, calculated fields and refresh",{
 intro:`<p>Grouping combines items (dates into months, numbers into bands). Calculated fields add new measures inside the pivot.</p>`,
 steps:["<b>Group dates</b>: right-click a date in the pivot → Group → tick Months and Years.","<b>Group numbers</b>: right-click a number field in Rows → Group → set Starting at, Ending at, By (e.g. 10000).","<b>Calculated field</b>: PivotTable Analyze → Fields, Items &amp; Sets → <b>Calculated Field</b>. Name: <i>Achievement</i>, Formula: <code>=Sales/Target</code>.","<b>Refresh</b> after the source data changes: right-click → Refresh, or Data → <b>Refresh All</b> (<kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>F5</kbd>).","<b>PivotChart</b>: PivotTable Analyze → PivotChart."],
 mistake:{wrong:"Expecting the pivot to update by itself",why:"PivotTables keep a copy (cache) of the data.",fix:"Refresh (or enable Refresh data when opening the file in PivotTable Options)."},
 questions:[["Combine dates into months:",["Group","Slicer","Sort"],0,"Right-click → Group."]]}),

/* ===== Case ===== */
RT("xcase","excel","basic","Case study","Case study: monthly ticket report",{
 intro:`<p>Put the chapters together. You receive a raw export of support tickets (Ticket ID, Created, Resolved, Priority, Status, Region) and must build a monthly report.</p>`,
 steps:["Press <kbd>Ctrl</kbd>+<kbd>T</kbd> to make the data a Table named <i>tblTickets</i>.","Clean: TRIM the text columns, check dates are real dates (right-aligned), remove duplicate Ticket IDs.","Add helper columns: <code>Age = IF([@Resolved]=\"\",INT(NOW()-[@Created]),INT([@Resolved]-[@Created]))</code> and an <b>Age Bucket</b> with IFS.","Add <b>SLA Due</b> = WORKDAY by priority, and <b>SLA Result</b> = Met / Breached.","Insert a PivotTable: Region in Rows, SLA Result in Columns, Count of Ticket ID in Values; Show Values As % of Row Total.","Add slicers for Priority and Month, a column chart, and conditional formatting on breached rows.","Freeze the header, set print titles, save as .xlsx."],
 notes:["All these formulas are ready to copy in the Formula Library (Ageing, SLA and Logic-based groups)."],
 questions:[["First step with a raw export:",["Make it a Table and clean it","Add charts","Print it"],0,"Clean, structured data comes first."]]}),

/* ===== Function reference: basic ===== */
XF("fx_sum","basic","Function reference","SUM","<p>Adds numbers. Use it for totals of any range.</p>","=SUM(number1, [number2], …)",[["number1","A cell, range or number to add"],["number2…","More cells or ranges (optional)"]],[["Total sales","=SUM(D2:D11)"],["Two ranges together","=SUM(D2:D11,E2:E11)"],["Separate cells","=SUM(D2,D5,D10)"]],{mistake:{wrong:"=SUM(D2+D3+D4)",why:"Works, but defeats the purpose and breaks when rows are added.",fix:"=SUM(D2:D4)"},practice:["e3","e25"]}),
XF("fx_avg","basic","Function reference","AVERAGE","<p>Returns the average (arithmetic mean). Empty cells and text are ignored; zeros are included.</p>","=AVERAGE(number1, [number2], …)",[["number1","Range of numbers"]],[["Average sales","=AVERAGE(D2:D11)"],["Average target","=AVERAGE(E2:E11)"]],{notes:["A blank and a 0 are different: AVERAGE skips blanks but counts zeros."],practice:["e4"]}),
XF("fx_minmax","basic","Function reference","MIN and MAX","<p><code>MIN</code> returns the smallest number, <code>MAX</code> the largest.</p>","=MIN(number1, …)   =MAX(number1, …)",[["number1","Range of numbers"]],[["Lowest sales","=MIN(D2:D11)"],["Highest sales","=MAX(D2:D11)"],["Earliest join date","=MIN(F2:F11)"]],{practice:["e5"]}),
XF("fx_count","basic","Function reference","COUNT, COUNTA, COUNTBLANK","<p>Three ways to count: numbers only, anything non-empty, or empty cells.</p>","=COUNT(range)  =COUNTA(range)  =COUNTBLANK(range)",[["range","Cells to count"]],[["Cells with numbers","=COUNT(D2:D11)"],["Non-empty cells (names)","=COUNTA(A2:A11)"],["Empty manager cells","=COUNTBLANK(I2:I11)"]],{mistake:{wrong:"=COUNT(A2:A11) to count names",why:"COUNT ignores text, so it returns 0.",fix:"=COUNTA(A2:A11)"},practice:["e6","e26"]}),
XF("fx_round","basic","Function reference","ROUND","<p>Rounds a number to a given number of decimals. Negative digits round to tens, hundreds, thousands.</p>","=ROUND(number, num_digits)",[["number","The value to round"],["num_digits","Decimals: 2 = two decimals, 0 = whole, -3 = thousands"]],[["Achievement to 2 decimals","=ROUND(D2/E2,2)"],["Average to nearest thousand","=ROUND(AVERAGE(D2:D11),-3)"]],{notes:["Formatting only hides decimals; ROUND actually changes the value used in later calculations."],practice:["e7","e31"]}),
XF("fx_if","basic","Function reference","IF","<p>Checks a condition and returns one value if TRUE and another if FALSE.</p>","=IF(logical_test, value_if_true, [value_if_false])",[["logical_test","A comparison such as D2>=E2"],["value_if_true","Result when the test is TRUE"],["value_if_false","Result when FALSE (optional)"]],[["Target met?",'=IF(D2>=E2,"Met","Missed")'],["Priya",'=IF(D3>=E3,"Met","Missed")'],["Bonus amount",'=IF(D2>50000,D2*5%,0)']],{mistake:{wrong:'=IF(D2>=E2,Met,Missed)',why:"Text without quotes is read as a name, giving #NAME?.",fix:'=IF(D2>=E2,"Met","Missed")'},practice:["e9"]}),
XF("fx_andor","basic","Function reference","AND and OR","<p><code>AND</code> is TRUE when all conditions are TRUE. <code>OR</code> is TRUE when at least one is. Usually used inside IF.</p>","=AND(logical1, [logical2], …)   =OR(logical1, …)",[["logical1","First condition"],["logical2…","More conditions"]],[["IT and above 45000",'=AND(B4="IT",D4>45000)'],["Pune or Delhi",'=OR(C2="Pune",C2="Delhi")'],["Inside IF",'=IF(AND(B4="IT",D4>45000),"Bonus","No bonus")']],{practice:["e10"]}),
XF("fx_countif","basic","Function reference","COUNTIF","<p>Counts cells that meet one condition. The condition can be a value, a comparison or a wildcard pattern.</p>","=COUNTIF(range, criteria)",[["range","Cells to check"],["criteria","\"IT\", \">50000\", \"A*\" or a cell"]],[["IT employees",'=COUNTIF(B2:B11,"IT")'],["Sales above 50000",'=COUNTIF(D2:D11,">50000")'],["Names starting with A",'=COUNTIF(A2:A11,"A*")'],["Above Aarav\'s target (cell)",'=COUNTIF(D2:D11,">"&E2)']],{notes:["Criteria are not case-sensitive: \"it\" also counts \"IT\"."],mistake:{wrong:'=COUNTIF(D2:D11,>50000)',why:"The comparison must be text in quotes.",fix:'=COUNTIF(D2:D11,">50000")'},practice:["e11","e12","e40"]}),
XF("fx_sumif","basic","Function reference","SUMIF","<p>Adds the values in one range where another range meets a condition.</p>","=SUMIF(range, criteria, [sum_range])",[["range","Cells to check"],["criteria","The condition"],["sum_range","Cells to add (if different from range)"]],[["Sales from Pune",'=SUMIF(C2:C11,"Pune",D2:D11)'],["HR sales",'=SUMIF(B2:B11,"HR",D2:D11)'],["Sum of sales above 50000",'=SUMIF(D2:D11,">50000")']],{practice:["e13"]}),
XF("fx_averageif","basic","Function reference","AVERAGEIF","<p>Averages the values where a condition is met.</p>","=AVERAGEIF(range, criteria, [average_range])",[["range","Cells to check"],["criteria","The condition"],["average_range","Cells to average"]],[["Average Sales dept",'=AVERAGEIF(B2:B11,"Sales",D2:D11)'],["Average in Delhi",'=AVERAGEIF(C2:C11,"Delhi",D2:D11)']],{practice:["e14"]}),
XF("fx_vlookup","basic","Function reference","VLOOKUP","<p>Searches for a value in the <b>first column</b> of a table and returns a value from another column in the same row.</p>","=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",[["lookup_value","What to find"],["table_array","The table; first column is searched"],["col_index_num","Which column to return (1 = first)"],["range_lookup","FALSE = exact match (use this), TRUE = approximate"]],[["Manager of IT",'=VLOOKUP("IT",H2:I4,2,FALSE)'],["Manager of HR",'=VLOOKUP("HR",H2:I4,2,FALSE)'],["Grade slab (approximate)",'=VLOOKUP(D2,K8:L10,2,TRUE)']],{notes:["Lock the table with $ (=VLOOKUP(A2,$H$2:$I$4,2,FALSE)) before filling down."],mistake:{wrong:'=VLOOKUP("IT",H2:I4,2)',why:"Without FALSE, Excel uses approximate match and may return a wrong row.",fix:'=VLOOKUP("IT",H2:I4,2,FALSE)'},practice:["e16","ea2"]}),
XF("fx_text3","basic","Function reference","LEFT, RIGHT, MID","<p>Extract part of a text: from the start, the end, or the middle.</p>","=LEFT(text, [num_chars])  =RIGHT(text, [num_chars])  =MID(text, start_num, num_chars)",[["text","The text or cell"],["num_chars","How many characters"],["start_num","Position to start (MID)"]],[["City code",'=LEFT(C2,3)'],["Last 3 letters",'=RIGHT(A2,3)'],["Middle",'=MID(A8,2,3)']],{practice:["e20","e51"]}),
XF("fx_lentrim","basic","Function reference","LEN and TRIM","<p><code>LEN</code> counts characters. <code>TRIM</code> removes extra spaces (leading, trailing and double spaces).</p>","=LEN(text)   =TRIM(text)",[["text","The text or cell"]],[["Length of a name","=LEN(A7)"],["Clean spaces",'=TRIM("  Aarav   Shah ")'],["Check for hidden spaces",'=LEN(A2)-LEN(TRIM(A2))']],{notes:["TRIM fixes most lookup failures caused by invisible spaces."],practice:["e52"]}),
XF("fx_case","basic","Function reference","UPPER, LOWER, PROPER","<p>Change capitalisation.</p>","=UPPER(text)  =LOWER(text)  =PROPER(text)",[["text","The text or cell"]],[["Capitals","=UPPER(A5)"],["Small letters","=LOWER(B2)"],["Title case",'=PROPER("mumbai west")']],{practice:["e21"]}),
XF("fx_concat","basic","Function reference","Joining text: &, CONCAT, TEXTJOIN","<p>Combine text from cells. <code>&amp;</code> is quickest, <code>TEXTJOIN</code> joins whole ranges with a separator.</p>","=A2&\" \"&B2   =CONCAT(text1, …)   =TEXTJOIN(delimiter, ignore_empty, text1, …)",[["delimiter","Separator, e.g. \", \""],["ignore_empty","TRUE skips blank cells"]],[["Name and city",'=A2&" - "&C2'],["CONCAT",'=CONCAT(A4," (",B4,")")'],["Join a range",'=TEXTJOIN(", ",TRUE,A2:A4)']],{practice:["e19","e50"]}),
XF("fx_today","basic","Function reference","TODAY and NOW","<p><code>TODAY()</code> returns today's date; <code>NOW()</code> returns the date and time. They update whenever the sheet recalculates.</p>","=TODAY()   =NOW()",[],[["Today's year","=YEAR(TODAY())"],["Days since Aarav joined","=TODAY()-F2"]],{steps:["Click a cell.","Type <code>=TODAY()</code> (keep the empty brackets) and press Enter.","Format the cell as a date if a number appears."],notes:["To stamp a fixed date that never changes, press Ctrl+; instead."],practice:["e56"]}),
XF("fx_ymd","basic","Function reference","YEAR, MONTH, DAY, WEEKDAY","<p>Extract parts of a date.</p>","=YEAR(date)  =MONTH(date)  =DAY(date)  =WEEKDAY(date, [return_type])",[["date","A date cell"],["return_type","2 = Monday is 1 … Sunday is 7"]],[["Year Aarav joined","=YEAR(F2)"],["Month Priya joined","=MONTH(F3)"],["Weekday (Mon=1)","=WEEKDAY(F2,2)"]],{practice:["e22","e60"]}),
XF("fx_datedif","basic","Function reference","DATEDIF","<p>Counts complete days, months or years between two dates. It is a hidden function: Excel doesn't suggest it while typing, but it works.</p>","=DATEDIF(start_date, end_date, unit)",[["start_date","Earlier date"],["end_date","Later date"],["unit","\"D\" days, \"M\" months, \"Y\" years, \"YM\" months after years"]],[["Months between joins",'=DATEDIF(F3,F2,"M")'],["Years between joins",'=DATEDIF(F6,F7,"Y")'],["Days",'=DATEDIF(F3,F2,"D")']],{mistake:{wrong:'=DATEDIF(F2,F3,"D") when F2 is later',why:"A start date after the end date gives #NUM!.",fix:"Put the earlier date first."},practice:["e58"]}),
XF("fx_networkdays","basic","Function reference","NETWORKDAYS and WORKDAY","<p><code>NETWORKDAYS</code> counts working days (Mon–Fri) between dates. <code>WORKDAY</code> returns the date after n working days. Both accept a holiday list.</p>","=NETWORKDAYS(start, end, [holidays])   =WORKDAY(start, days, [holidays])",[["start","Start date"],["end / days","End date, or number of working days to add"],["holidays","Optional range of holiday dates"]],[["Working days between joins","=NETWORKDAYS(F3,F2)"],["Business-day ageing","=NETWORKDAYS(F2,DATE(2026,3,31))-1"],["10 working days later","=WORKDAY(F2,10)"]],{notes:["Use NETWORKDAYS.INTL / WORKDAY.INTL for other weekends, e.g. \"0000001\" = Sunday only."],practice:["e61","ea17"]}),

/* ===== Function reference: advanced ===== */
XF("fx_countifs","advanced","Function reference","COUNTIFS and SUMIFS","<p>Count or sum with <b>several</b> conditions (all must be true). In SUMIFS the sum range comes first.</p>","=COUNTIFS(range1, criteria1, range2, criteria2, …)   =SUMIFS(sum_range, range1, criteria1, …)",[["range1, criteria1","First condition pair"],["range2, criteria2","Second condition pair"],["sum_range","Values to add (SUMIFS)"]],[["IT in Pune",'=COUNTIFS(B2:B11,"IT",C2:C11,"Pune")'],["Sales dept in Mumbai",'=SUMIFS(D2:D11,B2:B11,"Sales",C2:C11,"Mumbai")'],["Joined in 2021 or later",'=SUMIFS(D2:D11,F2:F11,">="&DATE(2021,1,1))']],{mistake:{wrong:'=SUMIFS(B2:B11,"Sales",D2:D11)',why:"The sum range must come first in SUMIFS (unlike SUMIF).",fix:'=SUMIFS(D2:D11,B2:B11,"Sales")'},practice:["e15","e41"]}),
XF("fx_avgifs","advanced","Function reference","AVERAGEIFS, MAXIFS, MINIFS","<p>Average, maximum or minimum with several conditions.</p>","=AVERAGEIFS(avg_range, range1, criteria1, …)",[["avg_range / max_range","Values"],["range1, criteria1","Condition pairs"]],[["IT not in Delhi",'=AVERAGEIFS(D2:D11,B2:B11,"IT",C2:C11,"<>Delhi")'],["Highest HR sales",'=MAXIFS(D2:D11,B2:B11,"HR")'],["Lowest in Pune",'=MINIFS(D2:D11,C2:C11,"Pune")']],{practice:["e42","e43"]}),
XF("fx_ifs","advanced","Function reference","IFS","<p>Tests conditions in order and returns the result of the first TRUE one. Cleaner than nested IFs.</p>","=IFS(test1, value1, [test2, value2], …)",[["test1, value1","First condition and its result"],["TRUE, value","Final catch-all (otherwise)"]],[["Grade",'=IFS(D5>=60000,"A",D5>=45000,"B",TRUE,"C")'],["Age bucket",'=IFS(DATE(2026,3,31)-F7<=365,"0-365",DATE(2026,3,31)-F7<=1095,"366-1095",TRUE,"1095+")']],{mistake:{wrong:'=IFS(D5>=45000,"B",D5>=60000,"A")',why:"The first TRUE test wins, so 61000 gets B.",fix:"Put the strictest test first."},practice:["e36","ea18"]}),
XF("fx_switch","advanced","Function reference","SWITCH","<p>Compares one value against a list and returns the matching result, with an optional default.</p>","=SWITCH(expression, value1, result1, …, [default])",[["expression","The value to compare"],["value, result","Pairs"],["default","Result when nothing matches"]],[["Department code",'=SWITCH(B3,"Sales","S","HR","H","IT","T")'],["SLA days by department",'=SWITCH(B2,"Sales",2,"HR",5,3)']],{practice:["e37","ea20"]}),
XF("fx_iferror","advanced","Function reference","IFERROR and IFNA","<p><code>IFERROR</code> replaces any error with your value. <code>IFNA</code> only catches #N/A (lookup not found), so real mistakes stay visible.</p>","=IFERROR(value, value_if_error)   =IFNA(value, value_if_na)",[["value","The formula to check"],["value_if_error","What to show instead"]],[["Missing department",'=IFERROR(VLOOKUP("Finance",H2:I4,2,FALSE),"Not found")'],["Only #N/A",'=IFNA(VLOOKUP("Finance",H2:I4,2,FALSE),"Not found")'],["Safe divide",'=IFERROR(D2/0,0)']],{practice:["e17"]}),
XF("fx_xlookup","advanced","Function reference","XLOOKUP","<p>The modern lookup (Excel 2021/365): searches any column, returns any column, and has a built-in not-found value.</p>","=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])",[["lookup_value","What to find"],["lookup_array","Column to search"],["return_array","Column to return"],["if_not_found","Text when missing"]],[["Karan's city",'=XLOOKUP("Karan",A2:A11,C2:C11)'],["Not found text",'=XLOOKUP("Zoya",A2:A11,C2:C11,"Not found")'],["Look left (city → name)",'=XLOOKUP("Delhi",C2:C11,A2:A11)']],{practice:["e45"]}),
XF("fx_indexmatch","advanced","Function reference","INDEX and MATCH","<p><code>MATCH</code> finds the position; <code>INDEX</code> returns the value at a position. Together they look up in any direction and survive inserted columns.</p>","=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))",[["return_range","Column to return"],["lookup_value","What to find"],["lookup_range","Column to search"],["0","Exact match"]],[["Position of Vikram",'=MATCH("Vikram",A2:A11,0)'],["Meera's sales",'=INDEX(D2:D11,MATCH("Meera",A2:A11,0))'],["Two-way: Meera's target",'=INDEX(A1:F11,MATCH("Meera",A1:A11,0),MATCH("Target",A1:F1,0))']],{practice:["e18","ea1"]}),
XF("fx_sumproduct","advanced","Function reference","SUMPRODUCT","<p>Multiplies ranges row by row and adds the results. Great for value totals and weighted averages.</p>","=SUMPRODUCT(array1, [array2], …)",[["array1, array2","Ranges of the same size"]],[["Stock value (price × qty)","=SUMPRODUCT(L2:L5,M2:M5)"],["Weighted average price","=ROUND(SUMPRODUCT(L2:L5,M2:M5)/SUM(M2:M5),2)"]],{practice:["e30","ea9"]}),
XF("fx_rank","advanced","Function reference","RANK, LARGE, SMALL","<p><code>RANK</code> gives a value's position; <code>LARGE</code>/<code>SMALL</code> return the k-th largest/smallest value.</p>","=RANK(number, ref, [order])   =LARGE(array, k)",[["number","Value to rank"],["ref","List to rank within"],["order","0 = highest first, 1 = lowest first"]],[["Meera's rank","=RANK(D9,D2:D11)"],["2nd highest","=LARGE(D2:D11,2)"],["3rd lowest","=SMALL(D2:D11,3)"]],{practice:["e29","e28"]}),
XF("fx_findsub","advanced","Function reference","FIND, SEARCH, SUBSTITUTE","<p><code>FIND</code>/<code>SEARCH</code> return the position of text (FIND is case-sensitive). <code>SUBSTITUTE</code> replaces text.</p>","=FIND(find_text, within_text)  =SUBSTITUTE(text, old, new, [instance])",[["find_text","Text to look for"],["within_text","Where to look"],["old, new","Text to replace and replacement"]],[["Position of e",'=FIND("e",A9)'],["Contains ee?",'=ISNUMBER(SEARCH("ee",A9))'],["Rename city",'=SUBSTITUTE(C2,"Mumbai","Bombay")']],{practice:["e53","e54","ea8"]}),
XF("fx_textfn","advanced","Function reference","TEXT","<p>Converts a number or date to text in a chosen format, so it can be joined into sentences.</p>","=TEXT(value, format_text)",[["value","Number or date"],["format_text","Codes like \"0.00\", \"#,##0\", \"dd/mm/yyyy\", \"mmm-yy\", \"0%\""]],[["Date as text",'=TEXT(F2,"dd/mm/yyyy")'],["Sentence",'=A2&" joined on "&TEXT(F2,"dd/mm/yyyy")']],{notes:["The result is text: you can't add it up afterwards."],practice:["e55","ea16"]}),
XF("fx_eomonth","advanced","Function reference","EDATE and EOMONTH","<p><code>EDATE</code> moves a date by whole months. <code>EOMONTH</code> returns the last day of a month.</p>","=EDATE(start_date, months)   =EOMONTH(start_date, months)",[["start_date","Date"],["months","0 = same month, 1 = next, -1 = previous"]],[["6 months later","=EDATE(F2,6)"],["Month end","=EOMONTH(F2,0)"],["First of next month","=EOMONTH(F2,0)+1"]],{practice:["e59"]}),
XF("fx_offset","advanced","Function reference","OFFSET","<p>Returns a range moved from a starting cell by rows and columns, with an optional size. Used for dynamic ranges.</p>","=OFFSET(reference, rows, cols, [height], [width])",[["reference","Starting cell"],["rows, cols","How far to move"],["height, width","Size of the result range"]],[["First 5 sales","=SUM(OFFSET(D2,0,0,5,1))"],["Last 5 sales","=AVERAGE(OFFSET(D2,5,0,5,1))"]],{notes:["OFFSET recalculates on every change (volatile). In big files prefer INDEX."],practice:["ea10"]}),
XF("fx_stats","advanced","Function reference","MEDIAN, MODE, STDEV","<p>Basic statistics: the middle value, the most common value, and how spread out values are.</p>","=MEDIAN(range)  =MODE(range)  =STDEV.S(range)",[["range","Numbers"]],[["Median sales","=MEDIAN(D2:D11)"],["Most common target","=MODE(E2:E11)"],["Spread of sales","=STDEV.S(D2:D11)"]],{practice:["e27"]}),
XF("fx_dynamic","advanced","Function reference","FILTER, SORT, UNIQUE","<p>Dynamic-array functions (Excel 365/2021) return many results that spill into neighbouring cells.</p>","=FILTER(array, include, [if_empty])  =SORT(array, [sort_index], [order])  =UNIQUE(array)",[["array","Data to return"],["include","TRUE/FALSE test per row, e.g. C2:C11=\"Pune\""],["sort_index / order","Column number; 1 ascending, -1 descending"]],[["Pune rows",'=FILTER(A2:D11,C2:C11="Pune")'],["Sorted by sales",'=SORT(A2:D11,4,-1)'],["Unique departments","=UNIQUE(B2:B11)"]],{noRun:true,notes:["These spill: keep the cells below and to the right empty to avoid #SPILL!.","Not available in Excel 2019 or older."]})
];
/* ---------- Rich theory: SQL (examples run live on the practice database) ---------- */
const QS = (id, level, group, title, intro, syntax, ex, o = {}) => RT(id, "sql", level, group, title, Object.assign({intro, syntax, ex:ex.map(([desc, code, run]) => ({desc, code, run:run !== false}))}, o));

const SQL_RICH = [
QS("qs1","basic","SQL basics","SQL syntax and rules",
`<p>A SQL <b>statement</b> is a command sent to the database. Every statement is built from <b>keywords</b> (SELECT, FROM, WHERE…), <b>table and column names</b>, and <b>values</b>.</p>
<ul><li>Keywords are not case-sensitive: <code>select</code> = <code>SELECT</code>. Write them in capitals for readability.</li><li>End each statement with a semicolon <code>;</code> (required when you run several statements together).</li><li>Text values go in <b>single quotes</b>: <code>'Pune'</code>. Numbers don't.</li><li>Comments: <code>-- one line</code> or <code>/* several lines */</code>.</li><li>Line breaks and spaces don't matter; use them to make queries readable.</li></ul>`,
"SELECT column1, column2\nFROM table_name\nWHERE condition;",
[["The tables in this course's practice database","SELECT name, department, city FROM employees;"],["A comment is ignored","-- this line is a comment\nSELECT COUNT(*) FROM employees;"]],
{terms:[["Statement","One complete SQL command"],["Keyword","Reserved word like SELECT"],["Clause","Part of a statement, e.g. WHERE …"]],questions:[["How is text written in SQL?",["'Pune'","\"Pune\" only","Pune"],0,"Single quotes are standard for text."],["Are SQL keywords case-sensitive?",["No","Yes","Only in WHERE"],0,"select and SELECT are the same."]],practice:["s1"]}),
QS("qs2","basic","SQL basics","SELECT","<p><code>SELECT</code> reads data. List the columns you want, or use <code>*</code> for all columns.</p>","SELECT column1, column2 FROM table_name;\nSELECT * FROM table_name;",
[["All columns","SELECT * FROM departments;"],["Chosen columns","SELECT name, salary FROM employees;"],["Calculated column","SELECT name, salary * 12 AS yearly FROM employees;"]],
{notes:["In real projects avoid SELECT * on big tables; list only the columns you need."],questions:[["Select every column:",["SELECT *","SELECT ALL()","SELECT columns"],0,"* means all columns."]],practice:["s1","s2"]}),
QS("qs3","basic","SQL basics","SELECT DISTINCT","<p><code>DISTINCT</code> removes duplicate rows from the result.</p>","SELECT DISTINCT column FROM table_name;",
[["Unique cities","SELECT DISTINCT city FROM employees;"],["Unique department + city pairs","SELECT DISTINCT department, city FROM employees;"],["Count of unique cities","SELECT COUNT(DISTINCT city) FROM employees;"]],
{mistake:{wrong:"SELECT DISTINCT city, name FROM employees;",why:"DISTINCT applies to the whole row, so every name makes the row unique again.",fix:"Select only the columns that should be unique."},practice:["s10","s27"]}),
QS("qs4","basic","SQL basics","WHERE","<p><code>WHERE</code> filters rows. Only rows where the condition is TRUE are returned.</p>","SELECT columns FROM table_name WHERE condition;",
[["Text condition","SELECT * FROM employees WHERE city = 'Pune';"],["Number condition","SELECT name, salary FROM employees WHERE salary > 50000;"],["Date condition","SELECT * FROM sales WHERE sale_date >= '2026-03-01';"]],
{table:{cols:["Operator","Meaning"],rows:[["=","Equal"],["&lt;&gt; or !=","Not equal"],["&gt;  &lt;","Greater / less than"],["&gt;=  &lt;=","Greater / less or equal"],["BETWEEN","In a range"],["LIKE","Pattern"],["IN","In a list"],["IS NULL","Empty value"]]},
mistake:{wrong:"WHERE city = Pune",why:"Without quotes Pune is treated as a column name.",fix:"WHERE city = 'Pune'"},practice:["s3","s4"]}),
QS("qs5","basic","SQL basics","AND, OR, NOT","<p>Combine conditions. <code>AND</code> needs both, <code>OR</code> needs either, <code>NOT</code> reverses. Use brackets when you mix AND and OR.</p>","WHERE condition1 AND condition2\nWHERE condition1 OR condition2\nWHERE NOT condition",
[["Both conditions","SELECT name FROM employees WHERE department = 'IT' AND city = 'Pune';"],["Either condition","SELECT name, city FROM employees WHERE city = 'Pune' OR city = 'Delhi';"],["Mixed with brackets","SELECT name FROM employees WHERE (city = 'Pune' OR city = 'Delhi') AND salary > 40000;"],["NOT","SELECT name FROM employees WHERE NOT department = 'Sales';"]],
{mistake:{wrong:"WHERE city = 'Pune' OR city = 'Delhi' AND salary > 40000",why:"AND runs before OR, so the salary rule applies only to Delhi.",fix:"Add brackets around the OR part."},practice:["s5","s21"]}),
QS("qs6","basic","SQL basics","ORDER BY","<p><code>ORDER BY</code> sorts the result. <code>ASC</code> (default) = smallest first; <code>DESC</code> = largest first. Sort by several columns separated by commas.</p>","SELECT columns FROM table_name ORDER BY column1 ASC, column2 DESC;",
[["Highest salary first","SELECT name, salary FROM employees ORDER BY salary DESC;"],["Department, then salary","SELECT department, name, salary FROM employees ORDER BY department, salary DESC;"],["By a calculated alias","SELECT name, salary * 12 AS yearly FROM employees ORDER BY yearly DESC;"]],
{practice:["s8","s25"]}),
QS("qs7","basic","SQL basics","TOP, LIMIT, FETCH FIRST","<p>Return only the first rows. Each database uses a different keyword.</p>","SELECT … ORDER BY … LIMIT 3;            -- MySQL, PostgreSQL, SQLite\nSELECT TOP 3 … ORDER BY …;            -- SQL Server\nSELECT … ORDER BY … FETCH FIRST 3 ROWS ONLY; -- Oracle, standard",
[["Top 3 earners","SELECT name, salary FROM employees ORDER BY salary DESC LIMIT 3;"],["Skip 3, take 3 (page 2)","SELECT name, salary FROM employees ORDER BY salary DESC LIMIT 3 OFFSET 3;"]],
{notes:["Without ORDER BY, \"top\" rows are in no guaranteed order."],practice:["s9","s26"]}),
QS("qs8","basic","SQL basics","NULL values","<p><b>NULL</b> means unknown or missing. It is not zero and not empty text. Comparisons with NULL are never TRUE, so use <code>IS NULL</code> / <code>IS NOT NULL</code>.</p>","WHERE column IS NULL\nWHERE column IS NOT NULL",
[["No bonus","SELECT name FROM employees WHERE bonus IS NULL;"],["Has a mentor","SELECT name, mentor_id FROM employees WHERE mentor_id IS NOT NULL;"],["This finds nothing","SELECT name FROM employees WHERE bonus = NULL;"]],
{mistake:{wrong:"WHERE bonus = NULL",why:"NULL = NULL is unknown, not TRUE, so no rows match.",fix:"WHERE bonus IS NULL"},practice:["s16","s24"]}),
QS("qs9","basic","SQL basics","Aliases (AS)","<p>An <b>alias</b> gives a column or table a temporary name for the query. Table aliases make joins short.</p>","SELECT column AS alias_name FROM table_name AS t;",
[["Column alias","SELECT name AS employee, salary AS monthly_pay FROM employees;"],["Table alias","SELECT e.name, e.city FROM employees AS e;"]],
{notes:["Use double quotes or [brackets] for aliases with spaces: AS \"Monthly Pay\"."],practice:["s13"]}),
QS("qs10","basic","Functions & grouping","MIN, MAX, COUNT, SUM, AVG","<p><b>Aggregate functions</b> turn many rows into one value. NULLs are skipped (except COUNT(*)).</p>","SELECT COUNT(*), SUM(col), AVG(col), MIN(col), MAX(col) FROM table_name;",
[["All in one","SELECT COUNT(*) AS people, SUM(salary) AS total, AVG(salary) AS average, MIN(salary) AS lowest, MAX(salary) AS highest FROM employees;"],["COUNT(column) skips NULL","SELECT COUNT(*) AS rows_, COUNT(bonus) AS with_bonus FROM employees;"],["With a filter","SELECT AVG(salary) FROM employees WHERE department = 'IT';"]],
{questions:[["COUNT(bonus) counts…",["Rows where bonus is not NULL","All rows","Bonus total"],0,"COUNT(column) skips NULL."]],practice:["s11","s12"]}),
QS("qs11","basic","Functions & grouping","LIKE and wildcards","<p><code>LIKE</code> searches for patterns in text. <code>%</code> = any number of characters, <code>_</code> = exactly one character.</p>","WHERE column LIKE 'pattern'",
[["Starts with A","SELECT name FROM employees WHERE name LIKE 'A%';"],["Ends with a","SELECT name FROM employees WHERE name LIKE '%a';"],["Second letter r","SELECT name FROM employees WHERE name LIKE '_r%';"],["Contains ee","SELECT name FROM employees WHERE name LIKE '%ee%';"]],
{table:{cols:["Pattern","Matches"],rows:[["'A%'","Starts with A"],["'%a'","Ends with a"],["'%an%'","Contains an"],["'_____'","Exactly 5 characters"],["'A_%'","Starts with A, at least 2 characters"]]},practice:["s7","s23"]}),
QS("qs12","basic","Functions & grouping","IN and BETWEEN","<p><code>IN</code> matches any value in a list (or a subquery). <code>BETWEEN</code> matches a range including both ends.</p>","WHERE column IN (value1, value2)\nWHERE column BETWEEN low AND high",
[["IN list","SELECT name, department FROM employees WHERE department IN ('HR', 'IT');"],["NOT IN","SELECT name FROM employees WHERE city NOT IN ('Mumbai');"],["BETWEEN numbers","SELECT name, salary FROM employees WHERE salary BETWEEN 40000 AND 50000;"],["BETWEEN dates","SELECT * FROM sales WHERE sale_date BETWEEN '2026-02-01' AND '2026-02-28';"]],
{mistake:{wrong:"NOT IN (SELECT employee_id FROM projects)",why:"projects has a NULL employee_id, and NOT IN with a NULL returns no rows.",fix:"Use NOT EXISTS, or add WHERE employee_id IS NOT NULL in the subquery."},practice:["s6","s22"]}),
QS("qs13","basic","Functions & grouping","GROUP BY","<p><code>GROUP BY</code> groups rows that have the same value and calculates an aggregate for each group.</p>","SELECT column, AGG(col) FROM table_name GROUP BY column;",
[["People per department","SELECT department, COUNT(*) AS people FROM employees GROUP BY department;"],["Salary per city","SELECT city, SUM(salary) AS total FROM employees GROUP BY city ORDER BY total DESC;"],["Two columns","SELECT department, city, COUNT(*) FROM employees GROUP BY department, city;"]],
{mistake:{wrong:"SELECT department, name, COUNT(*) FROM employees GROUP BY department;",why:"name is neither grouped nor aggregated (most databases raise an error).",fix:"Add name to GROUP BY or remove it from SELECT."},practice:["s14","s29","s30"]}),
QS("qs14","basic","Functions & grouping","HAVING","<p><code>HAVING</code> filters groups after GROUP BY, so it can use aggregates. <code>WHERE</code> filters rows before grouping.</p>","SELECT column, AGG(col) FROM table_name GROUP BY column HAVING AGG(col) condition;",
[["Departments with more than 3 people","SELECT department, COUNT(*) FROM employees GROUP BY department HAVING COUNT(*) > 3;"],["Cities with high average","SELECT city, AVG(salary) FROM employees GROUP BY city HAVING AVG(salary) > 45000;"],["WHERE and HAVING together","SELECT department, AVG(salary) FROM employees WHERE join_year >= 2020 GROUP BY department HAVING AVG(salary) > 40000;"]],
{practice:["s15","sa3"]}),
QS("qs15","basic","Functions & grouping","CASE","<p><code>CASE</code> is SQL's IF: it returns a value based on conditions, checked in order.</p>","CASE\n  WHEN condition1 THEN result1\n  WHEN condition2 THEN result2\n  ELSE result\nEND",
[["Salary band","SELECT name, salary, CASE WHEN salary >= 55000 THEN 'High' WHEN salary >= 40000 THEN 'Medium' ELSE 'Low' END AS band FROM employees;"],["Count inside groups","SELECT department, SUM(CASE WHEN salary >= 50000 THEN 1 ELSE 0 END) AS high_earners FROM employees GROUP BY department;"]],
{practice:["s17","s38"]}),
QS("qs16","basic","Functions & grouping","NULL functions (COALESCE)","<p>Replace NULL with a value. <code>COALESCE</code> works everywhere; databases also have their own (IFNULL, ISNULL, NVL).</p>","COALESCE(column, replacement)",
[["Bonus or 0","SELECT name, COALESCE(bonus, 0) AS bonus FROM employees;"],["Total pay with NULL bonus","SELECT name, salary + COALESCE(bonus, 0) AS total_pay FROM employees;"],["What happens without it","SELECT name, salary + bonus AS total_pay FROM employees;"]],
{notes:["Any arithmetic with NULL gives NULL: 52000 + NULL = NULL."],practice:["s35"]}),
QS("qs17","basic","Functions & grouping","String functions","<p>Common text functions. Names differ slightly by database.</p>","UPPER(text)  LOWER(text)  LENGTH(text)  SUBSTR(text, start, len)  TRIM(text)  REPLACE(text, old, new)",
[["Text functions","SELECT name, UPPER(name), LENGTH(name), SUBSTR(city, 1, 3) AS code FROM employees;"],["Join text","SELECT name || ' (' || department || ')' AS label FROM employees;"],["Replace","SELECT REPLACE(city, 'Mumbai', 'Bombay') FROM employees;"]],
{table:{cols:["Task","SQL Server","MySQL","PostgreSQL"],rows:[["Length","LEN()","LENGTH()","LENGTH()"],["Substring","SUBSTRING()","SUBSTRING()","SUBSTRING()"],["Join","+ or CONCAT()","CONCAT()","|| or CONCAT()"]]},practice:["s31","s32","s33","s34"]}),
QS("qs18","basic","Functions & grouping","Date functions","<p>Dates are stored as DATE/DATETIME types. This course database (SQLite) stores them as text in <code>YYYY-MM-DD</code> form and uses <code>strftime</code>.</p>","strftime('%Y', date)  date(date, '+1 month')  julianday(end) - julianday(start)",
[["Year and month","SELECT name, hire_date, strftime('%Y', hire_date) AS year, strftime('%m', hire_date) AS month FROM employees;"],["Sales per month","SELECT strftime('%m', sale_date) AS month, SUM(amount) FROM sales GROUP BY month;"],["Days employed until 2026","SELECT name, CAST(julianday('2026-01-01') - julianday(hire_date) AS INTEGER) AS days FROM employees;"]],
{table:{cols:["Task","SQL Server","MySQL","PostgreSQL"],rows:[["Today","GETDATE()","CURDATE() / NOW()","CURRENT_DATE / NOW()"],["Year","YEAR(d)","YEAR(d)","EXTRACT(YEAR FROM d)"],["Difference in days","DATEDIFF(day, a, b)","DATEDIFF(b, a)","b - a"],["Add months","DATEADD(month, 1, d)","DATE_ADD(d, INTERVAL 1 MONTH)","d + INTERVAL '1 month'"]]},practice:["s36","s37","sa13"]}),
QS("qs19","basic","Joins","INNER JOIN","<p>Returns rows that have matching values in both tables.</p>","SELECT …\nFROM table1 a\nINNER JOIN table2 b ON a.key = b.key;",
[["Employee and manager","SELECT e.name, d.manager FROM employees e INNER JOIN departments d ON e.department = d.department;"],["Sales with names","SELECT e.name, s.sale_date, s.amount FROM sales s JOIN employees e ON s.employee_id = e.id;"]],
{steps:["Pick the main table (FROM) and give it an alias.","Add JOIN with the second table and an alias.","Write ON with the matching columns.","Select columns using the aliases (e.name, d.manager)."],practice:["s18","s39"]}),
QS("qs20","basic","Joins","LEFT JOIN","<p>Returns all rows from the left table and the matching rows from the right table; NULL where there is no match.</p>","SELECT … FROM table1 a LEFT JOIN table2 b ON a.key = b.key;",
[["All employees and their projects","SELECT e.name, p.project_name FROM employees e LEFT JOIN projects p ON e.id = p.employee_id;"],["Employees without a project","SELECT e.name FROM employees e LEFT JOIN projects p ON e.id = p.employee_id WHERE p.project_id IS NULL;"],["Departments with no staff","SELECT d.department FROM departments d LEFT JOIN employees e ON d.department = e.department WHERE e.id IS NULL;"]],
{mistake:{wrong:"LEFT JOIN … WHERE p.project_name = 'Mobile App'",why:"A WHERE on the right table removes the NULL rows, turning it into an INNER JOIN.",fix:"Put that condition in the ON clause if you want to keep all left rows."},practice:["s19","s40"]}),
QS("qs21","basic","Joins","RIGHT JOIN and FULL JOIN","<p><code>RIGHT JOIN</code> keeps all rows of the right table. <code>FULL OUTER JOIN</code> keeps all rows from both tables.</p>","SELECT … FROM a RIGHT JOIN b ON …;\nSELECT … FROM a FULL OUTER JOIN b ON …;",
[["All projects (even unassigned)","SELECT e.name, p.project_name FROM employees e RIGHT JOIN projects p ON e.id = p.employee_id;"],["Everything from both","SELECT e.name, p.project_name FROM employees e FULL OUTER JOIN projects p ON e.id = p.employee_id;"]],
{notes:["MySQL has no FULL OUTER JOIN: combine a LEFT JOIN and a RIGHT JOIN with UNION."],practice:["s41"]}),
QS("qs22","basic","Joins","Self join","<p>A table joined to itself, using two aliases. Used for hierarchies (employee → mentor/manager).</p>","SELECT a.col, b.col FROM t a JOIN t b ON a.parent_id = b.id;",
[["Employee and mentor","SELECT e.name, m.name AS mentor FROM employees e JOIN employees m ON e.mentor_id = m.id;"],["Mentee counts","SELECT m.name, COUNT(*) AS mentees FROM employees e JOIN employees m ON e.mentor_id = m.id GROUP BY m.name;"]],
{practice:["s42","sa4"]}),
QS("qs23","basic","Joins","UNION and UNION ALL","<p>Stacks the results of two queries. Same number of columns, compatible types. <code>UNION</code> removes duplicates; <code>UNION ALL</code> keeps them (faster).</p>","SELECT col FROM a\nUNION\nSELECT col FROM b;",
[["UNION","SELECT name FROM employees WHERE city = 'Pune' UNION SELECT name FROM employees WHERE salary > 55000;"],["UNION ALL keeps duplicates","SELECT name FROM employees WHERE city = 'Pune' UNION ALL SELECT name FROM employees WHERE salary > 55000;"]],
{practice:["s46","s47"]}),
QS("qs24","basic","Change data","INSERT INTO","<p>Adds new rows. List the columns, then the values in the same order.</p>","INSERT INTO table_name (col1, col2) VALUES (val1, val2);",
[["Add a department, then view it","INSERT INTO departments (department, manager, floor) VALUES ('Admin', 'Farhan', 1);\nSELECT * FROM departments;"],["Several rows","INSERT INTO departments VALUES ('Legal', 'Isha', 5), ('Ops', 'Ravi', 2);\nSELECT * FROM departments;"]],
{notes:["The practice database resets every time, so experiments are safe."],practice:["s54"]}),
QS("qs25","basic","Change data","UPDATE","<p>Changes existing rows. <b>Always</b> include WHERE, or every row changes.</p>","UPDATE table_name SET col1 = value1 WHERE condition;",
[["10% raise for IT","UPDATE employees SET salary = salary * 1.1 WHERE department = 'IT';\nSELECT name, department, salary FROM employees;"],["Move Karan to Pune","UPDATE employees SET city = 'Pune' WHERE name = 'Karan';\nSELECT name, city FROM employees WHERE name = 'Karan';"]],
{mistake:{wrong:"UPDATE employees SET salary = 0;",why:"No WHERE: every employee's salary becomes 0.",fix:"Run the WHERE as a SELECT first to check which rows will change."},practice:["s55"]}),
QS("qs26","basic","Change data","DELETE","<p>Removes rows. Without WHERE it deletes everything (the table remains).</p>","DELETE FROM table_name WHERE condition;",
[["Delete small sales","DELETE FROM sales WHERE amount < 10000;\nSELECT * FROM sales;"],["Count left","DELETE FROM projects WHERE employee_id IS NULL;\nSELECT COUNT(*) FROM projects;"]],
{table:{cols:["Command","Removes","Undo with ROLLBACK?"],rows:[["DELETE … WHERE","Chosen rows","Yes (inside a transaction)"],["TRUNCATE TABLE","All rows, fast","Depends on the database"],["DROP TABLE","The whole table","No"]]},practice:["s56"]}),
QS("qs27","basic","Database design","CREATE TABLE","<p>Creates a new table with columns, data types and rules (constraints).</p>","CREATE TABLE table_name (\n  column1 TYPE constraint,\n  column2 TYPE\n);",
[["Create, insert and read","CREATE TABLE training (id INTEGER PRIMARY KEY, course TEXT NOT NULL, hours INTEGER CHECK (hours > 0));\nINSERT INTO training (id, course, hours) VALUES (1, 'Excel', 6), (2, 'SQL', 8);\nSELECT * FROM training;"]],
{practice:["s57"]}),
QS("qs28","basic","Database design","ALTER TABLE and DROP","<p><code>ALTER TABLE</code> changes a table's structure: add, rename or remove columns. <code>DROP TABLE</code> deletes the table.</p>","ALTER TABLE t ADD column TYPE;\nALTER TABLE t RENAME COLUMN old TO new;\nDROP TABLE t;",
[["Add a column","ALTER TABLE departments ADD budget INTEGER;\nUPDATE departments SET budget = 500000 WHERE department = 'IT';\nSELECT * FROM departments;"],["Drop a table","CREATE TABLE temp_list (x INTEGER);\nDROP TABLE temp_list;\nSELECT name FROM sqlite_master WHERE type = 'table';"]],
{}),
/* ===== Advanced ===== */
QS("qa1","advanced","Advanced querying","Subqueries","<p>A subquery is a query inside another. It can return a single value, a list, or a table.</p>","WHERE col > (SELECT AVG(col) FROM t)      -- single value\nWHERE id IN (SELECT id FROM t2)           -- list\nFROM (SELECT … ) AS sub                   -- table",
[["Above average salary","SELECT name, salary FROM employees WHERE salary > (SELECT AVG(salary) FROM employees);"],["People with projects","SELECT name FROM employees WHERE id IN (SELECT employee_id FROM projects);"],["Derived table","SELECT department, avg_sal FROM (SELECT department, AVG(salary) AS avg_sal FROM employees GROUP BY department) AS t WHERE avg_sal > 40000;"]],
{practice:["s20","s44"]}),
QS("qa2","advanced","Advanced querying","EXISTS, ANY and ALL","<p><code>EXISTS</code> is TRUE when a subquery returns at least one row. <code>ANY</code>/<code>ALL</code> compare a value with every value a subquery returns (SQL Server, PostgreSQL, MySQL).</p>","WHERE EXISTS (SELECT 1 FROM t WHERE t.key = outer.key)",
[["Employees with a big sale","SELECT name FROM employees e WHERE EXISTS (SELECT 1 FROM sales s WHERE s.employee_id = e.id AND s.amount > 15000);"],["Departments with no staff","SELECT department FROM departments d WHERE NOT EXISTS (SELECT 1 FROM employees e WHERE e.department = d.department);"]],
{notes:["ALL/ANY example (not supported in SQLite): WHERE salary > ALL (SELECT salary FROM employees WHERE department = 'HR')."],practice:["s45"]}),
QS("qa3","advanced","Advanced querying","WITH (CTE)","<p>A <b>Common Table Expression</b> names a subquery at the top of the statement. Several CTEs can be chained with commas.</p>","WITH name AS (SELECT …),\n     name2 AS (SELECT … FROM name)\nSELECT … FROM name2;",
[["Totals per employee","WITH totals AS (SELECT employee_id, SUM(amount) AS total FROM sales GROUP BY employee_id) SELECT e.name, t.total FROM totals t JOIN employees e ON e.id = t.employee_id ORDER BY t.total DESC;"],["Recursive: numbers 1–5","WITH RECURSIVE n(x) AS (SELECT 1 UNION ALL SELECT x + 1 FROM n WHERE x < 5) SELECT x FROM n;"]],
{practice:["s48","sa14"]}),
QS("qa4","advanced","Advanced querying","Window functions: ranking","<p><code>ROW_NUMBER</code>, <code>RANK</code>, <code>DENSE_RANK</code> and <code>NTILE</code> number rows without collapsing them.</p>","FUNCTION() OVER (PARTITION BY group_col ORDER BY sort_col DESC)",
[["Compare ranking functions","SELECT name, bonus, ROW_NUMBER() OVER (ORDER BY bonus DESC) AS rn, RANK() OVER (ORDER BY bonus DESC) AS rnk, DENSE_RANK() OVER (ORDER BY bonus DESC) AS drnk FROM employees WHERE bonus IS NOT NULL;"],["Rank inside each department","SELECT name, department, salary, RANK() OVER (PARTITION BY department ORDER BY salary DESC) AS dept_rank FROM employees;"],["Top earner per department","WITH r AS (SELECT name, department, salary, ROW_NUMBER() OVER (PARTITION BY department ORDER BY salary DESC) AS rn FROM employees) SELECT * FROM r WHERE rn = 1;"]],
{practice:["s49","s50","s51","sa2"]}),
QS("qa5","advanced","Advanced querying","Window functions: running totals, LAG, LEAD","<p>Aggregate windows (<code>SUM() OVER</code>) give running totals and percentages; <code>LAG</code>/<code>LEAD</code> read neighbouring rows.</p>","SUM(x) OVER (ORDER BY d)\nLAG(x) OVER (ORDER BY d)\nAVG(x) OVER (ORDER BY d ROWS BETWEEN 2 PRECEDING AND CURRENT ROW)",
[["Running total and previous sale","SELECT sale_date, amount, SUM(amount) OVER (ORDER BY sale_date) AS running, LAG(amount) OVER (ORDER BY sale_date) AS prev FROM sales;"],["% of total","SELECT name, salary, ROUND(100.0 * salary / SUM(salary) OVER (), 2) AS pct FROM employees;"],["3-sale moving average","SELECT sale_date, amount, AVG(amount) OVER (ORDER BY sale_date ROWS BETWEEN 2 PRECEDING AND CURRENT ROW) AS moving_avg FROM sales;"]],
{practice:["s52","s53","sa5","sa6"]}),
QS("qa6","advanced","Database objects","Views","<p>A <b>view</b> is a saved query you can select from like a table. Good for reuse and for hiding sensitive columns.</p>","CREATE VIEW view_name AS SELECT …;\nSELECT * FROM view_name;",
[["Create and use a view","CREATE VIEW dept_summary AS SELECT department, COUNT(*) AS people, AVG(salary) AS avg_salary FROM employees GROUP BY department;\nSELECT * FROM dept_summary WHERE people > 2;"]],
{practice:["sa15"]}),
QS("qa7","advanced","Database objects","Constraints in practice","<p>Constraints stop bad data. Try inserting data that breaks a rule and the database refuses it.</p>","PRIMARY KEY, FOREIGN KEY, UNIQUE, NOT NULL, CHECK, DEFAULT",
[["CHECK stops a negative value","CREATE TABLE pay (id INTEGER PRIMARY KEY, amount INTEGER CHECK (amount > 0));\nINSERT INTO pay VALUES (1, 500);\nINSERT INTO pay VALUES (2, -50);"],["DEFAULT fills a missing value","CREATE TABLE tasks (id INTEGER PRIMARY KEY, status TEXT DEFAULT 'Open');\nINSERT INTO tasks (id) VALUES (1);\nSELECT * FROM tasks;"]],
{notes:["The first example shows an error on purpose: the CHECK constraint rejected -50."]}),
QS("qa8","advanced","Database objects","Transactions in practice","<p>Group changes so they all succeed or all fail. <code>ROLLBACK</code> undoes, <code>COMMIT</code> saves.</p>","BEGIN;\n  … changes …\nROLLBACK;   -- or COMMIT;",
[["Delete then undo","BEGIN;\nDELETE FROM sales;\nROLLBACK;\nSELECT COUNT(*) AS sales_still_there FROM sales;"],["Commit a change","BEGIN;\nUPDATE employees SET bonus = 1000 WHERE bonus IS NULL;\nCOMMIT;\nSELECT name, bonus FROM employees;"]],
{practice:["sa16"]})
];
/* ---------- Rich theory: Power BI (DAX examples run live on the practice model) ---------- */
const PT = (id, level, group, title, o) => RT(id, "pbi", level, group, title, o);
const DX = (id, level, group, name, intro, syntax, args, ex, o = {}) => PT(id, level, group, name, Object.assign({intro, syntax, args,
  ex: ex.map(([desc, code, by, run]) => ({desc, code, by: by || "Region", run: run !== false})),
  steps: o.steps || ["Go to the <b>Report</b> or <b>Table</b> view and select your table in the Data pane.","Click <b>New measure</b> (Home or Modeling tab). A formula bar opens.",`Type a name, an equals sign and the formula, e.g. <code>${ex[0][1].replace(/</g,"&lt;")}</code>.`,"Press <kbd>Enter</kbd>. The measure appears with a calculator icon.","Drag it into a Card or a Matrix to see the result. In a matrix, each row recalculates for its own filter."]}, o));

const PBI_RICH = [
PT("pb1","basic","Getting started","Install Power BI Desktop and load data",{
 intro:`<p><b>Power BI Desktop</b> is free for Windows. Get it from the Microsoft Store (updates automatically) or from the Microsoft download site.</p>`,
 steps:["Install Power BI Desktop from the Microsoft Store and open it.","On the start screen click <b>Get data</b> → <b>Excel workbook</b>.","Choose your file. The <b>Navigator</b> shows its sheets and tables: tick the ones you need.","Click <b>Transform Data</b> to clean them first in Power Query (recommended), or <b>Load</b> to load them as they are.","Save the report (<kbd>Ctrl</kbd>+<kbd>S</kbd>) as a <b>.pbix</b> file."],
 notes:["Power BI Desktop runs on Windows only. On a Mac, use a Windows virtual machine or build in the Power BI Service (browser) with limited features.","Load Excel data that is shaped as tables (one header row, no merged cells) for the best result."],
 questions:[["Report file extension:",[".pbix",".xlsx",".pbi"],0,"Power BI Desktop saves .pbix files."],["Clean data before loading with…",["Transform Data","Publish","Bookmarks"],0,"Transform Data opens Power Query."]]}),
PT("pb2","basic","Getting started","The Power BI Desktop window",{
 intro:`<p>Get to know the screen before building. Three views sit on the left edge, and panes on the right.</p>`,
 table:{cols:["Part","What it's for"],rows:[["Report view","Canvas where you build pages of visuals"],["Table view","See rows of each table, create calculated columns"],["Model view","Tables and relationships"],["Ribbon (Home, Insert, Modeling, View)","Get data, new measure, visuals, themes"],["Data pane","Tables, columns and measures"],["Visualizations pane","Choose a visual; fill its wells (X-axis, Y-axis, Legend, Values); Format tab"],["Filters pane","Filters on this visual, this page, or all pages"],["Page tabs (bottom)","Add, rename and order report pages"]]},
 questions:[["Where do you see relationships?",["Model view","Report view","Filters pane"],0,"Model view shows tables and lines."],["Change colours of a visual in…",["Visualizations → Format","Data pane","Model view"],0,"Format your visual with the paint brush."]]}),
PT("pb3","basic","Power Query","Power Query Editor tour",{
 intro:`<p>Home → <b>Transform data</b> opens the Power Query Editor, where you clean data before it enters the model.</p>`,
 table:{cols:["Part","Use"],rows:[["Queries pane (left)","Each table/query"],["Data preview (centre)","First rows; column headers show data type icons"],["Applied Steps (right)","Every change you make, in order; click a step to see the data at that point"],["Formula bar","The M code of the selected step"],["Close & Apply","Load the cleaned data into the model"]]},
 steps:["Click a column header to select it; <kbd>Ctrl</kbd>+click for several.","Right-click the header to see transformations for that column.","Check the column quality: View → <b>Column quality</b> and <b>Column distribution</b> show errors, empties and unique values.","When finished, click <b>Close &amp; Apply</b>."],
 notes:["Delete a wrong step with the X next to it in Applied Steps. Steps after it may break if they depend on it."],
 questions:[["Where is every cleaning step listed?",["Applied Steps","Visualizations","Model view"],0,"Applied Steps records the steps."]]}),
PT("pb4","basic","Power Query","Common transformations step by step",{
 intro:`<p>These transformations fix most data problems. Each becomes a step that repeats automatically on refresh.</p>`,
 table:{cols:["Task","How"],rows:[["Remove columns","Select → right-click → Remove (or Remove Other Columns)"],["Rename","Double-click the header"],["Change data type","Click the type icon left of the header (123, ABC, calendar)"],["Filter rows","Header dropdown → untick values or Text/Number/Date Filters"],["Remove blank rows","Home → Remove Rows → Remove Blank Rows"],["Remove duplicates","Select key column → right-click → Remove Duplicates"],["Trim / clean text","Transform → Format → Trim / Clean / Capitalize Each Word"],["Replace values","Right-click → Replace Values"],["Split column","Transform → Split Column → By Delimiter"],["Fill down","Right-click → Fill → Down"],["Use first row as headers","Home → Use First Row as Headers"]]},
 steps:["Open Transform data.","Set the correct data type on every column first (dates as Date, numbers as Decimal/Whole Number).","Remove the columns you don't need: fewer columns = smaller, faster model.","Clean text (Trim), remove blanks and duplicates.","Close & Apply."],
 mistake:{wrong:"Leaving dates as text",why:"Text dates can't be used in time intelligence or sorted correctly.",fix:"Change the type to Date (choose Using Locale if the format is dd/mm/yyyy)."},
 questions:[["Imported dates appear as text. Fix with…",["Change type to Date","Rename","Fill down"],0,"Set the column type."]]}),
PT("pb5","basic","Power Query","Add Column: conditional and custom columns",{
 intro:`<p>Create new columns in Power Query without DAX.</p>`,
 steps:["<b>Conditional Column</b>: Add Column → Conditional Column. Example: if Amount is greater than 20000 then \"Big\" else \"Small\".","<b>Column From Examples</b>: type the result you want for a few rows; Power Query writes the logic (like Flash Fill).","<b>Custom Column</b>: Add Column → Custom Column, formula in M, e.g. <code>[Quantity] * [Price]</code>.","<b>Date parts</b>: select a date column → Add Column → Date → Month → Name of Month.","<b>Group By</b>: Transform → Group By → e.g. Region, Sum of Amount."],
 table:{cols:["M formula","Result"],rows:[["[Quantity] * [Price]","Gross amount"],["if [Amount] > 20000 then \"Big\" else \"Small\"","Size label"],["Text.Upper([Region])","NORTH"],["Date.Year([Date])","2026"],["Duration.Days([Resolved] - [Created])","Days to resolve"]]},
 notes:["Power Query columns are calculated during refresh and stored, like DAX calculated columns. Prefer them for row-level cleanup."],
 questions:[["Create a column by typing example results:",["Column From Examples","Group By","Merge"],0,"It infers the logic from your examples."]]}),
PT("pb6","basic","Data model","Create relationships",{
 intro:`<p>Relationships let filters flow from one table to another, e.g. selecting a Region in the Regions table filters Sales.</p>`,
 steps:["Open <b>Model view</b>.","Drag the key column from the dimension table (Regions[Region]) onto the matching column in the fact table (Sales[Region]).","Check the settings by double-clicking the line: Cardinality <b>Many to one (*:1)</b>, Cross filter direction <b>Single</b>, <b>Active</b> ticked.","Hide key columns on the fact table (right-click → Hide in report view) so report builders use the dimension columns."],
 table:{cols:["Symbol","Meaning"],rows:[["1 → *","One-to-many (the normal case)"],["Arrow on the line","Direction filters flow"],["Dotted line","Inactive relationship (use with USERELATIONSHIP)"]]},
 mistake:{wrong:"Building visuals with Region from the Sales table and Target from Regions with no relationship",why:"Unrelated tables don't filter each other, so every row shows the same total.",fix:"Create the relationship on Region."},
 questions:[["Normal relationship cardinality:",["Many-to-one","Many-to-many","None"],0,"Fact (many) to dimension (one)."]]}),
PT("pb7","basic","Data model","Create a Date table",{
 intro:`<p>Time intelligence (YTD, last month, last year) needs a proper <b>Date table</b>: one row per day with no gaps.</p>`,
 steps:["Modeling → <b>New table</b>.",`Paste: <code>Date = ADDCOLUMNS(CALENDAR(DATE(2024,1,1), DATE(2026,12,31)), "Year", YEAR([Date]), "Month", FORMAT([Date],"MMM"), "Month No", MONTH([Date]), "Quarter", "Q" &amp; QUARTER([Date]))</code>`,"Select the Month column → <b>Sort by column</b> → Month No (so months aren't alphabetical).","Table tools → <b>Mark as date table</b> → choose the Date column.","In Model view, relate Date[Date] to your fact table's date column.","Turn off auto date/time: File → Options → Current file → Data Load."],
 notes:["The ready-made formula is in the Formula Library (Power BI → Time intelligence → Date table)."],
 questions:[["Why sort Month by Month No?",["Otherwise months sort alphabetically","To save space","It's required for relationships"],0,"Apr, Aug, Dec… otherwise."]]}),
PT("pb8","basic","DAX","Create your first measure",{
 intro:`<p>A <b>measure</b> is a calculation that responds to filters. You create it once and reuse it in any visual.</p>`,
 steps:["In the Data pane, right-click the <b>Sales</b> table → <b>New measure</b>.","Type <code>Total Sales = SUM(Sales[Amount])</code> and press Enter.","Set the format: Measure tools → Format → Currency/Whole number, thousands separator.","Insert a <b>Card</b> visual and drag Total Sales into it.","Insert a <b>Matrix</b>, put Region in Rows and Total Sales in Values. Each region shows its own total.","Create more measures that reuse it: <code>Avg Order = DIVIDE([Total Sales], COUNTROWS(Sales))</code>."],
 ex:[{desc:"Your first measure, evaluated on the practice Sales table",code:"Total Sales = SUM(Sales[Amount])",by:"Region",run:true}],
 notes:["Keep measures in a dedicated table: Home → Enter data → name it _Measures, then move measures there (Measure tools → Home table)."],
 questions:[["Where do you click to create a measure?",["New measure","New page","Publish"],0,"Home/Modeling → New measure."]],practice:["p6"]}),
PT("pb9","basic","DAX","DAX syntax rules",{
 intro:`<p>DAX formulas follow a few rules. Getting these right avoids most errors.</p>`,
 table:{cols:["Rule","Example"],rows:[["Name, equals sign, formula","Total Sales = SUM(Sales[Amount])"],["Columns: Table[Column]","Sales[Amount]"],["Measures: [Measure] (no table)","[Total Sales]"],["Table names with spaces in single quotes","'Sales Data'[Amount]"],["Text in double quotes","\"North\""],["Comparison","=  &lt;&gt;  &gt;  &lt;  &gt;=  &lt;="],["AND / OR / join text","&amp;&amp;   ||   &amp;"],["List","Sales[Region] IN {\"North\", \"South\"}"],["Comments","// one line   /* several lines */"],["Variables","VAR x = … RETURN …"]]},
 ex:[{desc:"Measure with a text filter",code:"North Sales = CALCULATE(SUM(Sales[Amount]), Sales[Region] = \"North\")",by:"Region",run:true},{desc:"Using IN",code:"South and West = CALCULATE([Total Sales], Sales[Region] IN {\"South\", \"West\"})",by:"Region",run:true}],
 mistake:{wrong:"Total = SUM(Amount)",why:"Columns need their table name.",fix:"Total = SUM(Sales[Amount])"},
 questions:[["Reference a measure:",["[Total Sales]","Sales.Total","Total Sales()"],0,"Measures use square brackets only."]]}),
PT("pb10","basic","Visuals","Build your first report page",{
 intro:`<p>A good report page answers one question at a glance: KPIs on top, trend in the middle, details below.</p>`,
 steps:["Insert → <b>Card</b> visuals for KPIs (Total Sales, Orders, Avg Order). Place them in a row at the top.","Insert a <b>Line chart</b>: Date (month) on X-axis, Total Sales on Y-axis.","Insert a <b>Clustered bar chart</b>: Region on Y-axis, Total Sales on X-axis; sort descending.","Insert a <b>Slicer</b> with Category; set Style → Tile for buttons.","Insert a <b>Table</b> or <b>Matrix</b> with the details.","Add a title text box and apply a theme: View → Themes.","Align visuals: Format → Align → Distribute."],
 table:{cols:["Need","Visual"],rows:[["One number","Card / KPI"],["Compare categories","Bar / Column"],["Trend","Line / Area"],["Two measures","Line and clustered column (combo)"],["Share","Donut (few categories)"],["Details","Table / Matrix"],["Location","Map / Filled map"],["Target progress","Gauge / KPI"],["Filter","Slicer"]]},
 questions:[["Visual for one big number:",["Card","Matrix","Scatter"],0,"Cards show one value."]]}),
PT("pb11","basic","Visuals","Filters: visual, page and report level",{
 intro:`<p>The Filters pane has three levels. Slicers are on-page filters that users click.</p>`,
 table:{cols:["Level","Affects"],rows:[["Filters on this visual","Only the selected visual"],["Filters on this page","All visuals on the page"],["Filters on all pages","Every page in the report"],["Slicer","Visuals on the page (or synced pages)"],["Drillthrough filter","The detail page you jump to"]]},
 steps:["Select a visual and drag a field into <b>Filters on this visual</b>.","Choose Basic filtering (tick values), Advanced (conditions) or <b>Top N</b> (e.g. top 5 by Total Sales).","Sync a slicer across pages: View → <b>Sync slicers</b>.","Stop a visual reacting to a slicer: Format → <b>Edit interactions</b> → click the ⊘ icon on that visual."],
 questions:[["Show only the top 5 products in one chart:",["Visual-level Top N filter","Report filter","Bookmark"],0,"Top N works per visual."]]}),
PT("pb12","basic","Visuals","Drill down, drillthrough and tooltips",{
 intro:`<p>Let users explore: go deeper in a hierarchy, jump to a detail page, or see extra info on hover.</p>`,
 steps:["<b>Drill down</b>: put Year, Quarter, Month in the X-axis of a chart. Use the arrow icons on the visual to go down a level.","<b>Drillthrough</b>: create a detail page, drag Region into <b>Drill through</b> in the Visualizations pane. On other pages, right-click a region → Drill through → your page.","<b>Tooltip page</b>: create a small page, Page format → Page information → <b>Allow use as tooltip</b>, Canvas → Tooltip size. On a visual, Format → Tooltips → Type: Report page."],
 questions:[["Jump from a region to its detail page:",["Drillthrough","Drill down","Bookmark"],0,"Drillthrough passes the filter to another page."]]}),
PT("pb13","basic","Visuals","Bookmarks and buttons",{
 intro:`<p>Bookmarks save a view (filters, visible visuals); buttons let users navigate like an app.</p>`,
 steps:["View → <b>Bookmarks</b> pane → set up the page as you want → <b>Add</b>.","Insert → <b>Buttons</b> → Blank; in Format → Action, choose Type: <b>Bookmark</b> or <b>Page navigation</b>.","Hold <kbd>Ctrl</kbd> and click a button to test it in Desktop.","Use the <b>Selection</b> pane to show/hide visuals for each bookmark (e.g. switch between chart and table)."],
 questions:[["Test a button in Desktop:",["Ctrl+click","Double-click","Right-click"],0,"Ctrl+click runs the action while editing."]]}),
PT("pb14","basic","Sharing","Publish and refresh",{
 intro:`<p>Publishing sends the report and its semantic model to the Power BI Service so others can view it.</p>`,
 steps:["Sign in (top-right of Desktop) with your work account.","Home → <b>Publish</b> → choose a workspace (My workspace or a team workspace).","Open the link to the Service. The report and its semantic model are now online.","In the Service: semantic model → <b>Settings</b> → <b>Scheduled refresh</b> → set times (e.g. daily 8:00).","If the data is on company servers or local files, install and configure an <b>on-premises data gateway</b>.","Share: <b>Share</b> button, or publish an <b>App</b> from the workspace for many viewers."],
 notes:["Publishing and sharing need a Pro or Premium Per User licence (or Premium/Fabric capacity) in most organisations."],
 questions:[["Automatically update online data daily:",["Scheduled refresh","Bookmarks","Drillthrough"],0,"Configure it on the semantic model."]]}),

/* ===== DAX reference: basic ===== */
DX("dx_sum","basic","DAX reference","SUM, AVERAGE, MIN, MAX","<p>Aggregate one numeric column in the current filter context.</p>","SUM(Table[Column])  AVERAGE(…)  MIN(…)  MAX(…)",[["Table[Column]","A numeric column"]],[["Total","Total Sales = SUM(Sales[Amount])"],["Average order","Avg Amount = AVERAGE(Sales[Amount])"],["Largest order","Max Order = MAX(Sales[Amount])"]],{practice:["p6","p7"]}),
DX("dx_count","basic","DAX reference","COUNTROWS and DISTINCTCOUNT","<p><code>COUNTROWS</code> counts rows of a table; <code>DISTINCTCOUNT</code> counts unique values of a column.</p>","COUNTROWS(Table)   DISTINCTCOUNT(Table[Column])",[["Table","Table to count"],["Table[Column]","Column whose unique values are counted"]],[["Orders","Orders = COUNTROWS(Sales)"],["Products sold","Products = DISTINCTCOUNT(Sales[Product])"],["Orders per product","Orders per Product = DIVIDE(COUNTROWS(Sales), DISTINCTCOUNT(Sales[Product]))","Category"]],{practice:["p8","p9"]}),
DX("dx_divide","basic","DAX reference","DIVIDE","<p>Safe division: returns BLANK (or your alternate result) instead of an error when dividing by zero.</p>","DIVIDE(numerator, denominator, [alternate_result])",[["numerator","Top value"],["denominator","Bottom value"],["alternate_result","Returned when the denominator is 0 (optional)"]],[["Average order value","AOV = DIVIDE([Total Sales], [Orders])"],["Discount %","Discount % = DIVIDE(SUM(Sales[Discount]), SUMX(Sales, Sales[Quantity] * Sales[Price]))"]],{mistake:{wrong:"AOV = [Total Sales] / [Orders]",why:"Fails with an infinity/error when Orders is 0 in some filter.",fix:"AOV = DIVIDE([Total Sales], [Orders])"},practice:["p10"]}),
DX("dx_calc","basic","DAX reference","CALCULATE","<p>Evaluates an expression in a <b>modified filter context</b>. Filters you pass replace existing filters on the same column.</p>","CALCULATE(expression, filter1, [filter2], …)",[["expression","Usually a measure"],["filter","Table[Column] = value, a FILTER table, or modifiers like ALL"]],[["Electronics only","Electronics = CALCULATE([Total Sales], Sales[Category] = \"Electronics\")"],["Two filters","East Phones = CALCULATE([Total Sales], Sales[Product] = \"Phone\", Sales[Region] = \"East\")"],["Number filter","Bulk = CALCULATE([Total Sales], Sales[Quantity] >= 5)"]],{notes:["Notice in the matrix: a Region filter inside CALCULATE replaces each row's Region, so every row shows the same value."],practice:["p11","p21","p23"]}),
DX("dx_filter","basic","DAX reference","FILTER","<p>Returns a table of rows that meet a condition. Use it inside COUNTROWS, SUMX or CALCULATE.</p>","FILTER(Table, condition)",[["Table","Table to scan"],["condition","Row-by-row test"]],[["Count big orders","Big Orders = COUNTROWS(FILTER(Sales, Sales[Amount] > 20000))"],["Inside CALCULATE","Discounted Sales = CALCULATE([Total Sales], FILTER(Sales, Sales[Discount] > 0))"]],{notes:["Prefer simple column filters in CALCULATE (faster); use FILTER when the condition compares columns or uses measures."],practice:["p13","p27"]}),
DX("dx_all","basic","DAX reference","ALL and % of total","<p><code>ALL</code> removes filters from a table or column. Used for grand totals and percentages.</p>","ALL(Table)   ALL(Table[Column])",[["Table / Column","What to remove filters from"]],[["Grand total","Grand Total = CALCULATE([Total Sales], ALL(Sales))"],["% of total","Share = DIVIDE([Total Sales], CALCULATE([Total Sales], ALL(Sales)))"],["% within categories","Cat Share = DIVIDE([Total Sales], CALCULATE([Total Sales], ALL(Sales[Category])))","Category"]],{practice:["p14","p24"]}),
DX("dx_if","basic","DAX reference","IF and SWITCH","<p><code>IF</code> returns one of two results; <code>SWITCH(TRUE(), …)</code> handles many conditions cleanly.</p>","IF(condition, then, [else])   SWITCH(TRUE(), cond1, r1, cond2, r2, else)",[["condition","TRUE/FALSE test"],["then / else","Results"]],[["Status","Status = IF([Total Sales] > 150000, \"Top\", \"Growing\")"],["Band","Band = SWITCH(TRUE(), [Total Sales] >= 150000, \"High\", [Total Sales] >= 100000, \"Medium\", \"Low\")"]],{practice:["p15","p29"]}),
DX("dx_sumx","basic","DAX reference","SUMX and iterators","<p>X-functions go row by row through a table, calculate an expression on each row, then aggregate: SUMX, AVERAGEX, MINX, MAXX, COUNTX.</p>","SUMX(Table, expression)",[["Table","Rows to iterate"],["expression","Calculated for each row"]],[["Gross before discount","Gross = SUMX(Sales, Sales[Quantity] * Sales[Price])"],["Largest line","Max Gross = MAXX(Sales, Sales[Quantity] * Sales[Price])"],["Average per month","Avg per Month = AVERAGEX(VALUES(Sales[Month]), [Total Sales])"]],{practice:["p12","p26","pa2"]}),
DX("dx_related","basic","DAX reference","RELATED","<p>Brings a value from the related (one-side) table into the current row. Works in calculated columns and inside iterators.</p>","RELATED(OtherTable[Column])",[["OtherTable[Column]","Column on the one side of a relationship"]],[["Manager per sale (column)","Manager = RELATED(Regions[Manager])"],["Target of visible regions","Target = SUMX(VALUES(Sales[Region]), RELATED(Regions[Target]))"]],{steps:["Make sure a relationship exists (Model view).","For a column: Table view → select Sales → New column → type the formula.","For a measure: use RELATED inside an iterator such as SUMX."],practice:["p40","pa4"],exKind:{0:"column"}}),

/* ===== DAX reference: advanced ===== */
DX("dx_allx","advanced","DAX reference","ALLEXCEPT, ALLSELECTED, REMOVEFILTERS, KEEPFILTERS","<p>Fine control of filters inside CALCULATE.</p>","ALLEXCEPT(Table, Column)  ALLSELECTED(Table)  REMOVEFILTERS(Table)  KEEPFILTERS(condition)",[["ALLEXCEPT","Remove all filters except the listed columns"],["ALLSELECTED","Total of what the user selected in slicers"],["KEEPFILTERS","Add a filter without replacing the existing one"]],[["KEEPFILTERS","North Only = CALCULATE([Total Sales], KEEPFILTERS(Sales[Region] = \"North\"))"],["REMOVEFILTERS","All Sales = CALCULATE([Total Sales], REMOVEFILTERS(Sales))"],["ALLEXCEPT (keep Region)","Region Total = CALCULATE([Total Sales], ALLEXCEPT(Sales, Sales[Region]))"]],{practice:["p25"]}),
DX("dx_var","advanced","DAX reference","VAR and RETURN","<p>Variables store intermediate results: formulas become readable and each part is calculated once.</p>","VAR name = expression\nVAR name2 = expression\nRETURN result",[["VAR","Defines a variable"],["RETURN","The final result"]],[["Discount %","Discount % = VAR gross = SUMX(Sales, Sales[Quantity] * Sales[Price]) VAR disc = SUM(Sales[Discount]) RETURN DIVIDE(disc, gross)"],["Growth vs average","vs Avg = VAR me = [Total Sales] VAR avgAll = AVERAGEX(ALL(Sales[Region]), [Total Sales]) RETURN me - avgAll"]],{practice:["p30","pa6"]}),
DX("dx_rankx","advanced","DAX reference","RANKX","<p>Ranks items by an expression. Use ALL on the column so each item is compared with all others.</p>","RANKX(ALL(Table[Column]), expression, [value], [ASC|DESC], [Dense])",[["ALL(Table[Column])","Items to rank among"],["expression","Usually a measure"],["ASC/DESC","Order (DESC default: highest = 1)"]],[["Product rank","Product Rank = RANKX(ALL(Sales[Product]), [Total Sales])","Product"],["Lowest first","Low Rank = RANKX(ALL(Sales[Region]), [Orders], , ASC)"]],{mistake:{wrong:"Rank = RANKX(Sales, [Total Sales])",why:"Ranks rows of Sales, not products; every product row gets rank 1.",fix:"Rank = RANKX(ALL(Sales[Product]), [Total Sales])"},practice:["p33","p34"]}),
DX("dx_time","advanced","DAX reference","Time intelligence: TOTALYTD, DATEADD, PREVIOUSMONTH","<p>Shift or extend the date filter. In real models use a marked Date table; here the lab uses Sales[Date].</p>","TOTALYTD(expr, Dates, [year_end])  CALCULATE(expr, DATEADD(Dates, -1, MONTH))  CALCULATE(expr, PREVIOUSMONTH(Dates))",[["Dates","Date column of the Date table"],["year_end","\"03-31\" for an April–March financial year"]],[["Year to date","YTD = TOTALYTD([Total Sales], Sales[Date])","Month"],["Last month","Prev Month = CALCULATE([Total Sales], PREVIOUSMONTH(Sales[Date]))","Month"],["MoM %","MoM % = VAR cur = [Total Sales] VAR prev = CALCULATE([Total Sales], DATEADD(Sales[Date], -1, MONTH)) RETURN DIVIDE(cur - prev, prev)","Month"]],{practice:["p35","p36","p37"]}),
DX("dx_running","advanced","DAX reference","Running totals","<p>Keep everything up to the latest visible date using FILTER over ALL dates and MAX.</p>","CALCULATE([m], FILTER(ALL(Dates), Dates <= MAX(Dates)))",[["ALL(Dates)","All dates, ignoring the current filter"],["MAX(Dates)","Latest date in the current filter"]],[["Running sales","Running Sales = CALCULATE([Total Sales], FILTER(ALL(Sales[MonthNum]), Sales[MonthNum] <= MAX(Sales[MonthNum])))","Month"]],{practice:["pa1"]}),
DX("dx_selected","advanced","DAX reference","SELECTEDVALUE and HASONEVALUE","<p>Read the user's selection: dynamic titles, totals that only show on detail rows.</p>","SELECTEDVALUE(Table[Column], [alternate])   HASONEVALUE(Table[Column])",[["alternate","Returned when there is no single value (e.g. totals)"]],[["Dynamic title","Title = \"Region: \" & SELECTEDVALUE(Sales[Region], \"All\")"],["Hide total","Region Sales = IF(HASONEVALUE(Sales[Region]), [Total Sales])"]],{practice:["p31","pa5"]}),
DX("dx_concatx","advanced","DAX reference","VALUES and CONCATENATEX","<p><code>VALUES</code> returns the unique values in the current filter; <code>CONCATENATEX</code> joins them into text.</p>","VALUES(Table[Column])   CONCATENATEX(Table, expression, [delimiter])",[["delimiter","Separator, e.g. \", \""]],[["Products per region","Products = CONCATENATEX(VALUES(Sales[Product]), Sales[Product], \", \")"],["Count of products","Product Count = COUNTROWS(VALUES(Sales[Product]))"]],{practice:["p28","pa3"]}),
DX("dx_format","advanced","DAX reference","FORMAT and text","<p><code>FORMAT</code> converts numbers or dates to text with a format; <code>&amp;</code> joins text. The result is text, so use it for labels.</p>","FORMAT(value, \"format\")",[["format","\"0.0%\", \"#,0\", \"MMM YYYY\""]],[["Share as text","Share Text = FORMAT(DIVIDE([Total Sales], CALCULATE([Total Sales], ALL(Sales))), \"0.0%\")"],["Label","Label = \"Sales: \" & FORMAT([Total Sales], \"#,0\")"]],{practice:["p32"]}),
DX("dx_more","advanced","DAX reference","TOPN, SUMMARIZE, USERELATIONSHIP, SAMEPERIODLASTYEAR","<p>Functions you'll meet in real models. They need features beyond this lab's practice engine, so the examples are shown without results.</p>","TOPN(n, Table, expr)  SUMMARIZE(Table, col, \"Name\", expr)  USERELATIONSHIP(col1, col2)  SAMEPERIODLASTYEAR(Dates)",[["TOPN","Top n rows by an expression (returns a table)"],["SUMMARIZE","Grouped table"],["USERELATIONSHIP","Activates an inactive relationship in CALCULATE"],["SAMEPERIODLASTYEAR","Same dates one year earlier"]],[["Top 3 products' sales","Top 3 = CALCULATE([Total Sales], TOPN(3, ALL(Sales[Product]), [Total Sales]))","Region",false],["Summary table","Summary = SUMMARIZE(Sales, Sales[Region], \"Total\", SUM(Sales[Amount]))","Region",false],["By ship date","By Ship = CALCULATE([Total Sales], USERELATIONSHIP(Sales[ShipDate], 'Date'[Date]))","Region",false],["Last year","LY = CALCULATE([Total Sales], SAMEPERIODLASTYEAR('Date'[Date]))","Region",false]],{})
];
/* ---------- Theory order (W3-style sequence, grouped) ---------- */
const TH_POOL = {}; [...EXCEL_THEORY, ...EXCEL_THEORY_ADV, ...SQL_THEORY, ...SQL_THEORY_ADV, ...PBI_THEORY, ...PBI_THEORY_ADV, ...EXCEL_RICH, ...SQL_RICH, ...PBI_RICH].forEach(t => TH_POOL[t.id] = t);
const TH_ORDER = {
 excel:{basic:[
  ["Getting started",["te1","xg1","xg2","te2","xg3","xg4","xg5","xg6"]],
  ["Formulas",["te3","xf1","xf2","te4","xf3","te5"]],
  ["Formatting",["xm1","xm2","xm3","xm4"]],
  ["Working with data",["xd1","xd2","xd3","xd4","xd5"]],
  ["Conditional formatting",["xc1","xc2","xc3"]],
  ["Charts",["xh1","xh2"]],
  ["PivotTables",["te8","xp1","xp2","xp3"]],
  ["Function reference",["fx_sum","fx_avg","fx_minmax","fx_count","fx_round","fx_if","fx_andor","fx_countif","fx_sumif","fx_averageif","fx_vlookup","te7","fx_text3","fx_lentrim","fx_case","fx_concat","fx_today","fx_ymd","fx_datedif","fx_networkdays"]],
  ["Working smarter",["te10","te11"]],
  ["Case study",["xcase"]]],
 advanced:[
  ["Modern Excel",["tea1","tea2"]],
  ["Function reference",["fx_countifs","fx_avgifs","fx_ifs","fx_switch","fx_iferror","fx_xlookup","fx_indexmatch","fx_sumproduct","fx_rank","fx_findsub","fx_textfn","fx_eomonth","fx_offset","fx_stats","fx_dynamic"]],
  ["Analysis tools",["tea3","tea4","tea5","tea6"]],
  ["Automation",["tea7","tea8"]]]},
 sql:{basic:[
  ["SQL basics",["ts1","qs1","qs2","qs3","qs4","qs5","qs6","qs7","qs8","qs9","ts5"]],
  ["Functions & grouping",["qs10","qs11","qs12","qs13","qs14","ts7","qs15","qs16","qs17","qs18"]],
  ["Joins",["ts6","qs19","qs20","qs21","qs22","qs23"]],
  ["Change data",["ts4","qs24","qs25","qs26"]],
  ["Database design",["ts2","ts3","qs27","qs28","ts9","ts10","ts11","ts12"]]],
 advanced:[
  ["Advanced querying",["qa1","ts8","qa2","qa3","qa4","qa5","tsa1","tsa2"]],
  ["Database objects",["qa6","qa7","qa8","tsa5"]],
  ["Performance",["tsa3","tsa4"]],
  ["Beyond queries",["tsa6","tsa7","tsa8"]]]},
 pbi:{basic:[
  ["Power BI basics",["tp0","p1","pb1","pb2","tp1"]],
  ["Power Query",["p2","pb3","pb4","pb5","p16"]],
  ["Data model",["p4","p18","pb6","pb7"]],
  ["DAX",["p5","pb8","pb9","p19"]],
  ["DAX reference",["dx_sum","dx_count","dx_divide","dx_calc","dx_filter","dx_all","dx_if","dx_sumx","dx_related"]],
  ["Visuals",["p3","pb10","pb11","p17","pb12","pb13"]],
  ["Sharing",["tp2","pb14","tp3","tp4"]]],
 advanced:[
  ["Advanced DAX",["tpa1","tpa2"]],
  ["DAX reference",["dx_allx","dx_var","dx_rankx","dx_time","dx_running","dx_selected","dx_concatx","dx_format","dx_more"]],
  ["Advanced modelling",["tpa3","tpa4"]],
  ["Advanced Power Query",["tpa5"]],
  ["Advanced reports",["tpa6","tpa7","tpa8"]]]}
};
const THEORY_ORDERED = [];
for (const tr of ["excel","sql","pbi"]) for (const lv of ["basic","advanced"]) for (const [group, ids] of TH_ORDER[tr][lv]) for (const id of ids) {
  const t = TH_POOL[id]; if (!t) { console.warn("missing theory", id); continue; }
  const q = (t.questions || []).map(x => Array.isArray(x) ? {q:x[0], opts:x[1], a:x[2], why:x[3]} : x);
  THEORY_ORDERED.push(Object.assign({}, t, {group, level:lv, mode:"theory", track:tr, questions:q}));
}
/* ---------- Formula Library: every card has its own 1–3 boxes ---------- */
/* in: "key:Label:default|key2:Label:default"  f: v => formula  l: how it works  star: must know */
const F = (n, d, inp, f, l = "", star = false) => ({n, d, inp: inp ? inp.split("|").map(s => { const [k, lab, def] = s.split("~"); return {k, lab, def}; }) : [], f, l, star});


/* Bucket builders: "1,3,7,15,30" → 0-1, 2-3, 4-7, 8-15, 16-30, >30 */
function bucketLimits(limits) { return [...new Set(String(limits).split(/[,;\s]+/).map(Number).filter(n => !isNaN(n) && n >= 0))].sort((a, b) => a - b); }
function bucketLabels(L, unit, cont) { const out = []; let prev = null; L.forEach((n, i) => { out.push(i === 0 ? `0-${n}${unit}` : `${cont ? prev : prev + 1}-${n}${unit}`); prev = n; }); out.push(`>${prev}${unit}`); return out; }
function bucketIFS(age, limits, unit = " days", cont = false) { const L = bucketLimits(limits); if (!L.length) return `"Add bucket limits"`; const lab = bucketLabels(L, unit, cont); return `IFS(${L.map((n, i) => `${age}<=${n},"${lab[i]}"`).join(",")},TRUE,"${lab[lab.length - 1]}")`; }
function bucketIF(age, limits, unit = " days", cont = false) { const L = bucketLimits(limits); if (!L.length) return `"Add bucket limits"`; const lab = bucketLabels(L, unit, cont); let f = `"${lab[lab.length - 1]}"`; for (let i = L.length - 1; i >= 0; i--) f = `IF(${age}<=${L[i]},"${lab[i]}",${f})`; return f; }
function bucketSWITCH(age, limits, unit = " days") { const L = bucketLimits(limits); if (!L.length) return `"Add bucket limits"`; const lab = bucketLabels(L, unit, false); return `SWITCH(TRUE(), ${L.map((n, i) => `${age} <= ${n}, "${lab[i]}"`).join(", ")}, "${lab[lab.length - 1]}")`; }
function bucketCASE(age, limits, unit = " days") { const L = bucketLimits(limits); if (!L.length) return `'Add bucket limits'`; const lab = bucketLabels(L, unit, false); return `CASE\n${L.map((n, i) => `    WHEN ${age} <= ${n} THEN '${lab[i]}'`).join("\n")}\n    ELSE '${lab[lab.length - 1]}'\n  END`; }

/* ===================== EXCEL ===================== */
const LIB_EXCEL = [
{cat:"Math & totals", items:[
 F("SUM","Adds all numbers in a range.","r~Range~B2:B100",v=>`=SUM(${v.r})`,"Select the column of numbers.",true),
 F("AVERAGE","Average of a range.","r~Range~B2:B100",v=>`=AVERAGE(${v.r})`),
 F("MIN","Smallest number.","r~Range~B2:B100",v=>`=MIN(${v.r})`),
 F("MAX","Largest number.","r~Range~B2:B100",v=>`=MAX(${v.r})`),
 F("COUNT","Counts cells that contain numbers.","r~Range~B2:B100",v=>`=COUNT(${v.r})`),
 F("COUNTA","Counts non-empty cells (total rows/tickets).","r~Range~A2:A100",v=>`=COUNTA(${v.r})`,"",true),
 F("COUNTBLANK","Counts empty cells (missing data).","r~Range~C2:C100",v=>`=COUNTBLANK(${v.r})`),
 F("MEDIAN","Middle value; not pulled by outliers.","r~Range~B2:B100",v=>`=MEDIAN(${v.r})`),
 F("LARGE (2nd highest)","k-th largest value.","r~Range~B2:B100|k~Position (k)~2",v=>`=LARGE(${v.r},${v.k})`),
 F("SMALL (2nd lowest)","k-th smallest value.","r~Range~B2:B100|k~Position (k)~2",v=>`=SMALL(${v.r},${v.k})`),
 F("RANK","Rank of a value in a list (1 = highest).","c~Value cell~B2|r~List range~$B$2:$B$100",v=>`=RANK(${v.c},${v.r})`,"Lock the list with $ so it stays fixed when you fill down."),
 F("ROUND","Round to n decimals.","c~Number cell~B2|n~Decimals~2",v=>`=ROUND(${v.c},${v.n})`),
 F("ROUNDUP / ROUNDDOWN","Always round up (or down).","c~Number cell~B2|n~Decimals~0",v=>`=ROUNDUP(${v.c},${v.n})`,"Use ROUNDDOWN for the opposite."),
 F("INT","Drop the decimals.","c~Number cell~B2",v=>`=INT(${v.c})`),
 F("ABS","Remove the minus sign.","c~Number cell~B2",v=>`=ABS(${v.c})`),
 F("MOD (remainder)","Remainder after dividing.","c~Number cell~B2|n~Divide by~2",v=>`=MOD(${v.c},${v.n})`,"MOD(row,2)=0 finds even rows."),
 F("Percentage of total","This row's share of the total.","c~Value cell~B2|r~Total range~$B$2:$B$100",v=>`=${v.c}/SUM(${v.r})`,"Format the cell as %.",true),
 F("Percentage change","Growth from old to new value.","o~Old value cell~B2|n~New value cell~C2",v=>`=(${v.n}-${v.o})/${v.o}`,"Format as %. Negative = drop."),
 F("Running total","Cumulative sum down the rows.","c~First value cell~B2",v=>{const m=v.c.match(/^\$?([A-Za-z]+)\$?(\d+)$/);return m?`=SUM($${m[1].toUpperCase()}$${m[2]}:${m[1].toUpperCase()}${m[2]})`:`=SUM(${v.c})`},"The first cell is locked, the end moves as you fill down."),
 F("SUMPRODUCT","Multiply two columns row by row and add (e.g. qty × price).","a~First range~B2:B100|b~Second range~C2:C100",v=>`=SUMPRODUCT(${v.a},${v.b})`),
 F("Weighted average","Average weighted by quantity.","a~Values range~B2:B100|w~Weights range~C2:C100",v=>`=SUMPRODUCT(${v.a},${v.w})/SUM(${v.w})`),
 F("SUBTOTAL (visible rows only)","Total that ignores rows hidden by a filter.","r~Range~B2:B100",v=>`=SUBTOTAL(9,${v.r})`,"9 = SUM, 1 = AVERAGE, 3 = COUNTA.")
]},
{cat:"Date & time", items:[
 F("TODAY","Today's date.","",v=>`=TODAY()`,"Updates every day."),
 F("NOW","Current date and time.","",v=>`=NOW()`),
 F("DATE","Build a date from year, month, day.","y~Year~2026|m~Month~1|d~Day~1",v=>`=DATE(${v.y},${v.m},${v.d})`),
 F("DAY","Day number (1–31).","c~Date cell~A2",v=>`=DAY(${v.c})`),
 F("MONTH","Month number (1–12).","c~Date cell~A2",v=>`=MONTH(${v.c})`),
 F("YEAR","Year of a date.","c~Date cell~A2",v=>`=YEAR(${v.c})`),
 F("Day name","Full day name (Monday).","c~Date cell~A2",v=>`=TEXT(${v.c},"dddd")`),
 F("Short day name","Mon, Tue…","c~Date cell~A2",v=>`=TEXT(${v.c},"ddd")`),
 F("Month name","January, February…","c~Date cell~A2",v=>`=TEXT(${v.c},"mmmm")`),
 F("Month label (Jan-26)","For monthly pivots and charts.","c~Date cell~A2",v=>`=TEXT(${v.c},"mmm-yy")`,"",true),
 F("WEEKDAY","Day of week number (Mon = 1).","c~Date cell~A2",v=>`=WEEKDAY(${v.c},2)`,"With ,2 Monday = 1 … Sunday = 7."),
 F("WEEKNUM","Week number from 1 January.","c~Date cell~A2",v=>`=WEEKNUM(${v.c})`),
 F("ISOWEEKNUM","International week number (Monday start).","c~Date cell~A2",v=>`=ISOWEEKNUM(${v.c})`),
 F("Quarter","Calendar quarter Q1–Q4.","c~Date cell~A2",v=>`="Q"&ROUNDUP(MONTH(${v.c})/3,0)`),
 F("Indian financial year","FY label like 2026-27 (April–March).","c~Date cell~A2",v=>`=IF(MONTH(${v.c})>=4,YEAR(${v.c})&"-"&RIGHT(YEAR(${v.c})+1,2),YEAR(${v.c})-1&"-"&RIGHT(YEAR(${v.c}),2))`,"April or later belongs to the FY starting this year."),
 F("Indian FY quarter","Apr–Jun = Q1 … Jan–Mar = Q4.","c~Date cell~A2",v=>`="Q"&CHOOSE(MONTH(${v.c}),4,4,4,1,1,1,2,2,2,3,3,3)`),
 F("EDATE","Same date n months later.","c~Date cell~A2|n~Months~1",v=>`=EDATE(${v.c},${v.n})`),
 F("EOMONTH (month end)","Last date of the month.","c~Date cell~A2",v=>`=EOMONTH(${v.c},0)`),
 F("First day of month","1st of the same month.","c~Date cell~A2",v=>`=EOMONTH(${v.c},-1)+1`),
 F("WORKDAY","Date after n working days.","c~Start date cell~A2|n~Working days~5",v=>`=WORKDAY(${v.c},${v.n})`,"Skips Saturday and Sunday."),
 F("NETWORKDAYS","Working days between two dates.","a~Start date cell~A2|b~End date cell~B2",v=>`=NETWORKDAYS(${v.a},${v.b})`,"Counts both start and end day."),
 F("NETWORKDAYS.INTL (6-day week)","Working days with only Sunday off.","a~Start date cell~A2|b~End date cell~B2",v=>`=NETWORKDAYS.INTL(${v.a},${v.b},"0000001")`,"Code Mon→Sun; 1 = off. \"0000011\" = Sat+Sun off."),
 F("WORKDAY.INTL","Add working days with your own weekend.","c~Start date cell~A2|n~Working days~5",v=>`=WORKDAY.INTL(${v.c},${v.n},"0000011")`),
 F("SEQUENCE of dates","7 dates starting from a date (Excel 365).","c~Start date cell~A2|n~How many days~7",v=>`=SEQUENCE(${v.n},1,${v.c},1)`),
 F("DATEVALUE","Text that looks like a date → real date.","c~Text date cell~A2",v=>`=DATEVALUE(${v.c})`),
 F("TIMEVALUE","Text time → real time.","c~Text time cell~A2",v=>`=TIMEVALUE(${v.c})`),
 F("TIME","Build a time.","h~Hours~14|m~Minutes~30",v=>`=TIME(${v.h},${v.m},0)`),
 F("HOUR / MINUTE","Hour (0–23) of a date-time.","c~Date-time cell~A2",v=>`=HOUR(${v.c})`,"Use MINUTE(…) for minutes."),
 F("Date part only","Remove the time from a date-time.","c~Date-time cell~A2",v=>`=INT(${v.c})`,"Format the cell as a date."),
 F("Time part only","Keep only the time.","c~Date-time cell~A2",v=>`=MOD(${v.c},1)`,"Format the cell as time."),
 F("Combine date + time","Join a date column and a time column.","a~Date cell~A2|b~Time cell~B2",v=>`=${v.a}+${v.b}`),
 F("DATEDIF (days)","Complete days between two dates.","a~Start date cell~A2|b~End date cell~B2",v=>`=DATEDIF(${v.a},${v.b},"D")`,"Start must be earlier."),
 F("DATEDIF (months)","Complete months between two dates.","a~Start date cell~A2|b~End date cell~B2",v=>`=DATEDIF(${v.a},${v.b},"M")`),
 F("DATEDIF (years)","Complete years (age, tenure).","a~Start date cell~A2|b~End date cell~B2",v=>`=DATEDIF(${v.a},${v.b},"Y")`),
 F("DAYS","Days between two dates.","a~Start date cell~A2|b~End date cell~B2",v=>`=DAYS(${v.b},${v.a})`),
 F("YEARFRAC","Years with decimals.","a~Start date cell~A2|b~End date cell~B2",v=>`=YEARFRAC(${v.a},${v.b})`),
 F("Age from date of birth","Age in years today.","c~Birth date cell~A2",v=>`=DATEDIF(${v.c},TODAY(),"Y")`)
]},
{cat:"Logical", items:[
 F("IF","One test, two results.","c~Cell~A2|n~Is greater than~10",v=>`=IF(${v.c}>${v.n},"Yes","No")`,"",true),
 F("IFS","Several tests in order; TRUE = otherwise.","c~Score cell~A2",v=>`=IFS(${v.c}>90,"A",${v.c}>75,"B",TRUE,"C")`),
 F("Nested IF","IF inside IF for older Excel.","c~Score cell~A2",v=>`=IF(${v.c}>=90,"A",IF(${v.c}>=75,"B","C"))`),
 F("AND","TRUE when all conditions are TRUE.","c~Cell~A2",v=>`=AND(${v.c}>10,${v.c}<20)`),
 F("OR","TRUE when any condition is TRUE.","c~Cell~A2",v=>`=OR(${v.c}>10,${v.c}<5)`),
 F("NOT","Reverse TRUE/FALSE.","c~Cell~A2",v=>`=NOT(${v.c}>10)`),
 F("XOR","TRUE when exactly one is TRUE.","a~First cell~A2|b~Second cell~B2",v=>`=XOR(${v.a}>10,${v.b}>10)`),
 F("IF with AND","Two conditions must both be met.","a~Status cell~A2|b~Amount cell~B2",v=>`=IF(AND(${v.a}="Open",${v.b}>1000),"Check","OK")`),
 F("SWITCH","Map fixed values to results.","c~Code cell~A2",v=>`=SWITCH(${v.c},1,"Low",2,"Medium",3,"High","Unknown")`,"Last value is the default."),
 F("CHOOSE","Pick the n-th item.","c~Number cell~A2",v=>`=CHOOSE(${v.c},"Red","Blue","Green")`),
 F("ISBLANK","TRUE if the cell is empty.","c~Cell~A2",v=>`=ISBLANK(${v.c})`),
 F("ISNUMBER","TRUE if the cell holds a number.","c~Cell~A2",v=>`=ISNUMBER(${v.c})`),
 F("ISTEXT","TRUE if the cell holds text.","c~Cell~A2",v=>`=ISTEXT(${v.c})`),
 F("EXACT","Case-sensitive match.","a~First cell~A2|b~Second cell~B2",v=>`=EXACT(${v.a},${v.b})`),
 F("If blank show text","Friendly text for empty cells.","c~Cell~A2|t~Text to show~Not updated",v=>`=IF(${v.c}="","${v.t}",${v.c})`),
 F("If cell contains text","Yes when a word appears anywhere.","c~Cell~A2|t~Word~VPN",v=>`=IF(ISNUMBER(SEARCH("${v.t}",${v.c})),"Yes","No")`,"SEARCH ignores capitals.",true)
]},
{cat:"Lookup", items:[
 F("VLOOKUP","Find a key in the first column and return another column.","k~Lookup value cell~A2|r~Table range~Sheet2!A:B|n~Column number~2",v=>`=VLOOKUP(${v.k},${v.r},${v.n},FALSE)`,"FALSE = exact match.",true),
 F("Dynamic VLOOKUP (column by header)","Dynamic VLOOKUP: the column number is found from the header name, so it never breaks when columns move.","k~Lookup value cell~$A2|r~Table range~Sheet2!$A:$F|h~Header cell to return~B$1",v=>{const m=String(v.r).match(/^(.*!)?/);const p=m&&m[1]?m[1]:"";return `=VLOOKUP(${v.k},${v.r},MATCH(${v.h},${p}$1:$1,0),FALSE)`;},"MATCH finds which column of the table has the same header as your cell, and passes that number to VLOOKUP. Put the same headers in row 1 of both sheets, then fill right and down: one formula fills the whole report.",true),
 F("VLOOKUP many columns {2,3,4}","Dynamic VLOOKUP that returns several columns at once with an array {2,3,4}.","k~Lookup value cell~A2|r~Table range~Sheet2!A:F|c~Column numbers~2,3,4",v=>`=VLOOKUP(${v.k},${v.r},{${String(v.c).split(/[,;\s]+/).filter(Boolean).join(",")}},FALSE)`,"{2,3,4} is an array constant: VLOOKUP runs once for each number and the answers spill into the cells to the right. Excel 365/2021. In older Excel, select 3 cells first and press Ctrl+Shift+Enter. If your Excel uses ; between arguments, the numbers inside { } may need a backslash or full stop instead of commas."),
 F("SUM of VLOOKUP {2,3,4}","Add up several columns of the matched row (e.g. Jan+Feb+Mar).","k~Lookup value cell~A2|r~Table range~Sheet2!A:F|c~Column numbers~2,3,4",v=>`=SUM(VLOOKUP(${v.k},${v.r},{${String(v.c).split(/[,;\s]+/).filter(Boolean).join(",")}},FALSE))`,"VLOOKUP with {2,3,4} returns three values; SUM adds them. Works in every Excel version."),
 F("VLOOKUP with COLUMN() (fill right)","Column number grows by itself as you drag the formula to the right.","k~Lookup value cell~$A2|r~Table range~Sheet2!$A:$F|s~Start cell (column of the first value)~B1",v=>`=VLOOKUP(${v.k},${v.r},COLUMN(${v.s}),FALSE)`,"COLUMN(B1) = 2. Dragged one cell right it becomes COLUMN(C1) = 3, and so on. Lock the lookup column ($A2) and the table with $."),
 F("VLOOKUP to the left (CHOOSE {1,2})","Return a column that sits left of the search column.","k~Lookup value cell~A2|s~Search column~Sheet2!C:C|t~Return column~Sheet2!A:A",v=>`=VLOOKUP(${v.k},CHOOSE({1,2},${v.s},${v.t}),2,FALSE)`,"CHOOSE({1,2},…) builds a small 2-column table with the search column first, so VLOOKUP can look left. XLOOKUP or INDEX + MATCH do the same more simply."),
 F("XLOOKUP many columns","Return a whole block of columns for the match.","k~Lookup value cell~A2|s~Search column~Sheet2!A:A|t~Return columns~Sheet2!B:D",v=>`=XLOOKUP(${v.k},${v.s},${v.t},"Not Found")`,"Give a return range that is several columns wide; the results spill to the right. Excel 365/2021."),
 F("VLOOKUP approximate (slabs)","Grade or slab from a sorted table.","k~Value cell~A2|r~Slab table~$E$2:$F$5",v=>`=VLOOKUP(${v.k},${v.r},2,TRUE)`,"First column of the slab table must be sorted ascending."),
 F("HLOOKUP","Search the top row instead.","k~Lookup value cell~A2|r~Table range~Sheet2!A1:D5|n~Row number~2",v=>`=HLOOKUP(${v.k},${v.r},${v.n},FALSE)`),
 F("XLOOKUP","Modern lookup, any direction.","k~Lookup value cell~A2|s~Search column~Sheet2!A:A|t~Return column~Sheet2!B:B",v=>`=XLOOKUP(${v.k},${v.s},${v.t},"Not Found")`,"Excel 2021/365.",true),
 F("MATCH","Position of a value in a list.","k~Lookup value cell~A2|s~Search column~Sheet2!A:A",v=>`=MATCH(${v.k},${v.s},0)`),
 F("INDEX","Value at a position.","r~Column range~B:B|n~Row number~5",v=>`=INDEX(${v.r},${v.n})`),
 F("INDEX + MATCH","Lookup that works left or right.","k~Lookup value cell~A2|s~Search column~Sheet2!A:A|t~Return column~Sheet2!B:B",v=>`=INDEX(${v.t},MATCH(${v.k},${v.s},0))`),
 F("Two-way lookup","Value at a row label and a column header.","r~Whole table~Sheet2!A1:F50|a~Row value cell~A2|b~Column header cell~B1",v=>{const m=v.r.match(/^(.*!)?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i);const p=m?(m[1]||""):"";return m?`=INDEX(${v.r},MATCH(${v.a},${p}${m[2]}${m[3]}:${m[2]}${m[5]},0),MATCH(${v.b},${p}${m[2]}${m[3]}:${m[4]}${m[3]},0))`:`=INDEX(${v.r},MATCH(${v.a},…,0),MATCH(${v.b},…,0))`},"Row MATCH on the first column, column MATCH on the header row."),
 F("Lookup with Not Found","No #N/A in your report.","k~Lookup value cell~A2|r~Table range~Sheet2!A:B|n~Column number~2",v=>`=IFNA(VLOOKUP(${v.k},${v.r},${v.n},FALSE),"Not Found")`),
 F("Last value in a column","Latest entry in a list.","r~Column~A:A",v=>`=LOOKUP(2,1/(${v.r}<>""),${v.r})`),
 F("Lookup with wildcard","Find the first match that starts with text.","t~Starts with~Ar|r~Table range~Sheet2!A:B",v=>`=VLOOKUP("${v.t}*",${v.r},2,FALSE)`)
]},
{cat:"Count & sum with conditions", items:[
 F("COUNTIF","Count cells equal to a value.","r~Range~E2:E100|t~Value~Open",v=>`=COUNTIF(${v.r},"${v.t}")`,"",true),
 F("COUNTIF contains","Count cells containing a word.","r~Range~E2:E100|t~Word~VPN",v=>`=COUNTIF(${v.r},"*${v.t}*")`,"* = any characters."),
 F("COUNTIF greater than","Count numbers above a limit.","r~Range~B2:B100|n~Limit~50000",v=>`=COUNTIF(${v.r},">${v.n}")`),
 F("COUNTIF vs a cell","Rule taken from a cell (interactive).","r~Range~B2:B100|c~Limit cell~$H$1",v=>`=COUNTIF(${v.r},">"&${v.c})`),
 F("COUNTIFS","Count rows matching two rules.","a~First range~E2:E100|b~Second range~F2:F100",v=>`=COUNTIFS(${v.a},"Open",${v.b},"P1")`),
 F("Count OR","Count rows matching either value.","r~Range~C2:C100",v=>`=COUNTIF(${v.r},"Pune")+COUNTIF(${v.r},"Delhi")`),
 F("Count between dates","Rows with a date in a range.","r~Date range~A2:A100|a~From date cell~$H$1|b~To date cell~$H$2",v=>`=COUNTIFS(${v.r},">="&${v.a},${v.r},"<="&${v.b})`),
 F("Count unique values","Number of distinct entries.","r~Range~A2:A100",v=>`=SUMPRODUCT(1/COUNTIF(${v.r},${v.r}))`,"Range must have no blanks. In Excel 365: =COUNTA(UNIQUE(range))."),
 F("SUMIF","Sum values for one rule.","r~Check range~C2:C100|t~Value~East|s~Sum range~D2:D100",v=>`=SUMIF(${v.r},"${v.t}",${v.s})`,"",true),
 F("SUMIFS","Sum with several rules.","s~Sum range~D2:D100|a~Rule range 1~C2:C100|b~Rule range 2~E2:E100",v=>`=SUMIFS(${v.s},${v.a},"East",${v.b},"Closed")`),
 F("Sum between dates","Total in a date window.","s~Sum range~D2:D100|r~Date range~A2:A100|a~From date cell~$H$1",v=>`=SUMIFS(${v.s},${v.r},">="&${v.a},${v.r},"<"&EDATE(${v.a},1))`,"Sums one month starting at the From date."),
 F("AVERAGEIF","Average for one rule.","r~Check range~C2:C100|t~Value~East|s~Average range~D2:D100",v=>`=AVERAGEIF(${v.r},"${v.t}",${v.s})`),
 F("AVERAGEIFS","Average with several rules.","s~Average range~D2:D100|a~Rule range 1~C2:C100|b~Rule range 2~E2:E100",v=>`=AVERAGEIFS(${v.s},${v.a},"East",${v.b},"Open")`),
 F("MAXIFS","Largest value with a rule.","s~Max range~D2:D100|r~Rule range~C2:C100|t~Value~East",v=>`=MAXIFS(${v.s},${v.r},"${v.t}")`),
 F("MINIFS","Smallest value with a rule.","s~Min range~D2:D100|r~Rule range~C2:C100|t~Value~East",v=>`=MINIFS(${v.s},${v.r},"${v.t}")`)
]},
{cat:"Text", items:[
 F("Join with &","Join two cells with a separator.","a~First cell~A2|b~Second cell~B2",v=>`=${v.a}&" - "&${v.b}`),
 F("TEXTJOIN","Join a range with commas.","r~Range~A2:A10",v=>`=TEXTJOIN(", ",TRUE,${v.r})`,"TRUE skips empty cells."),
 F("LEFT","First n characters.","c~Text cell~A2|n~Characters~3",v=>`=LEFT(${v.c},${v.n})`),
 F("RIGHT","Last n characters.","c~Text cell~A2|n~Characters~4",v=>`=RIGHT(${v.c},${v.n})`),
 F("MID","Characters from the middle.","c~Text cell~A2|s~Start position~2|n~Characters~3",v=>`=MID(${v.c},${v.s},${v.n})`),
 F("LEN","Number of characters.","c~Text cell~A2",v=>`=LEN(${v.c})`),
 F("UPPER","ALL CAPS.","c~Text cell~A2",v=>`=UPPER(${v.c})`),
 F("LOWER","all small letters.","c~Text cell~A2",v=>`=LOWER(${v.c})`),
 F("PROPER","Title Case.","c~Text cell~A2",v=>`=PROPER(${v.c})`),
 F("TRIM","Remove extra spaces.","c~Text cell~A2",v=>`=TRIM(${v.c})`,"Fixes most lookup mismatches.",true),
 F("CLEAN","Remove hidden non-printing characters.","c~Text cell~A2",v=>`=CLEAN(TRIM(${v.c}))`),
 F("SUBSTITUTE","Replace a word.","c~Text cell~A2|o~Old text~-|n~New text~ ",v=>`=SUBSTITUTE(${v.c},"${v.o}","${v.n}")`),
 F("REPLACE","Replace by position.","c~Text cell~A2|s~Start~1|n~Characters~3",v=>`=REPLACE(${v.c},${v.s},${v.n},"XXX")`),
 F("FIND","Position of text (case-sensitive).","c~Text cell~A2|t~Find~@",v=>`=FIND("${v.t}",${v.c})`),
 F("SEARCH","Position of text (any case).","c~Text cell~A2|t~Find~vpn",v=>`=SEARCH("${v.t}",${v.c})`),
 F("Text before a character","Everything before the first dash.","c~Text cell~A2|t~Character~-",v=>`=IFERROR(LEFT(${v.c},FIND("${v.t}",${v.c})-1),${v.c})`),
 F("Text after a character","Everything after the first dash.","c~Text cell~A2|t~Character~-",v=>`=IFERROR(MID(${v.c},FIND("${v.t}",${v.c})+1,100),"")`),
 F("First name","Word before the first space.","c~Full name cell~A2",v=>`=IFERROR(LEFT(${v.c},FIND(" ",${v.c})-1),${v.c})`),
 F("Last name","Word after the first space.","c~Full name cell~A2",v=>`=IFERROR(MID(${v.c},FIND(" ",${v.c})+1,100),"")`),
 F("Email domain","Text after @.","c~Email cell~A2",v=>`=MID(${v.c},FIND("@",${v.c})+1,100)`),
 F("Count words","Words in a sentence.","c~Text cell~A2",v=>`=IF(TRIM(${v.c})="",0,LEN(TRIM(${v.c}))-LEN(SUBSTITUTE(${v.c}," ",""))+1)`),
 F("TEXT (number format)","Number as formatted text.","c~Number cell~A2",v=>`=TEXT(${v.c},"#,##0.00")`),
 F("Leading zeros","5 → 00005 (IDs, PIN codes).","c~Number cell~A2|n~Total digits~5",v=>`=TEXT(${v.c},REPT("0",${v.n}))`),
 F("VALUE","Text number → real number.","c~Text cell~A2",v=>`=VALUE(${v.c})`)
]},
{cat:"Ageing", items:[
 F("Age (24×7, days)","Calendar days since created, nights and weekends included.","c~Created date cell~A2",v=>`=INT(NOW()-${v.c})`,"NOW − created gives days with decimals; INT keeps completed days.",true),
 F("Age (business days)","Working days since created.","c~Created date cell~A2",v=>`=NETWORKDAYS(${v.c},NOW())-1`,"−1 removes the start day."),
 F("Basic ageing (TODAY − date)","Days pending.","c~Created date cell~A2",v=>`=TODAY()-${v.c}`),
 F("Exact ageing (DATEDIF)","Completed days until today.","c~Created date cell~A2",v=>`=DATEDIF(${v.c},TODAY(),"D")`),
 F("Months ageing (approx.)","Days ÷ 30.","c~Created date cell~A2",v=>`=INT((TODAY()-${v.c})/30)`),
 F("Age in years","Completed years.","c~Date cell~A2",v=>`=DATEDIF(${v.c},TODAY(),"Y")`),
 F("Financial ageing (YEARFRAC)","Years with decimals.","c~Date cell~A2",v=>`=YEARFRAC(${v.c},TODAY())`),
 F("Days till month end","Days left in the month.","c~Date cell~A2",v=>`=EOMONTH(${v.c},0)-${v.c}`),
 F("Reopen age","Age since the last reopen.","c~Reopen date cell~C2",v=>`=INT(NOW()-${v.c})`),
 F("Age bucket (IFS)","0–1, 2–3, 4–7, >7 days.","c~Age cell~C2",v=>`=IFS(${v.c}<=1,"0–1",${v.c}<=3,"2–3",${v.c}<=7,"4–7",TRUE,">7")`,"Smallest range first.",true),
 F("Age bucket (nested IF)","0-30, 31-60, 60+.","c~Age cell~C2",v=>`=IF(${v.c}<=30,"0-30",IF(${v.c}<=60,"31-60","60+"))`),
 F("Dynamic bucket (LOOKUP)","Bucket from two lists.","c~Age cell~C2",v=>`=LOOKUP(${v.c},{0,31,61},{"0-30","31-60","60+"})`,"Add more boundaries to both lists."),
 F("Age bucket from date directly","No separate age column needed.","c~Created date cell~A2",v=>`=IFS(TODAY()-${v.c}<=7,"0-7",TODAY()-${v.c}<=30,"8-30",TODAY()-${v.c}<=90,"31-90",TRUE,"90+")`),
 F("Overdue days","Days past due (0 if not overdue).","c~Due date cell~D2",v=>`=MAX(0,TODAY()-${v.c})`)
]},
 {cat:"Open ageing (one formula)", items:[
 F("Open age (days)","Today minus created date: how many days the item has been open.","c~Created date cell~A2",v=>`=TODAY()-${v.c}`,"TODAY() changes every day, so the age updates by itself. Format the cell as Number, not Date. If the created cell also has a time, use the date-time version below.",true),
 F("Open age (blank-safe)","Same, but stays empty when the created date is missing.","c~Created date cell~A2",v=>`=IF(${v.c}="","",TODAY()-${v.c})`),
 F("Open age from date-time (24×7)","Created cell has a time too: completed days till now.","c~Created date-time cell~A2",v=>`=IF(${v.c}="","",INT(NOW()-${v.c}))`),
 F("Open age bucket in one formula","Age and bucket together: no helper age column needed. Type your own bucket limits.","c~Created date cell~A2|b~Bucket limits (days)~1,3,7,15,30",v=>`=IF(${v.c}="","",${bucketIFS(`TODAY()-${v.c}`, v.b)})`,"Limits 1,3,7,15,30 give 0-1, 2-3, 4-7, 8-15, 16-30 and >30 days. Change the limits and the formula rebuilds. Works in Excel 2019 and later (IFS).",true),
 F("Open age bucket (any Excel version)","Same one-formula bucket with nested IF: works in every Excel version.","c~Created date cell~A2|b~Bucket limits (days)~1,3,7,15,30",v=>`=IF(${v.c}="","",${bucketIF(`TODAY()-${v.c}`, v.b)})`,"Nested IF checks from the smallest limit; the last label catches the rest."),
 F("Open age bucket with LET (Excel 365)","Calculates the age once and reuses it: shortest and fastest.","c~Created date cell~A2|b~Bucket limits (days)~1,3,7,15,30",v=>`=IF(${v.c}="","",LET(age,TODAY()-${v.c},${bucketIFS("age", v.b)}))`,"LET names the age once, so TODAY()-created isn't repeated in every test. Excel 365 / 2021 only."),
 F("Bucket only open items (by status)","Closed rows show \"Closed\"; open rows get an age bucket.","s~Status cell~E2|c~Created date cell~A2|b~Bucket limits (days)~1,3,7,15,30",v=>`=IF(${v.s}="Closed","Closed",IF(${v.c}="","",${bucketIFS(`TODAY()-${v.c}`, v.b)}))`,"Change \"Closed\" to the exact closed status text in your data (e.g. \"Resolved\").",true),
 F("Bucket only open items (by resolved date)","If a resolved date exists the item is closed; otherwise bucket its open age.","r~Resolved date cell~B2|c~Created date cell~A2|b~Bucket limits (days)~1,3,7,15,30",v=>`=IF(${v.r}<>"","Closed",IF(${v.c}="","",${bucketIFS(`TODAY()-${v.c}`, v.b)}))`),
 F("Open age bucket in business days","Counts only Mon–Fri, then buckets.","c~Created date cell~A2|b~Bucket limits (working days)~1,3,5,10",v=>`=IF(${v.c}="","",${bucketIFS(`(NETWORKDAYS(${v.c},TODAY())-1)`, v.b, " working days")}))`.replace(/\)\)$/,")"),"NETWORKDAYS counts both start and end day, so −1 gives completed working days."),
 F("Open age bucket in hours (24×7)","For hour-based SLAs: buckets like 0-4, 4-8, 8-24 hrs.","c~Created date-time cell~A2|b~Bucket limits (hours)~4,8,24,48",v=>`=IF(${v.c}="","",${bucketIFS(`(NOW()-${v.c})*24`, v.b, " hrs", true)})`,"(NOW − created) × 24 = hours open. Labels like 4-8 hrs mean more than 4 up to 8."),
 F("Count open items in an age range","Summary table without a helper column: items created between X and Y days ago.","r~Created date range~$A$2:$A$1000|a~From (days)~4|z~To (days)~7",v=>`=COUNTIFS(${v.r},"<="&TODAY()-${v.a},${v.r},">="&TODAY()-${v.z})`,"Example counts items aged 4 to 7 days. Build one row per bucket for a summary."),
 F("Count open items in an age range (status)","Same, counting only open items.","r~Created date range~$A$2:$A$1000|s~Status range~$E$2:$E$1000|a~From (days)~4|z~To (days)~7",v=>`=COUNTIFS(${v.r},"<="&TODAY()-${v.a},${v.r},">="&TODAY()-${v.z},${v.s},"<>Closed")`),
 F("Count items older than N days","Everything still open beyond the last bucket.","r~Created date range~$A$2:$A$1000|n~Older than (days)~30",v=>`=COUNTIFS(${v.r},"<"&TODAY()-${v.n})`),
 F("Bucket sort number","Number 1, 2, 3… per bucket so pivots sort buckets in the right order.","c~Created date cell~A2|b~Bucket limits (days)~1,3,7,15,30",v=>{const L=bucketLimits(v.b);return `=IF(${v.c}="","",IFS(${L.map((n,i)=>`TODAY()-${v.c}<=${n},${i+1}`).join(",")},TRUE,${L.length+1}))`},"Put this next to the bucket column and sort the pivot by it (or use a custom list).")
]},
{cat:"SLA", items:[
 F("Age in hours (24×7)","Hours since created.","c~Created date-time cell~A2",v=>`=ROUND((NOW()-${v.c})*24,2)`),
 F("Age in minutes (24×7)","Minutes since created.","c~Created date-time cell~A2",v=>`=ROUND((NOW()-${v.c})*1440,0)`),
 F("Age in business hours (approx.)","Working days × hours per day.","c~Created date-time cell~A2|h~Hours per day~9",v=>`=(NETWORKDAYS(${v.c},NOW())-1)*${v.h}`),
 F("SLA target date","Due after n working days.","c~Created date cell~A2|n~SLA working days~3",v=>`=WORKDAY(${v.c},${v.n})`,"",true),
 F("SLA due time (hours, 24×7)","Created + n hours.","c~Created date-time cell~A2|h~SLA hours~4",v=>`=${v.c}+${v.h}/24`),
 F("SLA status (now)","Breached or within SLA.","c~Due date-time cell~D2",v=>`=IF(NOW()>${v.c},"Breached","Within SLA")`,"",true),
 F("SLA breach check (date)","Same check using dates only.","c~Due date cell~D2",v=>`=IF(TODAY()>${v.c},"Breached","Within SLA")`),
 F("Remaining SLA time","Days left (negative = overdue).","c~Due date-time cell~D2",v=>`=${v.c}-NOW()`),
 F("Remaining SLA hours","Hours left.","c~Due date-time cell~D2",v=>`=ROUND((${v.c}-NOW())*24,1)`),
 F("Resolution time (days)","Created → closed in days.","a~Created cell~A2|b~Closed cell~B2",v=>`=INT(${v.b}-${v.a})`),
 F("Resolution time (hours)","Created → closed in hours.","a~Created cell~A2|b~Closed cell~B2",v=>`=ROUND((${v.b}-${v.a})*24,2)`,"",true),
 F("Resolution time (minutes)","Created → closed in minutes.","a~Created cell~A2|b~Closed cell~B2",v=>`=ROUND((${v.b}-${v.a})*1440,0)`),
 F("Met or breached (closed)","Compare closed date with due.","a~Closed cell~B2|b~Due cell~D2",v=>`=IF(${v.a}<=${v.b},"Met","Breached")`),
 F("Escalation flag","Escalate when 1 day or less is left.","c~Due date cell~D2",v=>`=IF((${v.c}-TODAY())<=1,"Escalate","OK")`),
 F("SLA %","Share within SLA.","r~Status range~F2:F100",v=>`=COUNTIF(${v.r},"Within SLA")/COUNTA(${v.r})`,"Format as %.",true),
 F("SLA met % (Met/Breached)","Met ÷ (Met + Breached).","r~SLA result range~F2:F100",v=>`=COUNTIF(${v.r},"Met")/(COUNTIF(${v.r},"Met")+COUNTIF(${v.r},"Breached"))`),
 F("Uptime %","Uptime from downtime minutes in a month.","d~Downtime minutes cell~B2|m~Days in month~30",v=>`=1-${v.d}/(${v.m}*1440)`,"Format as % with 2–3 decimals.")
]},
{cat:"Time conversion", items:[
 F("Days → hours","1 day = 24 hours.","c~Days cell~A2",v=>`=${v.c}*24`),
 F("Days → minutes","× 1440.","c~Days cell~A2",v=>`=${v.c}*1440`),
 F("Days → seconds","× 86400.","c~Days cell~A2",v=>`=${v.c}*86400`),
 F("Hours → days","÷ 24.","c~Hours cell~A2",v=>`=${v.c}/24`),
 F("Hours → minutes","× 60.","c~Hours cell~A2",v=>`=${v.c}*60`),
 F("Hours → seconds","× 3600.","c~Hours cell~A2",v=>`=${v.c}*3600`),
 F("Minutes → hours","÷ 60.","c~Minutes cell~A2",v=>`=${v.c}/60`),
 F("Minutes → days","÷ 1440.","c~Minutes cell~A2",v=>`=${v.c}/1440`),
 F("Minutes → seconds","× 60.","c~Minutes cell~A2",v=>`=${v.c}*60`),
 F("Seconds → minutes","÷ 60.","c~Seconds cell~A2",v=>`=${v.c}/60`),
 F("Seconds → hours","÷ 3600.","c~Seconds cell~A2",v=>`=${v.c}/3600`),
 F("Seconds → days","÷ 86400.","c~Seconds cell~A2",v=>`=${v.c}/86400`),
 F("Time difference in hours","End − start × 24.","a~Start cell~A2|b~End cell~B2",v=>`=(${v.b}-${v.a})*24`,"Excel stores 1 day = 1, so 12 h = 0.5.",true),
 F("Duration as hh:mm","Shows 27:30 instead of 03:30.","a~Start cell~A2|b~End cell~B2",v=>`=TEXT(${v.b}-${v.a},"[h]:mm")`),
 F("Minutes → hh:mm text","135 → 2:15.","c~Minutes cell~A2",v=>`=TEXT(${v.c}/1440,"[h]:mm")`)
]},
{cat:"Error handling", items:[
 F("#N/A → Not Found","Lookup value not found.","k~Lookup value cell~A2|s~Search column~B:B|t~Return column~C:C",v=>`=IFNA(XLOOKUP(${v.k},${v.s},${v.t}),"Not Found")`,"IFNA catches only #N/A.",true),
 F("#DIV/0! guard","Show blank instead of divide-by-zero.","a~Top cell~A1|b~Bottom cell~B1",v=>`=IF(${v.b}=0,"",${v.a}/${v.b})`),
 F("#SPILL! / any error → blank","Hide errors from dynamic arrays.","r~Range~A:A",v=>`=IFERROR(FILTER(${v.r},${v.r}>5),"")`),
 F("#VALUE! guard","Text mixed with numbers.","a~First cell~A1|b~Second cell~B1",v=>`=IFERROR(${v.a}+${v.b},"")`),
 F("#REF! message","Show a message instead of #REF!.","a~First cell~A1|b~Second cell~B1",v=>`=IFERROR(${v.a}+${v.b},"Reference Error")`),
 F("Safe divide → 0","Return 0 when the division fails.","a~Top cell~A2|b~Bottom cell~B2",v=>`=IFERROR(${v.a}/${v.b},0)`),
 F("Error check column","Flag rows whose formula errors.","c~Cell to check~C2",v=>`=IF(ISERROR(${v.c}),"Check","OK")`)
]},
{cat:"Conditional formatting", items:[
 F("Green when value meets target","Home → Conditional Formatting → New Rule → Use a formula.","a~First value cell~$O2|b~First target cell~$P2",v=>`=${v.a}>=${v.b}`,"Select all data rows first; keep $ before the letter.",true),
 F("Red when below target","Second rule with red fill.","a~First value cell~$O2|b~First target cell~$P2",v=>`=${v.a}<${v.b}`),
 F("Overdue open rows","Whole row red when open and past due.","a~First status cell~$E2|b~First due cell~$D2",v=>`=AND(${v.a}<>"Closed",${v.b}<NOW())`),
 F("Highlight duplicates","Repeated IDs.","r~Whole column~$A:$A|c~First cell~$A2",v=>`=COUNTIF(${v.r},${v.c})>1`),
 F("Highlight blanks","Missing values.","c~First cell~$B2",v=>`=${v.c}=""`),
 F("Highlight weekends","Dates on Sat/Sun.","c~First date cell~$A2",v=>`=WEEKDAY(${v.c},2)>5`),
 F("Highlight top 10%","Values above the 90th percentile.","c~First cell~$B2|r~Whole range~$B$2:$B$100",v=>`=${v.c}>=PERCENTILE(${v.r},0.9)`),
 F("Zebra stripes","Shade every other row.","",v=>`=MOD(ROW(),2)=0`)
]},
{cat:"Logic-based formulas", items:[
 F("Resolution bucket with blank check","Your Advance IF: blank → No Resolution, else bucket.","c~Age cell~AX2",v=>`=IF(${v.c}="","No Resolution",IFS(${v.c}<=1,"0–1 day",${v.c}<=3,"2–3 days",${v.c}<=7,"4–7 days",${v.c}<=14,"8–14 days",${v.c}>14,"15+ days"))`,"Check blank first, then bucket from smallest.",true),
 F("Age for open and closed","Open: till now. Closed: created → closed.","a~Created cell~A2|b~Closed cell~B2",v=>`=IF(${v.b}="",INT(NOW()-${v.a}),INT(${v.b}-${v.a}))`,"",true),
 F("SLA result for open and closed","Met / Breached / Breached (Open) / Within SLA (Open).","b~Closed cell~B2|d~Due cell~D2",v=>`=IF(${v.b}="",IF(NOW()>${v.d},"Breached (Open)","Within SLA (Open)"),IF(${v.b}>${v.d},"Breached","Met"))`,"",true),
 F("SLA due date by priority","P1 = 1, P2 = 2, P3 = 3, else 5 working days.","a~Created cell~A2|p~Priority cell~F2",v=>`=WORKDAY(${v.a},SWITCH(${v.p},"P1",1,"P2",2,"P3",3,5))`),
 F("SLA due time by priority (hours)","P1 4h, P2 8h, P3 24h, else 48h.","a~Created cell~A2|p~Priority cell~F2",v=>`=${v.a}+SWITCH(${v.p},"P1",4,"P2",8,"P3",24,48)/24`),
 F("Exact business hours (9–18, Mon–Fri)","Office hours between two date-times.","a~Start date-time~A2|b~End date-time~B2",v=>`=((NETWORKDAYS(${v.a},${v.b})-1)*(TIME(18,0,0)-TIME(9,0,0))+IF(NETWORKDAYS(${v.b},${v.b}),MEDIAN(MOD(${v.b},1),TIME(18,0,0),TIME(9,0,0)),TIME(18,0,0))-MEDIAN(NETWORKDAYS(${v.a},${v.a})*MOD(${v.a},1),TIME(18,0,0),TIME(9,0,0)))*24`,"Full days × 9h, plus the last day's part, minus the first day's part. Change 9 and 18 for your office hours.",true),
 F("Shift from time","Morning 6–14, Evening 14–22, Night.","c~Date-time cell~A2",v=>`=IFS(AND(HOUR(${v.c})>=6,HOUR(${v.c})<14),"Morning",AND(HOUR(${v.c})>=14,HOUR(${v.c})<22),"Evening",TRUE,"Night")`),
 F("Weekend or weekday","Flag weekend tickets.","c~Date cell~A2",v=>`=IF(WEEKDAY(${v.c},2)>5,"Weekend","Weekday")`),
 F("Business-hours flag","Inside or outside office hours.","c~Date-time cell~A2",v=>`=IF(AND(WEEKDAY(${v.c},2)<=5,MOD(${v.c},1)>=TIME(9,0,0),MOD(${v.c},1)<TIME(18,0,0)),"Business hours","Outside hours")`),
 F("Duplicate flag","2nd, 3rd… occurrence = Duplicate.","c~First ID cell~A2",v=>{const m=v.c.match(/^\$?([A-Za-z]+)\$?(\d+)$/);const L=m?m[1].toUpperCase():"A",R=m?m[2]:"2";return `=IF(COUNTIF($${L}$${R}:${L}${R},${L}${R})>1,"Duplicate","Unique")`},"Expanding range counts how many times the ID appeared so far."),
 F("Running count per ID","1, 2, 3… for repeats (e.g. reopen count).","c~First ID cell~A2",v=>{const m=v.c.match(/^\$?([A-Za-z]+)\$?(\d+)$/);const L=m?m[1].toUpperCase():"A",R=m?m[2]:"2";return `=COUNTIF($${L}$${R}:${L}${R},${L}${R})`}),
 F("Pending days only for open","Blank for closed rows.","s~Status cell~E2|a~Created cell~A2",v=>`=IF(${v.s}="Closed","",TODAY()-${v.a})`),
 F("Status from dates","Open / Closed from the resolved date.","b~Resolved date cell~B2",v=>`=IF(${v.b}="","Open","Closed")`),
 F("Achievement %","Actual ÷ target, safe.","a~Actual cell~B2|t~Target cell~C2",v=>`=IFERROR(${v.a}/${v.t},0)`),
 F("Target status with tolerance","Met / Near / Missed (within 5%).","a~Actual cell~B2|t~Target cell~C2",v=>`=IF(${v.a}>=${v.t},"Met",IF(${v.a}>=${v.t}*0.95,"Near","Missed"))`),
 F("Grade from score","A/B/C/D bands.","c~Score cell~B2",v=>`=IFS(${v.c}>=90,"A",${v.c}>=75,"B",${v.c}>=60,"C",TRUE,"D")`),
 F("Incentive slab","Payout % by achievement.","c~Achievement % cell~D2",v=>`=IFS(${v.c}>=1.2,15%,${v.c}>=1,10%,${v.c}>=0.9,5%,TRUE,0)`),
 F("Next working day if weekend","Move a weekend date to Monday.","c~Date cell~A2",v=>`=WORKDAY(${v.c}-1,1)`),
 F("Clean text","Trim spaces and fix capitals.","c~Text cell~A2",v=>`=PROPER(TRIM(${v.c}))`),
 F("Missing data check","Flag rows with an empty key field.","a~First key cell~A2|b~Second key cell~B2",v=>`=IF(OR(${v.a}="",${v.b}=""),"Incomplete","OK")`)
]}
];

/* ===================== DAX ===================== */
const D = (n, d, inp, f, l = "", star = false, note = "") => Object.assign(F(n, d, inp, f, l, star), {note});
const LIB_PBI = [
{cat:"Aggregation", items:[
 D("SUM","Adds values.","c~Column~Sales[Amount]",v=>`Total = SUM(${v.c})`,"",true),
 D("AVERAGE","Average value.","c~Column~Sales[Amount]",v=>`Average = AVERAGE(${v.c})`),
 D("MIN","Smallest value.","c~Column~Sales[Amount]",v=>`Minimum = MIN(${v.c})`),
 D("MAX","Largest value.","c~Column~Sales[Amount]",v=>`Maximum = MAX(${v.c})`),
 D("COUNT","Count numbers.","c~Column~Sales[Amount]",v=>`Count = COUNT(${v.c})`),
 D("COUNTA","Count non-blank.","c~Column~Sales[Status]",v=>`Count Non Blank = COUNTA(${v.c})`),
 D("COUNTBLANK","Count blanks.","c~Column~Tickets[Resolved Date]",v=>`Blanks = COUNTBLANK(${v.c})`),
 D("COUNTROWS","Count rows.","t~Table~Sales",v=>`Rows = COUNTROWS(${v.t})`,"",true),
 D("DISTINCTCOUNT","Unique count.","c~Column~Sales[Customer]",v=>`Unique Count = DISTINCTCOUNT(${v.c})`)
]},
{cat:"Logical", items:[
 D("IF","Conditional logic.","m~Measure~[Total]|n~Limit~100000",v=>`Flag = IF(${v.m} > ${v.n}, "High", "Normal")`),
 D("SWITCH","Match fixed values.","c~Column~Tickets[Priority]",v=>`Priority Name = SWITCH(${v.c}, "P1", "Critical", "P2", "High", "P3", "Medium", "Low")`,"",false,"New column"),
 D("SWITCH(TRUE())","Multiple conditions, first TRUE wins.","m~Measure~[Total]",v=>`Band = SWITCH(TRUE(), ${v.m} >= 500000, "Very High", ${v.m} >= 100000, "High", "Normal")`,"",true),
 D("AND ( && )","Both true.","a~Column 1~Tickets[Status]|b~Column 2~Tickets[Priority]",v=>`Open P1 = CALCULATE(COUNTROWS(Tickets), ${v.a} <> "Closed" && ${v.b} = "P1")`),
 D("OR ( || )","Either true.","c~Column~Sales[Region]",v=>`North or South = CALCULATE([Total], ${v.c} = "North" || ${v.c} = "South")`),
 D("IN","Value in a list.","c~Column~Sales[Region]",v=>`Selected Regions = CALCULATE([Total], ${v.c} IN {"North", "South"})`),
 D("NOT","Reverse logic.","c~Column~Tickets[Status]",v=>`Not Closed = CALCULATE(COUNTROWS(Tickets), NOT(${v.c} = "Closed"))`),
 D("ISBLANK","Check blank.","c~Column~Tickets[Resolved Date]",v=>`State = IF(ISBLANK(${v.c}), "Open", "Closed")`,"",false,"New column"),
 D("IFERROR","Handle error.","c~Column~Sales[Code]",v=>`Code Number = IFERROR(VALUE(${v.c}), 0)`,"",false,"New column"),
 D("COALESCE","First non-blank value.","m~Measure~[Total]",v=>`Total or Zero = COALESCE(${v.m}, 0)`)
]},
{cat:"Filter (CALCULATE)", items:[
 D("CALCULATE","Change the filter context.","m~Measure~[Total]|c~Column~Sales[Region]|v~Value~North",v=>`${v.v} Total = CALCULATE(${v.m}, ${v.c} = "${v.v}")`,"The most important DAX function.",true),
 D("CALCULATE with 2 filters","Both conditions apply.","m~Measure~[Total]|a~Column 1~Sales[Region]|b~Column 2~Sales[Category]",v=>`Filtered = CALCULATE(${v.m}, ${v.a} = "North", ${v.b} = "Electronics")`),
 D("FILTER","Row filtering (returns a table).","t~Table~Sales|c~Column~Sales[Amount]",v=>`Big Rows = COUNTROWS(FILTER(${v.t}, ${v.c} > 10000))`,"",true),
 D("ALL","Remove all filters.","m~Measure~[Total]|t~Table~Sales",v=>`Grand Total = CALCULATE(${v.m}, ALL(${v.t}))`,"",true),
 D("% of total","Share of the grand total.","m~Measure~[Total]|t~Table~Sales",v=>`% of Total = DIVIDE(${v.m}, CALCULATE(${v.m}, ALL(${v.t})))`,"Format as %.",true),
 D("ALL on one column","Remove one column's filter.","m~Measure~[Total]|c~Column~Sales[Category]",v=>`All Categories = CALCULATE(${v.m}, ALL(${v.c}))`),
 D("ALLEXCEPT","Keep only the chosen filter.","m~Measure~[Total]|t~Table~Sales|c~Keep column~Sales[Region]",v=>`Region Total = CALCULATE(${v.m}, ALLEXCEPT(${v.t}, ${v.c}))`),
 D("ALLSELECTED","Total of what's selected in slicers.","m~Measure~[Total]|t~Table~Sales",v=>`% of Selected = DIVIDE(${v.m}, CALCULATE(${v.m}, ALLSELECTED(${v.t})))`),
 D("REMOVEFILTERS","Clear filters.","m~Measure~[Total]|t~Table~Sales",v=>`Unfiltered = CALCULATE(${v.m}, REMOVEFILTERS(${v.t}))`),
 D("KEEPFILTERS","Add a filter safely.","m~Measure~[Total]|c~Column~Sales[Region]",v=>`North Only = CALCULATE(${v.m}, KEEPFILTERS(${v.c} = "North"))`)
]},
{cat:"Iterators", items:[
 D("SUMX","Row-by-row, then sum.","t~Table~Sales|a~Qty column~Sales[Qty]|b~Price column~Sales[Price]",v=>`Revenue = SUMX(${v.t}, ${v.a} * ${v.b})`,"",true),
 D("AVERAGEX","Row-by-row average.","t~Table~Sales|a~Qty column~Sales[Qty]|b~Price column~Sales[Price]",v=>`Avg Line Value = AVERAGEX(${v.t}, ${v.a} * ${v.b})`),
 D("COUNTX","Count non-blank results.","t~Table~Tickets|c~Column~Tickets[Resolved Date]",v=>`Resolved = COUNTX(${v.t}, ${v.c})`),
 D("MAXX","Row-by-row max.","t~Table~Sales|a~Qty column~Sales[Qty]|b~Price column~Sales[Price]",v=>`Biggest Line = MAXX(${v.t}, ${v.a} * ${v.b})`),
 D("MINX","Row-by-row min.","t~Table~Sales|a~Qty column~Sales[Qty]|b~Price column~Sales[Price]",v=>`Smallest Line = MINX(${v.t}, ${v.a} * ${v.b})`),
 D("Average per month","Iterate months, average a measure.","m~Measure~[Total]|c~Month column~'Date'[Month]",v=>`Avg per Month = AVERAGEX(VALUES(${v.c}), ${v.m})`),
 D("CONCATENATEX","Join values into text.","c~Column~Sales[Product]",v=>`Product List = CONCATENATEX(VALUES(${v.c}), ${v.c}, ", ")`)
]},
{cat:"Time intelligence", items:[
 D("DATEADD (last month)","Shift the period.","m~Measure~[Total]|d~Date column~'Date'[Date]",v=>`Last Month = CALCULATE(${v.m}, DATEADD(${v.d}, -1, MONTH))`,"Needs a Date table.",true),
 D("SAMEPERIODLASTYEAR","Year-over-year.","m~Measure~[Total]|d~Date column~'Date'[Date]",v=>`Last Year = CALCULATE(${v.m}, SAMEPERIODLASTYEAR(${v.d}))`),
 D("YoY %","Growth vs last year.","m~Measure~[Total]|d~Date column~'Date'[Date]",v=>`YoY % = \nVAR cur = ${v.m}\nVAR ly = CALCULATE(${v.m}, SAMEPERIODLASTYEAR(${v.d}))\nRETURN DIVIDE(cur - ly, ly)`),
 D("MoM %","Growth vs last month.","m~Measure~[Total]|d~Date column~'Date'[Date]",v=>`MoM % = \nVAR cur = ${v.m}\nVAR prev = CALCULATE(${v.m}, DATEADD(${v.d}, -1, MONTH))\nRETURN DIVIDE(cur - prev, prev)`,"",true),
 D("TOTALYTD","Year-to-date.","m~Measure~[Total]|d~Date column~'Date'[Date]",v=>`YTD = TOTALYTD(${v.m}, ${v.d})`),
 D("TOTALYTD (Indian FY)","YTD from April.","m~Measure~[Total]|d~Date column~'Date'[Date]",v=>`FYTD = TOTALYTD(${v.m}, ${v.d}, "03-31")`),
 D("TOTALMTD","Month-to-date.","m~Measure~[Total]|d~Date column~'Date'[Date]",v=>`MTD = TOTALMTD(${v.m}, ${v.d})`),
 D("TOTALQTD","Quarter-to-date.","m~Measure~[Total]|d~Date column~'Date'[Date]",v=>`QTD = TOTALQTD(${v.m}, ${v.d})`),
 D("PREVIOUSMONTH","Last month's value.","m~Measure~[Total]|d~Date column~'Date'[Date]",v=>`Prev Month = CALCULATE(${v.m}, PREVIOUSMONTH(${v.d}))`),
 D("Rolling 3 months","Last 3 months up to the visible date.","m~Measure~[Total]|d~Date column~'Date'[Date]",v=>`Rolling 3M = CALCULATE(${v.m}, DATESINPERIOD(${v.d}, MAX(${v.d}), -3, MONTH))`),
 D("Running total","Cumulative over dates.","m~Measure~[Total]|d~Date column~'Date'[Date]",v=>`Running Total = CALCULATE(${v.m}, FILTER(ALL(${v.d}), ${v.d} <= MAX(${v.d})))`),
 D("Date table","Create a calendar table.","s~Start date~2024,1,1|e~End date~2026,12,31",v=>`Date = \nADDCOLUMNS(\n  CALENDAR(DATE(${v.s}), DATE(${v.e})),\n  "Year", YEAR([Date]),\n  "Month", FORMAT([Date], "MMM"),\n  "Month No", MONTH([Date]),\n  "Quarter", "Q" & QUARTER([Date]))`,"New table → then Mark as date table.",true,"New table")
]},
{cat:"Date & ageing", items:[
 D("DATEDIFF","Difference in DAY / HOUR / MONTH.","a~Start column~Tickets[Created Date]|b~End column~Tickets[Resolved Date]",v=>`Resolution Days = DATEDIFF(${v.a}, ${v.b}, DAY)`,"",true,"New column"),
 D("Resolution hours","Created → resolved in hours.","a~Start column~Tickets[Created Date]|b~End column~Tickets[Resolved Date]",v=>`Resolution Hrs = DATEDIFF(${v.a}, ${v.b}, HOUR)`,"",false,"New column"),
 D("Age today","Days since created.","a~Created column~Tickets[Created Date]",v=>`Age Days = DATEDIFF(${v.a}, TODAY(), DAY)`,"",false,"New column"),
 D("Open age (days)","Today minus created, as a column.","a~Created column~Tickets[Created Date]",v=>`Open Age = DATEDIFF(${v.a}, TODAY(), DAY)`,"",true,"New column"),
 D("Open age bucket (one column)","Age and bucket in one calculated column, your own limits.","a~Created column~Tickets[Created Date]|b~Bucket limits (days)~1,3,7,15,30",v=>`Open Age Bucket = 
VAR age = DATEDIFF(${v.a}, TODAY(), DAY)
RETURN IF(ISBLANK(${v.a}), BLANK(), ${bucketSWITCH("age", v.b)})`,"",true,"New column"),
 D("Open-only age bucket","Closed rows say Closed; open rows get a bucket.","a~Created column~Tickets[Created Date]|r~Resolved column~Tickets[Resolved Date]|b~Bucket limits (days)~1,3,7,15,30",v=>`Open Bucket = 
VAR age = DATEDIFF(${v.a}, TODAY(), DAY)
RETURN IF(NOT ISBLANK(${v.r}), "Closed", ${bucketSWITCH("age", v.b)})`,"",false,"New column"),
 D("Age bucket","Bucket with SWITCH(TRUE()).","a~Created column~Tickets[Created Date]|b~Resolved column~Tickets[Resolved Date]",v=>`Age Bucket = \nVAR d = DATEDIFF(${v.a}, IF(ISBLANK(${v.b}), NOW(), ${v.b}), DAY)\nRETURN SWITCH(TRUE(), d <= 1, "0–1", d <= 3, "2–3", d <= 7, "4–7", ">7")`,"",true,"New column"),
 D("SLA result","Met / Breached / Open.","b~Resolved column~Tickets[Resolved Date]|d~Due column~Tickets[SLA Due]",v=>`SLA Result = \nIF(ISBLANK(${v.b}),\n   IF(NOW() > ${v.d}, "Breached (Open)", "Within SLA (Open)"),\n   IF(${v.b} > ${v.d}, "Breached", "Met"))`,"",true,"New column"),
 D("SLA Met %","Met ÷ closed.","c~SLA result column~Tickets[SLA Result]|t~Table~Tickets",v=>`SLA Met % = DIVIDE(CALCULATE(COUNTROWS(${v.t}), ${v.c} = "Met"), CALCULATE(COUNTROWS(${v.t}), ${v.c} IN {"Met", "Breached"}))`,"",true),
 D("Avg resolution hours","Average over closed rows.","t~Table~Tickets|a~Created column~Tickets[Created Date]|b~Resolved column~Tickets[Resolved Date]",v=>`Avg Resolution Hrs = AVERAGEX(FILTER(${v.t}, NOT ISBLANK(${v.b})), DATEDIFF(${v.a}, ${v.b}, HOUR))`),
 D("Open count","Rows without a resolved date.","t~Table~Tickets|b~Resolved column~Tickets[Resolved Date]",v=>`Open = CALCULATE(COUNTROWS(${v.t}), ISBLANK(${v.b}))`),
 D("Month label column","Jan-26 style label.","a~Date column~Tickets[Created Date]",v=>`Month = FORMAT(${v.a}, "MMM-YY")`,"",false,"New column"),
 D("Weekday / weekend","Weekend flag.","a~Date column~Tickets[Created Date]",v=>`Day Type = IF(WEEKDAY(${v.a}, 2) > 5, "Weekend", "Weekday")`,"",false,"New column"),
 D("TODAY / NOW","Current date / date-time.","",v=>`Refreshed On = NOW()`,"Updates when the data refreshes.")
]},
{cat:"Relationships & ranking", items:[
 D("RELATED","Fetch from the related one-side table.","c~Column in other table~Products[Category]",v=>`Category = RELATED(${v.c})`,"",true,"New column"),
 D("RELATEDTABLE","Rows from the many side.","t~Many-side table~Sales",v=>`Orders = COUNTROWS(RELATEDTABLE(${v.t}))`,"",false,"New column in dimension"),
 D("LOOKUPVALUE","Lookup without a relationship.","r~Return column~Targets[Target]|s~Search column~Targets[Region]|c~This table's column~Sales[Region]",v=>`Target = LOOKUPVALUE(${v.r}, ${v.s}, ${v.c})`,"",false,"New column"),
 D("USERELATIONSHIP","Use an inactive relationship.","m~Measure~[Total]|a~Fact date~Sales[Ship Date]|d~Date column~'Date'[Date]",v=>`By Ship Date = CALCULATE(${v.m}, USERELATIONSHIP(${v.a}, ${v.d}))`),
 D("RANKX","Rank items by a measure.","c~Column to rank~Sales[Product]|m~Measure~[Total]",v=>`Rank = RANKX(ALL(${v.c}), ${v.m})`,"",true),
 D("TOPN","Top N items' total.","c~Column~Sales[Product]|m~Measure~[Total]|n~N~5",v=>`Top ${v.n} Total = CALCULATE(${v.m}, TOPN(${v.n}, ALL(${v.c}), ${v.m}))`),
 D("VALUES","Distinct values in the current filter.","c~Column~Sales[Region]",v=>`Regions Shown = COUNTROWS(VALUES(${v.c}))`),
 D("SELECTEDVALUE","Selected slicer value.","c~Column~Sales[Region]",v=>`Title = "Sales – " & SELECTEDVALUE(${v.c}, "All Regions")`),
 D("HASONEVALUE","Show only on single-item rows.","c~Column~Sales[Region]|m~Measure~[Total]",v=>`Row Value = IF(HASONEVALUE(${v.c}), ${v.m})`)
]},
{cat:"Tables & text", items:[
 D("SUMMARIZE","Group rows.","t~Table~Sales|c~Group column~Sales[Region]",v=>`Summary = SUMMARIZE(${v.t}, ${v.c}, "Rows", COUNTROWS(${v.t}))`,"",false,"New table"),
 D("ADDCOLUMNS","Add a column to a table.","c~Column~Sales[Region]|m~Measure~[Total]",v=>`Region Table = ADDCOLUMNS(VALUES(${v.c}), "Total", ${v.m})`,"",false,"New table"),
 D("UNION","Append tables.","a~Table 1~Sales2025|b~Table 2~Sales2026",v=>`All Sales = UNION(${v.a}, ${v.b})`,"",false,"New table"),
 D("INTERSECT","Common rows.","a~Table 1~ListA|b~Table 2~ListB",v=>`Common = INTERSECT(${v.a}, ${v.b})`,"",false,"New table"),
 D("EXCEPT","Rows in the first only.","a~Table 1~ListA|b~Table 2~ListB",v=>`Only In A = EXCEPT(${v.a}, ${v.b})`,"",false,"New table"),
 D("CONCATENATE / &","Join text.","a~Column 1~Sales[Region]|b~Column 2~Sales[Product]",v=>`Label = ${v.a} & " - " & ${v.b}`,"",false,"New column"),
 D("FORMAT","Format a value as text.","m~Measure~[Total]",v=>`Total Text = FORMAT(${v.m}, "#,0")`),
 D("VAR / RETURN","Store and reuse values.","m~Measure~[Total]|c~Column~Sales[Status]",v=>`Closed Rate = \nVAR total = ${v.m}\nVAR closed = CALCULATE(${v.m}, ${v.c} = "Closed")\nRETURN DIVIDE(closed, total)`,"",true),
 D("DIVIDE","Safe division.","a~Numerator~[Profit]|b~Denominator~[Total]",v=>`Margin % = DIVIDE(${v.a}, ${v.b})`,"",true)
]}
];

/* ===================== SQL ===================== */
const S = (n, d, inp, f, l = "", star = false) => F(n, d, inp, f, l, star);
const DB = {
  now:{mssql:"GETDATE()", mysql:"NOW()", pg:"NOW()"},
  days:{mssql:c=>`DATEDIFF(day, ${c}, GETDATE())`, mysql:c=>`DATEDIFF(CURDATE(), ${c})`, pg:c=>`(CURRENT_DATE - ${c}::date)`},
  hours:{mssql:(a,b)=>`DATEDIFF(hour, ${a}, ${b})`, mysql:(a,b)=>`TIMESTAMPDIFF(HOUR, ${a}, ${b})`, pg:(a,b)=>`ROUND(EXTRACT(EPOCH FROM (${b} - ${a})) / 3600, 2)`},
  month:{mssql:c=>`FORMAT(${c}, 'yyyy-MM')`, mysql:c=>`DATE_FORMAT(${c}, '%Y-%m')`, pg:c=>`TO_CHAR(${c}, 'YYYY-MM')`},
  top:{mssql:(n,rest)=>`SELECT TOP ${n} ${rest}`, mysql:(n,rest,o)=>`SELECT ${rest}`, pg:(n,rest)=>`SELECT ${rest}`},
  lim:{mssql:n=>"", mysql:n=>`\nLIMIT ${n}`, pg:n=>`\nLIMIT ${n}`},
  concat:{mssql:(a,b)=>`CONCAT(${a}, ' - ', ${b})`, mysql:(a,b)=>`CONCAT(${a}, ' - ', ${b})`, pg:(a,b)=>`${a} || ' - ' || ${b}`}
};
let SQLD = "mssql";
const LIB_SQL = [
{cat:"Basics", items:[
 S("Select all rows","See every column.","t~Table~tickets",v=>`SELECT *\nFROM ${v.t};`),
 S("Select some columns","Only what you need.","t~Table~tickets|a~Columns~ticket_id, status",v=>`SELECT ${v.a}\nFROM ${v.t};`),
 S("Filter by text","Rows equal to a value.","t~Table~tickets|c~Column~status|x~Value~Open",v=>`SELECT *\nFROM ${v.t}\nWHERE ${v.c} = '${v.x}';`,"",true),
 S("Filter by number","Rows above a limit.","t~Table~sales|c~Column~amount|n~Limit~10000",v=>`SELECT *\nFROM ${v.t}\nWHERE ${v.c} > ${v.n};`),
 S("Filter a list (IN)","Any of several values.","t~Table~tickets|c~Column~priority",v=>`SELECT *\nFROM ${v.t}\nWHERE ${v.c} IN ('P1', 'P2');`),
 S("Contains text (LIKE)","Text search.","t~Table~tickets|c~Column~description|x~Word~VPN",v=>`SELECT *\nFROM ${v.t}\nWHERE ${v.c} LIKE '%${v.x}%';`),
 S("Missing values","Rows where a column is empty.","t~Table~tickets|c~Column~resolved_at",v=>`SELECT *\nFROM ${v.t}\nWHERE ${v.c} IS NULL;`),
 S("Sort","Largest first.","t~Table~sales|c~Column~amount",v=>`SELECT *\nFROM ${v.t}\nORDER BY ${v.c} DESC;`),
 S("Top N","Top rows by a column.","t~Table~sales|c~Column~amount|n~N~10",v=>`${DB.top[SQLD](v.n, "*")}\nFROM ${v.t}\nORDER BY ${v.c} DESC${DB.lim[SQLD](v.n)};`),
 S("Distinct values","Unique list.","t~Table~tickets|c~Column~region",v=>`SELECT DISTINCT ${v.c}\nFROM ${v.t};`),
 S("Replace NULL","Show a default.","t~Table~tickets|c~Column~assignee",v=>`SELECT COALESCE(${v.c}, 'Unassigned') AS ${v.c}\nFROM ${v.t};`)
]},
{cat:"Group & summarise", items:[
 S("Count by group","Rows per value.","t~Table~tickets|g~Group column~status",v=>`SELECT ${v.g}, COUNT(*) AS total\nFROM ${v.t}\nGROUP BY ${v.g}\nORDER BY total DESC;`,"",true),
 S("Sum by group","Total per value.","t~Table~sales|g~Group column~region|c~Sum column~amount",v=>`SELECT ${v.g}, SUM(${v.c}) AS total\nFROM ${v.t}\nGROUP BY ${v.g};`),
 S("Average by group","Average per value.","t~Table~sales|g~Group column~region|c~Column~amount",v=>`SELECT ${v.g}, ROUND(AVG(${v.c}), 2) AS average\nFROM ${v.t}\nGROUP BY ${v.g};`),
 S("Filter groups (HAVING)","Groups above a count.","t~Table~tickets|g~Group column~assignee|n~Minimum~10",v=>`SELECT ${v.g}, COUNT(*) AS total\nFROM ${v.t}\nGROUP BY ${v.g}\nHAVING COUNT(*) > ${v.n};`),
 S("Find duplicates","Values appearing more than once.","t~Table~tickets|c~Column~ticket_id",v=>`SELECT ${v.c}, COUNT(*) AS times\nFROM ${v.t}\nGROUP BY ${v.c}\nHAVING COUNT(*) > 1;`,"",true),
 S("Percent of total","Each group's share.","t~Table~tickets|g~Group column~region",v=>`SELECT ${v.g}, COUNT(*) AS total,\n       ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 2) AS pct\nFROM ${v.t}\nGROUP BY ${v.g};`),
 S("Pivot (CASE)","Values as columns.","t~Table~tickets|g~Row column~region|c~Pivot column~status",v=>`SELECT ${v.g},\n  SUM(CASE WHEN ${v.c} = 'Open' THEN 1 ELSE 0 END) AS open_cnt,\n  SUM(CASE WHEN ${v.c} = 'Closed' THEN 1 ELSE 0 END) AS closed_cnt\nFROM ${v.t}\nGROUP BY ${v.g};`),
 S("Monthly trend","Rows per month.","t~Table~tickets|d~Date column~created_at",v=>`SELECT ${DB.month[SQLD](v.d)} AS month, COUNT(*) AS total\nFROM ${v.t}\nGROUP BY ${DB.month[SQLD](v.d)}\nORDER BY month;`)
]},
{cat:"Ageing & SLA", items:[
 S("Age in days","Days since created.","t~Table~tickets|d~Created column~created_at",v=>`SELECT *, ${DB.days[SQLD](v.d)} AS age_days\nFROM ${v.t};`,"",true),
 S("Age bucket","0-1, 2-3, 4-7, >7.","t~Table~tickets|d~Created column~created_at",v=>`SELECT *,\n  CASE\n    WHEN ${DB.days[SQLD](v.d)} <= 1 THEN '0-1'\n    WHEN ${DB.days[SQLD](v.d)} <= 3 THEN '2-3'\n    WHEN ${DB.days[SQLD](v.d)} <= 7 THEN '4-7'\n    ELSE '>7'\n  END AS age_bucket\nFROM ${v.t};`,"",true),
 S("Open age bucket (your limits)","Open items bucketed by age, limits you choose.","t~Table~tickets|d~Created column~created_at|r~Resolved column~resolved_at|b~Bucket limits (days)~1,3,7,15,30",v=>`SELECT *,
  ${bucketCASE(DB.days[SQLD](v.d), v.b)} AS age_bucket
FROM ${v.t}
WHERE ${v.r} IS NULL;`,"",true),
 S("Open items per bucket","Count of open items in each age bucket.","t~Table~tickets|d~Created column~created_at|r~Resolved column~resolved_at|b~Bucket limits (days)~1,3,7,15,30",v=>`SELECT age_bucket, COUNT(*) AS open_items
FROM (
  SELECT ${bucketCASE(DB.days[SQLD](v.d), v.b)} AS age_bucket
  FROM ${v.t}
  WHERE ${v.r} IS NULL
) b
GROUP BY age_bucket;`),
 S("Resolution hours","Created → resolved.","t~Table~tickets|a~Created column~created_at|b~Resolved column~resolved_at",v=>`SELECT *, ${DB.hours[SQLD](v.a, v.b)} AS resolution_hours\nFROM ${v.t}\nWHERE ${v.b} IS NOT NULL;`),
 S("SLA result","Met / Breached / Open.","t~Table~tickets|b~Resolved column~resolved_at|d~Due column~sla_due_at",v=>`SELECT *,\n  CASE\n    WHEN ${v.b} IS NULL AND ${DB.now[SQLD]} > ${v.d} THEN 'Breached (Open)'\n    WHEN ${v.b} IS NULL THEN 'Within SLA (Open)'\n    WHEN ${v.b} > ${v.d} THEN 'Breached'\n    ELSE 'Met'\n  END AS sla_result\nFROM ${v.t};`,"",true),
 S("SLA %","Closed on time ÷ closed.","t~Table~tickets|b~Resolved column~resolved_at|d~Due column~sla_due_at",v=>`SELECT ROUND(100.0 * SUM(CASE WHEN ${v.b} <= ${v.d} THEN 1 ELSE 0 END) / COUNT(*), 2) AS sla_pct\nFROM ${v.t}\nWHERE ${v.b} IS NOT NULL;`),
 S("Older than N days","Open items past a limit.","t~Table~tickets|d~Created column~created_at|n~Days~7",v=>`SELECT *\nFROM ${v.t}\nWHERE ${DB.days[SQLD](v.d)} > ${v.n};`),
 S("Average resolution by group","Avg hours per team.","t~Table~tickets|g~Group column~assignee|a~Created column~created_at|b~Resolved column~resolved_at",v=>`SELECT ${v.g}, ROUND(AVG(${DB.hours[SQLD](v.a, v.b)}), 2) AS avg_hours\nFROM ${v.t}\nWHERE ${v.b} IS NOT NULL\nGROUP BY ${v.g};`)
]},
{cat:"Joins", items:[
 S("INNER JOIN","Matching rows from two tables.","a~Table 1~tickets|b~Table 2~users|k~Key column~user_id",v=>`SELECT *\nFROM ${v.a} x\nJOIN ${v.b} y ON x.${v.k} = y.${v.k};`,"",true),
 S("LEFT JOIN","All rows of table 1.","a~Table 1~users|b~Table 2~tickets|k~Key column~user_id",v=>`SELECT *\nFROM ${v.a} x\nLEFT JOIN ${v.b} y ON x.${v.k} = y.${v.k};`),
 S("Rows with no match","Anti-join.","a~Table 1~users|b~Table 2~tickets|k~Key column~user_id",v=>`SELECT x.*\nFROM ${v.a} x\nLEFT JOIN ${v.b} y ON x.${v.k} = y.${v.k}\nWHERE y.${v.k} IS NULL;`),
 S("Join + count","Count related rows per item.","a~Table 1~users|b~Table 2~tickets|k~Key column~user_id",v=>`SELECT x.${v.k}, COUNT(y.${v.k}) AS related\nFROM ${v.a} x\nLEFT JOIN ${v.b} y ON x.${v.k} = y.${v.k}\nGROUP BY x.${v.k};`)
]},
{cat:"Window functions", items:[
 S("Row number","1, 2, 3 by a column.","t~Table~sales|c~Order column~amount",v=>`SELECT *, ROW_NUMBER() OVER (ORDER BY ${v.c} DESC) AS rn\nFROM ${v.t};`),
 S("Rank within group","Rank restarts per group.","t~Table~sales|g~Group column~region|c~Order column~amount",v=>`SELECT *, RANK() OVER (PARTITION BY ${v.g} ORDER BY ${v.c} DESC) AS rnk\nFROM ${v.t};`,"",true),
 S("Top 1 per group","Best row in each group.","t~Table~sales|g~Group column~region|c~Order column~amount",v=>`WITH r AS (\n  SELECT *, ROW_NUMBER() OVER (PARTITION BY ${v.g} ORDER BY ${v.c} DESC) AS rn\n  FROM ${v.t}\n)\nSELECT * FROM r WHERE rn = 1;`,"",true),
 S("Latest row per ID","Newest record of each ID.","t~Table~tickets|k~ID column~ticket_id|d~Date column~updated_at",v=>`WITH r AS (\n  SELECT *, ROW_NUMBER() OVER (PARTITION BY ${v.k} ORDER BY ${v.d} DESC) AS rn\n  FROM ${v.t}\n)\nSELECT * FROM r WHERE rn = 1;`),
 S("Running total","Cumulative sum.","t~Table~sales|d~Date column~sale_date|c~Amount column~amount",v=>`SELECT ${v.d}, ${v.c},\n       SUM(${v.c}) OVER (ORDER BY ${v.d}) AS running_total\nFROM ${v.t};`),
 S("Previous value (LAG)","Change vs previous row.","t~Table~sales|d~Date column~sale_date|c~Amount column~amount",v=>`SELECT ${v.d}, ${v.c},\n       ${v.c} - LAG(${v.c}) OVER (ORDER BY ${v.d}) AS change\nFROM ${v.t};`),
 S("Second highest","Classic interview query.","t~Table~employees|c~Column~salary",v=>`SELECT MAX(${v.c}) AS second_highest\nFROM ${v.t}\nWHERE ${v.c} < (SELECT MAX(${v.c}) FROM ${v.t});`)
]},
{cat:"Change data", items:[
 S("Insert a row","Add a record.","t~Table~departments",v=>`INSERT INTO ${v.t} (name, manager)\nVALUES ('Admin', 'Farhan');`),
 S("Update rows","Change values (always with WHERE).","t~Table~tickets|c~Column~status|k~Where column~ticket_id",v=>`UPDATE ${v.t}\nSET ${v.c} = 'Closed'\nWHERE ${v.k} = 101;`),
 S("Delete rows","Remove records (always with WHERE).","t~Table~tickets|k~Where column~status",v=>`DELETE FROM ${v.t}\nWHERE ${v.k} = 'Cancelled';`),
 S("Delete duplicates (keep one)","SQL Server / PostgreSQL style.","t~Table~tickets|k~Duplicate column~ticket_id",v=>`WITH d AS (\n  SELECT *, ROW_NUMBER() OVER (PARTITION BY ${v.k} ORDER BY ${v.k}) AS rn\n  FROM ${v.t}\n)\nDELETE FROM d WHERE rn > 1;`,"Works directly in SQL Server. In PostgreSQL/MySQL delete by a unique id from this CTE.")
]}
];

const LIBRARY = {excel:LIB_EXCEL, pbi:LIB_PBI, sql:LIB_SQL};
function libSep(f, sep) {
  if (sep !== ";") return f;
  let out = "", q = false, brace = 0;
  for (const ch of f) { if (ch === '"') q = !q; if (!q && ch === "{") brace++; if (!q && ch === "}") brace--; out += (!q && ch === "," && brace === 0) ? ";" : ch; }
  return out;
}
/* ---------- Comparing answers ---------- */
function sameValue(a, b) {
  const bl = v => v === null || v === undefined;
  if (bl(a) && bl(b)) return true;
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) <= 1e-6 * Math.max(1, Math.abs(b));
  if (typeof a === "boolean" || typeof b === "boolean") return a === b;
  return String(a ?? "").trim().toLowerCase() === String(b ?? "").trim().toLowerCase();
}
function compareSql(got, want, ordered) {
  if (!got) return {ok:false, msg:"Your query didn't return a table. Start with SELECT."};
  if (got.columns.length !== want.columns.length)
    return {ok:false, msg:`Your result has ${got.columns.length} column${got.columns.length===1?"":"s"}, but the answer has ${want.columns.length}. Check which columns you put after SELECT.`};
  if (got.values.length !== want.values.length)
    return {ok:false, msg:`Your result has ${got.values.length} row${got.values.length===1?"":"s"}, but the answer has ${want.values.length}. Check your WHERE, JOIN or GROUP BY.`};
  const norm = r => r.map(v => typeof v === "number" ? Math.round(v * 1e6) / 1e6 : (v === null ? null : String(v).toLowerCase()));
  let g = got.values.map(norm).map(r => JSON.stringify(r)), w = want.values.map(norm).map(r => JSON.stringify(r));
  if (!ordered) { g.sort(); w.sort(); }
  for (let i = 0; i < w.length; i++) if (g[i] !== w[i])
    return {ok:false, msg: ordered ? "Right number of rows, but the values or the order differ. Check ORDER BY and DESC." : "Right number of rows and columns, but some values differ. Check the column order and your conditions."};
  return {ok:true};
}
function missingParts(src, must) {
  const up = src.toUpperCase();
  return (must || []).filter(m => m.startsWith("$") ? !up.includes(m) : !new RegExp("(^|[^A-Z.])" + m.replace(/\./g, "\\.") + "\\s*\\(").test(up));
}
/* ---------- Extra Excel functions for the practice sheet ---------- */
function xlCrit(v, c) {
  if (typeof c === "number") return v === c;
  const m = String(c).match(/^(<=|>=|<>|<|>|=)?(.*)$/), op = m[1] || "=", rhs = m[2];
  const n = Number(rhs), isNum = rhs.trim() !== "" && !isNaN(n);
  if (isNum && typeof v === "number") return {"=":v===n,"<>":v!==n,"<":v<n,">":v>n,"<=":v<=n,">=":v>=n}[op];
  if (op === "=" || op === "<>") {
    const re = new RegExp("^" + rhs.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i");
    const hit = re.test(v === null || v === undefined ? "" : String(v));
    return op === "=" ? hit : !hit;
  }
  return false;
}
function installXlPlugins(HFm) {
  const {HyperFormula, FunctionPlugin, FunctionArgumentType: T, CellError, ErrorType} = HFm;
  const flat = r => (r && r.data ? r.data.flat() : [r]).map(v => (v && typeof v === "object" && "val" in v) ? v.val : v);
  const num = v => (v && typeof v === "object" && "val" in v) ? v.val : v;
  class Extra extends FunctionPlugin {
    xlookup(ast, st) { return this.runFunction(ast.args, st, this.metadata("XLOOKUP"), (key, look, ret, nf) => {
      const L = flat(look), R = flat(ret);
      const i = L.findIndex(v => typeof key === "string" ? String(v ?? "").toLowerCase() === key.toLowerCase() : num(v) === num(key));
      if (i < 0) return nf !== undefined ? nf : new CellError(ErrorType.NA, "Value not found");
      return R[i] ?? new CellError(ErrorType.NA, "Value not found");
    }); }
    rank(ast, st) { return this.runFunction(ast.args, st, this.metadata("RANK"), (v, r, order) => {
      const L = flat(r).filter(x => typeof x === "number"); v = num(v);
      if (!L.includes(v)) return new CellError(ErrorType.NA, "Value not in list");
      return 1 + L.filter(x => order ? x < v : x > v).length;
    }); }
    value(ast, st) { return this.runFunction(ast.args, st, this.metadata("VALUE"), t => {
      const n = Number(String(t).replace(/,/g, "").trim()); return isNaN(n) || String(t).trim() === "" ? new CellError(ErrorType.VALUE, "Not a number") : n;
    }); }
    textjoin(ast, st) { return this.runFunction(ast.args, st, this.metadata("TEXTJOIN"), (d, ign, ...rs) => {
      const vals = rs.flatMap(flat).map(v => v === null || v === undefined ? "" : String(num(v)));
      return vals.filter(v => !ign || v !== "").join(d);
    }); }
    concat(ast, st) { return this.runFunction(ast.args, st, this.metadata("CONCAT"), (...rs) => rs.flatMap(flat).map(v => v === null || v === undefined ? "" : String(num(v))).join("")); }
    averageifs(ast, st) { return this.runFunction(ast.args, st, this.metadata("AVERAGEIFS"), (avg, ...pairs) => {
      const A = flat(avg); const keep = A.map((_, i) => { for (let p = 0; p < pairs.length; p += 2) if (!xlCrit(num(flat(pairs[p])[i]), pairs[p + 1])) return false; return true; });
      const v = A.filter((x, i) => keep[i] && typeof num(x) === "number").map(num);
      return v.length ? v.reduce((s, x) => s + x, 0) / v.length : new CellError(ErrorType.DIV_BY_ZERO, "No matching cells");
    }); }
    mode(ast, st) { return this.runFunction(ast.args, st, this.metadata("MODE"), (...rs) => {
      const v = rs.flatMap(flat).map(num).filter(x => typeof x === "number"), c = new Map(); let best = null, bc = 1;
      v.forEach(x => { const k = (c.get(x) || 0) + 1; c.set(x, k); if (k > bc) { bc = k; best = x; } });
      return best === null ? new CellError(ErrorType.NA, "No repeated value") : best;
    }); }
  }
  const R = {argumentType: T.RANGE}, A = {argumentType: T.ANY}, S = {argumentType: T.STRING}, N = {argumentType: T.NUMBER};
  Extra.implementedFunctions = {
    XLOOKUP: {method: "xlookup", parameters: [A, R, R, {argumentType: T.ANY, optionalArg: true}]},
    RANK: {method: "rank", parameters: [N, R, {argumentType: T.NUMBER, defaultValue: 0}]},
    VALUE: {method: "value", parameters: [S]},
    TEXTJOIN: {method: "textjoin", parameters: [S, {argumentType: T.BOOLEAN}, R], repeatLastArgs: 1},
    CONCAT: {method: "concat", parameters: [R], repeatLastArgs: 1},
    AVERAGEIFS: {method: "averageifs", parameters: [R, R, A], repeatLastArgs: 2},
    MODE: {method: "mode", parameters: [R], repeatLastArgs: 1}
  };
  Extra.aliases = {"RANK.EQ": "RANK", "MODE.SNGL": "MODE"};
  const names = {}; Object.keys(Extra.implementedFunctions).forEach(k => names[k] = k); names["RANK.EQ"] = "RANK.EQ"; names["MODE.SNGL"] = "MODE.SNGL";
  HyperFormula.registerFunctionPlugin(Extra, {enGB: names, enUS: names});
}
/* ---------- App ---------- */
const ADV_IDS = new Set(["e30","e34","e35","e36","e37","e15","e41","e42","e43","e44","e45","e46","e47","e48","e18","e49","e50","e53","e54","e55","e58","e59","e60","e61","e62","e63",
  "s38","s40","s41","s42","s43","s44","s45","s46","s47","s48","s49","s50","s51","s52","s53",
  "p14","p24","p25","p26","p27","p28","p29","p30","p31","p32","p33","p34","p35","p36","p37"]);
const lvl = l => l.level || (ADV_IDS.has(l.id) ? "advanced" : "basic");
const byLevel = arr => [...arr.filter(l => lvl(l) === "basic"), ...arr.filter(l => lvl(l) === "advanced")];
const LESSONS = [...byLevel([...EXCEL_LESSONS, ...EXCEL_ADV]), ...byLevel([...SQL_LESSONS, ...SQL_ADV]), ...byLevel([...PBI_LESSONS.filter(l => l.kind !== "quiz"), ...PBI_ADV])];
const THEORY_ALL = THEORY_ORDERED;
const ALL = [...THEORY_ALL, ...LESSONS];
const LEVELS = {basic:["Basic","Start here. No experience needed."], advanced:["Advanced","After Basic. Expert skills used at work and in interviews."]};
const TRACKS = {excel:"Excel", sql:"SQL", pbi:"Power BI"};
const $ = s => document.querySelector(s);
const esc = s => String(s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const byId = id => ALL.find(l => l.id === id);
const modeOf = l => l.mode || "practice";
const trackAll = (t, m = state.mode) => (m === "theory" ? THEORY_ALL : LESSONS).filter(l => l.track === t);
const trackLessons = (t, m = state.mode, v = state.level) => trackAll(t, m).filter(l => lvl(l) === v);

let store = {done:[], last:"te1", mode:"theory", lastBy:{}};
try { const s = JSON.parse(localStorage.getItem("dpl-progress") || "null"); if (s && Array.isArray(s.done)) store = s; } catch (e) {}
const done = new Set(store.done);
const startId = byId(store.last) ? store.last : "te1";
const state = {lib:null, libCat:"all", libFilter:"", id:startId, mode:modeOf(byId(startId)), level:lvl(byId(startId)), lastBy:store.lastBy || {}, hint:0, point:null, drag:false, ac:{items:[], i:0, start:0}};
function save() { try { localStorage.setItem("dpl-progress", JSON.stringify({done:[...done], last:state.id, lastBy:state.lastBy, lib:state.lib})); } catch (e) {} }
if (store.lib && TRACKS[store.lib]) { state.lib = store.lib; state.mode = "theory"; state.level = "library"; }
const curTrack = () => state.lib || byId(state.id).track;

/* Engines */
let HF = null;
try {
  installXlPlugins(HyperFormula);
  HF = HyperFormula.HyperFormula.buildFromArray(XL_SHEET, {licenseKey:"gpl-v3"});
  HF.addNamedExpression("TRUE", "=TRUE()"); HF.addNamedExpression("FALSE", "=FALSE()");
} catch (e) { HF = null; }
let SQLlib = null, sqlFailed = false;
if (typeof initSqlJs === "function") initSqlJs().then(s => { SQLlib = s; if (curTrack() === "sql" && !state.lib) renderLesson(); }).catch(() => { sqlFailed = true; });
else sqlFailed = true;
function freshDb() { const db = new SQLlib.Database(); db.run(SQL_SCHEMA); return db; }
const SQL_COLS = {}; SQL_SCHEMA.replace(/CREATE TABLE (\w+) \(([^)]+)\)/g, (_, t, cols) => { SQL_COLS[t] = cols.split(",").map(c => c.trim().split(/\s+/)[0]); });

/* Formatting */
function fmt(v) {
  if (v === null || v === undefined || v === "") return "(blank)";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return v.toLocaleString("en-IN", {maximumFractionDigits:4});
  return String(v);
}
function serialToDate(n) { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5); return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`; }
function tableHtml(columns, rows, o = {}) {
  let h = `<div class="tscroll"><table class="grid"><thead><tr>${columns.map(c => `<th${o.pick ? ` class="pickable" data-ins="${esc(o.pick(c))}" title="Click to add to your answer"` : ""}>${esc(c)}</th>`).join("")}</tr></thead><tbody>`;
  rows.slice(0, 200).forEach(r => { const cells = r._total ? r.cells : r; h += `<tr${r._total ? ' class="total"' : ""}>` + cells.map(v => v === null ? `<td class="null">${o.blank || "NULL"}</td>` : `<td class="${typeof v === "number" ? "n" : ""}">${esc(o.nice && typeof v === "number" ? fmt(v) : v)}</td>`).join("") + "</tr>"; });
  return h + "</tbody></table></div>";
}

/* Side + tabs */
function renderChrome() {
  const t = curTrack();
  $("#app").dataset.track = t;
  for (const m of ["theory","practice"]) $("#mode-" + m).setAttribute("aria-pressed", m === state.mode);
  $("label[for=pick]").textContent = state.mode === "theory" ? "Chapter" : "Lesson";
  for (const k of Object.keys(TRACKS)) {
    $("#tab-" + k).setAttribute("aria-selected", k === t);
    const ls = trackAll(k); $("#cnt-" + k).textContent = `${ls.filter(x => done.has(x.id)).length}/${ls.length}`;
  }
  const ls = trackLessons(t), n = ls.filter(x => done.has(x.id)).length;
  $("#levels").innerHTML = Object.keys(LEVELS).map(v => { const items = trackLessons(t, state.mode, v), d = items.filter(x => done.has(x.id)).length;
    return `<button type="button" class="lvl" data-level="${v}" aria-pressed="${v === state.level}"><b>${LEVELS[v][0]}${d === items.length && items.length ? " ✓" : ""}</b><small>${d}/${items.length} · ${LEVELS[v][1]}</small></button>`; }).join("")
    + (state.mode === "theory" ? `<button type="button" class="lvl lib" data-level="library" aria-pressed="${state.level === "library"}"><b>Formula Library</b><small>${libCount(t)} ready-to-use ${t === "sql" ? "queries" : "formulas"}: type your column, copy</small></button>` : "");
  if (state.level === "library") { renderLibSide(t); return; }
  let h = `<div class="progress"><div class="bar"><i style="width:${Math.round(n / ls.length * 100)}%"></i></div><small>${n === ls.length ? `${LEVELS[state.level][0]} complete ✓` : `${n} of ${ls.length} ${LEVELS[state.level][0].toLowerCase()} ${state.mode === "theory" ? "chapters read" : "lessons done"}`}</small></div>`;
  let g = null, opts = "";
  ls.forEach((x, i) => {
    if (x.group !== g) { if (g !== null) opts += "</optgroup>"; g = x.group; h += `<div class="grp">${esc(g)}</div>`; opts += `<optgroup label="${esc(g)}">`; }
    h += `<button class="ls${done.has(x.id) ? " done" : ""}" data-id="${x.id}" aria-current="${x.id === state.id}"><span class="ck">✓</span>${esc(x.title)}</button>`;
    opts += `<option value="${x.id}"${x.id === state.id ? " selected" : ""}>${i + 1}. ${esc(x.title)}${done.has(x.id) ? " ✓" : ""}</option>`;
  });
  $("#side").innerHTML = h; $("#pick").innerHTML = opts + "</optgroup>";
  const cur = $("#side .ls[aria-current='true']"), rail = $("#rail"); if (cur && rail) { const rt = rail.getBoundingClientRect(), ct = cur.getBoundingClientRect(); if (ct.top < rt.top + 40 || ct.bottom > rt.bottom - 20) rail.scrollTop += ct.top - rt.top - rt.height / 3; }
}

/* ---------- Excel sheet + cell picking ---------- */
const cellRC = ref => { const m = ref.match(/^([A-M])(\d+)$/); return [XL_COLS.indexOf(m[1]), +m[2]]; };
function rangeRef(a, b) { if (a === b) return a; const [c1, r1] = cellRC(a), [c2, r2] = cellRC(b); return XL_COLS[Math.min(c1, c2)] + Math.min(r1, r2) + ":" + XL_COLS[Math.max(c1, c2)] + Math.max(r1, r2); }
function cellsOf(ref) { const set = new Set(); const [a, b] = ref.split(":"); const [c1, r1] = cellRC(a), [c2, r2] = cellRC(b || a); for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) set.add(XL_COLS[c] + r); return set; }
function sheetHtml(hits = new Set()) {
  let h = `<div class="sheetwrap"><table class="sheet"><thead><tr><th class="corner"></th>${XL_COLS.map(c => `<th class="colh" data-col="${c}" title="Click to select ${c}2:${c}11">${c}</th>`).join("")}</tr></thead><tbody>`;
  XL_SHEET.forEach((row, r) => {
    h += `<tr class="${r === 0 ? "hdr" : ""}"><th>${r + 1}</th>` + row.map((v, c) => { const ref = XL_COLS[c] + (r + 1); return `<td data-ref="${ref}" class="${typeof v === "number" ? "n" : ""}${hits.has(ref) ? " hit" : ""}">${v === null ? "" : esc(v)}</td>`; }).join("") + "</tr>";
  });
  return h + "</tbody></table></div>";
}
function refsIn(f) {
  const hits = new Set(), clean = f.replace(/"[^"]*"/g, "");
  const re = /\$?([A-M])\$?(\d{1,2})(?:\s*:\s*\$?([A-M])\$?(\d{1,2}))?(?![A-Za-z(])/gi; let m;
  while ((m = re.exec(clean))) { const a = m[1].toUpperCase() + m[2], b = m[3] ? m[3].toUpperCase() + m[4] : a; if (+m[2] >= 1 && +m[2] <= 11 && (!m[4] || (+m[4] >= 1 && +m[4] <= 11))) cellsOf(a + ":" + b).forEach(x => hits.add(x)); }
  return hits;
}
function markPick(ref) { document.querySelectorAll(".sheet td.pick").forEach(td => td.classList.remove("pick")); if (!ref) return; cellsOf(ref).forEach(c => { const td = document.querySelector(`.sheet td[data-ref="${c}"]`); if (td) td.classList.add("pick"); }); }
function insertRef(ref, anchor) {
  const inp = $("#code"); let v = inp.value, s, e;
  if (!v.startsWith("=")) { v = "=" + v; inp.value = v; state.point = null; }
  if (state.point && inp.selectionStart === state.point.end && v.slice(state.point.start, state.point.end) === state.point.text) { s = state.point.start; e = state.point.end; }
  else { s = inp.selectionStart ?? v.length; e = inp.selectionEnd ?? s; if (document.activeElement !== inp) { s = e = v.length; } }
  inp.value = v.slice(0, s) + ref + v.slice(e);
  const caret = s + ref.length; inp.focus(); inp.setSelectionRange(caret, caret);
  state.point = {start:s, end:caret, text:ref, anchor:anchor || ref.split(":")[0], single:!ref.includes(":")};
  markPick(ref); hideAc(); updateArgTip();
}

/* ---------- Autocomplete ---------- */
const inQuotes = (s, q) => (s.split(q).length - 1) % 2 === 1;
function candidates(track, before) {
  if (track === "excel") {
    if (!before.startsWith("=") || inQuotes(before, '"')) return null;
    const m = before.match(/(?:^=|[(,+\-*/^&<>=:\s])([A-Za-z][A-Za-z0-9.]*)$/); if (!m) return null;
    const tok = m[1].toUpperCase();
    const items = Object.keys(XL_FUNCS).filter(n => n.startsWith(tok) && n !== tok + "(").sort((a, b) => a.length - b.length || a.localeCompare(b))
      .map(n => ({label:n, insert:["TRUE","FALSE"].includes(n) ? n : n + "(", detail:XL_FUNCS[n][1]}));
    return {start:before.length - m[1].length, items};
  }
  if (track === "sql") {
    if (inQuotes(before, "'")) return null;
    const dot = before.match(/(\w+)\.(\w*)$/);
    if (dot) { const alias = dot[1].toLowerCase(); const cols = [...new Set(Object.values(SQL_COLS).flat())];
      return {start:before.length - dot[2].length, items:cols.filter(c => c.startsWith(dot[2].toLowerCase())).map(c => ({label:c, insert:c, detail:"column"}))}; }
    const m = before.match(/(?:^|[^\w.])([A-Za-z_]\w*)$/); if (!m) return null;
    const tok = m[1].toLowerCase(), items = [];
    SQL_WORDS.filter(w => w.toLowerCase().startsWith(tok)).forEach(w => items.push({label:w, insert:w + " ", detail:"keyword"}));
    Object.keys(SQL_FUNCS).filter(f => f.toLowerCase().startsWith(tok)).forEach(f => items.push({label:f, insert:f + "(", detail:"function"}));
    SQL_TABLES.filter(t => t.startsWith(tok)).forEach(t => items.push({label:t, insert:t + " ", detail:"table"}));
    const cols = {}; Object.entries(SQL_COLS).forEach(([t, cs]) => cs.forEach(c => (cols[c] = cols[c] || []).push(t)));
    Object.keys(cols).filter(c => c.startsWith(tok)).forEach(c => items.push({label:c, insert:c, detail:"column in " + cols[c].join(", ")}));
    return {start:before.length - m[1].length, items:items.filter(i => i.label.toLowerCase() !== tok)};
  }
  if (track === "pbi") {
    if (inQuotes(before, '"')) return null;
    const tc = before.match(/([A-Za-z_]+)\[([^\]]*)$/);
    if (tc) { const T = PBI_TABLES[tc[1].toLowerCase()]; if (!T) return null;
      return {start:before.length - tc[2].length, items:T.cols.filter(c => c.toLowerCase().startsWith(tc[2].toLowerCase())).map(c => ({label:c, insert:c + "]", detail:T.name + " column"}))}; }
    const br = before.match(/(?:^|[^A-Za-z_\]])\[([^\]]*)$/);
    if (br) { const p = br[1].toLowerCase(); return {start:before.length - br[1].length, items:Object.keys(PBI_MEASURES).filter(k => k.toLowerCase().startsWith(p)).map(k => ({label:"[" + k + "]", insert:k + "]", detail:"measure: " + PBI_MEASURES[k]}))}; }
    const m = before.match(/(?:^|[^\w\]])([A-Za-z_]\w*)$/); if (!m) return null;
    const tok = m[1].toUpperCase(), items = [];
    Object.values(PBI_TABLES).filter(t => t.name.toUpperCase().startsWith(tok)).forEach(t => items.push({label:t.name, insert:t.name, detail:"table"}));
    Object.keys(DAX_FN).filter(f => f.startsWith(tok)).sort((a, b) => a.length - b.length).forEach(f => items.push({label:f, insert:["TRUE","FALSE","BLANK"].includes(f) ? f + "()" : f + "(", detail:DAX_FN[f] ? f + "(" + DAX_FN[f] + ")" : ""}));
    ["VAR","RETURN"].filter(k => k.startsWith(tok)).forEach(k => items.push({label:k, insert:k + " ", detail:"keyword"}));
    return {start:before.length - m[1].length, items:items.filter(i => i.label.toUpperCase() !== tok)};
  }
  return null;
}
function updateAc() {
  const inp = $("#code"), box = $("#ac"); if (!inp || !box) return;
  const l = byId(state.id), c = inp.selectionStart, before = inp.value.slice(0, c);
  const r = document.activeElement === inp && inp.selectionStart === inp.selectionEnd ? candidates(l.track, before) : null;
  if (!r || !r.items.length) return hideAc();
  state.ac = {items:r.items.slice(0, 8), i:0, start:r.start};
  box.innerHTML = state.ac.items.map((it, i) => `<button type="button" class="acit${i === 0 ? " on" : ""}" data-i="${i}" role="option"><b>${esc(it.label)}</b><span>${esc(it.detail)}</span></button>`).join("") + `<div class="achelp">Tab to insert · ↑↓ to choose · Esc to close</div>`;
  box.hidden = false;
}
function hideAc() { const box = $("#ac"); if (box) { box.hidden = true; } state.ac.items = []; }
function acceptAc(i = state.ac.i) {
  const it = state.ac.items[i]; if (!it) return;
  const inp = $("#code"), c = inp.selectionStart, v = inp.value;
  inp.value = v.slice(0, state.ac.start) + it.insert + v.slice(c);
  const caret = state.ac.start + it.insert.length; inp.focus(); inp.setSelectionRange(caret, caret);
  state.point = null; hideAc(); updateArgTip(); updateAc();
}
function moveAc(d) { const n = state.ac.items.length; state.ac.i = (state.ac.i + d + n) % n; document.querySelectorAll("#ac .acit").forEach((b, i) => b.classList.toggle("on", i === state.ac.i)); }

function updateArgTip() {
  const inp = $("#code"), tip = $("#argtip"); if (!inp || !tip) return;
  const l = byId(state.id), before = inp.value.slice(0, inp.selectionStart);
  let depth = 0, commas = 0, q = null, open = -1;
  for (let i = before.length - 1; i >= 0; i--) {
    const ch = before[i];
    if (q) { if (ch === q) q = null; continue; }
    if (ch === '"' || (l.track === "sql" && ch === "'")) { q = ch; continue; }
    if (ch === ")") depth++; else if (ch === "(") { if (depth === 0) { open = i; break; } depth--; } else if (ch === "," && depth === 0) commas++;
  }
  const name = open > 0 ? (before.slice(0, open).match(/([A-Za-z][A-Za-z0-9._]*)\s*$/) || [])[1] : null;
  const N = name && name.toUpperCase();
  const sig = l.track === "excel" ? (XL_FUNCS[N] || [])[0] : l.track === "pbi" ? DAX_FN[N] : SQL_FUNCS[N];
  if (!N || sig === undefined) { tip.hidden = true; return; }
  const parts = sig ? sig.split(", ") : [];
  const idx = Math.min(commas, parts.length - 1), rep = parts.length && /…/.test(parts[parts.length - 1]);
  const html = parts.map((p, i) => (i === idx || (rep && i === parts.length - 1 && commas >= parts.length - 1 && i === idx)) ? `<b>${esc(p)}</b>` : esc(p)).join(", ");
  const desc = l.track === "excel" ? XL_FUNCS[N][1] : "";
  tip.innerHTML = `<code>${esc(N)}(${html})</code>${desc ? ` <span>${esc(desc)}</span>` : ""}`; tip.hidden = false;
}

/* ---------- Data views ---------- */
function sqlTablesHtml() {
  if (sqlFailed) return `<p class="notice">The SQL engine couldn't load. Check your internet connection and reload the page.</p>`;
  if (!SQLlib) return `<p class="notice">Loading the practice database…</p>`;
  const db = freshDb(); let h = "";
  SQL_TABLES.forEach((t, i) => {
    const r = db.exec(`SELECT * FROM ${t}`)[0];
    h += `<details class="data"${i === 0 ? " open" : ""}><summary><span class="tname" data-ins="${t}" title="Click to add to your query">${t}</span><span>${r.values.length} rows · ${r.columns.join(", ")}</span></summary>${tableHtml(r.columns, r.values, {pick:c => c})}</details>`;
  });
  db.close(); return h;
}
function pbiTablesHtml(open) {
  return `<details class="data"${open ? " open" : ""}><summary><span class="tname" data-ins="Sales">Sales</span><span>${PBI_SALES.length} rows · fact table</span></summary>${tableHtml(PBI_COLS, PBI_SALES.map(r => PBI_COLS.map(c => r[c])), {pick:c => `Sales[${c}]`})}</details>
  <details class="data"><summary><span class="tname" data-ins="Regions">Regions</span><span>${PBI_REGIONS_T.length} rows · related to Sales on Region</span></summary>${tableHtml(REG_COLS, PBI_REGIONS_T.map(r => REG_COLS.map(c => r[c])), {pick:c => `Regions[${c}]`})}</details>
  <div class="chips"><span class="lab">Ready-made measures</span>${Object.keys(PBI_MEASURES).map(m => `<button type="button" class="chip" data-ins="[${esc(m)}]" title="${esc(PBI_MEASURES[m])}">[${esc(m)}]</button>`).join("")}</div>`;
}

/* ---------- Lesson ---------- */
function renderLesson() {
  if (state.home) { renderHome(); return; }
  if (state.lib) { renderLibrary(); return; }
  const l = byId(state.id), ls = trackLessons(l.track), idx = ls.indexOf(l);
  state.hint = 0; state.point = null; renderChrome();
  if (l.mode === "theory") { renderTheory(l, ls, idx); return; }
  let h = `<div class="eyebrow">${TRACKS[l.track]} practice · ${LEVELS[lvl(l)][0]} · ${esc(l.group)} · Lesson ${idx + 1} of ${ls.length}</div><h2>${esc(l.title)}</h2><div class="learn">${l.learn}</div>`;
  if (l.syntax || l.example) {
    h += `<div class="pattern">`;
    if (l.syntax) h += `<div><span class="lab">Pattern</span><pre>${esc(l.syntax)}</pre></div>`;
    if (l.example) h += `<div><span class="lab">Example</span><pre>${esc(l.example.code)}</pre>${l.example.note ? `<p>${esc(l.example.note)}</p>` : ""}${l.kind !== "quiz" ? `<button type="button" class="btn small" id="tryex">Try this example ▸</button>` : ""}</div>`;
    h += `</div>`;
  }
  h += `<section class="try">`;
  if (l.kind === "quiz") {
    h += `<h3>Check yourself</h3>` + l.questions.map((q, qi) => `<fieldset class="q" data-q="${qi}"><legend>${qi + 1}. ${esc(q.q)}</legend><div class="opts">${q.opts.map((o, oi) => `<button class="opt" data-q="${qi}" data-o="${oi}">${esc(o)}</button>`).join("")}</div><p class="why" hidden></p></fieldset>`).join("");
    h += `<div class="result" id="result"></div>`;
  } else {
    h += `<h3>Your turn</h3><p class="task">${esc(l.task)}</p>`;
    const tipLine = {excel:"Type <kbd>=su</kbd> then <kbd>Tab</kbd> to insert SUM( · click or drag cells to add them · click a column letter for the whole column · <kbd>Enter</kbd> runs",
      sql:"<kbd>Tab</kbd> completes words · click a table or column name to add it · <kbd>Ctrl</kbd>+<kbd>Enter</kbd> runs",
      pbi:"<kbd>Tab</kbd> completes · type <kbd>Sales[</kbd> for columns or <kbd>[</kbd> for measures · click a column name to add it · <kbd>Ctrl</kbd>+<kbd>Enter</kbd> runs"}[l.track];
    if (l.track === "excel") h += `<div id="sheet">${sheetHtml()}</div>`;
    if (l.track === "sql") h += `<div id="tables">${sqlTablesHtml()}</div>`;
    if (l.track === "pbi") h += pbiTablesHtml(l.kind === "column");
    h += `<div class="editor">`;
    h += l.track === "excel"
      ? `<div class="fbar"><span class="fx">fx</span><input id="code" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" placeholder="Type your formula, starting with =" aria-label="Your formula" aria-autocomplete="list"></div>`
      : `<textarea class="code" id="code" spellcheck="false" autocapitalize="off" autocorrect="off" placeholder="${l.track === "sql" ? "Type your SQL query here" : (l.kind === "column" ? "Column Name = your DAX formula" : "Measure Name = your DAX formula")}" aria-label="Your answer" aria-autocomplete="list"></textarea>`;
    h += `<div class="ac" id="ac" role="listbox" hidden></div></div><div class="argtip" id="argtip" hidden></div><p class="tipline">${tipLine}</p>`;
    h += `<div class="actions"><button class="btn primary" id="run">Run &amp; check</button><button class="btn" id="hint">Hint</button><button class="btn" id="reveal">Show answer</button></div><div class="hintbox" id="hintbox"></div><div class="result" id="result"></div>`;
  }
  h += `</section><div class="pager">${idx > 0 ? `<button class="btn" data-go="${ls[idx - 1].id}">← ${esc(ls[idx - 1].title)}</button>` : "<span></span>"}${idx < ls.length - 1 ? `<button class="btn" data-go="${ls[idx + 1].id}">${esc(ls[idx + 1].title)} →</button>` : nextTrackBtn(l.track)}</div>`;
  $("#lesson").innerHTML = h;
  const code = $("#code"); if (code && l.track === "excel") code.value = "=";
}
function quizHtml(l, title) {
  return `<h3>${title}</h3>` + l.questions.map((q, qi) => `<fieldset class="q" data-q="${qi}"><legend>${qi + 1}. ${esc(q.q)}</legend><div class="opts">${q.opts.map((o, oi) => `<button class="opt" data-q="${qi}" data-o="${oi}">${esc(o)}</button>`).join("")}</div><p class="why" hidden></p></fieldset>`).join("") + `<div class="result" id="result"></div>`;
}
function richBody(l) {
  let h = l.intro || "";
  if (l.syntax) h += `<h4>Syntax</h4><div class="fcode syntax"><pre>${esc(l.syntax)}</pre></div>`;
  if (l.args && l.args.length) h += `<table class="mini"><tr><th>${l.track === "sql" ? "Part" : "Argument"}</th><th>Meaning</th></tr>${l.args.map(([a, m]) => `<tr><td><code>${esc(a)}</code></td><td>${esc(m)}</td></tr>`).join("")}</table>`;
  if (l.steps && l.steps.length) h += `<h4>${l.track === "excel" ? "How to use it in Excel" : l.track === "pbi" ? "Step by step in Power BI" : "How it works"}</h4><ol class="steps">${l.steps.map(s => `<li><span>${s}</span></li>`).join("")}</ol>`;
  if (l.table) h += `<table class="mini"><tr>${l.table.cols.map(c => `<th>${c}</th>`).join("")}</tr>${l.table.rows.map(r => `<tr>${r.map(c => `<td>${c}</td>`).join("")}</tr>`).join("")}</table>`;
  if (l.ex && l.ex.length) {
    h += `<h4>Example${l.ex.length > 1 ? "s" : ""}</h4>`;
    if (l.track === "excel" && l.sheet !== false && l.ex.some(e => /[A-M]\d/.test(e.code))) {
      const hits = new Set(); l.ex.forEach(e => refsIn(e.code).forEach(x => hits.add(x)));
      h += `<details class="data exdata" open><summary>Example data<span>the practice sheet · highlighted cells are used below</span></summary>${sheetHtml(hits)}</details>`;
    }
    if (l.track === "sql") h += `<details class="data exdata"><summary>Example data<span>tables: ${SQL_TABLES.join(", ")}</span></summary><div id="thtables">${sqlTablesHtml()}</div></details>`;
    if (l.track === "pbi") h += `<details class="data exdata"><summary>Example data<span>Sales (24 rows) and Regions</span></summary>${pbiTablesHtml(true)}</details>`;
    l.ex.forEach((e, i) => { h += `<div class="exbox"><p class="exd"><b>Example ${i + 1}:</b> ${esc(e.desc)}</p><div class="fcode"><pre>${esc(e.code)}</pre><button type="button" class="btn small copy" data-copy="1">Copy</button></div>${exResult(l, e, i)}</div>`; });
  }
  (l.notes || []).forEach(n => { h += `<div class="callout note"><b>Note</b><p>${n}</p></div>`; });
  if (l.mistake) h += `<div class="callout bad"><b>Common mistake</b><p><code>${esc(l.mistake.wrong)}</code></p><p>${esc(l.mistake.why)}</p><p><b>Instead:</b> <code>${esc(l.mistake.fix)}</code></p></div>`;
  return h;
}
function exResult(l, e, i) {
  if (l.noRun || e.run === false) return `<p class="exr muted">Result not shown: this needs ${l.track === "excel" ? "Excel 365" : "a full Power BI model"}.</p>`;
  if (l.track === "excel") {
    if (!HF) return "";
    const r = xlEval(e.code); if (r.err) return `<p class="exr err">Result: ${esc(r.code)}</p>`;
    const isDate = /WORKDAY|EOMONTH|EDATE|^=DATE\(|MIN\(F/.test(e.code) && typeof r.v === "number" && r.v > 20000;
    return `<p class="exr">Result: <b>${esc(fmt(r.v))}</b>${isDate ? ` <span class="muted">(${serialToDate(r.v)} as a date)</span>` : ""}</p>`;
  }
  if (l.track === "sql") {
    if (sqlFailed) return `<p class="exr err">The SQL engine couldn't load.</p>`;
    if (!SQLlib) return `<p class="exr muted">Loading database…</p>`;
    const db = freshDb(); let out;
    try { const r = db.exec(e.code); const last = r[r.length - 1]; out = last ? `<div class="exr">${tableHtml(last.columns, last.values)}<span class="muted">${last.values.length} row${last.values.length === 1 ? "" : "s"}</span></div>` : `<p class="exr">Statement ran (no rows returned).</p>`; }
    catch (err) { out = `<p class="exr err">Database says: ${esc(err.message)}</p>`; }
    db.close(); return out;
  }
  try {
    const kind = e.kind || (l.exKind && l.exKind[i]);
    if (kind === "column") { const c = daxColumn(e.code); return `<div class="exr">${tableHtml(["OrderID","Region","Product", c.name || "New column"], PBI_SALES.slice(0, 6).map((r, k) => [r.OrderID, r.Region, r.Product, c.values[k] === null ? "(blank)" : c.values[k]]))}<span class="muted">first 6 of ${PBI_SALES.length} rows</span></div>`; }
    const m = daxMeasure(e.code, e.by || "Region"); const rows = m.groups.map(g => [g.key, dv(g.value)]); rows.push({_total:true, cells:["Total", dv(m.total)]});
    return `<div class="exr">${tableHtml([e.by || "Region", m.name || "Measure"], rows)}</div>`;
  } catch (err) { return `<p class="exr muted">Result not shown (${esc(err.message)})</p>`; }
}
function renderTheory(l, ls, idx) {
  let h = `<div class="eyebrow">${TRACKS[l.track]} theory · ${LEVELS[lvl(l)][0]} · ${esc(l.group)} · Chapter ${idx + 1} of ${ls.length}${done.has(l.id) ? " · read ✓" : ""}</div><h2>${esc(l.title)}</h2><article class="learn theory">${l.rich ? richBody(l) : l.body}</article>`;
  if (l.terms && l.terms.length) h += `<div class="terms"><span class="lab">Key terms</span><dl>${l.terms.map(([k, v]) => `<div><dt>${esc(k)}</dt><dd>${esc(v)}</dd></div>`).join("")}</dl></div>`;
  if (l.questions && l.questions.length) h += `<section class="try">${quizHtml(l, "Quick check")}</section>`;
  else h += `<section class="try"><div class="result" id="result">${done.has(l.id) ? `<div class="verdict ok"><b>Read ✓</b><p>You've marked this chapter as read.</p>${nextBtn()}</div>` : `<button type="button" class="btn primary" id="markread">Mark as read</button>`}</div></section>`;
  const prac = (l.practice || []).map(byId).filter(Boolean);
  if (prac.length) h += `<div class="practise"><span class="lab">Try it in Practice</span><div class="chips">${prac.map(p => `<button type="button" class="chip go" data-go="${p.id}">${esc(p.title)} →</button>`).join("")}</div></div>`;
  h += `<div class="pager">${idx > 0 ? `<button class="btn" data-go="${ls[idx - 1].id}">← ${esc(ls[idx - 1].title)}</button>` : "<span></span>"}${idx < ls.length - 1 ? `<button class="btn" data-go="${ls[idx + 1].id}">${esc(ls[idx + 1].title)} →</button>` : nextTrackBtn(l.track)}</div>`;
  $("#lesson").innerHTML = h;
}
function nextTrackBtn(t) {
  if (state.level === "basic" && trackLessons(t, state.mode, "advanced").length) return `<button class="btn" data-go="${trackLessons(t, state.mode, "advanced")[0].id}">Continue to Advanced →</button>`;
  if (state.mode === "theory") return `<button class="btn" data-go="${trackLessons(t, "practice", "basic")[0].id}">Practise ${TRACKS[t]} →</button>`;
  const order = ["excel", "sql", "pbi"], n = order[order.indexOf(t) + 1]; return n ? `<button class="btn" data-go="${trackLessons(n, "practice", "basic")[0].id}">Start ${TRACKS[n]} →</button>` : "<span></span>"; }
function go(id) { state.home = false; closeDrawer(); state.lib = null; const l = byId(id); state.id = id; state.mode = modeOf(l); state.level = lvl(l); state.lastBy[state.mode + ":" + l.track + ":" + state.level] = id; save(); renderLesson(); animateLesson(); scrollTopSoft(); }
function openLibrary(track, cat) { state.home = false; closeDrawer(); state.lib = track; state.mode = "theory"; state.level = "library"; if (cat) state.libCat = cat; save(); renderLibrary(); animateLesson(); scrollTopSoft(); }
function switchTo(mode, track, level = state.level) {
  if (level === "library" && mode === "practice") level = "basic";
  if (level === "library") { if (state.lib !== track) state.libCat = "all"; openLibrary(track); return; }
  state.lib = null; const key = mode + ":" + track + ":" + level, ls = trackLessons(track, mode, level); const id = state.lastBy[key] && byId(state.lastBy[key]) ? state.lastBy[key] : (ls.find(x => !done.has(x.id)) || ls[0]).id; go(id); }

function verdict(ok, msg) { return `<div class="verdict ${ok ? "ok" : "no"}"><b>${ok ? "Correct!" : "Not yet"}</b><p>${msg}</p>${ok ? nextBtn() : ""}</div>`; }
function nextBtn() { const l = byId(state.id), ls = trackLessons(l.track), i = ls.indexOf(l); return i < ls.length - 1 ? `<button class="btn primary" data-go="${ls[i + 1].id}">Next lesson →</button>` : nextTrackBtn(l.track).replace('class="btn"', 'class="btn primary"'); }
function markDone() {
  const id = state.id, was = done.has(id); done.add(id); save(); renderChrome();
  if (was) return;
  recordDay();
  const l = byId(id), lvItems = trackLessons(l.track, modeOf(l), lvl(l));
  if (lvItems.every(x => done.has(x.id))) { confetti(160); toast(`${TRACKS[l.track]} ${LEVELS[lvl(l)][0]} ${modeOf(l) === "theory" ? "theory" : "practice"} complete!`); }
  else { confetti(50); }
}

/* ---------- Checking ---------- */
const XL_ERRORS = {
  "#NAME?":"Excel doesn't recognise a name. Check the function's spelling, and put text inside \"double quotes\".",
  "#DIV/0!":"The formula divides by zero (or by an empty cell).",
  "#N/A":"A lookup couldn't find what you searched for. Check the spelling and the range.",
  "#VALUE!":"The formula mixes text and numbers in math, or an input has the wrong type.",
  "#REF!":"The formula points to a cell that doesn't exist.",
  "#NUM!":"A number in the formula is not valid for this function.",
  "#ERROR!":"The formula has a typo. Check your brackets, commas and quotes."
};
function xlEval(f) {
  try { const v = HF.calculateFormula(f, 0);
    if (Array.isArray(v)) return {err:"This formula gives back many cells. The answer should be a single value. Did you forget a function like SUM around the range?", code:"Many cells"};
    if (v && typeof v === "object" && "value" in v) return {err:XL_ERRORS[v.value] || "The formula gave an error.", code:v.value};
    return {v}; }
  catch (e) { return {err:"Excel can't read this formula. Check that every ( has a closing ), inputs are separated by commas, and text is in \"double quotes\".", code:"#ERROR!"}; }
}
function runExcel(l, grade = true) {
  if (!HF) return `<p class="notice">The Excel engine couldn't load. Check your internet connection and reload the page.</p>`;
  let f = $("#code").value.trim();
  if (!f || f === "=") return `<div class="verdict info"><p>Type a formula first. It starts with <code>=</code>.</p></div>`;
  if (!f.startsWith("=")) f = "=" + f;
  const opens = (f.match(/\(/g) || []).length, closes = (f.match(/\)/g) || []).length;
  if (opens > closes) { f += ")".repeat(opens - closes); }
  $("#sheet").innerHTML = sheetHtml(refsIn(f)); state.point = null;
  const got = xlEval(f), want = xlEval(l.answer).v;
  const shown = got.err ? esc(got.code) : esc(fmt(got.v)) + (l.dateResult && typeof got.v === "number" ? ` <small class="asdate">= ${serialToDate(got.v)} as a date</small>` : "");
  let out = `<div class="out"><span class="lab">Your formula gives${opens > closes ? " (closing brackets added for you)" : ""}</span><div class="big">${shown}</div>${got.err ? `<p class="errline">${got.err}</p>` : ""}</div>`;
  if (!grade) return out + PREVIEW;
  if (got.err) return out + verdict(false, "Fix the error above, then run again. Highlighted cells show what your formula reads.");
  if (!sameValue(got.v, want)) return out + verdict(false, "That's not the expected result. Look at the highlighted cells: are they the ones the task asks about? Try a hint if you're stuck.");
  const miss = missingParts(f, l.must);
  if (miss.length) return out + verdict(false, `The result is right, but this lesson practices <code>${esc(miss[0])}</code>. Try writing it that way.`);
  markDone(); return out + verdict(true, "Your formula works. The highlighted cells are the ones it used.");
}
function friendlySql(msg) {
  let m;
  if ((m = msg.match(/no such column: (.+)/))) return `There is no column called <code>${esc(m[1])}</code>. Check the spelling against the column names above.`;
  if ((m = msg.match(/no such table: (.+)/))) return `There is no table called <code>${esc(m[1])}</code>. Tables: ${SQL_TABLES.join(", ")}.`;
  if ((m = msg.match(/near "(.+?)": syntax error/))) return `SQL couldn't understand the query near <code>${esc(m[1])}</code>. Check spelling, commas and quotes around text.`;
  if (/incomplete input/.test(msg)) return "The query stops too early. Is a bracket or quote missing?";
  if (/misuse of aggregate/.test(msg)) return "COUNT, SUM, AVG can't be used in WHERE. Use HAVING after GROUP BY.";
  if (/ambiguous column name: (.+)/.test(msg)) return `The column <code>${esc(msg.split(": ")[1])}</code> exists in two tables. Put the table alias in front, like <code>e.${esc(msg.split(": ")[1])}</code>.`;
  return esc(msg);
}
function runSql(l, grade = true) {
  if (sqlFailed) return `<p class="notice">The SQL engine couldn't load. Check your internet connection and reload the page.</p>`;
  if (!SQLlib) return `<p class="notice">The practice database is still loading. Try again in a second.</p>`;
  const q = $("#code").value.trim();
  if (!q) return `<div class="verdict info"><p>Type a query first. Most queries start with <code>SELECT</code>.</p></div>`;
  let got, db = freshDb();
  if (!grade) { let r; try { r = db.exec(q); } catch (e) { db.close(); return verdict(false, friendlySql(e.message)); } db.close(); const last = r[r.length - 1];
    return (last ? `<div class="out"><span class="lab">Example result · ${last.values.length} row${last.values.length === 1 ? "" : "s"}</span>${tableHtml(last.columns, last.values)}</div>` : `<div class="out"><p>The example ran. It changes data, so there is no table to show.</p></div>`) + PREVIEW; }
  try { const r = db.exec(q); got = l.check ? db.exec(l.check)[0] : (r[r.length - 1] || {columns:[], values:[]}); }
  catch (e) { db.close(); return verdict(false, friendlySql(e.message)); }
  db.close();
  const db2 = freshDb(); const wr = db2.exec(l.answer); const want = l.check ? db2.exec(l.check)[0] : wr[wr.length - 1]; db2.close();
  const label = l.check ? `After your change: <code>${esc(l.check)}</code>` : `Your result · ${got.values.length} row${got.values.length === 1 ? "" : "s"}`;
  const out = `<div class="out"><span class="lab">${label}</span>${got.columns.length ? tableHtml(got.columns, got.values) : "<p>No rows came back.</p>"}</div>`;
  const c = compareSql(got.columns.length ? got : null, want, l.ordered);
  const kw = (l.mustKw || []).filter(k => !new RegExp("\\b" + k + "\\b", "i").test(q));
  if (c.ok && kw.length) return out + verdict(false, `The table looks right, but this lesson practices <code>${esc(kw[0])}</code>. Use it in your answer.`);
  if (!c.ok) return out + verdict(false, l.check && /column|row/.test(c.msg) ? "The table doesn't look right after your change. Check your WHERE condition and the values." : c.msg);
  markDone(); return out + verdict(true, l.check ? "Your change produced exactly the right table." : "Your query returns exactly the right rows.");
}
function dv(v) { return v === null || v === undefined ? "(blank)" : fmt(v); }
function runDax(l, grade = true) {
  const src = $("#code").value.trim();
  if (!src) return `<div class="verdict info"><p>Type your formula first, for example <code>Total = SUM(Sales[Amount])</code>.</p></div>`;
  if (l.kind === "column") {
    let c; try { c = daxColumn(src); } catch (e) { if (e instanceof DaxError) return verdict(false, esc(e.message)); throw e; }
    const want = daxColumn(l.answer), nm = c.name || "New Column";
    const out = `<div class="out"><span class="lab">Sales with your new column</span>${tableHtml(["OrderID","Region","Product","Quantity", nm], PBI_SALES.map((r, i) => [r.OrderID, r.Region, r.Product, r.Quantity, c.values[i] === null ? "(blank)" : c.values[i]]))}</div>`;
    if (!grade) return out + PREVIEW;
    const bad = want.values.findIndex((v, i) => !sameValue(c.values[i], v));
    if (bad >= 0) return out + verdict(false, `Row ${bad + 1} (order ${PBI_SALES[bad].OrderID}) has ${esc(dv(c.values[bad]))}, which isn't the expected value. Check your formula.`);
    const miss = missingParts(src, l.must); if (miss.length) return out + verdict(false, `The values are right, but this lesson practices <code>${esc(miss[0])}</code>.`);
    markDone(); return out + verdict(true, "Every row of your new column is correct.");
  }
  const by = l.by || "Region";
  let m; try { m = daxMeasure(src, by); } catch (e) { if (e instanceof DaxError) return verdict(false, esc(e.message)); throw e; }
  const want = daxMeasure(l.answer, by);
  const rows = m.groups.map(g => [g.key, dv(g.value)]); rows.push({_total:true, cells:["Total", dv(m.total)]});
  const out = `<div class="pbi-out"><div class="card"><span class="lab">Card</span><div class="big">${esc(dv(m.total))}</div><small>${esc(m.name || "Measure")}</small></div><div class="out"><span class="lab">Matrix by ${by}</span>${tableHtml([by, m.name || "Measure"], rows)}</div></div>`;
  if (!grade) return out + PREVIEW;
  const okTotal = sameValue(m.total, want.total), okRows = m.groups.every((g, i) => sameValue(g.value, want.groups[i].value));
  if (!okTotal || !okRows) return out + verdict(false, okTotal ? `The total is right, but some ${by} rows differ. Check how your filters work.` : "That's not the expected result. Check the columns and functions. Try a hint if you're stuck.");
  const miss = missingParts(src, l.must);
  if (miss.length) return out + verdict(false, `The numbers are right, but this lesson practices <code>${esc(miss[0])}</code>. Try writing it that way.`);
  markDone(); return out + verdict(true, `Your measure is right in the card and in every ${by} row.`);
}

const PREVIEW = `<div class="verdict info"><p>This is the example's result. Now change it to solve <b>Your turn</b>.</p></div>`;
function tryExample() {
  const l = byId(state.id), inp = $("#code"); if (!inp) return;
  let c = l.example.code; if (l.track === "excel" && !c.startsWith("=")) c = "=" + c;
  inp.value = c; hideAc();
  $("#result").innerHTML = ({excel:runExcel, sql:runSql, pbi:runDax})[l.track](l, false);
  inp.scrollIntoView({block:"center", behavior:"smooth"});
}
function doSearch(q) {
  const box = $("#sres"); q = q.trim().toLowerCase(); if (q.length < 2) { box.hidden = true; return; }
  const words = q.split(/\s+/);
  const hits = ALL.map(l => { const hay = [l.title, l.group, l.syntax || "", l.answer || "", (l.terms || []).map(t => t[0]).join(" "), (l.body || l.learn || "").replace(/<[^>]+>/g, " ")].join(" ").toLowerCase();
      const title = l.title.toLowerCase(); if (!words.every(w => hay.includes(w))) return null; return {l, score:(title.includes(q) ? 0 : 1) + (modeOf(l) === "theory" ? 0.5 : 0)}; })
    .filter(Boolean).sort((a, b) => a.score - b.score).slice(0, 9);
  const lh = []; for (const tr of Object.keys(LIBRARY)) for (const c of LIBRARY[tr]) for (const it of c.items) { const hay = (it.n + " " + it.d + " " + c.cat).toLowerCase(); if (words.every(w => hay.includes(w))) lh.push({tr, c:c.cat, it}); }
  const libHtml = lh.slice(0, 6).map(x => `<button type="button" class="sr" data-lib="${x.tr}|${esc(x.c)}|${esc(x.it.n)}"><b>${esc(x.it.n)}</b><span>${TRACKS[x.tr]} · Formula Library · ${esc(x.c)}</span></button>`).join("");
  box.innerHTML = (hits.length || lh.length) ? libHtml + hits.map(({l}) => `<button type="button" class="sr" data-res="${l.id}"><b>${esc(l.title)}</b><span>${TRACKS[l.track]} · ${modeOf(l) === "theory" ? "Theory" : "Practice"} · ${LEVELS[lvl(l)][0]}${done.has(l.id) ? " · ✓" : ""}</span></button>`).join("") : `<p class="none">No topics match “${esc(q)}”.</p>`;
  box.hidden = false;
}

/* ---------- Formula Library ---------- */
const libCount = t => LIBRARY[t].reduce((n, c) => n + c.items.length, 0);
function renderLibSide(t) {
  const cats = LIBRARY[t];
  let h = `<div class="progress"><small>${libCount(t)} ${t === "sql" ? "queries" : "formulas"} in ${cats.length} groups</small></div><div class="grp">Formula Library</div>`;
  h += `<button class="ls libcat" data-libcat="all" aria-current="${state.libCat === "all"}"><span class="ck n">${libCount(t)}</span>All</button>`;
  let opts = `<option value="all"${state.libCat === "all" ? " selected" : ""}>All (${libCount(t)})</option>`;
  cats.forEach(c => { h += `<button class="ls libcat" data-libcat="${esc(c.cat)}" aria-current="${state.libCat === c.cat}"><span class="ck n">${c.items.length}</span>${esc(c.cat)}</button>`;
    opts += `<option value="${esc(c.cat)}"${state.libCat === c.cat ? " selected" : ""}>${esc(c.cat)} (${c.items.length})</option>`; });
  $("#side").innerHTML = h; $("#pick").innerHTML = opts; $("label[for=pick]").textContent = "Group";
}
let libVals = {}, libOpt = {sep:",", db:"mssql"};
try { libVals = JSON.parse(localStorage.getItem("dpl-libv") || "{}") || {}; libOpt = Object.assign(libOpt, JSON.parse(localStorage.getItem("dpl-libo") || "{}")); } catch (e) {}
function saveLibVals() { try { localStorage.setItem("dpl-libv", JSON.stringify(libVals)); localStorage.setItem("dpl-libo", JSON.stringify(libOpt)); } catch (e) {} }
const cardKey = (t, cat, it) => t + "|" + cat + "|" + it.n;
function cardValues(key, it) { const v = {}; it.inp.forEach(x => { const s = libVals[key + "|" + x.k]; v[x.k] = (s === undefined || String(s).trim() === "") ? x.def : String(s).trim(); }); return v; }
function cardFormula(t, it, v) { SQLD = libOpt.db; let f; try { f = it.f(v); } catch (e) { f = "Check the boxes above."; } return t === "excel" ? libSep(f, libOpt.sep) : f; }
function renderLibrary() {
  const t = state.lib; renderChrome();
  let h = `<div class="eyebrow">${TRACKS[t]} theory · Formula Library</div><h2>${TRACKS[t]} ${t === "sql" ? "query" : "formula"} library</h2>
  <div class="learn"><p>${t === "excel" ? "Pick a formula, type your cell or range in the small boxes, and the formula is ready. To change a word or number inside the formula (like <b>&quot;Yes&quot;</b> or <b>90</b>), click it and type. Click <b>Copy</b> and paste it into Excel." : t === "pbi" ? "Type your table, column or measure name in the boxes. To change a word or number inside the formula, click it and type. Click <b>Copy</b>, then in Power BI choose <b>New measure</b> (or New column / New table where shown) and paste." : "Type your table and column names in the boxes. To change a word or number inside the query, click it and type. Click <b>Copy</b> and run it in your SQL tool."}</p></div>
  <div class="libtools"><input id="libfilter" type="search" placeholder="Find a formula, e.g. ageing, SLA, VLOOKUP…" value="${esc(state.libFilter)}" aria-label="Find a formula">
  ${t === "excel" ? `<label class="libopt">Separator <select id="libsep"><option value=","${libOpt.sep === "," ? " selected" : ""}>, comma</option><option value=";"${libOpt.sep === ";" ? " selected" : ""}>; semicolon</option></select></label>` : ""}
  ${t === "sql" ? `<label class="libopt">Database <select id="libdb"><option value="mssql"${libOpt.db === "mssql" ? " selected" : ""}>SQL Server</option><option value="mysql"${libOpt.db === "mysql" ? " selected" : ""}>MySQL</option><option value="pg"${libOpt.db === "pg" ? " selected" : ""}>PostgreSQL</option></select></label>` : ""}</div>
  <div id="libcards"></div>`;
  $("#lesson").innerHTML = h; renderLibCards();
}
function renderLibCards() {
  const t = state.lib, box = $("#libcards"); if (!box) return;
  const q = state.libFilter.trim().toLowerCase(), words = q.split(/\s+/).filter(Boolean);
  let out = "";
  for (const c of LIBRARY[t]) {
    if (state.libCat !== "all" && state.libCat !== c.cat && !q) continue;
    const items = c.items.filter(it => !words.length || words.every(w => (it.n + " " + it.d + " " + c.cat).toLowerCase().includes(w)));
    if (!items.length) continue;
    out += `<h3 class="libcat-h">${esc(c.cat)}</h3><div class="fcards">`;
    for (const it of items) {
      const key = cardKey(t, c.cat, it), v = cardValues(key, it);
      out += `<article class="fcard" data-n="${esc(it.n)}" data-key="${esc(key)}"><header><h4>${esc(it.n)}${it.star ? ` <span class="must">Must know</span>` : ""}</h4>${it.note ? `<span class="note">${esc(it.note)}</span>` : ""}</header>
        <p>${esc(it.d)}</p>
        ${it.inp.length ? `<div class="fins">${it.inp.map(x => `<label class="fin"><span>${esc(x.lab)}</span><input data-k="${esc(x.k)}" value="${esc(libVals[key + "|" + x.k] ?? "")}" placeholder="${esc(x.def)}" spellcheck="false" autocomplete="off"></label>`).join("")}</div>` : ""}
        <div class="fcode">${litPre(t, key, it, v)}<button type="button" class="btn small copy" data-copy="1">Copy</button></div>
        <div class="litbar"${libVals[key + "|@lit"] && Object.keys(libVals[key + "|@lit"]).length ? "" : " hidden"}><span>You changed text in this formula.</span><button type="button" class="linkbtn" data-litreset="1">Reset</button></div>
        ${it.l ? `<p class="logic"><b>How it works:</b> ${esc(it.l)}</p>` : ""}</article>`;
    }
    out += `</div>`;
  }
  box.innerHTML = out || `<p class="notice">No formulas match “${esc(q)}”.</p>`;
}
/* Editable text and numbers inside a library formula ("Yes", "No", 90, 75 …).
   Cell/column names come from the boxes; every other fixed text or number can be typed over right in the formula. */
function litTokens(t, f, inputVals) {
  const Q = t === "sql" ? "'" : '"', skipQ = t === "sql" ? '"' : "'", toks = [], ins = new Set(inputVals.flatMap(x => [String(x).trim(), ...String(x).split(/[,;\s]+/)]).filter(Boolean));
  let i = 0, buf = "";
  const push = (lit) => { if (buf) { toks.push({s:buf}); buf = ""; } toks.push(lit); };
  while (i < f.length) {
    const ch = f[i];
    if (ch === Q) { let j = i + 1, inner = ""; while (j < f.length) { if (f[j] === Q) { if (f[j + 1] === Q) { inner += Q; j += 2; continue; } break; } inner += f[j++]; }
      if (j >= f.length) { buf += f.slice(i); break; }
      if (ins.has(inner)) buf += f.slice(i, j + 1); else push({q:Q, v:inner}); i = j + 1; continue; }
    if (ch === skipQ || (ch === "[" && t !== "excel")) { const end = ch === "[" ? "]" : ch, j = f.indexOf(end, i + 1); const k = j < 0 ? f.length : j + 1; buf += f.slice(i, k); i = k; continue; }
    if (t === "sql" && ch === "-" && f[i + 1] === "-") { const j = f.indexOf("\n", i); const k = j < 0 ? f.length : j; buf += f.slice(i, k); i = k; continue; }
    if (/[0-9]/.test(ch) && !/[A-Za-z0-9_$.:!@#\]]/.test(f[i - 1] || "")) {
      const m = /^\d+(\.\d+)?/.exec(f.slice(i)), nx = f[i + m[0].length] || "";
      if (!/[A-Za-z0-9_:($\[]/.test(nx) && !ins.has(m[0])) { push({n:true, v:m[0]}); i += m[0].length; continue; }
      buf += m[0]; i += m[0].length; continue;
    }
    buf += ch; i++;
  }
  if (buf) toks.push({s:buf});
  return toks;
}
function litApply(toks, ov) {
  const lits = toks.filter(x => x.v !== undefined);
  lits.forEach((x, i) => { x.i = i; x.orig = x.v; });
  for (const [i, o] of Object.entries(ov || {})) { let x = lits[i]; if (!x || x.orig !== o.from) x = lits.find(y => y.orig === o.from && y.v === y.orig); if (x) x.v = o.to; }
  return toks;
}
const litText = toks => toks.map(x => x.s !== undefined ? x.s : x.q ? x.q + x.v.split(x.q).join(x.q + x.q) + x.q : (x.v.trim() === "" ? "0" : x.v)).join("");
let litCE = null;
function litPre(t, key, it, v) {
  const f = cardFormula(t, it, v);
  if (litCE === null) { const d = document.createElement("div"); try { d.contentEditable = "plaintext-only"; litCE = d.contentEditable === "plaintext-only" ? "plaintext-only" : "true"; } catch (e) { litCE = "true"; } }
  const toks = litApply(litTokens(t, f, Object.values(v)), libVals[key + "|@lit"]);
  const html = toks.map(x => x.s !== undefined ? esc(x.s) : `${x.q ? esc(x.q) : ""}<span class="lit${x.n ? " num" : ""}${x.v !== x.orig ? " changed" : ""}" contenteditable="${litCE}" spellcheck="false" role="textbox" aria-label="Edit ${x.n ? "number" : "text"} in formula" data-i="${x.i}" data-from="${esc(x.orig)}" title="Click to change">${esc(x.v)}</span>${x.q ? esc(x.q) : ""}`).join("");
  return `<pre data-f="${esc(litText(toks))}">${html}</pre>`;
}
function litCardRefs(card) { const key = card.dataset.key, [t, cat] = key.split("|"); const it = LIBRARY[t].find(c => c.cat === cat).items.find(x => x.n === card.dataset.n); return {key, t, it}; }
function litInput(el) {
  const card = el.closest(".fcard"), {key, t, it} = litCardRefs(card), ov = libVals[key + "|@lit"] = libVals[key + "|@lit"] || {};
  const val = el.textContent.replace(/\n/g, ""), from = el.dataset.from;
  if (el.classList.contains("num") && val !== "" && !/^-?\d*\.?\d*$/.test(val)) el.classList.add("bad"); else el.classList.remove("bad");
  if (val === from) delete ov[el.dataset.i]; else ov[el.dataset.i] = {from, to:val};
  el.classList.toggle("changed", val !== from);
  saveLibVals();
  const toks = litApply(litTokens(t, cardFormula(t, it, cardValues(key, it)), Object.values(cardValues(key, it))), ov);
  card.querySelector("pre").dataset.f = litText(toks);
  card.querySelector(".litbar").hidden = !Object.keys(ov).length;
}
function litReset(btn) {
  const card = btn.closest(".fcard"), {key, t, it} = litCardRefs(card);
  delete libVals[key + "|@lit"]; saveLibVals();
  card.querySelector("pre").outerHTML = litPre(t, key, it, cardValues(key, it)); card.querySelector(".litbar").hidden = true;
}
document.addEventListener("input", e => { if (e.target.classList && e.target.classList.contains("lit")) litInput(e.target); });
document.addEventListener("keydown", e => { if (e.target.classList && e.target.classList.contains("lit") && e.key === "Enter") { e.preventDefault(); e.target.blur(); } }, true);
document.addEventListener("paste", e => { if (e.target.classList && e.target.classList.contains("lit") && litCE === "true") { e.preventDefault(); document.execCommand("insertText", false, (e.clipboardData.getData("text/plain") || "").replace(/\s*\n\s*/g, " ")); } });
document.addEventListener("click", e => { const r = e.target.closest("[data-litreset]"); if (r) litReset(r); });

function libCardInput(el) {
  const card = el.closest(".fcard"), key = card.dataset.key, [t, cat] = key.split("|");
  const it = LIBRARY[t].find(c => c.cat === cat).items.find(x => x.n === card.dataset.n);
  libVals[key + "|" + el.dataset.k] = el.value; saveLibVals();
  card.querySelector("pre").outerHTML = litPre(t, key, it, cardValues(key, it));
}
function copyFormula(btn) {
  const pre = btn.closest(".fcode").querySelector("pre"); if (!pre) return;
  const text = pre.dataset.f ?? pre.textContent, done_ = () => { btn.textContent = "Copied ✓"; setTimeout(() => { btn.textContent = "Copy"; }, 1500); };
  const fallback = () => { const r = document.createRange(); r.selectNodeContents(pre); const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r); try { if (document.execCommand("copy")) { done_(); return; } } catch (e) {} btn.textContent = "Press Ctrl+C"; };
  let settled = false; const ok = () => { settled = true; done_(); }, bad = () => { if (!settled) { settled = true; fallback(); } };
  try { navigator.clipboard.writeText(text).then(ok, bad); setTimeout(bad, 1200); } catch (e) { bad(); }
}

/* ---------- Insert helpers ---------- */
function insertText(t) {
  const inp = $("#code"); if (!inp) return;
  let s = document.activeElement === inp ? inp.selectionStart : inp.value.length, e = document.activeElement === inp ? inp.selectionEnd : s;
  const v = inp.value, prev = v.slice(0, s);
  if (/[\w\])]$/.test(prev) && !/^[\]),]/.test(t)) t = " " + t;
  inp.value = prev + t + v.slice(e); const c = s + t.length; inp.focus(); inp.setSelectionRange(c, c); hideAc(); updateArgTip();
}

/* ---------- Events ---------- */
let lastPointer = "mouse";
document.addEventListener("pointerdown", e => {
  lastPointer = e.pointerType || "mouse";
  const acb = e.target.closest("#ac .acit"); if (acb) { e.preventDefault(); acceptAc(+acb.dataset.i); return; }
  if (e.pointerType !== "mouse" || e.button !== 0) return;
  const td = e.target.closest(".sheet td[data-ref]"), colh = e.target.closest(".sheet th.colh");
  if (td && $("#code")) { e.preventDefault(); const ref = td.dataset.ref;
    if (e.shiftKey && state.point) insertRef(rangeRef(state.point.anchor, ref), state.point.anchor); else insertRef(ref, ref);
    state.drag = true; return; }
  if (colh && $("#code")) { e.preventDefault(); const c = colh.dataset.col; insertRef(`${c}2:${c}11`, c + "2"); return; }
  if (e.target.closest("[data-ins]")) e.preventDefault();
});
document.addEventListener("pointerover", e => {
  if (!state.drag) return; const td = e.target.closest(".sheet td[data-ref]");
  if (td && state.point) insertRef(rangeRef(state.point.anchor, td.dataset.ref), state.point.anchor);
});
document.addEventListener("pointerup", () => { state.drag = false; });
document.addEventListener("click", e => {
  if (!e.target.closest(".themewrap")) closeThemePop();
  if (e.target.closest("[data-home]")) { goHome(); return; }
  if (e.target.closest("#menubtn")) { toggleMenu(); return; }
  if (e.target.closest("#scrim")) { closeDrawer(); return; }
  if (e.target.closest("#focusbtn")) { toggleFocus(); return; }
  if (e.target.closest("#themebtn")) { const p = $("#themepop"); p.hidden ? openThemePop() : closeThemePop(); return; }
  const tm = e.target.closest("[data-tmode]"); if (tm) { theme.mode = tm.dataset.tmode; saveTheme(); applyTheme(); return; }
  const ta = e.target.closest("[data-accent]"); if (ta) { theme.accent = ta.dataset.accent; saveTheme(); applyTheme(); return; }
  const ts = e.target.closest("[data-tsize]"); if (ts) { theme.size = ts.dataset.tsize; saveTheme(); applyTheme(); return; }
  if (e.target.closest("#kbdbtn")) { $("#kbdpop").hidden = false; return; }
  if (e.target.closest("#kbdclose") || e.target.id === "kbdpop") { $("#kbdpop").hidden = true; return; }
  const lo = e.target.closest("[data-libopen]"); if (lo) { openLibrary(lo.dataset.libopen); return; }
  const sw = e.target.closest("[data-start]"); if (sw) { const [tr, md] = sw.dataset.start.split("|"); state.home = false; switchTo(md, tr, "basic"); return; }
  if (lastPointer !== "mouse") {
    const td = e.target.closest(".sheet td[data-ref]"), colh = e.target.closest(".sheet th.colh");
    if (td && $("#code")) { const ref = td.dataset.ref, inp = $("#code");
      if (state.point && state.point.single && inp.selectionStart === state.point.end) insertRef(rangeRef(state.point.anchor, ref), state.point.anchor); else insertRef(ref, ref); return; }
    if (colh && $("#code")) { const c = colh.dataset.col; insertRef(`${c}2:${c}11`, c + "2"); return; }
  }
  const ins = e.target.closest("[data-ins]"); if (ins && $("#code")) { e.preventDefault(); insertText(ins.dataset.ins); return; }
  const tab = e.target.closest(".tab"); if (tab) { switchTo(state.mode, tab.dataset.t); return; }
  const mb = e.target.closest("[data-mode]"); if (mb) { switchTo(mb.dataset.mode, curTrack()); return; }
  const lb = e.target.closest("[data-level]"); if (lb) { switchTo(state.mode, curTrack(), lb.dataset.level); return; }
  const lc = e.target.closest("[data-libcat]"); if (lc) { state.libCat = lc.dataset.libcat; renderLibrary(); $("#main").scrollIntoView({block:"start", behavior:"smooth"}); return; }
  const cp = e.target.closest("[data-copy]"); if (cp) { copyFormula(cp); return; }
  const sl = e.target.closest("[data-lib]"); if (sl) { const [tr, cat, n] = sl.dataset.lib.split("|"); $("#search").value = ""; $("#sres").hidden = true; state.libFilter = ""; openLibrary(tr, cat); setTimeout(() => { const el = document.querySelector(`.fcard[data-n="${CSS.escape(n)}"]`); if (el) { el.scrollIntoView({block:"center"}); el.classList.add("flash"); } }, 50); return; }
  const sr = e.target.closest("[data-res]"); if (sr) { $("#search").value = ""; $("#sres").hidden = true; go(sr.dataset.res); return; }
  if (!e.target.closest(".search")) $("#sres").hidden = true;
  if (e.target.closest("#tryex")) { tryExample(); return; }
  if (e.target.closest("#markread")) { markDone(); $("#result").innerHTML = verdict(true, "Chapter marked as read."); return; }
  const lsb = e.target.closest(".ls"); if (lsb) { go(lsb.dataset.id); return; }
  const g = e.target.closest("[data-go]"); if (g) { go(g.dataset.go); return; }
  const l = byId(state.id);
  if (e.target.closest("#run")) { hideAc(); const fn = {excel:runExcel, sql:runSql, pbi:runDax}[l.track]; $("#result").innerHTML = fn(l); return; }
  if (e.target.closest("#hint")) { if (state.hint < l.hints.length) { $("#hintbox").insertAdjacentHTML("beforeend", `<p><b>Hint ${state.hint + 1}:</b> ${esc(l.hints[state.hint])}</p>`); state.hint++; } if (state.hint >= l.hints.length) $("#hint").disabled = true; return; }
  if (e.target.closest("#reveal")) { $("#hintbox").insertAdjacentHTML("beforeend", `<p class="answerline"><b>Answer:</b> <code>${esc(l.answer)}</code> <button class="btn" id="useans">Put it in the box</button></p>`); $("#reveal").disabled = true; return; }
  if (e.target.closest("#useans")) { $("#code").value = l.answer; $("#code").focus(); return; }
  if (!e.target.closest(".editor")) hideAc();
  const opt = e.target.closest(".opt");
  if (opt) {
    const qi = +opt.dataset.q, oi = +opt.dataset.o, q = l.questions[qi], fs = opt.closest(".q");
    fs.querySelectorAll(".opt").forEach(b => b.classList.remove("right", "wrong"));
    opt.classList.add(oi === q.a ? "right" : "wrong");
    const why = fs.querySelector(".why"); why.hidden = false; why.textContent = (oi === q.a ? "Correct. " : "Not quite. ") + q.why;
    fs.dataset.ok = oi === q.a ? "1" : "";
    if ([...document.querySelectorAll(".q")].every(f => f.dataset.ok === "1")) { markDone(); $("#result").innerHTML = verdict(true, l.mode === "theory" ? "Chapter complete. You answered every question correctly." : "You answered every question correctly."); }
  }
});
document.addEventListener("input", e => { if (e.target.id === "search") { doSearch(e.target.value); return; }
  if (e.target.id === "libfilter") { state.libFilter = e.target.value; renderLibCards(); return; }
  if (e.target.closest(".fin")) { libCardInput(e.target); return; } if (e.target.id !== "code") return; state.point = null; markPick(null); updateAc(); updateArgTip(); });
document.addEventListener("keyup", e => { if (e.target.id === "code" && ["ArrowLeft","ArrowRight","Home","End"].includes(e.key)) { updateAc(); updateArgTip(); } });
document.addEventListener("focusout", e => { if (e.target.id === "code") setTimeout(() => { if (document.activeElement !== $("#code")) hideAc(); }, 120); });
document.addEventListener("keydown", e => {
  if (e.target.id === "search") { if (e.key === "Enter") { const f = $("#sres .sr"); if (f) f.click(); } if (e.key === "Escape") $("#sres").hidden = true; return; }
  if (e.target.id !== "code") return;
  const l = byId(state.id), open = state.ac.items.length && !$("#ac").hidden;
  if (open && e.key === "ArrowDown") { e.preventDefault(); moveAc(1); return; }
  if (open && e.key === "ArrowUp") { e.preventDefault(); moveAc(-1); return; }
  if (open && (e.key === "Tab" || (e.key === "Enter" && !e.ctrlKey && !e.metaKey))) { e.preventDefault(); acceptAc(); return; }
  if (open && e.key === "Escape") { e.preventDefault(); hideAc(); return; }
  if (e.key === "Enter" && (l.track === "excel" || e.ctrlKey || e.metaKey)) { e.preventDefault(); $("#run").click(); }
});
$("#pick").addEventListener("change", e => { if (state.lib) { state.libCat = e.target.value; renderLibrary(); } else go(e.target.value); });
document.addEventListener("change", e => { if (e.target.id === "libsep") { libOpt.sep = e.target.value; saveLibVals(); renderLibCards(); } if (e.target.id === "libdb") { libOpt.db = e.target.value; saveLibVals(); renderLibCards(); } });

/* ---------- v2: home, theme, focus, shortcuts, celebrations ---------- */
const $app = $("#app");
function scrollTopSoft() { window.scrollTo({top:0, behavior:matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth"}); }
function animateLesson() { const el = $("#lesson"); el.classList.remove("enter"); void el.offsetWidth; el.classList.add("enter"); }
function goHome() { state.home = true; closeDrawer(); renderHome(); animateLesson(); scrollTopSoft(); try { localStorage.setItem("dpl-view", "home"); } catch (e) {} }
function closeDrawer() { $app.classList.remove("drawer"); $("#scrim").hidden = true; }
const isNarrow = () => matchMedia("(max-width: 960px)").matches;
let ui = {railOff:false, focus:false};
try { ui = Object.assign(ui, JSON.parse(localStorage.getItem("dpl-ui") || "{}")); } catch (e) {}
function saveUi() { try { localStorage.setItem("dpl-ui", JSON.stringify(ui)); } catch (e) {} }
function applyUi() { $app.classList.toggle("rail-off", ui.railOff || ui.focus); $app.classList.toggle("focus", ui.focus); $("#focusbtn").setAttribute("aria-pressed", ui.focus); }
function toggleMenu() { if (isNarrow()) { const open = !$app.classList.contains("drawer"); $app.classList.toggle("drawer", open); $("#scrim").hidden = !open; if (open) $("#rail").scrollTop = 0; } else { if (ui.focus) ui.focus = false; else ui.railOff = !ui.railOff; saveUi(); applyUi(); } }
function toggleFocus() { ui.focus = !ui.focus; saveUi(); applyUi(); toast(ui.focus ? "Focus mode on. Press F to exit." : "Focus mode off"); }

/* Theme */
const ACCENTS = {blue:{l:["#1F5FBF","#E7EEF9","#FFFFFF"], d:["#7DA6EE","#1C2A44","#0E121A"]}, teal:{l:["#0F766E","#DDF1EE","#FFFFFF"], d:["#4FC1B8","#123230","#0E121A"]},
  purple:{l:["#5B45B5","#ECE8F9","#FFFFFF"], d:["#A898F0","#262040","#0E121A"]}, crimson:{l:["#B4233C","#F9E4E8","#FFFFFF"], d:["#F08A9C","#3A1820","#0E121A"]},
  graphite:{l:["#334155","#E5E8ED","#FFFFFF"], d:["#C3CAD6","#262C37","#0E121A"]}};
let theme = {mode:"system", accent:"blue", size:"m"};
try { theme = Object.assign(theme, JSON.parse(localStorage.getItem("dpl-theme") || "{}")); } catch (e) {}
const root = document.documentElement; let hostTheme = root.getAttribute("data-theme"); let settingTheme = false;
function saveTheme() { try { localStorage.setItem("dpl-theme", JSON.stringify(theme)); } catch (e) {} }
function isDarkNow() { const a = root.getAttribute("data-theme"); return a ? a === "dark" : matchMedia("(prefers-color-scheme: dark)").matches; }
function setAccentVars() { const a = ACCENTS[theme.accent] || ACCENTS.blue, v = isDarkNow() ? a.d : a.l; root.style.setProperty("--accent", v[0]); root.style.setProperty("--accent-soft", v[1]); root.style.setProperty("--accent-ink", v[2]); }
function applyTheme() {
  settingTheme = true;
  if (theme.mode === "system") { if (hostTheme) root.setAttribute("data-theme", hostTheme); else root.removeAttribute("data-theme"); }
  else root.setAttribute("data-theme", theme.mode);
  settingTheme = false;
  setAccentVars();
  root.style.setProperty("--fs", {s:"15px", m:"16px", l:"17.5px"}[theme.size] || "16px");
  document.querySelectorAll("[data-tmode]").forEach(b => b.setAttribute("aria-pressed", b.dataset.tmode === theme.mode));
  document.querySelectorAll("[data-tsize]").forEach(b => b.setAttribute("aria-pressed", b.dataset.tsize === theme.size));
  $("#swatches").innerHTML = Object.keys(ACCENTS).map(k => `<button type="button" class="swatch" data-accent="${k}" aria-pressed="${k === theme.accent}" title="${k[0].toUpperCase() + k.slice(1)}" style="--sw:${(isDarkNow() ? ACCENTS[k].d : ACCENTS[k].l)[0]}"></button>`).join("");
}
matchMedia("(prefers-color-scheme: dark)").addEventListener("change", applyTheme);
new MutationObserver(() => { if (settingTheme) return; if (theme.mode === "system") hostTheme = root.getAttribute("data-theme"); setAccentVars(); }).observe(root, {attributes:true, attributeFilter:["data-theme"]});
function openThemePop() { $("#themepop").hidden = false; $("#themebtn").setAttribute("aria-expanded", "true"); }
function closeThemePop() { const p = $("#themepop"); if (p && !p.hidden) { p.hidden = true; $("#themebtn").setAttribute("aria-expanded", "false"); } }

/* Toast + confetti */
let toastT = null;
function toast(msg) { const t = $("#toast"); t.textContent = msg; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => { t.hidden = true; }, 2600); }
function confetti(n) {
  if (matchMedia("(prefers-reduced-motion: reduce)").matches) return;
  const c = $("#confetti"), ctx = c.getContext("2d"), W = c.width = innerWidth, H = c.height = innerHeight;
  const cols = [getComputedStyle(root).getPropertyValue("--accent").trim() || "#1F5FBF", "#1D7A4E", "#E5B455", "#2C5DA8", "#B4233C"];
  const ps = Array.from({length:n}, () => ({x:W / 2 + (Math.random() - .5) * W * .3, y:H * .35, vx:(Math.random() - .5) * 12, vy:-Math.random() * 11 - 4, r:Math.random() * 6 + 3, c:cols[Math.floor(Math.random() * cols.length)], a:Math.random() * 6, va:(Math.random() - .5) * .3}));
  const t0 = performance.now();
  (function frame(t) { const k = (t - t0) / 1400; ctx.clearRect(0, 0, W, H); if (k > 1) return;
    ps.forEach(p => { p.vy += .35; p.x += p.vx; p.y += p.vy; p.a += p.va; ctx.save(); ctx.globalAlpha = 1 - k; ctx.translate(p.x, p.y); ctx.rotate(p.a); ctx.fillStyle = p.c; ctx.fillRect(-p.r / 2, -p.r / 4, p.r, p.r / 2); ctx.restore(); });
    requestAnimationFrame(frame); })(t0);
}

/* Streak */
const dayKey = d => d.toISOString().slice(0, 10);
function recordDay() { try { const days = new Set(JSON.parse(localStorage.getItem("dpl-days") || "[]")); days.add(dayKey(new Date())); localStorage.setItem("dpl-days", JSON.stringify([...days].slice(-400))); } catch (e) {} }
function streak() { let days; try { days = new Set(JSON.parse(localStorage.getItem("dpl-days") || "[]")); } catch (e) { return 0; }
  const d = new Date(); if (!days.has(dayKey(d))) d.setDate(d.getDate() - 1); let n = 0; while (days.has(dayKey(d))) { n++; d.setDate(d.getDate() - 1); } return n; }

/* Home dashboard */
function ringSvg(pct) { const C = 2 * Math.PI * 28, off = C * (1 - pct / 100); return `<svg viewBox="0 0 68 68" aria-hidden="true"><circle class="bg" cx="34" cy="34" r="28"/><circle class="fg" cx="34" cy="34" r="28" stroke-dasharray="${C.toFixed(1)}" stroke-dashoffset="${C.toFixed(1)}" data-off="${off.toFixed(1)}"/></svg>`; }
function renderHome() {
  renderChrome(); $app.classList.add("home");
  const pc = (a, b) => b ? Math.round(a / b * 100) : 0;
  const th = THEORY_ALL.filter(x => done.has(x.id)).length, pr = LESSONS.filter(x => done.has(x.id)).length;
  const hr = new Date().getHours(), greet = hr < 12 ? "Good morning" : hr < 17 ? "Good afternoon" : "Good evening";
  const last = byId(store.last) && (th + pr) ? byId(state.id) : null;
  let h = `<div class="home-hero"><div><div class="eyebrow">${greet}</div><h2>${th + pr ? "Welcome back" : "Welcome to Data Practice Lab"}</h2><p>Learn Excel, SQL and Power BI from the basics to expert level. Read a theory chapter, practise it in a real sheet, database or data model, and use the formula library at work.</p></div></div>`;
  if (last) h += `<div class="continue"><p><small>Continue where you left off</small><b>${esc(last.title)}</b> <span class="muted">· ${TRACKS[last.track]} ${modeOf(last) === "theory" ? "theory" : "practice"}</span></p><button type="button" class="btn primary" data-go="${last.id}">Continue →</button></div>`;
  else h += `<div class="continue"><p><small>New here?</small><b>Start with Excel theory, chapter 1</b></p><button type="button" class="btn primary" data-go="te1">Start learning →</button></div>`;
  h += `<div class="stats"><div class="stat"><small>Theory chapters read</small><b>${th}</b> <span>of ${THEORY_ALL.length}</span></div><div class="stat"><small>Practice lessons done</small><b>${pr}</b> <span>of ${LESSONS.length}</span></div><div class="stat"><small>Day streak</small><b>${streak()}</b> <span>day${streak() === 1 ? "" : "s"}</span></div><div class="stat"><small>Formula library</small><b>${Object.keys(LIBRARY).reduce((n, t) => n + libCount(t), 0)}</b> <span>ready to copy</span></div></div>`;
  h += `<h3 class="home-h">Your tracks</h3><div class="tcards">`;
  for (const t of Object.keys(TRACKS)) {
    const tA = THEORY_ALL.filter(x => x.track === t), pA = LESSONS.filter(x => x.track === t);
    const td = tA.filter(x => done.has(x.id)).length, pd = pA.filter(x => done.has(x.id)).length;
    const nextT = tA.find(x => !done.has(x.id)), nextP = pA.find(x => !done.has(x.id));
    const ring = (pct, label, go) => `<button type="button" class="ring" ${go ? `data-go="${go}"` : ""} title="${label}">${ringSvg(pct)}<div class="pct">${pct}%</div><small>${label}</small></button>`;
    h += `<article class="tcard" data-t="${t}"><header><span class="dot">${{excel:"XL", sql:"SQL", pbi:"BI"}[t]}</span><div><h3>${TRACKS[t]}</h3><small>${tA.length} chapters · ${pA.length} lessons · ${libCount(t)} formulas</small></div></header>
      <div class="rings">${ring(pc(td, tA.length), `Theory<br>${td}/${tA.length}`, nextT && nextT.id)}${ring(pc(pd, pA.length), `Practice<br>${pd}/${pA.length}`, nextP && nextP.id)}${ring(pc(td + pd, tA.length + pA.length), "Overall", null)}</div>
      <div class="row"><button type="button" class="btn small" data-start="${t}|theory">Theory</button><button type="button" class="btn small" data-start="${t}|practice">Practice</button><button type="button" class="btn small" data-libopen="${t}">Formula Library</button></div></article>`;
  }
  h += `</div><h3 class="home-h">How to use this course</h3><ol class="steps"><li><span><b>Read Basic theory</b>: short chapters with steps and live examples.</span></li><li><span><b>Practise</b>: type real formulas, queries and measures; you get instant feedback and hints.</span></li><li><span><b>Go Advanced</b> when Basic is complete, for expert topics used at work and in interviews.</span></li><li><span><b>Use the Formula Library</b> at work: type your cell or column and copy a ready formula.</span></li></ol>`;
  $("#lesson").innerHTML = h;
  requestAnimationFrame(() => requestAnimationFrame(() => document.querySelectorAll(".ring .fg").forEach(c => c.setAttribute("stroke-dashoffset", c.dataset.off))));
}
/* Shell class follows home state */
const _renderChrome = renderChrome;
renderChrome = function () { _renderChrome(); $app.classList.toggle("home", !!state.home); document.getElementById("nav-home").setAttribute("aria-current", !!state.home); };

/* Shortcuts */
document.addEventListener("keydown", e => {
  if (e.key === "Escape") { closeThemePop(); $("#kbdpop").hidden = true; closeDrawer(); }
  if (e.ctrlKey || e.metaKey || e.altKey) return;
  const tg = e.target; if (["INPUT","TEXTAREA","SELECT"].includes(tg.tagName) || tg.isContentEditable) { if (e.key === "Escape") tg.blur(); return; }
  const k = e.key.toLowerCase();
  const pager = [...document.querySelectorAll(".pager [data-go]")];
  if (k === "n") { const b = pager.find(x => x.textContent.trim().endsWith("→")); if (b) { e.preventDefault(); b.click(); } }
  else if (k === "p") { const b = pager.find(x => x.textContent.trim().startsWith("←")); if (b) { e.preventDefault(); b.click(); } }
  else if (k === "h") { e.preventDefault(); goHome(); }
  else if (k === "/") { e.preventDefault(); $("#search").focus(); }
  else if (e.key === "\\") { e.preventDefault(); toggleMenu(); }
  else if (k === "f") { e.preventDefault(); toggleFocus(); }
  else if (k === "t") { e.preventDefault(); $("#themepop").hidden ? openThemePop() : closeThemePop(); }
  else if (e.key === "?") { e.preventDefault(); $("#kbdpop").hidden = false; }
});
matchMedia("(max-width: 960px)").addEventListener("change", () => closeDrawer());

applyTheme(); applyUi();
try { state.home = localStorage.getItem("dpl-view") !== "lesson" || true; } catch (e) { state.home = true; }
renderLesson(); animateLesson();

/* Hooks for the website layer (login, certificates, leaderboard) */
window.DPL = {done, get THEORY(){ return THEORY_ALL; }, get LESSONS(){ return LESSONS; }, TRACKS, LEVELS, lvl, modeOf, toast, confetti, goHome,
  stats() { const per = {}; for (const t of Object.keys(TRACKS)) { per[t] = {}; for (const v of ["basic","advanced"]) { const th = THEORY_ALL.filter(x => x.track === t && lvl(x) === v), pr = LESSONS.filter(x => x.track === t && lvl(x) === v);
      per[t][v] = {theory:[th.filter(x => done.has(x.id)).length, th.length], practice:[pr.filter(x => done.has(x.id)).length, pr.length]}; } }
    return {theory:THEORY_ALL.filter(x => done.has(x.id)).length, practice:LESSONS.filter(x => done.has(x.id)).length, per}; }};
