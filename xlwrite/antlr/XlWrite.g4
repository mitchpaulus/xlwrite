grammar XlWrite;

file : item* EOF ;

// One or more selections (their union), optional filters, then one or more
// formatting elements. There are no separators: every selection, filter and
// element starts with a distinct token, so the parser can always tell where
// one part ends and the next begins.
item : selection+ filter* element+ ;

// A selection without a sheet uses the sheet of the previous selection in the
// same item, or the active sheet if there is none.
selection : sheet? ref ;

sheet
  : STRING        # QuotedSheet   // "My Sheet" A1
  | SHEET_PREFIX  # PrefixSheet   // 'My Sheet'!A1 or Data!A1
  ;

ref
  : CELL              # CellRef
  | CELL_RANGE        # CellRangeRef
  | COL_RANGE         # ColumnRangeRef
  | ROW_RANGE         # RowRangeRef
  | TABLE tablePart?  # TableRef       // A1:* auto-detects the table bounds
  ;

tablePart
  : 'header'                  # TableHeader
  | 'body'                    # TableBody
  | ('col' | 'column') STRING # TableColumn
  | 'lastrow'                 # TableLastRow
  ;

// Filters narrow the selection when the script runs (static).
filter : PIPE condition ;

element
  : action  # ActionElement
  | cfRule  # CfElement
  ;

action
  : 'bold' toggle?                                # BoldAction
  | 'italic' toggle?                              # ItalicAction
  | 'underline' toggle?                           # UnderlineAction
  | ('strike' | 'strikethrough') toggle?          # StrikeAction
  | 'wrap' toggle?                                # WrapAction
  | ('fill' | 'bg') colorValue                    # FillAction
  | ('color' | 'fg') colorValue                   # FontColorAction
  | 'font' STRING                                 # FontAction
  | ('fontsize' | 'fs') number                    # FontSizeAction
  | ('width' | 'w') number                        # WidthAction
  | ('height' | 'h') number                       # HeightAction
  | ('format' | 'fmt') STRING                     # NumberFormatAction
  | 'align' hAlign                                # AlignAction
  | 'valign' vAlign                               # VAlignAction
  | 'border' borderSide* borderStyle? colorValue? # BorderAction
  ;

// Conditional formatting, written into the workbook as live Excel rules.
cfRule
  : 'cond' condition LCURLY action+ STOP? RCURLY      # CondRule
  | 'cond' 'scale' colorValue colorValue colorValue?  # ScaleRule
  | 'cond' 'databar' colorValue                       # DataBarRule
  | 'cond' 'icons' iconSet                            # IconsRule
  ;

// Shared by filters and conditional formatting.
condition
  : compareOp operand                  # CompareCond
  | 'between' operand operand          # BetweenCond
  | 'notbetween' operand operand       # NotBetweenCond
  | 'contains' STRING                  # ContainsCond
  | 'notcontains' STRING               # NotContainsCond
  | 'begins' STRING                    # BeginsCond
  | 'ends' STRING                      # EndsCond
  | 'blank'                            # BlankCond
  | 'nonblank'                         # NonBlankCond
  | 'error'                            # ErrorCond
  | 'noerror'                          # NoErrorCond
  | 'top' INT PERCENT?                 # TopCond
  | 'bottom' INT PERCENT?              # BottomCond
  | 'aboveavg'                         # AboveAvgCond
  | 'belowavg'                         # BelowAvgCond
  | 'duplicate'                        # DuplicateCond
  | 'unique'                           # UniqueCond
  | datePeriod                         # DateCond
  | 'formula' STRING                   # FormulaCond  // relative to the top-left cell
  ;

compareOp : '=' | '!=' | '<>' | '<' | '<=' | '>' | '>=' ;

operand
  : number  # NumberOperand
  | STRING  # StringOperand
  | CELL    # CellOperand
  ;

datePeriod
  : 'today' | 'yesterday' | 'tomorrow' | 'last7days'
  | 'thisweek' | 'lastweek' | 'nextweek'
  | 'thismonth' | 'lastmonth' | 'nextmonth'
  ;

iconSet : 'arrows' | 'flags' | 'traffic' | 'stars' | 'symbols' | 'ratings' ;

toggle : 'on' | 'off' | 'true' | 'false' ;

hAlign : 'left' | 'center' | 'right' ;
vAlign : 'top' | 'middle' | 'center' | 'bottom' ;

borderSide  : 'all' | 'outline' | 'inside' | 'top' | 'bottom' | 'left' | 'right' ;
borderStyle : 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'none' ;

colorValue
  : HEXCOLOR        # HexColor
  | 'rgb' INT INT INT # RgbColor
  | knownColor      # NamedColor
  ;

knownColor
  : 'red' | 'green' | 'blue' | 'black' | 'white' | 'gray' | 'grey'
  | 'orange' | 'yellow' | 'purple'
  ;

number : MINUS? (INT | FLOAT) ;

STOP : 'stop' ;
LCURLY : '{' ;
RCURLY : '}' ;
PIPE : '|' ;
PERCENT : '%' ;
MINUS : '-' ;

// Selectors are single tokens so they never compete with numeric arguments.
TABLE        : C ':' '*' ;
CELL_RANGE   : C ':' C ;
COL_RANGE    : '$'? COL ':' '$'? COL ;
ROW_RANGE    : '$'? ROW ':' '$'? ROW ;
CELL         : C ;
SHEET_PREFIX : '\'' (~'\'' | '\'\'')+ '\'!'
             | [a-zA-Z_] [a-zA-Z0-9_.]* '!'
             ;

fragment C   : '$'? COL '$'? ROW ;
fragment COL : [a-zA-Z] [a-zA-Z]? [a-zA-Z]? ;
fragment ROW : [0-9]+ ;

HEXCOLOR : '#' HEX HEX HEX HEX HEX HEX ;
fragment HEX : [0-9a-fA-F] ;

FLOAT : [0-9]+ '.' [0-9]* ;
INT   : [0-9]+ ;

STRING : '"' (ESC | .)*? '"' ;
fragment ESC : '\\"' | '\\\\' ;

WS            : [ \t\r\n]+ -> skip ;
BLOCK_COMMENT : '/*' .*? '*/' -> skip ;
LINE_COMMENT  : '//' ~[\r\n]* -> skip ;

// Lowest priority, so misspelled keywords are reported as a whole word by the parser.
UNKNOWN_WORD : [a-zA-Z_] [a-zA-Z0-9_]* ;
