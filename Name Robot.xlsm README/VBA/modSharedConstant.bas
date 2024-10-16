Attribute VB_Name = "modSharedConstant"
'@IgnoreModule UndeclaredVariable, ConstantNotUsed, ImplicitlyTypedConst
'@Folder "Lambda.Editor.Shared"

Option Explicit

' Table related constants
Public Const TABLE_ALL_MARKER As String = "[#All]"
Public Const TABLE_HEADERS_MARKER As String = "[#Headers]"
Public Const TABLE_TOTALS_MARKER As String = "[#Totals]"
Public Const TABLE_DATA_MARKER As String = "[#Data]"

' Keyword related constants
Public Const CALC_ERR_KEYWORD As String = "#CALC!"
Public Const REF_ERR_KEYWORD As String = "#REF!"
Public Const NAME_ERR_KEYWORD As String = "#NAME?"
Public Const TRUE_KEYWORD As String = "TRUE"
Public Const FALSE_KEYWORD As String = "FALSE"


' Excel Formulas related constants
Public Const LET_FX_NAME As String = "LET"
Public Const LAMBDA_FX_NAME As String = "LAMBDA"
Public Const ISOMITTED_FX_NAME As String = "ISOMITTED"
Public Const DROP_FX_NAME As String = "DROP"
Public Const TAKE_FX_NAME As String = "TAKE"
Public Const VSTACK_FX_NAME As String = "VSTACK"
Public Const HSTACK_FX_NAME As String = "HSTACK"
Public Const OR_FX_NAME As String = "OR"
Public Const AND_FX_NAME As String = "AND"
Public Const ISBLANK_FX_NAME As String = "ISBLANK"
Public Const CHOOSECOLS_FX_NAME As String = "CHOOSECOLS"
Public Const COLUMNS_FX_NAME As String = "COLUMNS"
Public Const IF_FX_NAME As String = "IF"
Public Const FILTER_FX_NAME As String = "FILTER"
Public Const ROWS_FX_NAME As String = "ROWS"
Public Const NA_FX_NAME As String = "NA"
Public Const SWITCH_FX_NAME As String = "SWITCH"
Public Const MAP_FX_NAME As String = "MAP"
Public Const BYROW_FX_NAME As String = "BYROW"
Public Const TOROW_FX_NAME As String = "TOROW"
Public Const CHOOSEROWS_FX_NAME As String = "CHOOSEROWS"
Public Const ISNUMBER_FX_NAME As String = "ISNUMBER"
Public Const XMATCH_FX_NAME As String = "XMATCH"
Public Const EXPAND_FX_NAME As String = "EXPAND"
Public Const TYPE_FX_NAME As String = "TYPE"
Public Const SORTBY_FX_NAME As String = "SORTBY"
Public Const IFERROR_FX_NAME As String = "IFERROR"
Public Const INDEX_FX_NAME As String = "INDEX"
Public Const SEQUENCE_FX_NAME As String = "SEQUENCE"
Public Const OFFSET_FX_NAME As String = "OFFSET"
Public Const EQUAL_LET_FIRST_PAREN As String = "=LET("
Public Const LET_AND_OPEN_PAREN As String = "LET("
Public Const LAMBDA_AND_OPEN_PAREN As String = "LAMBDA("
Public Const LEFT_BRACE As String = "{"
Public Const RIGHT_BRACE As String = "}"
Public Const LEFT_BRACKET As String = "["
Public Const RIGHT_BRACKET As String = "]"
Public Const ARRAY_CONST_COLUMN_SEPARATOR As String = ","
Public Const ARRAY_CONST_ROW_SEPARATOR As String = ";"
Public Const LIST_SEPARATOR As String = ","

' @Defined Independent Const

Public Const APP_NAME As String = "Lambda Robot"
Public Const DOLLAR_SIGN As String = "$"
Public Const HASH_SIGN As String = "#"
Public Const EQUAL_SIGN As String = "="
Public Const LET_PARTS_VALUE_COL_INDEX As Long = 5
Public Const LAMBDA_PARTS_VALUE_COL_INDEX As Long = 5
Public Const LAMBDA_PARTS_PARAMETER_INDEX_COL_INDEX As Long = 4
Public Const INPUT_CELL_BACKGROUND_COLOR As Long = 13434879
Public Const INPUT_CELL_FONT_COLOR As Long = 16711680
Public Const FONT_COLOR_INDEX As Long = -65536
Public Const MAX_ALLOWED_LET_STEP_NAME_LENGTH As Long = 255
Public Const MAX_LENGTH_OF_FORMULA As Long = 8192
Public Const NEW_LINE As String = vbNewLine
Public Const KEY_VALUE_SEPARATOR As String = " - "
Public Const METADATA_IDENTIFIER As String = "\\"
Public Const THREE_SPACE As String = "   "
Public Const QUOTES As String = """"
Public Const SINGLE_QUOTE As String = "'"
Public Const COMMA As String = ","
Public Const EXCLAMATION_SIGN As String = "!"
Public Const INVOKE_TEXT As String = "<INVOKE>"
Public Const DOT As String = "."
Public Const UNDER_SCORE As String = "_"
Public Const DOUBLE_QUOTE As String = """"
Public Const TILE_FX_NAME As String = "TILE"

Public Const ONE_SPACE As String = " "
Public Const FIRST_PARENTHESIS_OPEN As String = "("
Public Const FIRST_PARENTHESIS_CLOSE As String = ")"
Public Const LEFT_SQUARE_BRACKET As String = "["
Public Const RIGHT_SQUARE_BRACKET As String = "]"

Public Const MAXIMUM_ALLOWABLE_DEPENDENCY_LEVEL As Long = 1048576
Public Const LAST_STEP_NAME As String = "Result"

Public Const ARGUMENT_SEPARATOR As String = "Argument Seperator"
Public Const LET_STEP_NAME_TOKEN As String = "Let Step Name"
Public Const LET_STEP_USED_NAME_TOKEN As String = "Local Name"
Public Const NAMED_RANGE_TOKEN As String = "Named Range"
Public Const FIRST_PAREN_CLOSE_TOKEN As String = "Right Paren"
Public Const LAMBDA_NAME_NOTE_PREFIX As String = "Editing Lambda: "
Public Const NAMED_RANGE_NOTE_PREFIX As String = "Editing Named Range: "
Public Const LAMBDA_NAME_AUDIT_PREFIX As String = "Auditing Lambda: "
Public Const QUOTES_AND_FIRST_PAREN_CLOSE As String = """)"
Public Const LETSTEP_PREFIX As String = "LETStep"
Public Const LETSTEPREF_PREFIX As String = LETSTEP_PREFIX & "Ref"

' Gist Export for github integration
Public Const VBA_SUB_FOLDER_NAME As String = "VBA Code"
Public Const POWER_QUERY_SUB_FOLDER_NAME As String = "Power Queries"
Public Const LAMBDA_SUB_FOLDER_NAME As String = "Lambdas"
Public Const POWER_QUERY_FILE_EXTENSION As String = ".pq"
Public Const LAMBDA_FILE_EXTENSION As String = ".lambda"

' @Calculated Constant.
Public Const LETSTEP_UNDERSCORE_PREFIX As String = LETSTEP_PREFIX & "_"
Public Const LETSTEPREF_UNDERSCORE_PREFIX As String = LETSTEPREF_PREFIX & "_"

Public Const SHEET_NAME_SEPARATOR As String = EXCLAMATION_SIGN
Public Const THREE_NEW_LINE As String = vbNewLine & vbNewLine & vbNewLine
Public Const DYNAMIC_CELL_REFERENCE_SIGN As String = HASH_SIGN

Public Const LAMBDA_NAME_LET_VAR As String = METADATA_IDENTIFIER & "LambdaName"
Public Const COMMAND_NAME_LET_VAR As String = METADATA_IDENTIFIER & "CommandName"
Public Const DESCRIPTION_LET_VAR As String = METADATA_IDENTIFIER & "Description"
Public Const PARAMETERS_LET_VAR As String = METADATA_IDENTIFIER & "Parameters"
Public Const DEPENDENCIES_LET_VAR As String = METADATA_IDENTIFIER & "Dependencies"
Public Const CUSTOMPROPERTIES_LET_VAR As String = METADATA_IDENTIFIER & "CustomProperties"
Public Const SOURCE_NAME_LET_VAR As String = METADATA_IDENTIFIER & "Source"
Public Const GIST_URL_LET_VAR As String = METADATA_IDENTIFIER & "gistURL"

Public Const QUOTES_COMMA_NEWLINE As String = QUOTES & COMMA & NEW_LINE

