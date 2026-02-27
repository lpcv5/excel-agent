"""Excel COM constants — mapped from Excel VBA enumerations."""

# XlListObjectSourceType
xlSrcRange = 1
xlSrcQuery = 3

# XlListObjectHasHeaders
xlYes = 1
xlNo = 2
xlGuess = 0

# XlCalculation
xlCalculationAutomatic = -4105
xlCalculationManual = -4135

# XlDirection
xlUp = -4162
xlDown = -4121
xlToLeft = -4159
xlToRight = -4161

# XlReferenceStyle
xlA1 = 1
xlR1C1 = -4150

# XlCellType
xlCellTypeLastCell = 11

# XlChartType
xlColumnClustered = 51
xlLine = 4
xlPie = 5
xlBarClustered = 57
xlArea = 1
xlXYScatter = -4169

# XlRowCol
xlColumns = 2
xlRows = 1

# XlChartLocation
xlLocationAsNewSheet = 1
xlLocationAsObject = 2

# Common HRESULT error codes
HRESULT_MAP: dict[int, str] = {
    -2147418111: "RPC_E_CALL_REJECTED (Excel is busy, retry)",
    -2147417848: "RPC_E_DISCONNECTED (Excel process terminated)",
    -2146827284: "TYPE_MISMATCH",
    -2147024809: "E_INVALIDARG (invalid argument)",
    -2146826281: "#REF! error",
    -2146826273: "#VALUE! error",
    -2146826259: "#NAME? error",
    -2146826246: "#N/A error",
    -2146826252: "#DIV/0! error",
    -2146826288: "#NULL! error",
}

# Transient COM errors that are safe to retry
TRANSIENT_HRESULTS = {-2147418111, -2147417848}
