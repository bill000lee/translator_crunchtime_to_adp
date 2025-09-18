Attribute VB_Name = "basImport"
Option Explicit

Private Const WORKSHEET_NAME As String = "DataIn"
Private Const DEFAULT_PATH As String = "C:\ADP\"

Sub Import()
    Dim ws As Worksheet
    Dim filePath As String
    Dim lastRow As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "Preparing to import data..."
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets(WORKSHEET_NAME)
    
    ' Clear existing data in the worksheet
    ws.Cells.Clear
    
    ' Set up headers
    SetupWorksheetHeaders ws
    
    ' Get file path
    filePath = SelectCSVFile()
    If filePath = "" Then
        MsgBox "No file selected.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Validate file extension
    If Right(LCase(filePath), 4) <> ".csv" Then
        MsgBox "Selected file is not a CSV file.", vbExclamation
        GoTo CleanExit
    End If
    
    Application.StatusBar = "Importing data from " & filePath & "..."
    
    ' Import the CSV file
    ImportCSVFile ws, filePath
    
    ' Text to Columns
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If lastRow > 1 Then ' Only process if there's data
        ws.Range("A2:A" & lastRow).TextToColumns Destination:=ws.Range("A2"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False
    End If
    
    ' Auto-fit columns
    ws.Columns.AutoFit
    
    MsgBox "Data imported and formatted successfully!", vbInformation
    
CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub SetupWorksheetHeaders(ws As Worksheet)
    Dim headers As Variant
    Dim i As Integer
    
    headers = Array("OwnershipEntity", "PayrollExportCode", "WeekEndingDate", "PayrollID", _
                   "EmployeePositionCode", "GLNumber", "DateIn", "DateOut", _
                   "TimeIn", "TimeOut", "PayRate")
    
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
End Sub

Private Function SelectCSVFile() As String
    Dim fileDialog As fileDialog
    Dim filePath As String
    
    ' Open file dialog to select CSV file
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Select Tab Delimited CSV File"
        .InitialFileName = DEFAULT_PATH
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            filePath = ""
        End If
    End With
    
    SelectCSVFile = filePath
End Function

Private Sub ImportCSVFile(ws As Worksheet, filePath As String)
    On Error GoTo ErrorHandler
    
    ' Import the CSV file
    With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A2"))
        .TextFileParseType = xlDelimited
        .TextFileTabDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileConsecutiveDelimiter = False
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "ImportCSVFile", "Error importing CSV file: " & Err.Description
End Sub