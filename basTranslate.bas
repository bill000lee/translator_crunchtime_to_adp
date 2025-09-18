Attribute VB_Name = "basTranslate"
Option Explicit

    ' At the module level
    Private Const COL_OWNERSHIP_ENTITY As Integer = 1
    Private Const COL_PAYROLL_EXPORT_CODE As Integer = 2
    Private Const COL_WEEK_ENDING_DATE As Integer = 3
    Private Const COL_EMPLOYEE_CODE As Integer = 4
    ' More column constants...
    Private Const RECORD_TYPE_ELEMENT As String = "E"
    Private Const DEFAULT_PAY_CLASS_CODE As String = "Y99"
    Private Const ERROR_PAY_CLASS_CODE As String = "ERR"


Sub Translate()
    ' Define all variables with appropriate types
    Dim wsTimeSheetData As Worksheet
    Dim wsPayrollElements As Worksheet
    Dim wsLookup As Worksheet
    Dim wsADP As Worksheet
    Dim wsHolidays As Worksheet
    Dim wsTimeSheetData As Worksheet
    Dim wsPayrollElements As Worksheet
    Dim hoursSummary As Object
    
    ' Initialize worksheets and dictionary
    If Not InitializeWorksheets(wsTimeSheetData, wsPayrollElements, wsLookup, wsADP, wsHolidays) Then Exit Sub
    Set hoursSummary = CreateObject("Scripting.Dictionary")
    
    ' Setup output worksheet
    SetupOutputWorksheet wsPayrollElements
    
    ' Process input data
    ProcessInputData wsTimeSheetData, wsPayrollElements, wsLookup, wsADP, wsHolidays, hoursSummary
    
    ' Output results
    OutputResults wsPayrollElements, hoursSummary
    
    MsgBox "Data converted, grouped, and sorted successfully!", vbInformation
End Sub

Private Function ValidateInputData(wsTimeSheetData As Worksheet) As Boolean
    ' Check if input worksheet has data
    If wsTimeSheetData.Cells(2, 1).Value = "" Then
        MsgBox "No data found in the input worksheet.", vbExclamation
        ValidateInputData = False
        Exit Function
    End If
    
    ' Check for required columns
    Dim requiredColumns As Variant
    requiredColumns = Array("OwnershipEntity", "PayrollExportCode", "WeekEndingDate", "EmployeeCode", "DateIn", "DateOut", "TimeIn", "TimeOut", "PayRate")
    ' Additional validation logic
    
    ValidateInputData = True
End Function

Private Function InitializeWorksheets(ByRef wsTimeSheetData As Worksheet, ByRef wsPayrollElements As Worksheet, _
                                     ByRef wsLookup As Worksheet, ByRef wsADP As Worksheet, _
                                     ByRef wsHolidays As Worksheet) As Boolean
    On Error Resume Next
    Set wsTimeSheetData = ThisWorkbook.Sheets("DataIn")
    Set wsPayrollElements = ThisWorkbook.Sheets("ElementsOut")
    Set wsLookup = ThisWorkbook.Sheets("Lookup")
    Set wsADP = ThisWorkbook.Sheets("ADP Pay Class")
    Set wsHolidays = ThisWorkbook.Sheets("Holidays")
    
    If Err.Number <> 0 Then
        MsgBox "Error initializing worksheets: " & Err.Description, vbCritical
        InitializeWorksheets = False
        Exit Function
    End If
    
    On Error GoTo 0
    InitializeWorksheets = True
End Function

Private Function ConvertYYMMDDToDate(yymmdd As String) As Date
    On Error Resume Next
    ConvertYYMMDDToDate = DateSerial(2000 + CInt(Mid(yymmdd, 1, 2)), _
                                     CInt(Mid(yymmdd, 3, 2)), _
                                     CInt(Mid(yymmdd, 5, 2)))
    If Err.Number <> 0 Then
        ' Handle invalid date format
        Err.Clear
        ConvertYYMMDDToDate = DateSerial(1900, 1, 1) ' Default date
    End If
    On Error GoTo 0
End Function

Private Function ConvertDateToYYMMDD(dateValue As Date) As String
    ConvertDateToYYMMDD = Format(dateValue, "YYMMDD")
End Function

Private Function ConvertDateToDDMMYY(dateValue As Date) As String
    ConvertDateToDDMMYY = Format(dateValue, "DDMMYY")
End Function

Private Function DeterminePayrollCode(payRate As Variant, dayOfWeek As Integer, _
                                     isHoliday As Boolean, wsADP As Worksheet) As String
    On Error Resume Next
    
    Select Case dayOfWeek
        Case 1 To 5 ' Monday to Friday
            If isHoliday Then
                DeterminePayrollCode = Application.VLookup(payRate, wsADP.Range("A:I"), 9, False)
            Else
                DeterminePayrollCode = Application.VLookup(payRate, wsADP.Range("A:I"), 6, False)
            End If
        Case 6 ' Saturday
            If isHoliday Then
                DeterminePayrollCode = Application.VLookup(payRate, wsADP.Range("B:I"), 8, False)
            Else
                DeterminePayrollCode = Application.VLookup(payRate, wsADP.Range("B:G"), 6, False)
            End If
        Case 7 ' Sunday
            If isHoliday Then
                DeterminePayrollCode = Application.VLookup(payRate, wsADP.Range("C:I"), 7, False)
            Else
                DeterminePayrollCode = Application.VLookup(payRate, wsADP.Range("C:H"), 6, False)
            End If
    End Select
    
    If IsError(DeterminePayrollCode) Then DeterminePayrollCode = "ERR"
    On Error GoTo 0
End Function

Private Function BuildDictionaryKey(companyCode As Variant, employeeCode As Variant, _
                                   payrollEntryDate As String, payrollCode As String, _
                                   payClassCode As String, costCentre As String, _
                                   fromDate As String, toDate As String, _
                                   weekSortKey As String, dateSortKey As String) As String
    BuildDictionaryKey = companyCode & "|" & employeeCode & "|" & "E" & "|" & _
                         payrollEntryDate & "|" & payrollCode & "|" & payClassCode & "|" & _
                         costCentre & "|" & fromDate & "|" & toDate & "|" & _
                         "" & "|" & weekSortKey & "|" & dateSortKey
End Function

