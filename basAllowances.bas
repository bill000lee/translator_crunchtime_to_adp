Attribute VB_Name = "basAllowances"
Sub Allowances()
    Dim wsDataIn As Worksheet
    Dim wsLookup As Worksheet
    Dim wsAllowancesOut As Worksheet
    Dim wsHolidays As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim key As String
    Dim dictKey As Variant
    Dim parts As Variant
    Dim i As Long

    Dim companyCode As Variant
    Dim employeeCode As String
    Dim payrollExportCode As String
    Dim recordType As String
    Dim entryDate As String
    Dim glNumber As Long
    Dim dateSortKey As String
    Dim suffix As Variant
    Dim DateIn As Date
    Dim DateOut As Date
    Dim dateFrom As String
    Dim dateTo As String
    Dim dateTimeIn As Date
    Dim dateTimeOut As Date
    Dim dayOfWeek As Integer
    Dim allowanceHours As Double
    Dim allowanceCode As String
    
    ' Set worksheets
    Set wsDataIn = ThisWorkbook.Sheets("DataIn")
    Set wsLookup = ThisWorkbook.Sheets("Lookup")
    Set wsAllowancesOut = ThisWorkbook.Sheets("AllowancesOut")
    Set wsHolidays = ThisWorkbook.Sheets("Holidays")
    
    ' Clear existing data in the output worksheet
    wsAllowancesOut.Cells.Clear
    
    ' Add Headers for the columns in AllowancesOut
    wsAllowancesOut.Cells(1, 1).Value = "Company Code"
    wsAllowancesOut.Cells(1, 2).Value = "Employee Code"
    wsAllowancesOut.Cells(1, 3).Value = "Record Type"
    wsAllowancesOut.Cells(1, 4).Value = "Entry Date"
    wsAllowancesOut.Cells(1, 5).Value = "Allowance Code"
    wsAllowancesOut.Cells(1, 6).Value = "Amount/Units"
    wsAllowancesOut.Cells(1, 7).Value = "Cost Centre"
    wsAllowancesOut.Cells(1, 8).Value = "Notation 1"
    wsAllowancesOut.Cells(1, 9).Value = "Notation 2"
    wsAllowancesOut.Cells(1, 10).Value = "From Date"
    wsAllowancesOut.Cells(1, 11).Value = "To Date"
    wsAllowancesOut.Cells(1, 12).Value = "Week Sort Key"
    wsAllowancesOut.Cells(1, 13).Value = "Date Sort Key"
       
    ' Convert sort columns to text
    wsAllowancesOut.Columns(2).NumberFormat = "@"
    wsAllowancesOut.Columns(4).NumberFormat = "@"
    wsAllowancesOut.Columns(5).NumberFormat = "@"
    wsAllowancesOut.Columns(6).NumberFormat = "@"
    wsAllowancesOut.Columns(10).NumberFormat = "@"
    wsAllowancesOut.Columns(11).NumberFormat = "@"
    wsAllowancesOut.Columns(12).NumberFormat = "0"
    wsAllowancesOut.Columns(13).NumberFormat = "0"
    
    ' Create dictionary for summary data
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Get the last row of DataIn
    lastRow = wsDataIn.Cells(wsDataIn.Rows.Count, "A").End(xlUp).row

    ' Loop through DataIn and summarize data
    For i = 2 To lastRow
        
        ' Lookup the Company Code using the OwnershipEntity
        companyCode = Application.VLookup(wsDataIn.Cells(i, 1).Value, wsLookup.Range("CompanyCode"), 2, False)
        If IsError(companyCode) Then companyCode = "ERROR"
        
        ' Get the Employee Code
        employeeCode = wsDataIn.Cells(i, 4).Value
        
        ' Insert the Record Type
        recordType = "A"
        
        ' Insert the Entry Date in format DDMMYY
        entryDate = Format(DateSerial(2000 + Mid(wsDataIn.Cells(i, 3).Value, 1, 2), Mid(wsDataIn.Cells(i, 3).Value, 3, 2), Mid(wsDataIn.Cells(i, 3).Value, 5, 2)), "DDMMYY")
        
        ' Calculate the allowance code and allowance amount/units
        DateIn = DateSerial(2000 + Mid(wsDataIn.Cells(i, 7).Value, 1, 2), Mid(wsDataIn.Cells(i, 7).Value, 3, 2), Mid(wsDataIn.Cells(i, 7).Value, 5, 2))
        DateOut = DateSerial(2000 + Mid(wsDataIn.Cells(i, 8).Value, 1, 2), Mid(wsDataIn.Cells(i, 8).Value, 3, 2), Mid(wsDataIn.Cells(i, 8).Value, 5, 2))
        dateTimeIn = DateIn + wsDataIn.Cells(i, 9).Value
        dateTimeOut = DateOut + wsDataIn.Cells(i, 10).Value
        dayOfWeek = Weekday(dateTimeIn, vbMonday)
        allowanceUnits = 0
        
        ' Get DateFrom and DateTo
        fromDate = Format(DateSerial(2000 + Mid(wsDataIn.Cells(i, 7).Value, 1, 2), Mid(wsDataIn.Cells(i, 7).Value, 3, 2), Mid(wsDataIn.Cells(i, 7).Value, 5, 2)), "DDMMYY")
        toDate = Format(DateSerial(2000 + Mid(wsDataIn.Cells(i, 8).Value, 1, 2), Mid(wsDataIn.Cells(i, 8).Value, 3, 2), Mid(wsDataIn.Cells(i, 8).Value, 5, 2)), "DDMMYY")
        
        ' Check if the date is a holiday
        isHoliday = Not IsError(Application.VLookup(payrollExportCode & Format(DateIn, "YYMMDD"), wsHolidays.Range("A:B"), 2, False))
        
        Select Case dayOfWeek
            Case 1 To 5 ' Monday to Friday
                    ' Allowance Code A101 calculation
                    If Hour(dateTimeIn) < 6 And isHoliday = False Then
                        allowanceCode = "A101"
                        ' Calculate allowance units
                        Do While dateTimeIn < dateTimeOut
                            If Hour(dateTimeIn) < 6 Then
                                allowanceUnits = allowanceUnits + 1
                            End If
                            dateTimeIn = dateTimeIn + TimeSerial(1, 0, 0)
                        Loop
                        allowanceUnits = allowanceUnits * 100
                    End If
                    ' Allowance Code A100 calculation
                    If Hour(dateTimeIn) > 22 And isHoliday = False Then
                        allowanceCode = "A100"
                        ' Calculate allowance units
                        Do While dateTimeIn < dateTimeOut
                            If Hour(dateTimeOut) > 22 Then
                                allowanceUnits = allowanceUnits + 1
                            End If
                            dateTimeIn = dateTimeIn + TimeSerial(1, 0, 0)
                        Loop
                        allowanceUnits = allowanceUnits * 100
                    End If
            Case 6 ' Saturday
            Case 7 ' Sunday
        End Select
        
        If allowanceUnits > 0 Then
             ' Lookup and Concatenate the suffix and PayrollExportCode to form the Cost Centre
            glNumber = CLng(wsDataIn.Cells(i, 6).Value)
            suffix = Application.VLookup(glNumber, wsLookup.Range("CostCodeSuffix"), 2, False)
            If IsError(suffix) Then suffix = "ERROR"
            payrollExportCode = wsDataIn.Cells(i, 2).Value
            costCentre = suffix & payrollExportCode
            
            ' Create the week and date sort keys
            weekSortKey = wsDataIn.Cells(i, 3).Value * 1
            dateSortKey = Format(DateSerial(2000 + Mid(wsDataIn.Cells(i, 7).Value, 1, 2), Mid(wsDataIn.Cells(i, 7).Value, 3, 2), Mid(wsDataIn.Cells(i, 7).Value, 5, 2)), "YYYYMMDD") * 1
            
            'Create dictionary key
            key = companyCode & "|" & employeeCode & "|" & recordType & "|" & entryDate & "|" & allowanceCode & "|" & costCentre & "|" & "" & "|" & "" & "|" & fromDate & "|" & toDate & "|" & weekSortKey & "|" & dateSortKey
            
            
            ' Check if the key exists in the dictionary
            If dict.exists(key) Then
                ' Sum the number of hours
                dict(key) = dict(key) + allowanceUnits
            Else
                ' Add a new entry to the dictionary
                dict.Add key, allowanceUnits
            End If
        End If
    Next i
    
    ' Output the aggregated data to AllowancesOut
    i = 2
    
    For Each dictKey In dict.keys
        parts = Split(dictKey, "|")
        
        If parts(5) > 0 And Left(parts(6), 1) <> "M" Then  ' Only include rows where AllowanceUnits > 0 and Cost Centre does not start with "M
            wsAllowancesOut.Cells(i, 1).Value = parts(0)    ' Company Code
            wsAllowancesOut.Cells(i, 2).Value = parts(1)    ' Employee Code
            wsAllowancesOut.Cells(i, 3).Value = parts(2)    ' Record Type
            wsAllowancesOut.Cells(i, 4).Value = parts(3)    ' Entry Date
            wsAllowancesOut.Cells(i, 5).Value = parts(4)    ' Allowance Code
            wsAllowancesOut.Cells(i, 6).Value = dict(dictKey)    ' Allowance Units
            wsAllowancesOut.Cells(i, 7).Value = parts(5)    ' Cost Centre
            wsAllowancesOut.Cells(i, 8).Value = parts(6)    ' Notation 1
            wsAllowancesOut.Cells(i, 9).Value = parts(7)    ' Notation 2
            wsAllowancesOut.Cells(i, 10).Value = parts(8)    ' Date From
            wsAllowancesOut.Cells(i, 11).Value = parts(9)    ' Date To
            wsAllowancesOut.Cells(i, 12).Value = parts(10)    ' Week Sort Key
            wsAllowancesOut.Cells(i, 13).Value = parts(11)    ' Date Sort Key
            i = i + 1
        End If
    Next dictKey

    ' Auto-fit columns
    wsAllowancesOut.Columns.AutoFit
    
    MsgBox "Data has been successfully transformed, grouped, summarized, and exported to the AllowancesOut worksheet."
End Sub

