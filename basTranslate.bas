Sub Translate()
    Dim wsIn As Worksheet, wsOut As Worksheet, wsLookup As Worksheet, wsADP As Worksheet, wsHolidays As Worksheet
    Dim lastRow As Long, i As Long, dayOfWeek As Integer
    Dim dateTimeIn As Date, dateTimeOut As Date, NumberOfHours As Double
    Dim entryDate As String, DateIn As Date, DateOut As Date, fromDate As String, toDate As String
    Dim companyCode As Variant, suffix As Variant, glNumber As Long
    Dim payrollExportCode As String, costCentre As String, payrollCode As String, payClassCode As String, payRate As Variant
    Dim key As String, dict As Object, dictKey As Variant, parts() As String
    Dim lookupScript As Variant, lookupElement As Variant, isHoliday As Boolean
    Dim employeeCode As Variant
    
    ' Set the input, output, lookup, ADP Pay Class, and Holidays worksheets
    On Error GoTo ErrHandler
    Set wsIn = ThisWorkbook.Sheets("DataIn")
    Set wsOut = ThisWorkbook.Sheets("ElementsOut")
    Set wsLookup = ThisWorkbook.Sheets("Lookup")
    Set wsADP = ThisWorkbook.Sheets("ADP Pay Class")
    Set wsHolidays = ThisWorkbook.Sheets("Holidays")
    On Error GoTo 0
    
    ' Clear existing data in the output worksheet
    wsOut.Cells.Clear
    'Sets the columns to type text
    wsOut.Columns("A:H").NumberFormat = "@"

    
    ' Add headers for the columns in DataOut
    wsOut.Cells(1, 1).Value = "Company Code"
    wsOut.Cells(1, 2).Value = "Employee Code"
    wsOut.Cells(1, 3).Value = "Record Type"
    wsOut.Cells(1, 4).Value = "Entry Date"
    wsOut.Cells(1, 5).Value = "Payroll Code"  'Element
    wsOut.Cells(1, 6).Value = "Number of Hours"
    wsOut.Cells(1, 7).Value = "Pay Class Code"
    wsOut.Cells(1, 8).Value = "Cost Centre"
    wsOut.Cells(1, 9).Value = "From Date"
    wsOut.Cells(1, 10).Value = "To Date"
    wsOut.Cells(1, 11).Value = "Text"
    wsOut.Cells(1, 12).Value = "Week Sort Key"
    wsOut.Cells(1, 13).Value = "Date Sort Key"
    
    ' Convert some columns to text
    wsOut.Columns(2).NumberFormat = "@"
    wsOut.Columns(4).NumberFormat = "@"
    wsOut.Columns(5).NumberFormat = "@"
    wsOut.Columns(6).NumberFormat = "@"
    wsOut.Columns(9).NumberFormat = "@"
    wsOut.Columns(10).NumberFormat = "@"
    wsOut.Columns(12).NumberFormat = "0"
    wsOut.Columns(13).NumberFormat = "0"
    
    ' Get the last row of data in the input worksheet
    lastRow = wsIn.Cells(wsIn.Rows.Count, "A").End(xlUp).row
    
    ' Initialize the dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop through each row in the input worksheet
    For i = 2 To lastRow
    On Error Resume Next ' Continue on error
    
    ' Parse DateIn and DateOut columns formatted as YYMMDD
    DateIn = DateSerial(2000 + Mid(wsIn.Cells(i, 7).Value, 1, 2), Mid(wsIn.Cells(i, 7).Value, 3, 2), Mid(wsIn.Cells(i, 7).Value, 5, 2))
    DateOut = DateSerial(2000 + Mid(wsIn.Cells(i, 8).Value, 1, 2), Mid(wsIn.Cells(i, 8).Value, 3, 2), Mid(wsIn.Cells(i, 8).Value, 5, 2))
    employeeCode = wsIn.Cells(i, 4).Value
    ' Combine DateIn and TimeIn to a datetime
    dateTimeIn = DateIn + wsIn.Cells(i, 9).Value
    ' Combine DateOut and TimeOut to a datetime
    dateTimeOut = DateOut + wsIn.Cells(i, 10).Value
    ' Calculate the number of hours
    NumberOfHours = Round((dateTimeOut - dateTimeIn) * 24 * 10000, 0)
    
    If NumberOfHours > 0 Then
        ' Get the day of the week for DateIn
        dayOfWeek = Weekday(DateIn, vbMonday)
        ' Convert WeekEndingDate from YYMMDD to DDMMYY and copy to Entry Date column in DataOut
        entryDate = Format(DateSerial(2000 + Mid(wsIn.Cells(i, 3).Value, 1, 2), Mid(wsIn.Cells(i, 3).Value, 3, 2), Mid(wsIn.Cells(i, 3).Value, 5, 2)), "DDMMYY")
        ' Convert DateIN and DateOut from YYMMDD to DDMMYY and copy to fromDate and toDate
        fromDate = Format(DateSerial(2000 + Mid(wsIn.Cells(i, 7).Value, 1, 2), Mid(wsIn.Cells(i, 7).Value, 3, 2), Mid(wsIn.Cells(i, 7).Value, 5, 2)), "DDMMYY")
        toDate = Format(DateSerial(2000 + Mid(wsIn.Cells(i, 8).Value, 1, 2), Mid(wsIn.Cells(i, 8).Value, 3, 2), Mid(wsIn.Cells(i, 8).Value, 5, 2)), "DDMMYY")

        ' Lookup the Company Code using the OwnershipEntity
        companyCode = Application.VLookup(wsIn.Cells(i, 1).Value, wsLookup.Range("CompanyCode"), 2, False)
        If IsError(companyCode) Then companyCode = ""
        
        ' Lookup the payRate
        payRate = wsIn.Cells(i, 11).Value
        ' Lookup script for the suffix
        Select Case dayOfWeek
            Case 1 To 5 ' Monday to Friday
                suffix = Application.VLookup(payRate, wsADP.Range("A:J"), 10, False)
            Case 6 ' Saturday
                suffix = Application.VLookup(payRate, wsADP.Range("B:J"), 9, False)
            Case 7 ' Sunday
                suffix = Application.VLookup(payRate, wsADP.Range("C:J"), 8, False)
        End Select
        
        ''glNumber = CLng(wsIn.Cells(i, 6).Value)
        ''suffix = Application.VLookup(glNumber, wsLookup.Range("CostCodeSuffix"), 2, False)
        
        
        If IsError(suffix) Then suffix = ""
        ' Concatenate the suffix and PayrollExportCode to form the Cost Centre
        payrollExportCode = wsIn.Cells(i, 2).Value
        costCentre = suffix & payrollExportCode
        
        
        
        ' Lookup script for the Pay Class Code
        Select Case dayOfWeek
            Case 1 To 5 ' Monday to Friday
                lookupScript = Application.VLookup(payRate, wsADP.Range("A:D"), 4, False)
            Case 6 ' Saturday
                lookupScript = Application.VLookup(payRate, wsADP.Range("B:D"), 3, False)
            Case 7 ' Sunday
                lookupScript = Application.VLookup(payRate, wsADP.Range("C:D"), 2, False)
        End Select
        
        ' Handle cases where the payRate is either empty or zero
        If IsEmpty(payRate) Or payRate = 0 Then
            payClassCode = "Y99"
        Else
            payRate = CDbl(Format(payRate, "0.00")) ' Convert PayRate to a double with two decimal places
            ' Perform the VLOOKUP
            payClassCode = lookupScript
            If IsError(payClassCode) Then
                payClassCode = "ERR" ' Handle the case where VLookup doesn't find a match
                wsIn.Cells(i, 11).Interior.Color = RGB(255, 0, 0) ' Highlight the cell causing the issue
            End If
            If IsError(payClassCode) Then
                payClassCode = "ERR" ' Handle the case where VLookup doesn't find a match
                wsIn.Cells(i, 11).Interior.Color = RGB(255, 0, 0) ' Highlight the cell causing the issue
            End If
        End If
        
        ' Create the week and date sort keys
        weekSortKey = wsIn.Cells(i, 3).Value * 1
        dateSortKey = Format(DateSerial(2000 + Mid(wsIn.Cells(i, 7).Value, 1, 2), Mid(wsIn.Cells(i, 7).Value, 3, 2), Mid(wsIn.Cells(i, 7).Value, 5, 2)), "YYYYMMDD") * 1

        ' Determine the Payroll Code based on the day of the week
        Select Case dayOfWeek
            Case 1 To 5 ' Monday to Friday
                ' Check if the date is a holiday
                isHoliday = Not IsError(Application.VLookup(payrollExportCode & Format(DateIn, "YYMMDD"), wsHolidays.Range("A:B"), 2, False))
                If isHoliday Then
                    payrollCode = Application.VLookup(payRate, wsADP.Range("A:I"), 9, False)
                    key = companyCode & "|" & employeeCode & "|" & "E" & "|" & entryDate & "|" & payrollCode & "|" & payClassCode & "|" & costCentre & "|" & fromDate & "|" & toDate & "|" & "" & "|" & weekSortKey & "|" & dateSortKey
                Else
                    payrollCode = Application.VLookup(payRate, wsADP.Range("A:I"), 6, False)
                    key = companyCode & "|" & employeeCode & "|" & "E" & "|" & entryDate & "|" & payrollCode & "|" & payClassCode & "|" & costCentre & "|" & fromDate & "|" & toDate & "|" & "" & "|" & weekSortKey & "|" & dateSortKey
                End If
                lookupScript = Application.VLookup(payRate, wsADP.Range("A:D"), 4, False)
            Case 6 ' Saturday
                ' Check if the date is a holiday
                isHoliday = Not IsError(Application.VLookup(payrollExportCode & Format(DateIn, "YYMMDD"), wsHolidays.Range("A:B"), 2, False))
                If isHoliday Then
                    payrollCode = Application.VLookup(payRate, wsADP.Range("B:I"), 8, False)
                    key = companyCode & "|" & employeeCode & "|" & "E" & "|" & entryDate & "|" & payrollCode & "|" & payClassCode & "|" & costCentre & "|" & fromDate & "|" & toDate & "|" & "" & "|" & weekSortKey & "|" & dateSortKey
                Else
                    payrollCode = Application.VLookup(payRate, wsADP.Range("B:G"), 6, False)
                    key = companyCode & "|" & employeeCode & "|" & "E" & "|" & entryDate & "|" & payrollCode & "|" & payClassCode & "|" & costCentre & "|" & fromDate & "|" & toDate & "|" & "" & "|" & weekSortKey & "|" & dateSortKey
                End If
                lookupScript = Application.VLookup(payRate, wsADP.Range("B:D"), 3, False)
            Case 7 ' Sunday
                ' Check if the date is a holiday
                isHoliday = Not IsError(Application.VLookup(payrollExportCode & Format(DateIn, "YYMMDD"), wsHolidays.Range("A:B"), 2, False))
                If isHoliday Then
                    payrollCode = Application.VLookup(payRate, wsADP.Range("C:I"), 7, False)
                    key = companyCode & "|" & employeeCode & "|" & "E" & "|" & entryDate & "|" & payrollCode & "|" & payClassCode & "|" & costCentre & "|" & fromDate & "|" & toDate & "|" & "" & "|" & weekSortKey & "|" & dateSortKey
                Else
                    payrollCode = Application.VLookup(payRate, wsADP.Range("C:H"), 6, False)
                    key = companyCode & "|" & employeeCode & "|" & "E" & "|" & entryDate & "|" & payrollCode & "|" & payClassCode & "|" & costCentre & "|" & fromDate & "|" & toDate & "|" & "" & "|" & weekSortKey & "|" & dateSortKey
                End If
                lookupScript = Application.VLookup(payRate, wsADP.Range("C:D"), 2, False)
        End Select
        
        ' Check if the key exists in the dictionary
        If dict.exists(key) Then
            ' Sum the number of hours
            dict(key) = dict(key) + NumberOfHours
        Else
            ' Add a new entry to the dictionary
            dict.Add key, NumberOfHours
        End If
    End If
    
    On Error GoTo 0 ' Reset error handling
    Next i

    
    ' Output the aggregated data to DataOut
    i = 2
    For Each dictKey In dict.keys
        parts = Split(dictKey, "|")
        If UBound(parts) < 8 Then ' Corrected error check
            MsgBox "Error: Key does not have enough parts: " & dictKey, vbCritical
            Exit Sub
        End If
        wsOut.Cells(i, 1).Value = parts(0) ' Company Code
        wsOut.Cells(i, 2).Value = parts(1) ' Employee Code
        wsOut.Cells(i, 3).Value = parts(2) ' Record Type
        wsOut.Cells(i, 4).Value = parts(3) ' Entry Date
        wsOut.Cells(i, 5).Value = parts(4) ' Payroll Code
        wsOut.Cells(i, 6).Value = dict(dictKey) ' Number of Hours
        wsOut.Cells(i, 7).Value = parts(5) ' Pay Class Code
        wsOut.Cells(i, 8).Value = parts(6) ' Cost Centre
        wsOut.Cells(i, 9).Value = parts(7) ' From Date
        wsOut.Cells(i, 10).Value = parts(8) ' To Date
        wsOut.Cells(i, 11).Value = parts(9) ' Blank
        wsOut.Cells(i, 12).Value = parts(10) ' Week Sort Key
        wsOut.Cells(i, 13).Value = parts(11) ' Date Sort Key
        i = i + 1
    Next dictKey
    
    ' Auto-fit columns
    wsOut.Columns.AutoFit
    
    MsgBox "Data converted, grouped, and sorted successfully!", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
