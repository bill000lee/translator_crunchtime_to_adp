Attribute VB_Name = "basExport"
Sub Export()
    Dim wsDataOut As Worksheet
    Dim wsAllowancesOut As Worksheet
    Dim wsTemp As Worksheet
    Dim exportPath As String
    Dim companyCode As String
    Dim exportDate As String
    Dim fileName As String
    Dim lastRowDataOut As Long
    Dim lastRowAllowancesOut As Long
    Dim lastColDataOut As Long
    Dim lastColAllowancesOut As Long
    Dim lastRowTemp As Long
    Dim i As Long, j As Long
    Dim cellValue As String
    Dim fileNum As Integer
    Dim line As String
    
    ' Set the output worksheets
    Set wsDataOut = ThisWorkbook.Sheets("ElementsOut")
    Set wsAllowancesOut = ThisWorkbook.Sheets("AllowancesOut")
    
    ' Create a temporary worksheet
    Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTemp.Name = "TempSheet"
    
    ' Convert sort columns to text
    wsTemp.Columns(2).NumberFormat = "@"
    wsTemp.Columns(4).NumberFormat = "@"
    wsTemp.Columns(5).NumberFormat = "@"
    wsTemp.Columns(6).NumberFormat = "@"
    
    ' Get the Company Code from the first row of data
    companyCode = wsDataOut.Cells(2, 1).Value
    
    ' Get the current date in YYYYMMDD format
    exportDate = Format(Date, "YYYYMMDD")
    
    ' Create the file name with .dat extension
    fileName = "paymast" & ".dat"
    
    ' Set the export path to C:\ADP folder
    exportPath = "C:\ADP\" & fileName
    
    ' Check if the folder exists, if not, create it
    If Dir("C:\ADP\", vbDirectory) = "" Then
        MkDir "C:\ADP\"
    End If
    
    ' Get the last row and column of data for both sheets
    lastRowDataOut = wsDataOut.Cells(wsDataOut.Rows.Count, "A").End(xlUp).row
    lastColDataOut = wsDataOut.Cells(1, wsDataOut.Columns.Count).End(xlToLeft).Column
    lastRowAllowancesOut = wsAllowancesOut.Cells(wsAllowancesOut.Rows.Count, "A").End(xlUp).row
    lastColAllowancesOut = wsAllowancesOut.Cells(1, wsAllowancesOut.Columns.Count).End(xlToLeft).Column
    
    ' Copy data from both sheets to the temporary sheet
    wsDataOut.Range("A1").Resize(lastRowDataOut, lastColDataOut).Copy wsTemp.Range("A1")
    wsAllowancesOut.Range("A2").Resize(lastRowAllowancesOut - 1, lastColAllowancesOut).Copy wsTemp.Range("A" & lastRowDataOut + 1)
    
    ' Get the last row of the temporary sheet
    lastRowTemp = wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).row
    
    ' Sort the combined data by three columns
    wsTemp.Sort.SortFields.Clear
    wsTemp.Range("A1").Resize(lastRowTemp, lastColDataOut).Sort Key1:=Range("B1"), Order1:=xlAscending, _
                                                                Key2:=Range("L1"), Order2:=xlAscending, _
                                                                Key3:=Range("M1"), Order2:=xlAscending, _
                                                                Header:=xlYes
                                                                
    ' Open the file for writing
    fileNum = FreeFile
    Open exportPath For Output As #fileNum
    
    ' Loop through each row and column of TempSheet to write data to the file
    For i = 2 To lastRowTemp ' Start from row 2 to skip the header
        line = ""
        ' test = wsTemp.Cells(i, 3).Value
        For j = 1 To lastColDataOut - 2  ' Don't ecport the sort columns
            cellValue = wsTemp.Cells(i, j).Value
            If j = lastColDataOut - 2 Then
                line = line & cellValue
            Else
                line = line & cellValue & ","
            End If
        Next j
        Print #fileNum, line
    Next i
    
    ' Close the file
    Close #fileNum
    
    ' Delete the temporary sheet
    Application.DisplayAlerts = False
    wsTemp.Delete
    Application.DisplayAlerts = True
    
    MsgBox "Data exported successfully to C:\ADP\", vbInformation
End Sub

