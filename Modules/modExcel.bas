Attribute VB_Name = "modExcel"
Public Sub Recordset2Excel(rstSource As ADODB.Recordset)

    Dim xlsApp As Excel.Application
    Dim xlsWBook As Excel.Workbook
    Dim xlsWSheet As Excel.Worksheet
    Dim i, j As Long
    
    ' Get or Create Excel Object
    On Error Resume Next
    Set xlsApp = GetObject(, "Excel.Application")


    If Err.Number <> 0 Then
        Set xlsApp = New Excel.Application
            Err.Clear
    End If

    
    ' Create WorkSheet
    Set xlsWBook = xlsApp.Workbooks.Add
    Set xlsWSheet = xlsWBook.ActiveSheet
    
    ' Show Excel
    xlsApp.Visible = True
    
    
    ' Export ColumnHeaders

    For j = 0 To rstSource.Fields.Count
        xlsWSheet.Cells(2, j + 1) = rstSource.Fields(j).Name
    Next j

    
    ' Export Data
    rstSource.MoveFirst


    For i = 1 To rstSource.RecordCount


        For j = 0 To rstSource.Fields.Count
            xlsWSheet.Cells(i + 2, j + 1) = rstSource.Fields(j).Value
        Next j

        rstSource.MoveNext
    Next i

    rstSource.MoveFirst
    ' Autofit column headers

    For i = 1 To rstSource.Fields.Count
        xlsWSheet.Columns(i).AutoFit
    Next i

    ' Move to first cell to unselect
    xlsWSheet.Range("A1").Select
    
    
    
    Set xlsApp = Nothing
    Set xlsWBook = Nothing
    Set xlsWSheet = Nothing
End Sub

