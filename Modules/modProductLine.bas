Attribute VB_Name = "modProductLine"
Public Function GetProductSales(Scode As String, ReqDate As Date, PrdLine As String) As Double
Dim StartDate As Date
Dim EndDate As Date

StartDate = CDate(Month(ReqDate) & "/01/" & Year(ReqDate))
EndDate = CDate(Month(ReqDate) & "/" & LastDayInMonth(Year(ReqDate), Month(ReqDate)) & "/" & Year(ReqDate))

    strSQL = "SELECT Sum(PRODDTA.F4211.SDAEXP) "
    strSQL = strSQL & "AS SumOfSDAEXP, PRODDTA.F0101.ABAC07 "
    strSQL = strSQL & "FROM PRODDTA.F4211 INNER JOIN PRODDTA.F0101 "
    strSQL = strSQL & "ON PRODDTA.F4211.SDAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE PRODDTA.F4211.SDDGL >= " & JulianDate(StartDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDGL <= " & JulianDate(ReqDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDSRP1 = 'FG1' "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDCT = 'RI' "
    strSQL = strSQL & "AND LEFT(TRIM(PRODDTA.F4211.SDLITM),1) = '" & PrdLine & "' "
    strSQL = strSQL & "OR PRODDTA.F4211.SDDGL >= " & JulianDate(StartDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDGL <= " & JulianDate(ReqDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDSRP1 = 'FG1' "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDCT = 'CN' "
    strSQL = strSQL & "AND LEFT(TRIM(PRODDTA.F4211.SDLITM),1) = '" & PrdLine & "' "
    strSQL = strSQL & "GROUP BY PRODDTA.F0101.ABAC07 "
    strSQL = strSQL & "HAVING PRODDTA.F0101.ABAC07 ='" & Scode & "'"
    
    GetProductSales = 0
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       GetProductSales = tmpRset(0) / 100
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
End Function

