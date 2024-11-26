Attribute VB_Name = "modDailySales"
Public Function GetAgingBalSalesMan(Scode As String) As Double
    strSQL = "SELECT Sum(PRODDTA.F03B11.RPAAP) "
    strSQL = strSQL & "AS OpenAging, PRODDTA.F0101.ABAC07 "
    strSQL = strSQL & "FROM PRODDTA.F03B11 INNER JOIN PRODDTA.F0101 "
    strSQL = strSQL & "ON PRODDTA.F03B11.RPAN8 = PRODDTA.F0101.ABAN8 GROUP "
    strSQL = strSQL & "BY PRODDTA.F0101.ABAC07 HAVING (((PRODDTA.F0101.ABAC07)='" & Scode & "'))"
    
    GetAgingBalSalesMan = 0
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       GetAgingBalSalesMan = tmpRset(0) / 100
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
End Function
Public Function GetDaySales(Scode As String, ReqDate As Date) As Double
    strSQL = "SELECT Sum(PRODDTA.F4211.SDAEXP) "
    strSQL = strSQL & "AS SumOfSDAEXP, PRODDTA.F0101.ABAC07 "
    strSQL = strSQL & "FROM PRODDTA.F4211 INNER JOIN PRODDTA.F0101 "
    strSQL = strSQL & "ON PRODDTA.F4211.SDAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE (((PRODDTA.F4211.SDDGL)=" & JulianDate(ReqDate) & ")) "
    strSQL = strSQL & "AND PRODDTA.F4211.SDSRP1 = 'FG1' "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDCT = 'RI' "
    strSQL = strSQL & "OR (((PRODDTA.F4211.SDDGL)=" & JulianDate(ReqDate) & ")) "
    strSQL = strSQL & "AND PRODDTA.F4211.SDSRP1 = 'FG1' "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDCT = 'CN' "
    strSQL = strSQL & "OR (((PRODDTA.F4211.SDDGL)=0) AND ((PRODDTA.F4211.SDIVD)=" & JulianDate(ReqDate) & ")) "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDCT = 'RI' "
    strSQL = strSQL & "AND PRODDTA.F4211.SDSRP1 = 'FG1' "
    strSQL = strSQL & "OR (((PRODDTA.F4211.SDDGL)=0) AND ((PRODDTA.F4211.SDIVD)=" & JulianDate(ReqDate) & ")) "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDCT = 'CN' "
    strSQL = strSQL & "AND PRODDTA.F4211.SDSRP1 = 'FG1' "
    strSQL = strSQL & "GROUP BY PRODDTA.F0101.ABAC07 "
    strSQL = strSQL & "HAVING (((PRODDTA.F0101.ABAC07)='" & Scode & "')) "
    
    GetDaySales = 0
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       GetDaySales = tmpRset(0) / 100
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
    
End Function
Public Function GetMonthSales(Scode As String, ReqDate As Date) As Double
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
    strSQL = strSQL & "OR PRODDTA.F4211.SDDGL >= " & JulianDate(StartDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDGL <= " & JulianDate(ReqDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDSRP1 = 'FG1' "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDCT = 'CN' "
    strSQL = strSQL & "OR PRODDTA.F4211.SDDGL =0 AND PRODDTA.F4211.SDIVD >= " & JulianDate(StartDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDIVD <= " & JulianDate(ReqDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDSRP1 = 'FG1' "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDCT = 'RI' "
    strSQL = strSQL & "OR PRODDTA.F4211.SDDGL =0 AND PRODDTA.F4211.SDIVD >= " & JulianDate(StartDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDIVD <= " & JulianDate(ReqDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDSRP1 = 'FG1' "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDCT = 'CN' "
    strSQL = strSQL & "GROUP BY PRODDTA.F0101.ABAC07 "
    strSQL = strSQL & "HAVING PRODDTA.F0101.ABAC07 ='" & Scode & "'"
    
    GetMonthSales = 0
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       GetMonthSales = tmpRset(0) / 100
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
End Function
Public Function GetDayColl(Scode As String, ReqDate As Date) As Double
Dim StartDate As Date
Dim EndDate As Date

StartDate = CDate(Month(ReqDate) & "/01/" & Year(ReqDate))
EndDate = CDate(Month(ReqDate) & "/" & LastDayInMonth(Year(ReqDate), Month(ReqDate)) & "/" & Year(ReqDate))

    strSQL = "SELECT Sum(PRODDTA.F03B14.RZPAAP) "
    strSQL = strSQL & "AS SumOfRZPAAP "
    strSQL = strSQL & "FROM PRODDTA.F03B14 INNER JOIN PRODDTA.F0101 "
    strSQL = strSQL & "ON PRODDTA.F03B14.RZAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE PRODDTA.F03B14.RZDGJ=" & JulianDate(ReqDate) & " "
    strSQL = strSQL & "GROUP BY PRODDTA.F0101.ABAC07 HAVING PRODDTA.F0101.ABAC07 ='" & Scode & "'"
    
    GetDayColl = 0
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       GetDayColl = tmpRset(0) / 100
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
End Function


Public Function GetMonthColl(Scode As String, ReqDate As Date) As Double
Dim StartDate As Date
Dim EndDate As Date

StartDate = CDate(Month(ReqDate) & "/01/" & Year(ReqDate))
EndDate = CDate(Month(ReqDate) & "/" & LastDayInMonth(Year(ReqDate), Month(ReqDate)) & "/" & Year(ReqDate))

    strSQL = "SELECT Sum(PRODDTA.F03B14.RZPAAP) "
    strSQL = strSQL & "AS SumOfRZPAAP "
    strSQL = strSQL & "FROM PRODDTA.F03B14 INNER JOIN PRODDTA.F0101 "
    strSQL = strSQL & "ON PRODDTA.F03B14.RZAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE PRODDTA.F03B14.RZDGJ >=" & JulianDate(StartDate) & " "
    strSQL = strSQL & "AND PRODDTA.F03B14.RZDGJ <=" & JulianDate(ReqDate) & " "
    strSQL = strSQL & "GROUP BY PRODDTA.F0101.ABAC07 HAVING PRODDTA.F0101.ABAC07 ='" & Scode & "'"
    
    GetMonthColl = 0
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       GetMonthColl = (tmpRset(0) / 100)
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
End Function
Public Function GetMonthDNCN(Scode As String, ReqDate As Date) As Double
Dim StartDate As Date
Dim EndDate As Date

StartDate = CDate(Month(ReqDate) & "/01/" & Year(ReqDate))
EndDate = CDate(Month(ReqDate) & "/" & LastDayInMonth(Year(ReqDate), Month(ReqDate)) & "/" & Year(ReqDate))

    strSQL = "SELECT Sum(PRODDTA.F03B11.RPAG) "
    strSQL = strSQL & "AS SumOfRPAG "
    strSQL = strSQL & "FROM PRODDTA.F03B11 INNER JOIN PRODDTA.F0101 "
    strSQL = strSQL & "ON PRODDTA.F03B11.RPAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE PRODDTA.F03B11.RPDGJ >= " & JulianDate(StartDate) & " "
    strSQL = strSQL & "AND PRODDTA.F03B11.RPDGJ <= " & JulianDate(EndDate) & " "
    strSQL = strSQL & "AND PRODDTA.F03B11.RPDCT = 'RM' Or PRODDTA.F03B11.RPDCT = 'DN'  "
    strSQL = strSQL & "AND PRODDTA.F03B11.RPDGJ >= " & JulianDate(StartDate) & " "
    strSQL = strSQL & "AND PRODDTA.F03B11.RPDGJ <= " & JulianDate(ReqDate) & " "
    strSQL = strSQL & "GROUP BY PRODDTA.F0101.ABAC07 HAVING PRODDTA.F0101.ABAC07 ='" & Scode & "'"

    GetMonthDNCN = 0
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       GetMonthDNCN = tmpRset(0) / 100
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
End Function

Public Function GetDayDNCN(Scode As String, ReqDate As Date) As Double
Dim StartDate As Date
Dim EndDate As Date

StartDate = CDate(Month(ReqDate) & "/01/" & Year(ReqDate))
EndDate = CDate(Month(ReqDate) & "/" & LastDayInMonth(Year(ReqDate), Month(ReqDate)) & "/" & Year(ReqDate))

    strSQL = "SELECT Sum(PRODDTA.F03B11.RPAG) "
    strSQL = strSQL & "AS SumOfRPAG "
    strSQL = strSQL & "FROM PRODDTA.F03B11 INNER JOIN PRODDTA.F0101 "
    strSQL = strSQL & "ON PRODDTA.F03B11.RPAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE PRODDTA.F03B11.RPDGJ = " & JulianDate(ReqDate) & " "
    strSQL = strSQL & "AND PRODDTA.F03B11.RPDCT = 'RM' Or PRODDTA.F03B11.RPDCT = 'DN'  "
    strSQL = strSQL & "AND PRODDTA.F03B11.RPDGJ = " & JulianDate(ReqDate) & " "
    strSQL = strSQL & "GROUP BY PRODDTA.F0101.ABAC07 HAVING PRODDTA.F0101.ABAC07 ='" & Scode & "'"

    GetDayDNCN = 0
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       GetDayDNCN = tmpRset(0) / 100
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
End Function

Public Function GetMonthBlanketSales(Scode As String, ReqDate As Date) As Double
Dim StartDate As Date
Dim EndDate As Date

StartDate = CDate(Month(ReqDate) & "/01/" & Year(ReqDate))
EndDate = CDate(Month(ReqDate) & "/" & LastDayInMonth(Year(ReqDate), Month(ReqDate)) & "/" & Year(ReqDate))

    strSQL = "SELECT Sum(PRODDTA.F4211.SDAEXP) "
    strSQL = strSQL & "AS SumOfSDAEXP, PRODDTA.F0101.ABAC07 "
    strSQL = strSQL & "FROM PRODDTA.F4211 INNER JOIN PRODDTA.F0101 "
    strSQL = strSQL & "ON PRODDTA.F4211.SDAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE PRODDTA.F4211.SDTRDJ >= " & JulianDate(StartDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDTRDJ <= " & JulianDate(EndDate) & " "
    strSQL = strSQL & "AND PRODDTA.F4211.SDSRP1 = 'FG1' "
    strSQL = strSQL & "AND PRODDTA.F4211.SDDCTO = 'SB' "
    strSQL = strSQL & "AND PRODDTA.F4211.SDNXTR = '515' "
    strSQL = strSQL & "GROUP BY PRODDTA.F0101.ABAC07 "
    strSQL = strSQL & "HAVING PRODDTA.F0101.ABAC07 ='" & Scode & "'"
    
    GetMonthBlanketSales = 0
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       GetMonthBlanketSales = tmpRset(0) / 100
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
End Function

