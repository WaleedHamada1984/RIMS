Attribute VB_Name = "modBUSalesReport"
Public Function GetYTDSalesBUC(Line As String, Yr As Integer, Mth As Integer, MCU As String)
    strSQL = "SELECT PRODDTA.F4101.IMGLPT AS LINE, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, Sum(PRODDTA.F55STAB.ASSOQS)/100000 "
    strSQL = strSQL & "AS QTY, Sum(PRODDTA.F55STAB.ASAEXP)/100 AS VAL, Sum(PRODDTA.F55STAB.ASPQOR)/100000 "
    strSQL = strSQL & "AS PCS, Sum(PRODDTA.F55STAB.ASSOCN)/100000 AS GROSSWT, Sum(PRODDTA.F55STAB.ASSOBK)/100000 "
    strSQL = strSQL & "AS NETWT "
    strSQL = strSQL & "FROM PRODDTA.F55STAB "
    strSQL = strSQL & "INNER JOIN PRODDTA.F4101 ON PRODDTA.F55STAB.ASLITM = PRODDTA.F4101.IMLITM "
    strSQL = strSQL & "INNER JOIN PRODDTA.F0101 ON PRODDTA.F55STAB.ASAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE (PRODDTA.F55STAB.ASMNTH)<=" & Mth & " AND "
    strSQL = strSQL & "Trim(PRODDTA.F0101.ABMCU)='" & MCU & "' "
    strSQL = strSQL & "GROUP BY PRODDTA.F4101.IMGLPT, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR HAVING "
    strSQL = strSQL & "(PRODDTA.F55STAB.ASYEAR)= " & Yr & " "
    strSQL = strSQL & "AND PRODDTA.F4101.IMGLPT = '" & Line & "' "
    strSQL = strSQL & "ORDER BY PRODDTA.F4101.IMGLPT, PRODDTA.F55STAB.ASYEAR "
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
'    MsgBox strSQL
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       Cur_Yr_Qty = tmpRset.Fields("QTY")
       Cur_Yr_Val = Round(tmpRset.Fields("VAL"), 0)
       Cur_Yr_Nwt = tmpRset.Fields("NETWT")
       Cur_Yr_Gwt = tmpRset.Fields("GROSSWT")
       Cur_Yr_Pcs = tmpRset.Fields("PCS")
    Else
       Cur_Yr_Qty = 0
       Cur_Yr_Val = 0
       Cur_Yr_Nwt = 0
       Cur_Yr_Gwt = 0
       Cur_Yr_Pcs = 0
    End If

    tmpRset.Close
    Set tmpRset = Nothing
    
End Function
Public Function GetMTDSalesBUC(Line As String, Yr As Integer, Mth As Integer, MCU As String)
    
    strSQL = "SELECT PRODDTA.F4101.IMGLPT AS LINE, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, PRODDTA.F55STAB.ASMNTH, Sum(PRODDTA.F55STAB.ASSOQS)/100000 "
    strSQL = strSQL & "AS QTY, Sum(PRODDTA.F55STAB.ASAEXP)/100 AS VAL, Sum(PRODDTA.F55STAB.ASPQOR)/100000 "
    strSQL = strSQL & "AS PCS, Sum(PRODDTA.F55STAB.ASSOCN)/100000 AS GROSSWT, Sum(PRODDTA.F55STAB.ASSOBK)/100000 "
    strSQL = strSQL & "AS NETWT "
    strSQL = strSQL & "FROM PRODDTA.F55STAB "
    strSQL = strSQL & "INNER JOIN PRODDTA.F4101 ON PRODDTA.F55STAB.ASLITM = PRODDTA.F4101.IMLITM "
    strSQL = strSQL & "INNER JOIN PRODDTA.F0101 ON PRODDTA.F55STAB.ASAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE Trim(PRODDTA.F0101.ABMCU)='" & MCU & "' "
    strSQL = strSQL & "GROUP BY PRODDTA.F4101.IMGLPT, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, PRODDTA.F55STAB.ASMNTH HAVING "
    strSQL = strSQL & "(PRODDTA.F55STAB.ASYEAR)= " & Yr & " "
    strSQL = strSQL & "AND PRODDTA.F4101.IMGLPT = '" & Line & "' "
    strSQL = strSQL & "AND (PRODDTA.F55STAB.ASMNTH)=" & Mth & " "

    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       Cur_Yr_Qty = tmpRset.Fields("QTY")
       Cur_Yr_Val = Round(tmpRset.Fields("VAL"), 0)
       Cur_Yr_Nwt = tmpRset.Fields("NETWT")
       Cur_Yr_Gwt = tmpRset.Fields("GROSSWT")
       Cur_Yr_Pcs = tmpRset.Fields("PCS")
    Else
       Cur_Yr_Qty = 0
       Cur_Yr_Val = 0
       Cur_Yr_Nwt = 0
       Cur_Yr_Gwt = 0
       Cur_Yr_Pcs = 0
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
    
End Function



Public Function GetYTDSalesBUP(Line As String, Yr As Integer, Mth As Integer, MCU As String)
Dim Cur_Year As Integer, Prv_Year As Integer, rpt_Rset As Recordset

    
    strSQL = "SELECT PRODDTA.F4101.IMGLPT AS LINE, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, Sum(PRODDTA.F55STAB.ASSOQS)/100000 "
    strSQL = strSQL & "AS QTY, Sum(PRODDTA.F55STAB.ASAEXP)/100 AS VAL, Sum(PRODDTA.F55STAB.ASPQOR)/100000 "
    strSQL = strSQL & "AS PCS, Sum(PRODDTA.F55STAB.ASSOCN)/100000 AS GROSSWT, Sum(PRODDTA.F55STAB.ASSOBK)/100000 "
    strSQL = strSQL & "AS NETWT "
    strSQL = strSQL & "FROM PRODDTA.F55STAB "
    strSQL = strSQL & "INNER JOIN PRODDTA.F4101 ON PRODDTA.F55STAB.ASLITM = PRODDTA.F4101.IMLITM "
    strSQL = strSQL & "INNER JOIN PRODDTA.F0101 ON PRODDTA.F55STAB.ASAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE (PRODDTA.F55STAB.ASMNTH)<=" & Mth & " AND "
    strSQL = strSQL & "Trim(PRODDTA.F0101.ABMCU)='" & MCU & "' "
    strSQL = strSQL & "GROUP BY PRODDTA.F4101.IMGLPT, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR HAVING "
    strSQL = strSQL & "(PRODDTA.F55STAB.ASYEAR)= " & Yr & " "
    strSQL = strSQL & "AND PRODDTA.F4101.IMGLPT = '" & Line & "' "
    strSQL = strSQL & "ORDER BY PRODDTA.F4101.IMGLPT, PRODDTA.F55STAB.ASYEAR "
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
'    MsgBox strSQL
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       Prv_Yr_Qty = tmpRset.Fields("QTY")
       Prv_Yr_Val = Round(tmpRset.Fields("VAL"), 0)
       Prv_Yr_Nwt = tmpRset.Fields("NETWT")
       Prv_Yr_Gwt = tmpRset.Fields("GROSSWT")
       Prv_Yr_Pcs = tmpRset.Fields("PCS")
    Else
       Prv_Yr_Qty = 0
       Prv_Yr_Val = 0
       Prv_Yr_Nwt = 0
       Prv_Yr_Gwt = 0
       Prv_Yr_Pcs = 0
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
    
End Function

Public Function GetMTDSalesBUP(Line As String, Yr As Integer, Mth As Integer, MCU As String)
Dim Cur_Year As Integer, Prv_Year As Integer, rpt_Rset As Recordset

    strSQL = "SELECT PRODDTA.F4101.IMGLPT AS LINE, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, PRODDTA.F55STAB.ASMNTH, Sum(PRODDTA.F55STAB.ASSOQS)/100000 "
    strSQL = strSQL & "AS QTY, Sum(PRODDTA.F55STAB.ASAEXP)/100 AS VAL, Sum(PRODDTA.F55STAB.ASPQOR)/100000 "
    strSQL = strSQL & "AS PCS, Sum(PRODDTA.F55STAB.ASSOCN)/100000 AS GROSSWT, Sum(PRODDTA.F55STAB.ASSOBK)/100000 "
    strSQL = strSQL & "AS NETWT "
    strSQL = strSQL & "FROM PRODDTA.F55STAB "
    strSQL = strSQL & "INNER JOIN PRODDTA.F4101 ON PRODDTA.F55STAB.ASLITM = PRODDTA.F4101.IMLITM "
    strSQL = strSQL & "INNER JOIN PRODDTA.F0101 ON PRODDTA.F55STAB.ASAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE Trim(PRODDTA.F0101.ABMCU)='" & MCU & "' "
    strSQL = strSQL & "GROUP BY PRODDTA.F4101.IMGLPT, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, PRODDTA.F55STAB.ASMNTH HAVING "
    strSQL = strSQL & "(PRODDTA.F55STAB.ASYEAR)= " & Yr & " "
    strSQL = strSQL & "AND PRODDTA.F4101.IMGLPT = '" & Line & "' "
    strSQL = strSQL & "AND (PRODDTA.F55STAB.ASMNTH)=" & Mth & " "

    strSQL = strSQL & "ORDER BY PRODDTA.F4101.IMGLPT, PRODDTA.F55STAB.ASYEAR, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASMNTH "
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       Prv_Yr_Qty = tmpRset.Fields("QTY")
       Prv_Yr_Val = Round(tmpRset.Fields("VAL"), 0)
       Prv_Yr_Nwt = tmpRset.Fields("NETWT")
       Prv_Yr_Gwt = tmpRset.Fields("GROSSWT")
       Prv_Yr_Pcs = tmpRset.Fields("PCS")
    Else
       Prv_Yr_Qty = 0
       Prv_Yr_Val = 0
       Prv_Yr_Nwt = 0
       Prv_Yr_Gwt = 0
       Prv_Yr_Pcs = 0
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
    
End Function


