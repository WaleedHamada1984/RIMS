Attribute VB_Name = "modFinReport"
Public Prv_Yr_Qty As Double
Public Prv_Yr_Val As Double
Public Prv_Yr_Pcs As Double
Public Prv_Yr_Nwt As Double
Public Prv_Yr_Gwt As Double

Public Cur_Yr_Qty As Double
Public Cur_Yr_Val As Double
Public Cur_Yr_Pcs As Double
Public Cur_Yr_Nwt As Double
Public Cur_Yr_Gwt As Double

Public pb_Cond_Stmt As String, pb_RepType As String

Public Function GetYTDSalesC(Line As String, Yr As Integer, Mth As Integer)
    strSQL = "SELECT PRODDTA.F4101.IMGLPT AS LINE, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, Sum(PRODDTA.F55STAB.ASSOQS)/100000 "
    strSQL = strSQL & "AS QTY, Sum(PRODDTA.F55STAB.ASAEXP)/100 AS VAL, Sum(PRODDTA.F55STAB.ASPQOR)/100000 "
    strSQL = strSQL & "AS PCS, Sum(PRODDTA.F55STAB.ASSOCN)/100000 AS GROSSWT, Sum(PRODDTA.F55STAB.ASSOBK)/100000 "
    strSQL = strSQL & "AS NETWT "
    strSQL = strSQL & "FROM PRODDTA.F55STAB "
    strSQL = strSQL & "INNER JOIN PRODDTA.F4101 ON PRODDTA.F55STAB.ASLITM = PRODDTA.F4101.IMLITM "
    strSQL = strSQL & "INNER JOIN PRODDTA.F0101 ON PRODDTA.F55STAB.ASAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE (PRODDTA.F55STAB.ASMNTH)<=" & Mth & " "
    strSQL = strSQL & "AND RTRIM(LTRIM(PRODDTA.F0101.ABAC03)) " & pb_Cond_Stmt & "' "
    strSQL = strSQL & "GROUP BY PRODDTA.F4101.IMGLPT, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR HAVING "
    strSQL = strSQL & "(PRODDTA.F55STAB.ASYEAR)= " & Yr & " "
    strSQL = strSQL & "AND PRODDTA.F4101.IMGLPT = '" & Line & "' "
    strSQL = strSQL & "ORDER BY PRODDTA.F4101.IMGLPT, PRODDTA.F55STAB.ASYEAR "
    
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       Cur_Yr_Qty = tmpRset.Fields("QTY")
       Cur_Yr_Val = Round(tmpRset.Fields("VAL") / 1000, 0)
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
    If Trim(Line) = "F112" Or Trim(Line) = "F111" Or Trim(Line) = "F203" Or _
       Trim(Line) = "F204" Or Trim(Line) = "F402" Or Trim(Line) = "F404" Or _
       Trim(Line) = "F401" Or Trim(Line) = "F403" Or Trim(Line) = "F501" Or _
       Trim(Line) = "F503" Or Trim(Line) = "F502" Then
       Cur_Yr_Nwt = 0
   '    Cur_Yr_Gwt = 0

    End If

    tmpRset.Close
    Set tmpRset = Nothing
    
End Function
Public Function GetMTDSalesC(Line As String, Yr As Integer, Mth As Integer)
    
    strSQL = "SELECT PRODDTA.F4101.IMGLPT AS LINE, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, PRODDTA.F55STAB.ASMNTH, Sum(PRODDTA.F55STAB.ASSOQS)/100000 "
    strSQL = strSQL & "AS QTY, Sum(PRODDTA.F55STAB.ASAEXP)/100 AS VAL, Sum(PRODDTA.F55STAB.ASPQOR)/100000 "
    strSQL = strSQL & "AS PCS, Sum(PRODDTA.F55STAB.ASSOCN)/100000 AS GROSSWT, Sum(PRODDTA.F55STAB.ASSOBK)/100000 "
    strSQL = strSQL & "AS NETWT "
    strSQL = strSQL & "FROM PRODDTA.F55STAB "
    strSQL = strSQL & "INNER JOIN PRODDTA.F4101 ON PRODDTA.F55STAB.ASLITM = PRODDTA.F4101.IMLITM "
    strSQL = strSQL & "INNER JOIN PRODDTA.F0101 ON PRODDTA.F55STAB.ASAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE (PRODDTA.F55STAB.ASMNTH)=" & Mth & " "
    strSQL = strSQL & "AND RTRIM(LTRIM(PRODDTA.F0101.ABAC03)) " & pb_Cond_Stmt & "' "
    strSQL = strSQL & "GROUP BY PRODDTA.F4101.IMGLPT, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, PRODDTA.F55STAB.ASMNTH HAVING "
    strSQL = strSQL & "(PRODDTA.F55STAB.ASYEAR)= " & Yr & " "
    strSQL = strSQL & "AND PRODDTA.F4101.IMGLPT = '" & Line & "' "
    strSQL = strSQL & "ORDER BY PRODDTA.F4101.IMGLPT, PRODDTA.F55STAB.ASYEAR, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASMNTH"
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       Cur_Yr_Qty = tmpRset.Fields("QTY")
       Cur_Yr_Val = Round(tmpRset.Fields("VAL") / 1000, 0)
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
    If Trim(Line) = "F112" Or Trim(Line) = "F111" Or Trim(Line) = "F203" Or _
       Trim(Line) = "F204" Or Trim(Line) = "F402" Or Trim(Line) = "F404" Or _
       Trim(Line) = "F401" Or Trim(Line) = "F403" Or Trim(Line) = "F501" Or _
       Trim(Line) = "F503" Or Trim(Line) = "F502" Then
       Cur_Yr_Nwt = 0
  '     Cur_Yr_Gwt = 0

    End If
    tmpRset.Close
    Set tmpRset = Nothing
    
End Function



Public Function GetYTDSalesP(Line As String, Yr As Integer, Mth As Integer)
Dim Cur_Year As Integer, Prv_Year As Integer, rpt_Rset As Recordset

    
    strSQL = "SELECT PRODDTA.F4101.IMGLPT AS LINE, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, Sum(PRODDTA.F55STAB.ASSOQS)/100000 "
    strSQL = strSQL & "AS QTY, Sum(PRODDTA.F55STAB.ASAEXP)/100 AS VAL, Sum(PRODDTA.F55STAB.ASPQOR)/100000 "
    strSQL = strSQL & "AS PCS, Sum(PRODDTA.F55STAB.ASSOCN)/100000 AS GROSSWT, Sum(PRODDTA.F55STAB.ASSOBK)/100000 "
    strSQL = strSQL & "AS NETWT "
    strSQL = strSQL & "FROM PRODDTA.F55STAB "
    strSQL = strSQL & "INNER JOIN PRODDTA.F4101 ON PRODDTA.F55STAB.ASLITM = PRODDTA.F4101.IMLITM "
    strSQL = strSQL & "INNER JOIN PRODDTA.F0101 ON PRODDTA.F55STAB.ASAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE (PRODDTA.F55STAB.ASMNTH)<=" & Mth & " "
    strSQL = strSQL & "AND RTRIM(LTRIM(PRODDTA.F0101.ABAC03)) " & pb_Cond_Stmt & "' "
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
       Prv_Yr_Val = Round(tmpRset.Fields("VAL") / 1000, 0)
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
    
    If Trim(Line) = "F112" Or Trim(Line) = "F111" Or Trim(Line) = "F203" Or _
       Trim(Line) = "F204" Or Trim(Line) = "F402" Or Trim(Line) = "F404" Or _
       Trim(Line) = "F401" Or Trim(Line) = "F403" Or Trim(Line) = "F501" Or _
       Trim(Line) = "F503" Or Trim(Line) = "F502" Then
       Prv_Yr_Nwt = 0
 '      Prv_Yr_Gwt = 0

    End If
    
    
    tmpRset.Close
    Set tmpRset = Nothing
    
End Function

Public Function GetMTDSalesP(Line As String, Yr As Integer, Mth As Integer)
Dim Cur_Year As Integer, Prv_Year As Integer, rpt_Rset As Recordset

    
    strSQL = "SELECT PRODDTA.F4101.IMGLPT AS LINE, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, PRODDTA.F55STAB.ASMNTH, Sum(PRODDTA.F55STAB.ASSOQS)/100000 "
    strSQL = strSQL & "AS QTY, Sum(PRODDTA.F55STAB.ASAEXP)/100 AS VAL, Sum(PRODDTA.F55STAB.ASPQOR)/100000 "
    strSQL = strSQL & "AS PCS, Sum(PRODDTA.F55STAB.ASSOCN)/100000 AS GROSSWT, Sum(PRODDTA.F55STAB.ASSOBK)/100000 "
    strSQL = strSQL & "AS NETWT "
    strSQL = strSQL & "FROM PRODDTA.F55STAB "
    strSQL = strSQL & "INNER JOIN PRODDTA.F4101 ON PRODDTA.F55STAB.ASLITM = PRODDTA.F4101.IMLITM "
    strSQL = strSQL & "INNER JOIN PRODDTA.F0101 ON PRODDTA.F55STAB.ASAN8 = PRODDTA.F0101.ABAN8 "
    strSQL = strSQL & "WHERE (PRODDTA.F55STAB.ASMNTH)= " & Mth & " "
    strSQL = strSQL & "AND RTRIM(LTRIM(PRODDTA.F0101.ABAC03)) " & pb_Cond_Stmt & "' "
    strSQL = strSQL & "GROUP BY PRODDTA.F4101.IMGLPT, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, PRODDTA.F55STAB.ASMNTH HAVING "
    strSQL = strSQL & "(PRODDTA.F55STAB.ASYEAR)= " & Yr & " "
    strSQL = strSQL & "AND PRODDTA.F4101.IMGLPT = '" & Line & "' "
    strSQL = strSQL & "ORDER BY PRODDTA.F4101.IMGLPT, PRODDTA.F55STAB.ASYEAR, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASMNTH "
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       Prv_Yr_Qty = tmpRset.Fields("QTY")
       Prv_Yr_Val = Round(tmpRset.Fields("VAL") / 1000, 0)
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
    If Trim(Line) = "F112" Or Trim(Line) = "F111" Or Trim(Line) = "F203" Or _
       Trim(Line) = "F204" Or Trim(Line) = "F402" Or Trim(Line) = "F404" Or _
       Trim(Line) = "F401" Or Trim(Line) = "F403" Or Trim(Line) = "F501" Or _
       Trim(Line) = "F503" Or Trim(Line) = "F502" Then
       Prv_Yr_Nwt = 0
 '     Prv_Yr_Gwt = 0

    End If
    tmpRset.Close
    Set tmpRset = Nothing
    
End Function


