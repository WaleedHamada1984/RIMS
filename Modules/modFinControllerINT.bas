Attribute VB_Name = "modFinControllerINT"
Public Function GetNetSalesGLINT(Catg As String, Yr As Integer, Company As String)
Dim Inum As Integer, FldNam As String

    strSQL = "SELECT Sum(PRODDTA.F0902.GBAPYC) "
    strSQL = strSQL & "AS SumOfGBAPYC, Sum(PRODDTA.F0902.GBAN01) AS SumOfGBAN01, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN02) AS SumOfGBAN02, Sum(PRODDTA.F0902.GBAN03) "
    strSQL = strSQL & "AS SumOfGBAN03, Sum(PRODDTA.F0902.GBAN04) AS SumOfGBAN04, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN05) AS SumOfGBAN05, Sum(PRODDTA.F0902.GBAN06) "
    strSQL = strSQL & "AS SumOfGBAN06, Sum(PRODDTA.F0902.GBAN07) AS SumOfGBAN07, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN08) AS SumOfGBAN08, Sum(PRODDTA.F0902.GBAN09) "
    strSQL = strSQL & "AS SumOfGBAN09, Sum(PRODDTA.F0902.GBAN10) AS SumOfGBAN10, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN11) AS SumOfGBAN11, Sum(PRODDTA.F0902.GBAN12) "
    strSQL = strSQL & "AS SumOfGBAN12, PRODDTA.F0901.GMCO "
    strSQL = strSQL & "FROM PRODDTA.F0901 INNER JOIN PRODDTA.F0902 "
    strSQL = strSQL & "ON (PRODDTA.F0901.GMCO = PRODDTA.F0902.GBCO) AND "
    strSQL = strSQL & "(PRODDTA.F0901.GMAID = PRODDTA.F0902.GBAID) "
    strSQL = strSQL & "WHERE (((PRODDTA.F0902.GBCTRY)=20) AND "
    strSQL = strSQL & "((PRODDTA.F0902.GBFY)=" & Yr & ") AND ((PRODDTA.F0901.GMR021)='" & Catg & "') "
    strSQL = strSQL & "AND PRODDTA.F0901.GMR022 <> 'INT' "
    strSQL = strSQL & "AND ((PRODDTA.F0902.GBLT)='" & Ltype & "')) GROUP BY PRODDTA.F0901.GMCO "
    strSQL = strSQL & "HAVING (((PRODDTA.F0901.GMCO)='" & Company & "'))"
    
    
    For Inum = 0 To 12
        GL_Value(Inum) = 0
    Next Inum
        
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    
    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       For Inum = 1 To 12
           If Inum < 10 Then
              FldNam = "SumOfGBAN0" & Trim(Str(Inum))
           Else
              FldNam = "SumOfGBAN" & Trim(Str(Inum))
           End If
          GL_Value(Inum) = tmpRset.Fields(FldNam)
       Next Inum
    End If
    tmpRset.Close
    Set tmpRset = Nothing
End Function

Public Function GetTloaderGLINT(Cond As String, Yr As Integer, Company As String)
Dim Inum As Integer, FldNam As String

    strSQL = "SELECT Sum(PRODDTA.F0902.GBAPYC) "
    strSQL = strSQL & "AS SumOfGBAPYC, Sum(PRODDTA.F0902.GBAN01) AS SumOfGBAN01, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN02) AS SumOfGBAN02, Sum(PRODDTA.F0902.GBAN03) "
    strSQL = strSQL & "AS SumOfGBAN03, Sum(PRODDTA.F0902.GBAN04) AS SumOfGBAN04, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN05) AS SumOfGBAN05, Sum(PRODDTA.F0902.GBAN06) "
    strSQL = strSQL & "AS SumOfGBAN06, Sum(PRODDTA.F0902.GBAN07) AS SumOfGBAN07, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN08) AS SumOfGBAN08, Sum(PRODDTA.F0902.GBAN09) "
    strSQL = strSQL & "AS SumOfGBAN09, Sum(PRODDTA.F0902.GBAN10) AS SumOfGBAN10, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN11) AS SumOfGBAN11, Sum(PRODDTA.F0902.GBAN12) "
    strSQL = strSQL & "AS SumOfGBAN12, PRODDTA.F0901.GMCO "
    strSQL = strSQL & "FROM PRODDTA.F0901 INNER JOIN PRODDTA.F0902 "
    strSQL = strSQL & "ON (PRODDTA.F0901.GMCO = PRODDTA.F0902.GBCO) AND "
    strSQL = strSQL & "(PRODDTA.F0901.GMAID = PRODDTA.F0902.GBAID) "
    strSQL = strSQL & "WHERE (((PRODDTA.F0902.GBCTRY)=20) AND "
    strSQL = strSQL & "((PRODDTA.F0902.GBFY)=" & Yr & ") AND ((PRODDTA.F0901.GMR021) " & Cond & ") "
    strSQL = strSQL & "AND PRODDTA.F0901.GMR022 <> 'INT' "
    strSQL = strSQL & "AND ((PRODDTA.F0902.GBLT)='" & Ltype & "')) GROUP BY PRODDTA.F0901.GMCO "
    strSQL = strSQL & "HAVING (((PRODDTA.F0901.GMCO)='" & Company & "'))"
    
    For Inum = 0 To 12
        GL_Value(Inum) = 0
    Next Inum
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    
    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       For Inum = 1 To 12
           If Inum < 10 Then
              FldNam = "SumOfGBAN0" & Trim(Str(Inum))
           Else
              FldNam = "SumOfGBAN" & Trim(Str(Inum))
           End If
          GL_Value(Inum) = tmpRset.Fields(FldNam)
       Next Inum
    End If
    tmpRset.Close
    Set tmpRset = Nothing
End Function

Public Function GetOverHeadGLINT(Catg As String, Yr As Integer, Company As String)
Dim Inum As Integer, FldNam As String

    strSQL = "SELECT Sum(PRODDTA.F0902.GBAPYC) "
    strSQL = strSQL & "AS SumOfGBAPYC, Sum(PRODDTA.F0902.GBAN01) AS SumOfGBAN01, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN02) AS SumOfGBAN02, Sum(PRODDTA.F0902.GBAN03) "
    strSQL = strSQL & "AS SumOfGBAN03, Sum(PRODDTA.F0902.GBAN04) AS SumOfGBAN04, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN05) AS SumOfGBAN05, Sum(PRODDTA.F0902.GBAN06) "
    strSQL = strSQL & "AS SumOfGBAN06, Sum(PRODDTA.F0902.GBAN07) AS SumOfGBAN07, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN08) AS SumOfGBAN08, Sum(PRODDTA.F0902.GBAN09) "
    strSQL = strSQL & "AS SumOfGBAN09, Sum(PRODDTA.F0902.GBAN10) AS SumOfGBAN10, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN11) AS SumOfGBAN11, Sum(PRODDTA.F0902.GBAN12) "
    strSQL = strSQL & "AS SumOfGBAN12, PRODDTA.F0901.GMCO "
    strSQL = strSQL & "FROM PRODDTA.F0901 INNER JOIN PRODDTA.F0902 "
    strSQL = strSQL & "ON (PRODDTA.F0901.GMCO = PRODDTA.F0902.GBCO) AND "
    strSQL = strSQL & "(PRODDTA.F0901.GMAID = PRODDTA.F0902.GBAID) "
    strSQL = strSQL & "WHERE (((PRODDTA.F0902.GBCTRY)=20) AND "
    strSQL = strSQL & "((PRODDTA.F0902.GBFY)=" & Yr & ") AND ((PRODDTA.F0901.GMR022)='" & Catg & "') "
    strSQL = strSQL & "AND PRODDTA.F0901.GMR022 <> 'INT' "
    strSQL = strSQL & "AND ((PRODDTA.F0902.GBLT)='" & Ltype & "')) GROUP BY PRODDTA.F0901.GMCO "
    strSQL = strSQL & "HAVING (((PRODDTA.F0901.GMCO)='" & Company & "'))"
    
    
    For Inum = 0 To 12
        GL_Value(Inum) = 0
    Next Inum
        
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    
    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       For Inum = 1 To 12
           If Inum < 10 Then
              FldNam = "SumOfGBAN0" & Trim(Str(Inum))
           Else
              FldNam = "SumOfGBAN" & Trim(Str(Inum))
           End If
          GL_Value(Inum) = tmpRset.Fields(FldNam)
       Next Inum
    End If
    tmpRset.Close
    Set tmpRset = Nothing
End Function
Public Function GetGL12SlotsINT(Cond As String, Yr As Integer, Company As String)
Dim Inum As Integer, FldNam As String

    strSQL = "SELECT Sum(PRODDTA.F0902.GBAPYC) "
    strSQL = strSQL & "AS SumOfGBAPYC, Sum(PRODDTA.F0902.GBAN01) AS SumOfGBAN01, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN02) AS SumOfGBAN02, Sum(PRODDTA.F0902.GBAN03) "
    strSQL = strSQL & "AS SumOfGBAN03, Sum(PRODDTA.F0902.GBAN04) AS SumOfGBAN04, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN05) AS SumOfGBAN05, Sum(PRODDTA.F0902.GBAN06) "
    strSQL = strSQL & "AS SumOfGBAN06, Sum(PRODDTA.F0902.GBAN07) AS SumOfGBAN07, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN08) AS SumOfGBAN08, Sum(PRODDTA.F0902.GBAN09) "
    strSQL = strSQL & "AS SumOfGBAN09, Sum(PRODDTA.F0902.GBAN10) AS SumOfGBAN10, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN11) AS SumOfGBAN11, Sum(PRODDTA.F0902.GBAN12) "
    strSQL = strSQL & "AS SumOfGBAN12, PRODDTA.F0901.GMCO "
    strSQL = strSQL & "FROM PRODDTA.F0901 INNER JOIN PRODDTA.F0902 "
    strSQL = strSQL & "ON (PRODDTA.F0901.GMCO = PRODDTA.F0902.GBCO) AND "
    strSQL = strSQL & "(PRODDTA.F0901.GMAID = PRODDTA.F0902.GBAID) "
    strSQL = strSQL & "WHERE  " & Cond & " "
    strSQL = strSQL & "AND PRODDTA.F0901.GMR022 <> 'INT' "
    strSQL = strSQL & "GROUP BY PRODDTA.F0901.GMCO "
    strSQL = strSQL & "HAVING (((PRODDTA.F0901.GMCO)='" & Company & "'))"
    
    
    For Inum = 0 To 12
        GL_Value(Inum) = 0
    Next Inum
        
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    
    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       For Inum = 1 To 12
           If Inum < 10 Then
              FldNam = "SumOfGBAN0" & Trim(Str(Inum))
           Else
              FldNam = "SumOfGBAN" & Trim(Str(Inum))
           End If
          GL_Value(Inum) = tmpRset.Fields(FldNam)
       Next Inum
    End If
    tmpRset.Close
    Set tmpRset = Nothing
End Function

