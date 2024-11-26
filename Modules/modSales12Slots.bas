Attribute VB_Name = "modSales12Slots"
Dim tmp_Rset01 As Recordset
Public Function GenSalesByBU(ByVal yr As Integer) As Boolean
    Dim ItmCode As String
    
    GenSalesByBU = False
    SQLConn.Execute "DELETE FROM tblSales"
    
    Set tmp_Rset01 = New Recordset
    tmp_Rset01.CursorLocation = adUseClient
    tmp_Rset01.Open "SELECT * From tblSales", SQLConn, adOpenStatic, adLockOptimistic
    
    
    strSQL = "SELECT PRODDTA.F0101.ABMCU, PRODDTA.F0006.MCDL01, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASLITM, PRODDTA.F55STAB.ASMNTH, PRODDTA.F4101.IMDSC1, Sum(PRODDTA.F55STAB.ASSOQS)/100000 "
    strSQL = strSQL & "AS SLSQTY, Sum(PRODDTA.F55STAB.ASAEXP)/100 AS SLSVAL "
    strSQL = strSQL & "FROM ((PRODDTA.F55STAB INNER JOIN PRODDTA.F0101 "
    strSQL = strSQL & "ON PRODDTA.F55STAB.ASAN8 = PRODDTA.F0101.ABAN8) INNER JOIN PRODDTA.F0006 "
    strSQL = strSQL & "ON PRODDTA.F0101.ABMCU = PRODDTA.F0006.MCMCU) INNER JOIN PRODDTA.F4101 "
    strSQL = strSQL & "ON (PRODDTA.F55STAB.ASITM = PRODDTA.F4101.IMITM) AND (PRODDTA.F55STAB.ASLITM "
    strSQL = strSQL & "= PRODDTA.F4101.IMLITM) "
    strSQL = strSQL & "WHERE (((PRODDTA.F0006.MCSTYL)='IS')) AND PRODDTA.F4101.IMSRP1 = 'FG1' GROUP BY "
    strSQL = strSQL & "PRODDTA.F0101.ABMCU, PRODDTA.F0006.MCDL01, PRODDTA.F55STAB.ASYEAR, PRODDTA.F55STAB.ASLITM, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASMNTH, PRODDTA.F4101.IMDSC1 HAVING (((PRODDTA.F55STAB.ASYEAR)=" & Trim(Str(yr)) & " "
    strSQL = strSQL & ")) "
    strSQL = strSQL & "ORDER BY PRODDTA.F0101.ABMCU, PRODDTA.F55STAB.ASLITM, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASMNTH "
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic
    GenSalesByBU = False
    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       Do
          If tmpRset.Fields("ASLITM") <> ItmCode Then
             tmp_Rset01.AddNew
             ItmCode = tmpRset.Fields("ASLITM")
             tmp_Rset01.Fields("BUnit") = (tmpRset.Fields("ABMCU"))
             tmp_Rset01.Fields("BUDesc") = tmpRset.Fields("MCDL01")
             tmp_Rset01.Fields("ItCode") = Trim(tmpRset.Fields("ASLITM"))
             tmp_Rset01.Fields("ItDesc") = tmpRset.Fields("IMDSC1")
          End If
          
          Select Case tmpRset.Fields("ASMNTH")
          Case Is = 1
               tmp_Rset01.Fields("M01") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV01") = tmpRset.Fields("SLSVAL")
          Case Is = 2
               tmp_Rset01.Fields("M02") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV02") = tmpRset.Fields("SLSVAL")
          Case Is = 3
               tmp_Rset01.Fields("M03") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV03") = tmpRset.Fields("SLSVAL")
          Case Is = 4
               tmp_Rset01.Fields("M04") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV04") = tmpRset.Fields("SLSVAL")
          Case Is = 5
               tmp_Rset01.Fields("M05") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV05") = tmpRset.Fields("SLSVAL")
          Case Is = 6
               tmp_Rset01.Fields("M06") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV06") = tmpRset.Fields("SLSVAL")
          Case Is = 7
               tmp_Rset01.Fields("M07") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV07") = tmpRset.Fields("SLSVAL")
          Case Is = 8
               tmp_Rset01.Fields("M08") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV08") = tmpRset.Fields("SLSVAL")
          Case Is = 9
               tmp_Rset01.Fields("M09") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV09") = tmpRset.Fields("SLSVAL")
          Case Is = 10
               tmp_Rset01.Fields("M10") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV10") = tmpRset.Fields("SLSVAL")
          Case Is = 11
               tmp_Rset01.Fields("M11") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV11") = tmpRset.Fields("SLSVAL")
          Case Is = 12
               tmp_Rset01.Fields("M12") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV12") = tmpRset.Fields("SLSVAL")
          End Select
          
          tmpRset.MoveNext
          If Not tmpRset.EOF Then
            If tmpRset.Fields("ASLITM") <> ItmCode Then
               tmp_Rset01.Update
            End If
          Else
            tmp_Rset01.Update
          End If
       Loop Until tmpRset.EOF
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
    tmp_Rset01.Close
    Set tmp_Rset01 = Nothing
    GenSalesByBU = True
End Function
Public Function GenSalesByCUST(CAn8 As Long) As Boolean
    Dim ItmCode As String
    
    GenSalesByCUST = False
    SQLConn.Execute "DELETE FROM tblSalesCUST"
    
    Set tmp_Rset01 = New Recordset
    tmp_Rset01.CursorLocation = adUseClient
    tmp_Rset01.Open "SELECT * From tblSalesCUST", SQLConn, adOpenStatic, adLockOptimistic
    
    
    strSQL = "SELECT PRODDTA.F0101.ABAN8, PRODDTA.F0101.ABALPH, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASLITM, PRODDTA.F55STAB.ASMNTH, PRODDTA.F4101.IMDSC1, Sum(PRODDTA.F55STAB.ASSOQS)/100000 "
    strSQL = strSQL & "AS SLSQTY, Sum(PRODDTA.F55STAB.ASAEXP)/100 AS SLSVAL "
    strSQL = strSQL & "FROM ((PRODDTA.F55STAB INNER JOIN PRODDTA.F0101 "
    strSQL = strSQL & "ON PRODDTA.F55STAB.ASAN8 = PRODDTA.F0101.ABAN8) INNER JOIN PRODDTA.F0006 "
    strSQL = strSQL & "ON PRODDTA.F0101.ABMCU = PRODDTA.F0006.MCMCU) INNER JOIN PRODDTA.F4101 "
    strSQL = strSQL & "ON (PRODDTA.F55STAB.ASITM = PRODDTA.F4101.IMITM) AND (PRODDTA.F55STAB.ASLITM "
    strSQL = strSQL & "= PRODDTA.F4101.IMLITM) "
    strSQL = strSQL & "WHERE (((PRODDTA.F0006.MCSTYL)='IS')) AND PRODDTA.F4101.IMSRP1 = 'FG1' "
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "PRODDTA.F0101.ABAN8, PRODDTA.F0101.ABALPH, PRODDTA.F55STAB.ASYEAR, PRODDTA.F55STAB.ASLITM, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASMNTH, PRODDTA.F4101.IMDSC1 HAVING (((PRODDTA.F55STAB.ASYEAR)=9 "
    strSQL = strSQL & ")) AND PRODDTA.F0101.ABAN8 = " & CAn8 & " "
    strSQL = strSQL & "ORDER BY PRODDTA.F0101.ABAN8, PRODDTA.F55STAB.ASLITM, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASMNTH "
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    MsgBox strSQL
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic
    GenSalesByCUST = False
    If tmpRset.RecordCount > 0 Then
       tmpRset.MoveFirst
       Do
          If tmpRset.Fields("ASLITM") <> ItmCode Then
             tmp_Rset01.AddNew
             ItmCode = tmpRset.Fields("ASLITM")
             tmp_Rset01.Fields("AN8") = (tmpRset.Fields("ABAN8"))
             tmp_Rset01.Fields("CustomerName") = tmpRset.Fields("ABALPH")
             tmp_Rset01.Fields("ItCode") = tmpRset.Fields("ASLITM")
             tmp_Rset01.Fields("ItDesc") = tmpRset.Fields("IMDSC1")
          End If
          
          Select Case tmpRset.Fields("ASMNTH")
          Case Is = 1
               tmp_Rset01.Fields("M01") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV01") = tmpRset.Fields("SLSVAL")
          Case Is = 2
               tmp_Rset01.Fields("M02") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV02") = tmpRset.Fields("SLSVAL")
          Case Is = 3
               tmp_Rset01.Fields("M03") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV03") = tmpRset.Fields("SLSVAL")
          Case Is = 4
               tmp_Rset01.Fields("M04") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV04") = tmpRset.Fields("SLSVAL")
          Case Is = 5
               tmp_Rset01.Fields("M05") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV05") = tmpRset.Fields("SLSVAL")
          Case Is = 6
               tmp_Rset01.Fields("M06") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV06") = tmpRset.Fields("SLSVAL")
          Case Is = 7
               tmp_Rset01.Fields("M07") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV07") = tmpRset.Fields("SLSVAL")
          Case Is = 8
               tmp_Rset01.Fields("M08") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV08") = tmpRset.Fields("SLSVAL")
          Case Is = 9
               tmp_Rset01.Fields("M09") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV09") = tmpRset.Fields("SLSVAL")
          Case Is = 10
               tmp_Rset01.Fields("M10") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV10") = tmpRset.Fields("SLSVAL")
          Case Is = 11
               tmp_Rset01.Fields("M11") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV11") = tmpRset.Fields("SLSVAL")
          Case Is = 12
               tmp_Rset01.Fields("M12") = tmpRset.Fields("SLSQTY")
               tmp_Rset01.Fields("MV12") = tmpRset.Fields("SLSVAL")
          End Select
          
          tmpRset.MoveNext
          If Not tmpRset.EOF Then
            If tmpRset.Fields("ASLITM") <> ItmCode Then
               tmp_Rset01.Update
            End If
          Else
            tmp_Rset01.Update
          End If
       Loop Until tmpRset.EOF
    End If
    
    tmpRset.Close
    Set tmpRset = Nothing
    tmp_Rset01.Close
    Set tmp_Rset01 = Nothing
    GenSalesByCUST = True
End Function


