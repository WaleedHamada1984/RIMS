Attribute VB_Name = "modAcctsSales"
Public Function AreaNetSales(MCU As String, LineCode As String, Yr As Integer, Mth As Integer) As Double
    strSQL = "SELECT Sum(PRODDTA.F0902.GBAN01) AS SumOfGBAN01, Sum(PRODDTA.F0902.GBAN02) AS SumOfGBAN02, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN03) AS SumOfGBAN03, Sum(PRODDTA.F0902.GBAN04) AS SumOfGBAN04, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN05) AS SumOfGBAN05, Sum(PRODDTA.F0902.GBAN06) AS SumOfGBAN06, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN07) AS SumOfGBAN07, Sum(PRODDTA.F0902.GBAN08) AS SumOfGBAN08, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN09) AS SumOfGBAN09, Sum(PRODDTA.F0902.GBAN10) AS SumOfGBAN10, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN11) AS SumOfGBAN11, Sum(PRODDTA.F0902.GBAN12) AS SumOfGBAN12, "
    strSQL = strSQL & "Sum(PRODDTA.F0902.GBAN13) AS SumOfGBAN13, Sum(PRODDTA.F0902.GBAN14) AS SumOfGBAN14 "
    strSQL = strSQL & "FROM PRODDTA.F0902 "
    strSQL = strSQL & "WHERE (((Trim(PRODDTA.F0902.GBMCU))='" & MCU & "') AND ((PRODDTA.F0902.GBOBJ)='6101') AND "
    strSQL = strSQL & "((PRODDTA.F0902.GBSUB)='" & GetAccount(LineCode, "SO", 4230) & "') AND ((PRODDTA.F0902.GBCTRY)=20) AND "
    strSQL = strSQL & "((PRODDTA.F0902.GBFY)=" & Yr & ") AND ((PRODDTA.F0902.GBLT)='AA')) OR "
    strSQL = strSQL & "(((Trim(PRODDTA.F0902.GBMCU))='" & MCU & "') AND ((PRODDTA.F0902.GBOBJ)='6201') AND "
    strSQL = strSQL & "((PRODDTA.F0902.GBSUB)='" & GetAccount(LineCode, "SO", 4230) & "') AND ((PRODDTA.F0902.GBCTRY)=20) AND "
    strSQL = strSQL & "((PRODDTA.F0902.GBFY)=" & Yr & ") AND ((PRODDTA.F0902.GBLT)='AA')) OR "
    strSQL = strSQL & "(((Trim(PRODDTA.F0902.GBMCU))='" & MCU & "') AND ((PRODDTA.F0902.GBOBJ)='6301') AND "
    strSQL = strSQL & "((PRODDTA.F0902.GBSUB)='" & GetAccount(LineCode, "SO", 4230) & "') AND ((PRODDTA.F0902.GBCTRY)=20) AND "
    strSQL = strSQL & "((PRODDTA.F0902.GBFY)=" & Yr & ") AND ((PRODDTA.F0902.GBLT)='AA'))"
    
    Set tmpRset = New Recordset
    tmpRset.CursorLocation = adUseClient
    
    tmpRset.Open strSQL, DbConn, adOpenStatic, adLockOptimistic
     
    AreaNetSales = 0
    
    If Not tmpRset.EOF And Not tmpRset.BOF Then
    Select Case Mth
           Case Is = 1
                AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN01"))
           Case Is = 2
                AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN02"))
           Case Is = 3
                AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN03"))
           Case Is = 4
                AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN04"))
           Case Is = 5
                AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN05"))
           Case Is = 6
                AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN06"))
           Case Is = 7
                AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN07"))
           Case Is = 8
                AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN08"))
           Case Is = 9
                AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN09"))
           Case Is = 10
                AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN10"))
           Case Is = 11
                If Not IsNull(tmpRset.Fields("SumOfGBAN11")) Then
                   AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN11"))
                Else
                   AreaNetSales = 0
                End If
           Case Is = 12
                If Not IsNull(tmpRset.Fields("SumOfGBAN12")) Then
                  AreaNetSales = NullToZero(tmpRset.Fields("SumOfGBAN12"))
                Else
                  AreaNetSales = 0
                End If
   End Select
   End If
   
   tmpRset.Close
   Set tmpRset = Nothing
End Function

Public Function GetAccount(Line As String, Dct As String, TableNo As Integer) As String
    Dim rstAcc As Recordset
    
    strSQL = "SELECT PRODDTA.F4095.MLANUM, PRODDTA.F4095.MLCO, PRODDTA.F4095.MLDCTO, "
    strSQL = strSQL & "Trim(PRODDTA.F4095.MLDCT) AS DCT, PRODDTA.F4095.MLGLPT, PRODDTA.F4095.MLCOST, "
    strSQL = strSQL & "PRODDTA.F4095.MLMCU, PRODDTA.F4095.MLOBJ, PRODDTA.F4095.MLSUB "
    strSQL = strSQL & "FROM PRODDTA.F4095 "
    strSQL = strSQL & "WHERE (((PRODDTA.F4095.MLANUM)=" & TableNo & ") AND "
    strSQL = strSQL & "((Trim(PRODDTA.F4095.MLDCT))='" & Dct & "') AND ((PRODDTA.F4095.MLGLPT)='" & Line & "'))"
    
    Set rstAcc = New Recordset
    rstAcc.CursorLocation = adUseClient
    
    rstAcc.Open strSQL, DbConn, adOpenStatic, adLockOptimistic

    GetAccount = ""
    
    If Not rstAcc.EOF And Not rstAcc.BOF Then
      GetAccount = rstAcc.Fields("MLSUB")
    End If
    
    rstAcc.Close
    Set rstAcc = Nothing
End Function
