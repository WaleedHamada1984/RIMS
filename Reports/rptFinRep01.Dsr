VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rptFinRep01 
   ClientHeight    =   11475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15420
   OleObjectBlob   =   "rptFinRep01.dsx":0000
End
Attribute VB_Name = "rptFinRep01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adors As ADODB.Recordset
Private Sub Report_Initialize()
Dim SQL As String
strSQL = "SELECT LineTable.MainGrp, LineTable.MainGrpDesc, "
strSQL = strSQL & "LineTable.LineSerial, LineTable.ShowWt, FINREPSLS01.LineDesc, FINREPSLS01.NWTCurYearM, "
strSQL = strSQL & "FINREPSLS01.NWTPrvYearM, FINREPSLS01.NWTCurYearYTD, "
strSQL = strSQL & "FINREPSLS01.NWTPrvYearYTD, FINREPSLS01.QTYCurYearM, "
strSQL = strSQL & "FINREPSLS01.QTYPrvYearM, FINREPSLS01.QTYCurYearYTD, "
strSQL = strSQL & "FINREPSLS01.QTYPrvYearYTD, FINREPSLS01.VALCurYearM, "
strSQL = strSQL & "FINREPSLS01.VALPrvYearM, FINREPSLS01.VALCurYearYTD, "
strSQL = strSQL & "FINREPSLS01.VALPrvYearYTD "
strSQL = strSQL & "FROM LineTable INNER JOIN FINREPSLS01 "
strSQL = strSQL & "ON LineTable.Line = FINREPSLS01.Line "
strSQL = strSQL & "ORDER BY LineTable.MainGrp, LineTable.LineSerial "
strSQL = strSQL & ""
Debug.Print strSQL
Set adors = CreateObject("adodb.recordset")
    adors.Open strSQL, SQLConn, adOpenStatic, adLockOptimistic
    Database.SetDataSource adors
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
    If Field3.Value > 0 Then
       Var01.SetText Round(((Field2.Value - Field3.Value) / Field3.Value) * 100, 1) & " %"
    End If
    
    If Field6.Value > 0 Then
       Var02.SetText Round(((Field5.Value - Field6.Value) / Field6.Value) * 100, 1) & " %"
    End If
    
    If Field9.Value > 0 Then
       Var03.SetText Round(((Field8.Value - Field9.Value) / Field9.Value) * 100, 1) & " %"
    End If
    
    If Field12.Value > 0 Then
       Var04.SetText Round(((Field11.Value - Field12.Value) / Field12.Value) * 100, 1) & " %"
    End If
    
    If Field15.Value > 0 Then
       Var05.SetText Round(((Field14.Value - Field15.Value) / Field15.Value) * 100, 1) & " %"
    End If
    
    If Field18.Value > 0 Then
       Var06.SetText Round(((Field17.Value - Field18.Value) / Field18.Value) * 100, 1) & " %"
    End If
    
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
    If pb_RepType = "INT" Then
       RepCaption.SetText "Excluding Inter Company"
    ElseIf pb_RepType = "EXP" Then
       RepCaption.SetText "Export"
    ElseIf pb_RepType = "EAS" Then
       RepCaption.SetText "Eastern Province"
    End If
End Sub
