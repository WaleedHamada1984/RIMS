VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rptSlsCollSummary 
   ClientHeight    =   9495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   OleObjectBlob   =   "rptSlsCollSummary.dsx":0000
End
Attribute VB_Name = "rptSlsCollSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adors As ADODB.Recordset
Private Sub Report_Initialize()
Dim SQL As String
SQL = "SELECT * From DailyReport"
    Set adors = CreateObject("adodb.recordset")
    adors.Open SQL, SQLConn, adOpenStatic, adLockOptimistic
    Database.SetDataSource adors
End Sub

'Private Sub Report_Terminate()
'    adors.Close
'    Set adors = Nothing
'End Sub
'
Private Sub Section6_Format(ByVal pFormattingInfo As Object)

End Sub

