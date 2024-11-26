VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rptPCS 
   ClientHeight    =   9495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13095
   OleObjectBlob   =   "rptPCS.dsx":0000
End
Attribute VB_Name = "rptPCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adors As ADODB.Recordset
Private Sub Report_Initialize()
Dim SQL As String
strSQL = "SELECT * From BrandSales"
Set adors = CreateObject("adodb.recordset")
    adors.Open strSQL, SQLConn, adOpenStatic, adLockOptimistic
    Database.SetDataSource adors
End Sub


