VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrdLine 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmPrdLine.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox comboReport 
      Height          =   315
      ItemData        =   "frmPrdLine.frx":000C
      Left            =   4200
      List            =   "frmPrdLine.frx":0016
      TabIndex        =   4
      Text            =   "All areas"
      Top             =   240
      Width           =   3135
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8295
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   15855
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   20709379
      CurrentDate     =   39354
   End
   Begin VB.Label Label1 
      Caption         =   "Select a Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrdLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DRepTbl As Recordset
Dim RepTemplate As Recordset, ISerial As Long, Ttype As String, fldName As String
Private Sub Command1_Click()
Dim Lmonth As Date

If LastDayInMonth(Year(DTPicker1.Value), Month(DTPicker1.Value) - 1) <= Day(DTPicker1.Value) Then
    Lmonth = Month(DTPicker1.Value) - 1 & "/" & LastDayInMonth(Year(DTPicker1.Value), Month(DTPicker1.Value) - 1) & "/" & Year(DTPicker1.Value)
Else
    Lmonth = Month(DTPicker1.Value) - 1 & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value)
End If

SQLConn.Execute "DELETE FROM tblPrdLine01"

Set DRepTbl = New Recordset
DRepTbl.CursorLocation = adUseClient
DRepTbl.Open "SELECT * From DailyRep Order By MainArea,AreaCode,SmanCode", SQLConn, adOpenStatic, adLockOptimistic

Set RepTemplate = New Recordset
RepTemplate.CursorLocation = adUseClient
RepTemplate.Open "SELECT * From tblPrdLine01", SQLConn, adOpenStatic, adLockOptimistic


If DRepTbl.RecordCount > 0 Then
   DRepTbl.MoveFirst
   Lmonth = Trim(Str(Month(DTPicker1.Value))) & "/" & LastDayInMonth(Year(DTPicker1.Value), Month(DTPicker1.Value)) & "/" & Trim(Str(Year(DTPicker1.Value)))
   
   Do
      
     frmMain.StBar.Panels(2).Text = "Processing Record For Salesman (Sales) - " & DRepTbl.Fields("SmanCode")
     DoEvents
     DoEvents
     
     RepTemplate.AddNew
     RepTemplate.Fields("SCode") = DRepTbl.Fields("SmanCode")
     If DRepTbl.Fields("AreaCode") <> "05" Then
     RepTemplate.Fields("LINED") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "D") + GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "U")
     RepTemplate.Fields("LINEL") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "L")
     RepTemplate.Fields("LINEH") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "H") + GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "O")
     RepTemplate.Fields("LINEF") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "F") + GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "J")
     RepTemplate.Fields("LINEA") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "A")
     RepTemplate.Fields("LINEE") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "E")
     RepTemplate.Fields("LINEP") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "P")
     RepTemplate.Fields("LINET") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "T")
     RepTemplate.Fields("LINE9") = GetMonthSales(DRepTbl.Fields("SmanCode"), Lmonth) - _
                                   (RepTemplate.Fields("LINED") + RepTemplate.Fields("LINEL") + RepTemplate.Fields("LINEH") + _
                                    RepTemplate.Fields("LINEF") + RepTemplate.Fields("LINEA") + RepTemplate.Fields("LINEE") + _
                                    RepTemplate.Fields("LINEP") + RepTemplate.Fields("LINET"))
     Else
     RepTemplate.Fields("LINED") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "D") + GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "U") + GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "L")
     RepTemplate.Fields("LINEL") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "N")
     RepTemplate.Fields("LINEH") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "H") + GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "O")
     RepTemplate.Fields("LINEF") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "F") + GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "J")
     RepTemplate.Fields("LINEA") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "A")
     RepTemplate.Fields("LINEE") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "E")
     RepTemplate.Fields("LINEP") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "P")
     RepTemplate.Fields("LINET") = GetProductSales(DRepTbl.Fields("SmanCode"), Lmonth, "T")
     RepTemplate.Fields("LINE9") = GetMonthSales(DRepTbl.Fields("SmanCode"), Lmonth) - _
                                   (RepTemplate.Fields("LINED") + RepTemplate.Fields("LINEL") + RepTemplate.Fields("LINEH") + _
                                    RepTemplate.Fields("LINEF") + RepTemplate.Fields("LINEA") + RepTemplate.Fields("LINEE") + _
                                    RepTemplate.Fields("LINEP") + RepTemplate.Fields("LINET"))
     End If
     
     RepTemplate.Update
     DRepTbl.MoveNext
     
   Loop Until DRepTbl.EOF
     
     frmMain.StBar.Panels(2).Text = "Process Finished"
   
End If

DRepTbl.Close
RepTemplate.Close

Set DRepTbl = Nothing
Set RepTemplate = Nothing

End Sub


Private Sub Command2_Click()
Set tmpRset = New Recordset
tmpRset.CursorLocation = adUseClient
tmpRset.Open "SELECT * From tblPrdLine01", SQLConn, adOpenStatic, adLockOptimistic

Recordset2Excel tmpRset

tmpRset.Close
Set tmpRset = Nothing

End Sub

Private Sub Form_Load()
'Dim RPT As New rptSlsCollSummary
'CRViewer1.ReportSource = RPT
'CRViewer1.ViewReport
DTPicker1.Value = Date - 1
End Sub

