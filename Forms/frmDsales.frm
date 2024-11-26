VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDsales 
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15225
   Icon            =   "frmDsales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10890
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7935
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   15015
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
         Name            =   "Arial"
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
      Format          =   20643843
      CurrentDate     =   39354
   End
   Begin VB.Label Label1 
      Caption         =   "Select a Date"
      BeginProperty Font 
         Name            =   "Arial"
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
Attribute VB_Name = "frmDsales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DRepTbl As Recordset
Dim RepTemplate As Recordset
Private Sub Command1_Click()
Dim Lmonth As Date, LM As Long, LY As Long

If Month(DTPicker1.Value) = 1 Then
   LM = 12
   LY = Year(DTPicker1.Value) - 1
Else
   LM = Month(DTPicker1.Value) - 1
   LY = Year(DTPicker1.Value)
End If
   
If LastDayInMonth(LY, LM) <= Day(DTPicker1.Value) Then
   Lmonth = LM & "/" & LastDayInMonth(LY, LM) & "/" & LY
Else
    If Month(DTPicker1.Value) > 1 Then
       Lmonth = Month(DTPicker1.Value) - 1 & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value)
    Else
       Lmonth = "12/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) - 1
    End If
End If

SQLConn.Execute "DELETE FROM DailyReport"

Set DRepTbl = New Recordset
DRepTbl.CursorLocation = adUseClient
DRepTbl.Open "SELECT * From DailyRep Order By MainArea,AreaCode,SmanCode", SQLConn, adOpenStatic, adLockOptimistic

Set RepTemplate = New Recordset
RepTemplate.CursorLocation = adUseClient
RepTemplate.Open "SELECT * From DailyReport", SQLConn, adOpenStatic, adLockOptimistic


If DRepTbl.RecordCount > 0 Then
   DRepTbl.MoveFirst
   
   Do
      
     frmMain.StBar.Panels(2).Text = "Processing Record For Salesman - " & DRepTbl.Fields("SmanCode")
     DoEvents
     DoEvents
     
     RepTemplate.AddNew
     RepTemplate.Fields("MainArea") = DRepTbl.Fields("MainArea")
     RepTemplate.Fields("AreaName") = DRepTbl.Fields("AreaName")
     RepTemplate.Fields("AreaCode") = DRepTbl.Fields("AreaCode")
     RepTemplate.Fields("AreaDesc") = DRepTbl.Fields("AreaDesc")
     RepTemplate.Fields("SmanCode") = DRepTbl.Fields("SmanCode")
     RepTemplate.Fields("SmanDesc") = DRepTbl.Fields("SmanDesc")
     RepTemplate.Fields("TgtSales") = DRepTbl.Fields("SlsTgt")
     RepTemplate.Fields("TgtColl") = DRepTbl.Fields("ColTgt")
     
     frmMain.StBar.Panels(2).Text = "Calculating Aging Balance For Salesman - " & DRepTbl.Fields("SmanCode")
     DoEvents
     DoEvents
     
     RepTemplate.Fields("OpAging") = GetAgingBalSalesMan(DRepTbl.Fields("SmanCode"))
     
     frmMain.StBar.Panels(2).Text = "Calculating Day Sales For Salesman - " & DRepTbl.Fields("SmanCode")
     DoEvents
     DoEvents
     
     RepTemplate.Fields("DaySales") = GetDaySales(DRepTbl.Fields("SmanCode"), DTPicker1.Value)
     
     frmMain.StBar.Panels(2).Text = "Calculating MTD Sales For Salesman - " & DRepTbl.Fields("SmanCode")
     DoEvents
     DoEvents
     
     RepTemplate.Fields("MTDSales") = GetMonthSales(DRepTbl.Fields("SmanCode"), DTPicker1.Value)
     RepTemplate.Fields("LMTDSales") = GetMonthSales(DRepTbl.Fields("SmanCode"), Lmonth)

     frmMain.StBar.Panels(2).Text = "Calculating Day Collection For Salesman - " & DRepTbl.Fields("SmanCode")
     DoEvents
     DoEvents
     
     RepTemplate.Fields("DayColl") = GetDayColl(DRepTbl.Fields("SmanCode"), DTPicker1.Value)
     
     frmMain.StBar.Panels(2).Text = "Calculating MTD Collection For Salesman - " & DRepTbl.Fields("SmanCode")
     DoEvents
     DoEvents
     
     RepTemplate.Fields("MtdColl") = GetMonthColl(DRepTbl.Fields("SmanCode"), DTPicker1.Value)
     RepTemplate.Fields("LMTDColl") = GetMonthColl(DRepTbl.Fields("SmanCode"), Lmonth)
     
     frmMain.StBar.Panels(2).Text = "Calculating DN/CN Values For Salesman - " & DRepTbl.Fields("SmanCode")
     DoEvents
     DoEvents
     
     RepTemplate.Fields("DNCN") = GetMonthDNCN(DRepTbl.Fields("SmanCode"), DTPicker1.Value)
     DoEvents
     DoEvents
     
     RepTemplate.Fields("BlanketOrder") = GetMonthBlanketSales(DRepTbl.Fields("SmanCode"), DTPicker1.Value)
     DoEvents
     DoEvents

     RepTemplate.Update
     
     DoEvents
     DoEvents
     DRepTbl.MoveNext
     
   Loop Until DRepTbl.EOF
     
     frmMain.StBar.Panels(2).Text = "Process Finished"
   
End If

DRepTbl.Close
RepTemplate.Close

Set DRepTbl = Nothing
Set RepTemplate = Nothing

Dim RPT As New rptSlsCollSummary
CRViewer1.ReportSource = RPT
CRViewer1.ViewReport
End Sub


Private Sub Form_Load()
Dim RPT As New rptSlsCollSummary
CRViewer1.ReportSource = RPT
CRViewer1.ViewReport
DTPicker1.Value = Date
End Sub
