VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTran12SlotsCust 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.TextBox CustID 
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   1695
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
      Left            =   6720
      TabIndex        =   1
      Top             =   120
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
Attribute VB_Name = "frmTran12SlotsCust"
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

SQLConn.Execute "DELETE FROM TranList"

Set DRepTbl = New Recordset
DRepTbl.CursorLocation = adUseClient
DRepTbl.Open "SELECT * From DailyRep Order By MainArea,AreaCode,SmanCode", SQLConn, adOpenStatic, adLockOptimistic

Set RepTemplate = New Recordset
RepTemplate.CursorLocation = adUseClient
RepTemplate.Open "SELECT * From TranList", SQLConn, adOpenStatic, adLockOptimistic


If DRepTbl.RecordCount > 0 Then
   DRepTbl.MoveFirst
   
   Do
      
     frmMain.StBar.Panels(2).Text = "Processing Record For Salesman (Sales) - " & DRepTbl.Fields("SmanCode")
     DoEvents
     DoEvents
     
     RepTemplate.AddNew
     RepTemplate.Fields("MArea") = DRepTbl.Fields("MainArea")
     RepTemplate.Fields("MAreaName") = DRepTbl.Fields("AreaName")
     RepTemplate.Fields("ACode") = DRepTbl.Fields("AreaCode")
     RepTemplate.Fields("AName") = DRepTbl.Fields("AreaDesc")
     RepTemplate.Fields("SCode") = DRepTbl.Fields("SmanCode")
     RepTemplate.Fields("SName") = DRepTbl.Fields("SmanDesc")
     RepTemplate.Fields("RecType") = "SALE"
    
     For ISerial = 1 To 12
         Lmonth = Trim(Str(ISerial)) & "/" & LastDayInMonth(Year(Date), ISerial) & "/" & Trim(Str(Year(Date)))
         fldName = "M" & Trim(Str(ISerial))
         RepTemplate.Fields(fldName) = GetMonthSales(DRepTbl.Fields("SmanCode"), Lmonth)
     Next ISerial
     RepTemplate.Update
     
     frmMain.StBar.Panels(2).Text = "Processing Record For Salesman (Collection) - " & DRepTbl.Fields("SmanCode")
     DoEvents
     DoEvents
     
          RepTemplate.AddNew
     RepTemplate.Fields("MArea") = DRepTbl.Fields("MainArea")
     RepTemplate.Fields("MAreaName") = DRepTbl.Fields("AreaName")
     RepTemplate.Fields("ACode") = DRepTbl.Fields("AreaCode")
     RepTemplate.Fields("AName") = DRepTbl.Fields("AreaDesc")
     RepTemplate.Fields("SCode") = DRepTbl.Fields("SmanCode")
     RepTemplate.Fields("SName") = DRepTbl.Fields("SmanDesc")
     RepTemplate.Fields("RecType") = "COLL"
    
     For ISerial = 1 To 12
         Lmonth = Trim(Str(ISerial)) & "/" & LastDayInMonth(Year(Date), ISerial) & "/" & Trim(Str(Year(Date)))
         fldName = "M" & Trim(Str(ISerial))
         RepTemplate.Fields(fldName) = GetMonthColl(DRepTbl.Fields("SmanCode"), Lmonth)
     Next ISerial
     RepTemplate.Update

     frmMain.StBar.Panels(2).Text = "Processing Record For Salesman (DN/CN) - " & DRepTbl.Fields("SmanCode")
     DoEvents
     DoEvents
     
          RepTemplate.AddNew
     RepTemplate.Fields("MArea") = DRepTbl.Fields("MainArea")
     RepTemplate.Fields("MAreaName") = DRepTbl.Fields("AreaName")
     RepTemplate.Fields("ACode") = DRepTbl.Fields("AreaCode")
     RepTemplate.Fields("AName") = DRepTbl.Fields("AreaDesc")
     RepTemplate.Fields("SCode") = DRepTbl.Fields("SmanCode")
     RepTemplate.Fields("SName") = DRepTbl.Fields("SmanDesc")
     RepTemplate.Fields("RecType") = "DNCN"
    
     For ISerial = 1 To 12
         Lmonth = Trim(Str(ISerial)) & "/" & LastDayInMonth(Year(Date), ISerial) & "/" & Trim(Str(Year(Date)))
         fldName = "M" & Trim(Str(ISerial))
         RepTemplate.Fields(fldName) = GetMonthDNCN(DRepTbl.Fields("SmanCode"), Lmonth)
     Next ISerial
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


Private Sub Form_Load()
'Dim RPT As New rptSlsCollSummary
'CRViewer1.ReportSource = RPT
'CRViewer1.ViewReport
DTPicker1.Value = Date - 1
End Sub
