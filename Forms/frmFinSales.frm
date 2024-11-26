VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmFinSales 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmFinSales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Width           =   8055
      Begin VB.OptionButton Option1 
         Caption         =   "Eastern Province"
         Height          =   195
         Index           =   2
         Left            =   5520
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Export"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Excluding Inter Company"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.TextBox txtMonth 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   210
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   120
      TabIndex        =   1
      Top             =   840
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
   Begin VB.Label Label1 
      Caption         =   "Month"
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
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmFinSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Cur_Year As Integer, Prv_Year As Integer, rpt_Rset As Recordset
Dim RepTemplate As Recordset, Drep As Recordset, Fill_PCS As Boolean
Dim MthNo As Integer

If Not IsNumeric(txtMonth) Then
   txtMonth.SetFocus
   Exit Sub
End If


If Option1(0) = True Then
   pb_Cond_Stmt = " <> 'INT"
   pb_RepType = "INT"
ElseIf Option1(1) = True Then
   pb_Cond_Stmt = " = 'EXP"
   pb_RepType = "EXP"
ElseIf Option1(2) = True Then
   pb_Cond_Stmt = " = 'EAS"
   pb_RepType = "EAS"
End If

SQLConn.Execute "DELETE FROM FINREPSLS01"

Set RepTemplate = New Recordset
RepTemplate.CursorLocation = adUseClient
RepTemplate.Open "SELECT * From FINREPSLS01", SQLConn, adOpenStatic, adLockOptimistic

Set Drep = New Recordset
Drep.CursorLocation = adUseClient
Drep.Open "SELECT * From LineTable", SQLConn, adOpenStatic, adLockOptimistic


    Cur_Year = Year(Date) - 2000
    Prv_Year = Cur_Year - 1
    MthNo = txtMonth
    
    
    If Drep.RecordCount > 0 Then
    Drep.MoveFirst
    Do
         RepTemplate.AddNew
         RepTemplate.Fields("Line") = Trim(Drep.Fields("Line"))
         RepTemplate.Fields("LineDesc") = Drep.Fields("LineDesc")
         
         If Trim(Drep.Fields("Line")) = "F201" Or Trim(Drep.Fields("Line")) = "F202" Or _
            Trim(Drep.Fields("Line")) = "F203" Or Trim(Drep.Fields("Line")) = "F204" Then
            Fill_PCS = True
         Else
            Fill_PCS = False
         End If
         
         GetYTDSalesC Drep.Fields("Line"), Cur_Year, txtMonth
         GetYTDSalesP Drep.Fields("Line"), Prv_Year, txtMonth
         
         RepTemplate.Fields("NWTPrvYearYTD") = Prv_Yr_Nwt
         RepTemplate.Fields("NWTCurYearYTD") = Cur_Yr_Nwt
         
         If Fill_PCS Then
            RepTemplate.Fields("QTYPrvYearYTD") = Prv_Yr_Pcs
            RepTemplate.Fields("QTYCurYearYTD") = Cur_Yr_Pcs
         Else
            RepTemplate.Fields("QTYPrvYearYTD") = Prv_Yr_Qty
            RepTemplate.Fields("QTYCurYearYTD") = Cur_Yr_Qty
         End If
         
         RepTemplate.Fields("VALPrvYearYTD") = Prv_Yr_Val
         RepTemplate.Fields("VALCurYearYTD") = Cur_Yr_Val
         
         GetMTDSalesC Drep.Fields("Line"), Cur_Year, MthNo
         GetMTDSalesP Drep.Fields("Line"), Prv_Year, MthNo
         
         RepTemplate.Fields("NWTPrvYearM") = Prv_Yr_Nwt
         RepTemplate.Fields("NWTCurYearM") = Cur_Yr_Nwt
         
         If Fill_PCS Then
            RepTemplate.Fields("QTYPrvYearM") = Prv_Yr_Pcs
            RepTemplate.Fields("QTYCurYearM") = Cur_Yr_Pcs
         Else
            RepTemplate.Fields("QTYPrvYearM") = Prv_Yr_Qty
            RepTemplate.Fields("QTYCurYearM") = Cur_Yr_Qty
         End If
         
         RepTemplate.Fields("VALPrvYearM") = Prv_Yr_Val
         RepTemplate.Fields("VALCurYearM") = Cur_Yr_Val
         RepTemplate.Update
         Drep.MoveNext
    Loop Until Drep.EOF
    End If
    
    Drep.Close
    RepTemplate.Close
    
    Set Drep = Nothing
    Set RepTemplate = Nothing
    
Dim RPT As New rptFinRep01
CRViewer1.ReportSource = RPT
CRViewer1.ViewReport
    
End Sub

Private Sub Form_Load()
Dim RPT As New rptFinRep01
CRViewer1.ReportSource = RPT
CRViewer1.ViewReport
If Option1(0) = True Then
   pb_Cond_Stmt = " <> 'INT"
   pb_RepType = "INT"
ElseIf Option1(1) = True Then
   pb_Cond_Stmt = " = 'EXP"
   pb_RepType = "EXP"
ElseIf Option1(2) = True Then
   pb_Cond_Stmt = " = 'EAS"
   pb_RepType = "EAS"
End If

End Sub


Private Sub Option1_Click(Index As Integer)
If Option1(0) = True Then
   pb_Cond_Stmt = " <> 'INT"
   pb_RepType = "INT"
ElseIf Option1(1) = True Then
   pb_Cond_Stmt = " = 'EXP"
   pb_RepType = "EXP"
ElseIf Option1(2) = True Then
   pb_Cond_Stmt = " = 'EAS"
   pb_RepType = "EAS"
End If

End Sub
