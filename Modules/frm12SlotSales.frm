VERSION 5.00
Begin VB.Form frm12SlotSales 
   Caption         =   "12 Slots Sales"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   Icon            =   "frm12SlotSales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   11070
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtYear 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Export to Excel"
      Height          =   1095
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Fiscal Year"
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
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frm12SlotSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim AA As Boolean
Command1.Enabled = False
AA = GenSalesByBU(txtYear)
Command1.Enabled = True

End Sub

Private Sub Command2_Click()
Set tmpRset = New Recordset
tmpRset.CursorLocation = adUseClient
tmpRset.Open "SELECT * From tblSales", SQLConn, adOpenStatic, adLockOptimistic

Recordset2Excel tmpRset

tmpRset.Close
Set tmpRset = Nothing
End Sub

Private Sub Form_Load()
txtYear = Year(Now) - 2000
End Sub
