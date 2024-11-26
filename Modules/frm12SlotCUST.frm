VERSION 5.00
Begin VB.Form frm12SlotCUST 
   Caption         =   "12 Slots Sales"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   11070
   WindowState     =   2  'Maximized
   Begin VB.TextBox CustAN8 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Export to Excel"
      Height          =   1095
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Address Number"
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
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "frm12SlotCUST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim AA As Boolean
Command1.Enabled = False
AA = GenSalesByCUST(Val(CustAN8))
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Set tmpRset = New Recordset
tmpRset.CursorLocation = adUseClient
tmpRset.Open "SELECT * From tblSalesCUST", SQLConn, adOpenStatic, adLockOptimistic

Recordset2Excel tmpRset

tmpRset.Close
Set tmpRset = Nothing
End Sub
