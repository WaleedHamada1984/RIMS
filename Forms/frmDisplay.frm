VERSION 5.00
Begin VB.Form frmDisplay 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblConnection 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   7815
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    lblConnection.Caption = "Connecting iSeries ... "
    DbConnection
    
    lblConnection.Caption = "Connecting SQL Serve ... "
    SQLConnection
    
    Unload Me
End Sub

