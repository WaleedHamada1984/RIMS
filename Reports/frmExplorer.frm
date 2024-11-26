VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmExplorer 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13140
   Icon            =   "frmExplorer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   13140
   Begin MSComCtl2.DTPicker fromDate 
      Height          =   315
      Left            =   6240
      TabIndex        =   5
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   20709379
      CurrentDate     =   39364
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7095
      Left            =   5040
      TabIndex        =   1
      Top             =   1320
      Width           =   11295
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   14631
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker toDate 
      Height          =   315
      Left            =   9720
      TabIndex        =   7
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   20709379
      CurrentDate     =   39364
   End
   Begin VB.Label Label3 
      Caption         =   "To Date"
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
      Left            =   8640
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "From Date"
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
      Left            =   5040
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   5040
      X2              =   16320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblProcess 
      Height          =   255
      Left            =   10680
      TabIndex        =   4
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Report:"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   5040
      X2              =   16320
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
