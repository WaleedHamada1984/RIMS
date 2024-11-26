VERSION 5.00
Begin VB.Form frmIncStmt 
   Caption         =   "Income Statement"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14145
   Icon            =   "frmIncStmt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   14145
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   840
      TabIndex        =   15
      Top             =   1800
      Width           =   6255
      Begin VB.TextBox TXTCompany 
         Height          =   285
         Left            =   3960
         TabIndex        =   18
         Text            =   "00050"
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton OPTXA 
         Caption         =   "XA"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OPTAA 
         Caption         =   "AA"
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Company"
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   840
      TabIndex        =   12
      Top             =   1080
      Width           =   6255
      Begin VB.OptionButton OptRegular 
         Caption         =   "Regular"
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton OptBudget 
         Caption         =   "Budget"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton expExcel 
      Caption         =   "Export To Excel"
      Height          =   1215
      Left            =   3720
      TabIndex        =   11
      Top             =   5520
      Width           =   3255
   End
   Begin VB.TextBox tMonth 
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton genCont 
      Caption         =   "Process"
      Height          =   1215
      Left            =   480
      TabIndex        =   8
      Top             =   5520
      Width           =   3135
   End
   Begin VB.TextBox tYear 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   6255
      Begin VB.OptionButton Option2 
         Caption         =   "Excluding Inter Company"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Total Company"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Export To Excel"
      Height          =   1095
      Left            =   3960
      TabIndex        =   1
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Month"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Sales && Contribution"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   5040
      Width           =   6495
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   7080
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label1 
      Caption         =   "Year"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   4440
      Width           =   735
   End
End
Attribute VB_Name = "frmIncStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp_SQL As String
Dim ThisYear As Integer, LastYear As Integer
Private Sub Command1_Click()
Dim rsIncStmt As Recordset, iLOOP As Integer, Fnam As String

If OptBudget Then
   Ltype = "BA"
Else
    If OPTXA Then
        Ltype = "XA"
    Else
       Ltype = "AA"
    End If
End If

ThisYear = tYear
LastYear = tYear - 1

SQLConn.Execute "DELETE From FinINCSTMT"
    
    Set rsIncStmt = New Recordset
    rsIncStmt.CursorLocation = adUseClient
    rsIncStmt.Open "SELECT * FROM FinINCSTMT", SQLConn, adOpenStatic, adLockOptimistic
' ---------------------------------------------------------------------------------
' SALES -   TISSUE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetNetSalesGL "10", ThisYear, TXTCompany.Text
    Else
       GetNetSalesGLINT "10", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 1
    rsIncStmt("RepCode") = "SALES"
    rsIncStmt("RepDesc") = "Tissue"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' SALES -   Disposable
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetNetSalesGL "20", ThisYear, TXTCompany.Text
    Else
       GetNetSalesGLINT "20", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 2
    rsIncStmt("RepCode") = "SALES"
    rsIncStmt("RepDesc") = "Disposable"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' SALES -   Aluminum & Cling
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetNetSalesGL "30", ThisYear, TXTCompany.Text
    Else
       GetNetSalesGLINT "30", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 3
    rsIncStmt("RepCode") = "SALES"
    rsIncStmt("RepDesc") = "Aluminum & Cling"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' SALES -   PE & Others
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetTloaderGL ">= '40' AND PRODDTA.F0901.GMR021 <= '50' ", ThisYear, TXTCompany.Text
    Else
       GetTloaderGLINT ">= '40' AND PRODDTA.F0901.GMR021 <= '50' ", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 4
    rsIncStmt("RepCode") = "SALES"
    rsIncStmt("RepDesc") = "PE & Others"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' TL  -   TISSUE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetTloaderGL ">= '1003' AND PRODDTA.F0901.GMR021 <= '1004' ", ThisYear, TXTCompany.Text
    Else
       GetTloaderGLINT ">= '1003' AND PRODDTA.F0901.GMR021 <= '1004' AND PRODDTA.F0901.GMSUB >= '100300' AND PRODDTA.F0901.GMSUB <= '100312'  Or (PRODDTA.F0901.GMSUB)>='100400' And (PRODDTA.F0901.GMSUB)<='100412' ", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 5
    rsIncStmt("RepCode") = "TL"
   ' rsIncStmt("RepDesc") = "Tissue - Trade Loader"
    rsIncStmt("RepDesc") = "Trade Loader - All Items"
     
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' REBATE & GONDOLAS -   TISSUE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetTloaderGL ">= '1005' AND PRODDTA.F0901.GMR021 <= '1006' ", ThisYear, TXTCompany.Text
    Else
       GetTloaderGLINT ">= '1005' AND PRODDTA.F0901.GMR021 <= '1006' AND PRODDTA.F0901.GMSUB >= '1005' AND PRODDTA.F0901.GMSUB <= '100512' AND PRODDTA.F0902.GBLT ='" & Ltype & "' Or  PRODDTA.F0902.GBFY = " & ThisYear & "  AND PRODDTA.F0901.GMSUB >= '100600' AND PRODDTA.F0901.GMSUB <= '100612'  ", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 6
    rsIncStmt("RepCode") = "GOND"
    'rsIncStmt("RepDesc") = "Tissue  - Rebates & Gondolas"
    rsIncStmt("RepDesc") = "Rebates & Gondolas - ALL Items"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' TL  -   DISPOSABLE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetTloaderGL ">= '1003' AND PRODDTA.F0901.GMR021 <= '1004' AND PRODDTA.F0901.GMSUB >= '100321' AND PRODDTA.F0901.GMSUB <= '100324'  Or (PRODDTA.F0901.GMSUB)>='100421' And (PRODDTA.F0901.GMSUB)<='100424' ", ThisYear, TXTCompany.Text
    Else
       GetTloaderGLINT ">= '1003' AND PRODDTA.F0901.GMR021 <= '1004' AND PRODDTA.F0901.GMSUB >= '100321' AND PRODDTA.F0901.GMSUB <= '100324'  Or (PRODDTA.F0901.GMSUB)>='100421' And (PRODDTA.F0901.GMSUB)<='100424' ", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 7
    rsIncStmt("RepCode") = "TL"
    rsIncStmt("RepDesc") = "Disposable - Trade Loader"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' REBATE & GONDOLAS -   DISPOSABLE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetTloaderGL ">= '1005' AND PRODDTA.F0901.GMR021 <= '1006' AND PRODDTA.F0901.GMSUB >= '100521' AND PRODDTA.F0901.GMSUB <= '100524' AND PRODDTA.F0902.GBLT ='" & Ltype & "'  Or  PRODDTA.F0902.GBFY = " & ThisYear & "  AND PRODDTA.F0901.GMSUB >= '100621' AND PRODDTA.F0901.GMSUB <= '100624'  ", ThisYear, TXTCompany.Text
    Else
       GetTloaderGLINT ">= '1005' AND PRODDTA.F0901.GMR021 <= '1006' AND PRODDTA.F0901.GMSUB >= '100521' AND PRODDTA.F0901.GMSUB <= '100524' AND PRODDTA.F0902.GBLT ='" & Ltype & "' Or  PRODDTA.F0902.GBFY = " & ThisYear & "  AND PRODDTA.F0901.GMSUB >= '100621' AND PRODDTA.F0901.GMSUB <= '100624'  ", ThisYear, TXTCompany.Text
    End If
    
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 8
    rsIncStmt("RepCode") = "GOND"
    rsIncStmt("RepDesc") = "Disposable  - Rebates & Gondolas"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' TL  -   ALUMINUM & CLING
' ---------------------------------------------------------------------------------

    If Option1 = True Then
       GetTloaderGL ">= '1003' AND PRODDTA.F0901.GMR021 <= '1004' AND PRODDTA.F0901.GMSUB >= '100331' AND PRODDTA.F0901.GMSUB <= '100332'  Or (PRODDTA.F0901.GMSUB)>='100431' And (PRODDTA.F0901.GMSUB)<='100432' ", ThisYear, TXTCompany.Text
    Else
       GetTloaderGLINT ">= '1003' AND PRODDTA.F0901.GMR021 <= '1004' AND PRODDTA.F0901.GMSUB >= '100331' AND PRODDTA.F0901.GMSUB <= '100332'  Or (PRODDTA.F0901.GMSUB)>='100431' And (PRODDTA.F0901.GMSUB)<='100432' ", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 9
    rsIncStmt("RepCode") = "TL"
    rsIncStmt("RepDesc") = "Aluminium & Cling - Trade Loader"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' REBATE & GONDOLAS -   ALUMINIUM & CLING
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetTloaderGL ">= '1005' AND PRODDTA.F0901.GMR021 <= '1006' AND PRODDTA.F0901.GMSUB >= '100531' AND PRODDTA.F0901.GMSUB <= '100532' AND PRODDTA.F0902.GBLT ='" & Ltype & "'  Or  PRODDTA.F0902.GBFY = " & ThisYear & "  AND PRODDTA.F0901.GMSUB >= '100631' AND PRODDTA.F0901.GMSUB <= '100632'  ", ThisYear, TXTCompany.Text
    Else
       GetTloaderGLINT ">= '1005' AND PRODDTA.F0901.GMR021 <= '1006' AND PRODDTA.F0901.GMSUB >= '100531' AND PRODDTA.F0901.GMSUB <= '100532' AND PRODDTA.F0902.GBLT ='" & Ltype & "' Or  PRODDTA.F0902.GBFY = " & ThisYear & "  AND PRODDTA.F0901.GMSUB >= '100631' AND PRODDTA.F0901.GMSUB <= '100632'  ", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 10
    rsIncStmt("RepCode") = "GOND"
    rsIncStmt("RepDesc") = "Aluminium & Cling  - Rebates & Gondolas"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' TL  -   PE & OTHERS
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetTloaderGL ">= '1003' AND PRODDTA.F0901.GMR021 <= '1004' AND PRODDTA.F0901.GMSUB >= '100341' AND PRODDTA.F0901.GMSUB <= '100399'  Or (PRODDTA.F0901.GMSUB)>='100441' And (PRODDTA.F0901.GMSUB)<='100499' ", ThisYear, TXTCompany.Text
    Else
       GetTloaderGLINT ">= '1003' AND PRODDTA.F0901.GMR021 <= '1004' AND PRODDTA.F0901.GMSUB >= '100341' AND PRODDTA.F0901.GMSUB <= '100399'  Or (PRODDTA.F0901.GMSUB)>='100441' And (PRODDTA.F0901.GMSUB)<='100499' ", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 11
    rsIncStmt("RepCode") = "TL"
    rsIncStmt("RepDesc") = "PE & Others - Trade Loader"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' REBATE & GONDOLAS -   PE & OTHERS
' ---------------------------------------------------------------------------------
    
    If Option1 = True Then
       GetTloaderGL ">= '1005' AND PRODDTA.F0901.GMR021 <= '1006' AND PRODDTA.F0901.GMSUB >= '100541' AND PRODDTA.F0901.GMSUB <= '100553' AND PRODDTA.F0902.GBLT ='" & Ltype & "'  Or  PRODDTA.F0902.GBFY = " & ThisYear & "  AND PRODDTA.F0901.GMSUB >= '100641' AND PRODDTA.F0901.GMSUB <= '100653'  ", ThisYear, TXTCompany.Text
    Else
       GetTloaderGLINT ">= '1005' AND PRODDTA.F0901.GMR021 <= '1006' AND PRODDTA.F0901.GMSUB >= '100541' AND PRODDTA.F0901.GMSUB <= '100553' AND PRODDTA.F0902.GBLT ='" & Ltype & "' Or  PRODDTA.F0902.GBFY = " & ThisYear & "  AND PRODDTA.F0901.GMSUB >= '100641' AND PRODDTA.F0901.GMSUB <= '100653'  ", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 12
    rsIncStmt("RepCode") = "GOND"
    rsIncStmt("RepDesc") = "PE & Others  - Rebates & Gondolas"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' COGS -   TISSUE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetNetSalesGL "110", ThisYear, TXTCompany.Text
    Else
       GetNetSalesGLINT "110", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 13
    rsIncStmt("RepCode") = "COGS"
    rsIncStmt("RepDesc") = "Tissue"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' COGS -   DISPOSABLE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetNetSalesGL "120", ThisYear, TXTCompany.Text
    Else
       GetNetSalesGLINT "120", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 14
    rsIncStmt("RepCode") = "COGS"
    rsIncStmt("RepDesc") = "Disposable"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' COGS -   ALUMINUM & CLING
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetNetSalesGL "130", ThisYear, TXTCompany.Text
    Else
       GetNetSalesGLINT "130", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 15
    rsIncStmt("RepCode") = "COGS"
    rsIncStmt("RepDesc") = "Aluminum & Cling"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' COGS -   PE & OTHERS
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetTloaderGL ">= '140' AND PRODDTA.F0901.GMR021 <= '150'", ThisYear, TXTCompany.Text
    Else
       GetTloaderGLINT ">= '140' AND PRODDTA.F0901.GMR021 <= '150'", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 16
    rsIncStmt("RepCode") = "COGS"
    rsIncStmt("RepDesc") = "PE & Others"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' OVERHEAD -   PLANT INDIRECT EXPENSE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetOverHeadGL "PLT", ThisYear, TXTCompany.Text
    Else
       GetOverHeadGLINT "PIN", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 17
    rsIncStmt("RepCode") = "OH"
    rsIncStmt("RepDesc") = "Plant Indirect Expenses"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' OVERHEAD -   ADMINISTRATION EXPENSE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetOverHeadGL "FNA", ThisYear, TXTCompany.Text
    Else
       GetOverHeadGLINT "FNA", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 18
    rsIncStmt("RepCode") = "OH"
    rsIncStmt("RepDesc") = "Administration Expenses"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' OVERHEAD -   LOGISTIC EXPENSE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetOverHeadGL "LOG", ThisYear, TXTCompany.Text
    Else
       GetOverHeadGLINT "LOG", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 19
    rsIncStmt("RepCode") = "OH"
    rsIncStmt("RepDesc") = "Logistic Expenses"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' OVERHEAD -   MARKETING EXPENSE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetOverHeadGL "MAK", ThisYear, TXTCompany.Text
    Else
       GetOverHeadGLINT "MAK", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 20
    rsIncStmt("RepCode") = "OH"
    rsIncStmt("RepDesc") = "Marketing Expenses"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' OVERHEAD -   EGYPT EXPENSE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetOverHeadGL "EGP", ThisYear, TXTCompany.Text
    Else
       GetOverHeadGLINT "EPS", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 21
    rsIncStmt("RepCode") = "OH"
    rsIncStmt("RepDesc") = "Egypt Selling Expenses"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' OVERHEAD -   GULF EXPENSE
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       GetOverHeadGL "GUF", ThisYear, TXTCompany.Text
    Else
       GetOverHeadGLINT "GUF", ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 22
    rsIncStmt("RepCode") = "OH"
    rsIncStmt("RepDesc") = "Gulf Selling Expenses"
    
    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
    
' ---------------------------------------------------------------------------------
' OVERHEAD -   INLAND FREIGHT
' ---------------------------------------------------------------------------------
    
    tmp_SQL = "PRODDTA.F0901.GMOBJ = '9002' AND PRODDTA.F0901.GMSUB = '0701' AND PRODDTA.F0902.GBLT = '" & Ltype & "'  "
    tmp_SQL = tmp_SQL & "AND PRODDTA.F0902.GBFY = " & ThisYear & " AND "
    tmp_SQL = tmp_SQL & "LTRIM(PRODDTA.F0901.GMMCU) <> '0201' AND LTRIM(PRODDTA.F0901.GMMCU) <> '0202' AND LTRIM(PRODDTA.F0901.GMMCU) <> '0301' AND  "
    tmp_SQL = tmp_SQL & "LTRIM(PRODDTA.F0901.GMMCU) <> '0302' AND LTRIM(PRODDTA.F0901.GMMCU) <> '0303' AND "
    tmp_SQL = tmp_SQL & "LTRIM(PRODDTA.F0901.GMMCU) <> '1300' AND LTRIM(PRODDTA.F0901.GMMCU) <> '1310' "
    
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 23
    rsIncStmt("RepCode") = "OH"
    rsIncStmt("RepDesc") = "Inland Freight"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update

' ---------------------------------------------------------------------------------
' PROMOTION
' ---------------------------------------------------------------------------------
    
    tmp_SQL = "PRODDTA.F0901.GMR021 = '1002' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    tmp_SQL = tmp_SQL & "OR PRODDTA.F0901.GMR021 = '1007' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 24
    rsIncStmt("RepCode") = "PR"
    rsIncStmt("RepDesc") = "Promotions"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update

' ---------------------------------------------------------------------------------
' ADVERTISING
' ---------------------------------------------------------------------------------
    
    tmp_SQL = "PRODDTA.F0901.GMR021 = '1001' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 25
    rsIncStmt("RepCode") = "PR"
    rsIncStmt("RepDesc") = "Advertising"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update

' ---------------------------------------------------------------------------------
' TISSUE - SALES RETURN
' ---------------------------------------------------------------------------------
    
    tmp_SQL = "PRODDTA.F0901.GMR021 = '10' AND PRODDTA.F0901.GMOBJ = '6201' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 29
    rsIncStmt("RepCode") = "SRTN"
    rsIncStmt("RepDesc") = "Tissue"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update
' ---------------------------------------------------------------------------------
' DISPOSABLE - SALES RETURN
' ---------------------------------------------------------------------------------
    
    tmp_SQL = "PRODDTA.F0901.GMR021 = '20' AND PRODDTA.F0901.GMOBJ = '6201' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 30
    rsIncStmt("RepCode") = "SRTN"
    rsIncStmt("RepDesc") = "Disposable"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update

' ---------------------------------------------------------------------------------
' ALUMINUM & CLING - SALES RETURN
' ---------------------------------------------------------------------------------
    
    tmp_SQL = "PRODDTA.F0901.GMR021 = '30' AND PRODDTA.F0901.GMOBJ = '6201' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 31
    rsIncStmt("RepCode") = "SRTN"
    rsIncStmt("RepDesc") = "Aluminum & Cling"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update

' ---------------------------------------------------------------------------------
' PE & OTHERS - SALES RETURN
' ---------------------------------------------------------------------------------
    
    tmp_SQL = "PRODDTA.F0901.GMR021 >= '40' AND PRODDTA.F0901.GMR021 <= '50' AND PRODDTA.F0901.GMOBJ = '6201' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 32
    rsIncStmt("RepCode") = "SRTN"
    rsIncStmt("RepDesc") = "PE & Others"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update

' ---------------------------------------------------------------------------------
' DEPRECIATION
' ---------------------------------------------------------------------------------
    
    tmp_SQL = "PRODDTA.F0901.GMR021 = '9010' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 34
    rsIncStmt("RepCode") = "DEP"
    rsIncStmt("RepDesc") = "Depreciation"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update

' ---------------------------------------------------------------------------------
' BANK CHARGES / INTEREST
' ---------------------------------------------------------------------------------
    
    tmp_SQL = "PRODDTA.F0901.GMR021 = '9011' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 36
    rsIncStmt("RepCode") = "INT"
    rsIncStmt("RepDesc") = "Interest"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update

' ---------------------------------------------------------------------------------
' OTHER EXPENSE / OTHER GAINS AND LOSSES
' ---------------------------------------------------------------------------------
    
    tmp_SQL = "PRODDTA.F0901.GMR021 = '9012' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 38
    rsIncStmt("RepCode") = "EXP"
    rsIncStmt("RepDesc") = "Other Gains & Losses"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update

' ---------------------------------------------------------------------------------
' COMMON GROUP / CONSUMER DIVISION EXPENSES
' ---------------------------------------------------------------------------------
    If Option1 = True Then
       tmp_SQL = "PRODDTA.F0901.GMOBJ = '9014' AND LTRIM(PRODDTA.F0901.GMMCU) = '1400' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
'       tmp_SQL = tmp_SQL & " AND LTRIM(PRODDTA.F0901.GMMCU) = '1400' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    Else
       tmp_SQL = "PRODDTA.F0901.GMOBJ = '9014'  "
       tmp_SQL = tmp_SQL & " AND LTRIM(PRODDTA.F0901.GMMCU) = '1400' "
       tmp_SQL = tmp_SQL & "AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    End If
    
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 40
    rsIncStmt("RepCode") = "CDE"
    rsIncStmt("RepDesc") = "Consumer Division Expenses"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update

' ---------------------------------------------------------------------------------
' MANAGEMENT EXPENSES
' ---------------------------------------------------------------------------------
    If Option1 = True Then
        tmp_SQL = "PRODDTA.F0901.GMR021 >= '9014' AND PRODDTA.F0901.GMR021 <= '9015' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
    Else
       tmp_SQL = "PRODDTA.F0901.GMR021 >= '9014' AND PRODDTA.F0901.GMR021 <= '9015' AND PRODDTA.F0902.GBLT = '" & Ltype & "' AND PRODDTA.F0902.GBFY = " & ThisYear & " "
       tmp_SQL = tmp_SQL & " AND LTRIM(PRODDTA.F0901.GMMCU) <> '1400' "
       tmp_SQL = tmp_SQL & " AND LTRIM(PRODDTA.F0901.GMMCU) <> '1101' "
       tmp_SQL = tmp_SQL & " AND LTRIM(PRODDTA.F0901.GMMCU) <> '1102' "
    End If
    If Option1 = True Then
       GetGL12Slots tmp_SQL, ThisYear, TXTCompany.Text
    Else
       GetGL12SlotsINT tmp_SQL, ThisYear, TXTCompany.Text
    End If
    rsIncStmt.AddNew
    rsIncStmt("RepSerial") = 42
    rsIncStmt("RepCode") = "MGTE"
    rsIncStmt("RepDesc") = "Management Expenses"

    For iLOOP = 1 To 12
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           rsIncStmt.Fields(Fnam) = Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsIncStmt.Update

' ---------------------------------------------------------------------------------

End Sub

Private Sub Command2_Click()
Set tmpRset = New Recordset
tmpRset.CursorLocation = adUseClient
tmpRset.Open "SELECT * From FinINCSTMT", SQLConn, adOpenStatic, adLockOptimistic

Recordset2Excel tmpRset

tmpRset.Close
Set tmpRset = Nothing

End Sub


Private Sub expExcel_Click()
Set tmpRset = New Recordset
tmpRset.CursorLocation = adUseClient
tmpRset.Open "SELECT * From FinSLSCONT", SQLConn, adOpenStatic, adLockOptimistic

Recordset2Excel tmpRset

tmpRset.Close
Set tmpRset = Nothing

End Sub

Private Sub Form_Load()
tYear = Year(Now) - 2000
tMonth = Month(Now) - 1
End Sub

Private Sub genCont_Click()
Dim SqlCont As Recordset
Dim rsCont As Recordset, iLOOP As Integer, Fnam As String
Dim SLSTotal As Double, Fill_PCS As Boolean

If OptBudget Then
   Ltype = "BA"
Else
   Ltype = "AA"
End If


ThisYear = tYear
LastYear = tYear - 1

    Set SqlCont = New Recordset
    SqlCont.CursorLocation = adUseClient
    SqlCont.Open "SELECT * FROM ContTable Order By Serial", SQLConn, adOpenStatic, adLockOptimistic


SQLConn.Execute "DELETE From FinSLSCONT"
    
    Set rsCont = New Recordset
    rsCont.CursorLocation = adUseClient
    rsCont.Open "SELECT * FROM FinSLSCONT", SQLConn, adOpenStatic, adLockOptimistic
    
    SqlCont.MoveFirst
    
    Do While Not SqlCont.EOF
' ---------------------------------------------------------------------------------
' TISSUE
' ---------------------------------------------------------------------------------
             tmp_SQL = "GMOBJ = '6101' AND GMSUB = '" & SqlCont.Fields("LineCode") & "' AND GBFY = " & tYear & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '6201' AND GMSUB = '" & SqlCont.Fields("LineCode") & "' AND GBFY = " & tYear & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '6301' AND GMSUB = '" & SqlCont.Fields("LineCode") & "' AND GBFY = " & tYear & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102'  "
   
   GetGL12Slots tmp_SQL, tYear, TXTCompany.Text
    
    rsCont.AddNew
    rsCont("RepSerial") = SqlCont.Fields("Serial")
    rsCont("RepCode") = "CONT"
    rsCont("RepDesc") = SqlCont.Fields("Description")
    SLSTotal = 0
    For iLOOP = 1 To tMonth
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           SLSTotal = SLSTotal + Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsCont.Fields("SLSTYEAR") = SLSTotal
    
             tmp_SQL = "GMOBJ = '6101' AND GMSUB = '" & SqlCont.Fields("LineCode") & "' AND GBFY = " & tYear - 1 & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '6201' AND GMSUB = '" & SqlCont.Fields("LineCode") & "' AND GBFY = " & tYear - 1 & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '6301' AND GMSUB = '" & SqlCont.Fields("LineCode") & "' AND GBFY = " & tYear - 1 & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102'  "

    
    GetGL12Slots tmp_SQL, tYear - 1, TXTCompany.Text
    
    SLSTotal = 0
    For iLOOP = 1 To tMonth
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           SLSTotal = SLSTotal + Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsCont.Fields("SLSPYEAR") = SLSTotal
'------------------------------------------------
' TRADELOADER & REBATE
'-------------------------------------------------
             tmp_SQL = "GMOBJ = '9002' AND GMSUB = '1003" & LTrim(SqlCont.Fields("TLCode")) & "' AND GBFY = " & tYear & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '9002' AND GMSUB = '1004" & LTrim(SqlCont.Fields("TLCode")) & "' AND GBFY = " & tYear & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '9002' AND GMSUB = '1005" & LTrim(SqlCont.Fields("TLCode")) & "' AND GBFY = " & tYear & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '9002' AND GMSUB = '1006" & LTrim(SqlCont.Fields("TLCode")) & "' AND GBFY = " & tYear & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '9002' AND GMSUB = '1007" & LTrim(SqlCont.Fields("TLCode")) & "' AND GBFY = " & tYear & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' "
  
   GetGL12Slots tmp_SQL, tYear, TXTCompany.Text
    
    SLSTotal = 0
    For iLOOP = 1 To tMonth
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           SLSTotal = SLSTotal + Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsCont.Fields("TLRTYEAR") = SLSTotal
   
             tmp_SQL = "GMOBJ = '9002' AND GMSUB = '1003" & LTrim(SqlCont.Fields("TLCode")) & "' AND GBFY = " & tYear - 1 & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '9002' AND GMSUB = '1004" & LTrim(SqlCont.Fields("TLCode")) & "' AND GBFY = " & tYear - 1 & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '9002' AND GMSUB = '1005" & LTrim(SqlCont.Fields("TLCode")) & "' AND GBFY = " & tYear - 1 & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '9002' AND GMSUB = '1006" & LTrim(SqlCont.Fields("TLCode")) & "' AND GBFY = " & tYear - 1 & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' OR "
   tmp_SQL = tmp_SQL & "GMOBJ = '9002' AND GMSUB = '1007" & LTrim(SqlCont.Fields("TLCode")) & "' AND GBFY = " & tYear - 1 & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' "
  
     
   GetGL12Slots tmp_SQL, tYear - 1, TXTCompany.Text
    
    SLSTotal = 0
    For iLOOP = 1 To tMonth
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           SLSTotal = SLSTotal + Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsCont.Fields("TLRPYEAR") = SLSTotal
'------------------------------------------------
' AFTER TRADELOADER & REBATE
'-------------------------------------------------
   tmp_SQL = "GMOBJ = '8100' AND GMSUB = '" & SqlCont.Fields("LineCode") & "' AND GBFY = " & tYear & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' "
  
   GetGL12Slots tmp_SQL, tYear, TXTCompany.Text
    
    SLSTotal = 0
    For iLOOP = 1 To tMonth
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           SLSTotal = SLSTotal + Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsCont.Fields("CNTTYEAR") = SLSTotal
   
    tmp_SQL = "GMOBJ = '8100' AND GMSUB = '" & SqlCont.Fields("LineCode") & "' AND GBFY = " & tYear - 1 & " AND GBLT = '" & Ltype & "' AND LTRIM(GBMCU)<> '1101' AND LTRIM(GBMCU)<> '1102' "
  
     
   GetGL12Slots tmp_SQL, tYear - 1, TXTCompany.Text
    
    SLSTotal = 0
    For iLOOP = 1 To tMonth
        If iLOOP < 10 Then
           Fnam = "Mth0" & Trim(Str(iLOOP))
        Else
           Fnam = "Mth" & Trim(Str(iLOOP))
        End If
           SLSTotal = SLSTotal + Round(GL_Value(iLOOP) / 100, 0)
    Next iLOOP
    rsCont.Fields("CNTPYEAR") = SLSTotal
 
'Tonnage
     pb_Cond_Stmt = " <> 'INT"

    GetYTDSalesC SqlCont.Fields("GLCODE"), tYear, tMonth
    GetYTDSalesP SqlCont.Fields("GLCODE"), tYear - 1, tMonth
    
    If Trim(SqlCont.Fields("GLCODE")) = "F201" Or Trim(SqlCont.Fields("GLCODE")) = "F202" Or _
       Trim(SqlCont.Fields("GLCODE")) = "F203" Or Trim(SqlCont.Fields("GLCODE")) = "F204" Then
       Fill_PCS = True
    Else
       Fill_PCS = False
    End If
         
    If Not Fill_PCS Then
       rsCont.Fields("TONTYEAR") = Cur_Yr_Nwt
       rsCont.Fields("TONPYEAR") = Prv_Yr_Nwt
    Else
       rsCont.Fields("TONTYEAR") = Cur_Yr_Pcs
       rsCont.Fields("TONPYEAR") = Prv_Yr_Pcs
    End If
    If Trim(SqlCont.Fields("GLCODE")) = "F112" Or Trim(SqlCont.Fields("GLCODE")) = "F111" Or _
       Trim(SqlCont.Fields("GLCODE")) = "F401" Or Trim(SqlCont.Fields("GLCODE")) = "F503" Or _
       Trim(SqlCont.Fields("GLCODE")) = "F501" Or Trim(SqlCont.Fields("GLCODE")) = "F403" Or _
       Trim(SqlCont.Fields("GLCODE")) = "F404" Or Trim(SqlCont.Fields("GLCODE")) = "F402" Or _
       Trim(SqlCont.Fields("GLCODE")) = "F502" Then
       rsCont.Fields("TONTYEAR") = Cur_Yr_Gwt
       rsCont.Fields("TONPYEAR") = Prv_Yr_Gwt
    End If
 
    rsCont.Update
    SqlCont.MoveNext
    
    Loop
    
    SqlCont.Close
    Set SqlCont = Nothing
    
    rsCont.Close
    Set rsCont = Nothing
End Sub
