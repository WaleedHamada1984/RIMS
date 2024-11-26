VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmProductPCS 
   Caption         =   "Sales Analysis (PCS)"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmProductPCS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbYear 
      Height          =   315
      ItemData        =   "frmProductPCS.frx":000C
      Left            =   6840
      List            =   "frmProductPCS.frx":0022
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7695
      Left            =   4080
      TabIndex        =   4
      Top             =   960
      Width           =   11175
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
   Begin VB.ListBox List2 
      Height          =   3885
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   4680
      Width           =   3855
   End
   Begin VB.ListBox List1 
      Height          =   3885
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Year"
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
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblCntr 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   240
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   15720
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Business Unit (Area)"
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
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Product Line (Brand)"
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
      Top             =   4440
      Width           =   1935
   End
End
Attribute VB_Name = "frmProductPCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcess_Click()
Dim i As Integer, SQL01 As String, SQL02 As String, iCntr As Double
Dim tmpRset01 As Recordset, Brd As String
SQL01 = ""

SQLConn.Execute "DELETE From BrandSales"
For i = 0 To List1.ListCount - 1
    
    If List1.Selected(i) Then
       List1.ListIndex = i
       If SQL01 = "" Then
          SQL01 = "RTRIM(LTRIM(PRODDTA.F0101.ABMCU)) = '" & Trim(Left(List1.Text, 12)) & "' "
       Else
          SQL01 = SQL01 & "OR RTRIM(LTRIM(PRODDTA.F0101.ABMCU)) = '" & Trim(Left(List1.Text, 12)) & "' "
       End If
     End If
Next i

SQL02 = ""
For i = 0 To List2.ListCount - 1
    
    If List2.Selected(i) Then
       List2.ListIndex = i
       If SQL02 = "" Then
          SQL02 = "RTRIM(LTRIM(PRODDTA.F4101.IMSRP9)) = '" & Trim(Left(List2.Text, 8)) & "' "
          SQL02 = SQL02 & "AND PRODDTA.F55STAB.ASYEAR = " & cmbYear - 2000 & " "
       Else
          SQL02 = SQL02 & "OR RTRIM(LTRIM(PRODDTA.F4101.IMSRP9)) = '" & Trim(Left(List2.Text, 8)) & "' "
          SQL02 = SQL02 & "AND PRODDTA.F55STAB.ASYEAR = " & cmbYear - 2000 & " "
       End If
     End If
Next i

If SQL01 <> "" And SQL02 <> "" Then
    
    Set tmp_Rset01 = New Recordset
    tmp_Rset01.CursorLocation = adUseClient
    tmp_Rset01.Open "SELECT * From BrandSales", SQLConn, adOpenStatic, adLockOptimistic

    
    strSQL = "SELECT PRODDTA.F4101.IMSRP9, PRODDTA.F55STAB.ASMNTH, "
    strSQL = strSQL & "Sum(ASSOQS)/100000 AS QTY,  Sum(ASAEXP)/100000 AS "
    strSQL = strSQL & "VAL, Sum(ASPQOR)/100000 AS PCS,  Sum(ASSOCN)/100000 "
    strSQL = strSQL & "AS GROSSWT, Sum(ASSOBK)/100000 AS NETWT "
    strSQL = strSQL & "FROM PRODDTA.F55STAB INNER JOIN PRODDTA.F4101 "
    strSQL = strSQL & "ON PRODDTA.F55STAB.ASLITM = PRODDTA.F4101.IMLITM "
    strSQL = strSQL & "INNER JOIN PRODDTA.F0101 ON PRODDTA.F55STAB.ASAN8 "
    strSQL = strSQL & "=  PRODDTA.F0101.ABAN8 WHERE "
    strSQL = strSQL & SQL01
    strSQL = strSQL & "GROUP BY PRODDTA.F4101.IMSRP9, "
    strSQL = strSQL & "PRODDTA.F55STAB.ASYEAR, PRODDTA.F55STAB.ASMNTH HAVING "
    strSQL = strSQL & SQL02
    strSQL = strSQL & "ORDER BY PRODDTA.F4101.IMSRP9, PRODDTA.F55STAB.ASMNTH"
    
    Set tmpRset01 = New Recordset
    tmpRset01.CursorLocation = adUseClient
    tmpRset01.Open strSQL, DbConn, adOpenStatic, adLockOptimistic
          DoEvents
          DoEvents
          DoEvents

    iCntr = 0
    lblCntr = iCntr
    If tmpRset01.RecordCount > 0 Then
       Do
          If br <> tmpRset01.Fields("IMSRP9") Then
          br = tmpRset01.Fields("IMSRP9")
          tmp_Rset01.AddNew
          tmp_Rset01.Fields("BrandCode") = tmpRset01.Fields("IMSRP9")
          tmp_Rset01.Fields("BrandDesc") = GetLineDesc(Trim(tmpRset01.Fields("IMSRP9")))
          End If
          
          Select Case tmpRset01.Fields("ASMNTH")
          Case Is = 1
               tmp_Rset01.Fields("M01") = tmpRset01.Fields("PCS")
          Case Is = 2
               tmp_Rset01.Fields("M02") = tmpRset01.Fields("PCS")
          Case Is = 3
               tmp_Rset01.Fields("M03") = tmpRset01.Fields("PCS")
          Case Is = 4
               tmp_Rset01.Fields("M04") = tmpRset01.Fields("PCS")
          Case Is = 5
               tmp_Rset01.Fields("M05") = tmpRset01.Fields("PCS")
          Case Is = 6
               tmp_Rset01.Fields("M06") = tmpRset01.Fields("PCS")
          Case Is = 7
               tmp_Rset01.Fields("M07") = tmpRset01.Fields("PCS")
          Case Is = 8
               tmp_Rset01.Fields("M08") = tmpRset01.Fields("PCS")
          Case Is = 9
               tmp_Rset01.Fields("M09") = tmpRset01.Fields("PCS")
          Case Is = 10
               tmp_Rset01.Fields("M10") = tmpRset01.Fields("PCS")
          Case Is = 11
               tmp_Rset01.Fields("M11") = tmpRset01.Fields("PCS")
          Case Is = 12
               tmp_Rset01.Fields("M12") = tmpRset01.Fields("PCS")
          End Select
          tmpRset01.MoveNext
          If Not tmpRset01.EOF Then
            If br <> tmpRset01.Fields("IMSRP9") Then
               tmp_Rset01.Update
            End If
          Else
             tmp_Rset01.Update
          End If
          iCntr = iCntr + 1
          lblCntr.Caption = iCntr
          DoEvents
          DoEvents
          DoEvents
       Loop Until tmpRset01.EOF
     End If
       tmpRset01.Close
       tmp_Rset01.Close
       Set tmpRset01 = Nothing
       Set tmp_Rset01 = Nothing
Else
    MsgBox "You have to select atleast one item from each list boxes for report processing", vbInformation, "Information"
End If
Dim RPT As New rptPCS
CRViewer1.ReportSource = RPT
CRViewer1.ViewReport


End Sub

Private Sub Form_Load()
Set tmpRset = New Recordset
tmpRset.CursorLocation = adUseClient
tmpRset.Open "SELECT MCMCU,MCDL01 From PRODDTA.F0006 WHERE MCSTYL = 'IS' ORDER BY MCMCU", DbConn, adOpenStatic, adLockOptimistic

cmbYear.Text = Year(Date)
If tmpRset.RecordCount > 0 Then
   tmpRset.MoveFirst
   Do
     List1.AddItem tmpRset.Fields("MCMCU") & " - " & tmpRset.Fields("MCDL01")
     tmpRset.MoveNext
   Loop Until tmpRset.EOF
End If
tmpRset.Close
Set tmpRset = Nothing

'--
Set tmpRset = New Recordset
tmpRset.CursorLocation = adUseClient
tmpRset.Open "SELECT DRKY,DRDL01 From PRODCTL.F0005 WHERE DRSY = '41' AND DRRT = '09' ORDER BY DRKY", DbConn, adOpenStatic, adLockOptimistic

If tmpRset.RecordCount > 0 Then
   tmpRset.MoveFirst
   Do
     List2.AddItem tmpRset.Fields("DRKY") & " - " & tmpRset.Fields("DRDL01")
     tmpRset.MoveNext
   Loop Until tmpRset.EOF
End If
tmpRset.Close
Set tmpRset = Nothing
Dim RPT As New rptPCS
CRViewer1.ReportSource = RPT
CRViewer1.ViewReport
End Sub
