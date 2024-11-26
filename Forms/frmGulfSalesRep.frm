VERSION 5.00
Begin VB.Form frmGulfSalesRep 
   Caption         =   "Sales Report - Export"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   12105
   Begin VB.TextBox txtMonth 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   3855
   End
End
Attribute VB_Name = "frmGulfSalesRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcess_Click()
Dim Cur_Year As Integer, Prv_Year As Integer, rpt_Rset As Recordset
Dim RepTemplate As Recordset, Drep As Recordset, Fill_PCS As Boolean
Dim MthNo As Integer, iLOOP As Integer, p_Mth As Integer
Dim CtryArray(5) As String
Dim LineAcc As String

SQLConn.Execute "DELETE FROM GULFREP01"

Set RepTemplate = New Recordset
RepTemplate.CursorLocation = adUseClient
RepTemplate.Open "SELECT * From GULFREP01", SQLConn, adOpenStatic, adLockOptimistic

Set Drep = New Recordset
Drep.CursorLocation = adUseClient
Drep.Open "SELECT * From LineTable", SQLConn, adOpenStatic, adLockOptimistic


    Cur_Year = (Year(Date) - 2000)
    Prv_Year = Cur_Year - 1
    MthNo = txtMonth
     
    If MthNo > 1 Then
       p_Mth = MthNo - 1
    Else
       p_Mth = 12
    End If
    
    If p_Mth < 12 Then
       Prv_Year = Cur_Year
    End If
    
    CtryArray(0) = "1001"
    CtryArray(1) = "1002"
    CtryArray(2) = "1003"
    CtryArray(3) = "1004"
    CtryArray(4) = "1005"
    
    If Drep.RecordCount > 0 Then
    For iLOOP = 0 To 4
    Drep.MoveFirst
    Do
         RepTemplate.AddNew
         RepTemplate.Fields("Line") = Trim(Drep.Fields("Line"))
         RepTemplate.Fields("LineDesc") = Drep.Fields("LineDesc")
         RepTemplate.Fields("Country") = CtryArray(iLOOP)
         
         If Trim(Drep.Fields("Line")) = "F201" Or Trim(Drep.Fields("Line")) = "F202" Or _
            Trim(Drep.Fields("Line")) = "F203" Or Trim(Drep.Fields("Line")) = "F204" Then
            Fill_PCS = True
         Else
            Fill_PCS = False
         End If
         
         GetYTDSalesBUC Drep.Fields("Line"), Cur_Year, txtMonth, CtryArray(iLOOP)
         GetYTDSalesBUP Drep.Fields("Line"), Prv_Year, txtMonth, CtryArray(iLOOP)
         
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
         
         
         GetMTDSalesBUC Drep.Fields("Line"), Cur_Year, MthNo, CtryArray(iLOOP)
         GetMTDSalesBUP Drep.Fields("Line"), Prv_Year, p_Mth, CtryArray(iLOOP)
         
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
         
         LineAcc = GetAccount(Trim(Drep.Fields("Line")), "SO", 4230)
         
         RepTemplate.Fields("ValFinPM") = AreaNetSales(CtryArray(iLOOP), Trim(Drep.Fields("Line")), Prv_Year, p_Mth)
         RepTemplate.Fields("ValFinCM") = AreaNetSales(CtryArray(iLOOP), Trim(Drep.Fields("Line")), Cur_Year, MthNo)
          
         RepTemplate.Update
         Drep.MoveNext
    Loop Until Drep.EOF
    Next iLOOP
    End If
    Drep.Close
    RepTemplate.Close
    
    Set Drep = Nothing
    Set RepTemplate = Nothing
End Sub



