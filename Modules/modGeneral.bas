Attribute VB_Name = "modGeneral"
Public strSQL As String
Public DbConn As Connection, SQLConn As Connection
Public tmpRset As Recordset
Public Function DbConnection()
   If Not ReadIniFile("Settings.ini") Then
      End
   End If
Dim strConn As String

Set DbConn = New Connection
strConn = "uid=" & User_DB2 & ";pwd=" & PWD_DB2 & ";Provider=" & Provider_DB2 & ";Persist Security Info=False;Initial Catalog=" & Pub_DataLib & ";Data Source=" & Server_DB2
'strConn = "Provider=" & Provider_DB2 & ";Server=" & Server_DB2 & ";Database=" & pub_datalib & ";Uid=" & User_DB2 & ";Pwd=" & PWD_DB2 & ";"
'strConn = "Driver={SQL Server Native Client}"
'strConn = "UID=PSFT;PWD=PSFT;DSN=JDE"
'DbConn.ConnectionString = "Provider=IBMDA400;Data source=172.16.8.40;User Id=SATISHM;Password=KANCHANA"
DbConn.Open strConn
DbConn.CursorLocation = adUseClient
End Function
Public Function SQLConnection()
Dim strConn As String
Set SQLConn = New Connection

    strConn = "uid=" & User_SQL & ";pwd=" & PWD_SQL & ";Provider=" & Provider_SQL & ";Persist Security Info=False;Initial Catalog=" & DB_SQL & ";Data Source=" & Server_SQL
    SQLConn.Open strConn
End Function
Function JulianDate(DateParam As Date) As Variant
Dim varDate As String
varDate = Format(DateParam, "yyyymmdd")
JulianDate = _
        (Val(Left(varDate, 4)) - 2000 + 100) * 1000 + _
        CDate(Mid(varDate, 5, 2) & _
        "/" & Right(varDate, 2) & "/" & Left(varDate, 4)) - _
        CDate("1/1/" & Left(varDate, 4)) + 1
End Function

Function LastDayInMonth(YearValue As Long, MonthValue As Long) As Long
    LastDayInMonth = Day(DateSerial(YearValue, MonthValue + 1, 0))
End Function
Public Function GetLineDesc(LineCd As String) As String

Set tmpRset = New Recordset
tmpRset.CursorLocation = adUseClient
tmpRset.Open "SELECT DRDL01 From PRODCTL.F0005 WHERE DRSY = '41' AND DRRT = '09' AND RTRIM(LTRIM(DRKY)) ='" & Trim(LineCd) & "'", DbConn, adOpenStatic, adLockOptimistic
GetLineDesc = ""
If tmpRset.RecordCount > 0 Then
   tmpRset.MoveFirst
   GetLineDesc = tmpRset(0)
End If
tmpRset.Close
Set tmpRset = Nothing
End Function
Public Function NullToZero(GetVal As Variant) As Double

If IsNull(GetVal) Then
   NullToZero = 0
Else
   NullToZero = GetVal
End If
End Function
