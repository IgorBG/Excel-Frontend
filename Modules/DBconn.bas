Attribute VB_Name = "DBconn"
Option Explicit
Public Function GetNewConnToAccess(ByVal FullPathToMDBfile As String, Optional getOpened As Boolean = True) As ADODB.Connection
    Dim Conn As ADODB.Connection
Call SetNewConnToAccess(Conn, FullPathToMDBfile, getOpened)
Set GetNewConnToAccess = Conn
Set Conn = Nothing
End Function

Sub SetNewConnToAccess(ByRef Conn As ADODB.Connection, ByVal FullPathToMDBfile As String, Optional getOpened As Boolean = True)
    Dim ErrMsg As String
    Dim breakPoint As Long
    Dim RS As Variant
breakPoint = 100
On Error GoTo Err_Handler
If Not Conn Is Nothing Then GoTo Err_Handler
   Set Conn = New ADODB.Connection
    Conn.Provider = "Microsoft.Jet.OLEDB.4.0;"
    Conn.ConnectionString = "Data Source=" & FullPathToMDBfile
breakPoint = 200
    If getOpened Then Conn.Open
'Тестване на връзката
'    RS = GetRSData(conn, "SELECT 1")
Exit Sub

Err_Handler:
Select Case breakPoint
    Case 100: ErrMsg = "Връзката към базата данни не е била затворена през предишните операции. Рестартирайте програмата"
    Case 200: ErrMsg = "Програмата не открива връзката с базата данни по адрес: " & FullPathToMDBfile & ". Вероятно адресът, на който се намира БД трябва да се поднови в лист настройки."
    Case Else:: ErrMsg = "Грешка при опит за свързване с базата данни на адрес:: " & FullPathToMDBfile & "."
End Select
Call EmergencyExit(ErrMsg)
End Sub
Sub CloseTheConnToAccess(ByRef Conn As ADODB.Connection)
If Conn Is Nothing Then
    Debug.Print "Връзката не е била отворена, че да се затвори"
    Else: Set Conn = Nothing
End If
End Sub

Public Function GetFirstRecordFromRSData(ByRef Conn As ADODB.Connection, ByVal sqlCmd As String, Optional ConnCloseAfter As Boolean) As Variant
Dim r As Variant
    r = GetRSData(Conn, sqlCmd, ConnCloseAfter)
    If IsArray(r) Then GetFirstRecordFromRSData = r(LBound(r, 1), LBound(r, 2))
End Function

Public Function GetRSData(ByRef Conn As ADODB.Connection, ByVal sqlCmd As String, Optional ConnCloseAfter As Boolean) As Variant
Dim RS As ADODB.Recordset
'On Error GoTo ErrHandler
    Set RS = New ADODB.Recordset
        RS.ActiveConnection = Conn
        RS.Open sqlCmd
            If Not RS.EOF And Not RS.BOF Then
                GetRSData = RS.GetRows
            Else:
                GetRSData = Empty
            End If
        RS.Close
    Set RS = Nothing
    If ConnCloseAfter Then Call CloseTheConnToAccess(Conn)
Exit Function
ErrHandler:
    Debug.Print err.Number
    GetRSData = Empty
    Set RS = Nothing
End Function

Public Function AdaptedQuerry(ByVal SQLQuerry As String, Optional ByVal Criteria As String = vbNullString) As String
AdaptedQuerry = Replace(SQLQuerry, "@Criteria", "'%" & Criteria & "%'")
End Function

Public Function CreateCommandParameter(ByVal name As String, ByVal value As Variant, Optional ByVal numPrecision As Integer = 4, Optional ByVal numScale As Integer = 4) As ADODB.Parameter
' copied from https://codereview.stackexchange.com/questions/144063/passing-multiple-parameters-to-an-sql-query
    Dim result As New ADODB.Parameter
    result.Direction = adParamInput
    result.name = name
    result.value = value

    Select Case VarType(value)
        Case VbVarType.vbBoolean
            result.Type = adBoolean

        Case VbVarType.vbDate
            result.Type = adDate

        Case VbVarType.vbCurrency
            result.Type = adCurrency
            result.Precision = numPrecision
            result.NumericScale = numScale

        Case VbVarType.vbDouble
            result.Type = adDouble
            result.Precision = numPrecision
            result.NumericScale = numScale

        Case VbVarType.vbSingle
            result.Type = adSingle
            result.Precision = numPrecision
            result.NumericScale = numScale

        Case VbVarType.vbByte, VbVarType.vbInteger, VbVarType.vbLong
            result.Type = adInteger

        Case VbVarType.vbString
            result.Type = adVarChar

        Case Else
            err.Raise 5, Description:="Data type not supported"
    End Select

    Set CreateCommandParameter = result
End Function

