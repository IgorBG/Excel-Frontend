Attribute VB_Name = "DBconn"
Option Explicit
Public Function GetNewConnToAccess(ByVal FullPathToMDBfile As String, Optional getOpened As Boolean = True) As ADODB.Connection
    Dim conn As ADODB.Connection
Call SetNewConnToAccess(conn, FullPathToMDBfile, getOpened)
Set GetNewConnToAccess = conn
Set conn = Nothing
End Function

Sub SetNewConnToAccess(ByRef conn As ADODB.Connection, ByVal FullPathToMDBfile As String, Optional getOpened As Boolean = True)
    Dim ErrMsg As String
    Dim PP As Long
    Dim RS As Variant
PP = 100
On Error GoTo Err_Handler
If Not conn Is Nothing Then GoTo Err_Handler
   Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.Jet.OLEDB.4.0;"
    conn.ConnectionString = "Data Source=" & FullPathToMDBfile
PP = 200
    If getOpened Then conn.Open
'Testing the connection
PP = 250
'    RS = GetRSData(conn, "SELECT 1")
Exit Sub

Err_Handler:
Select Case PP
    Case 100: ErrMsg = "Връзката към базата данни не е била затворена през предишните операции. Рестартирайте програмата"
    Case 200: ErrMsg = "Програмата не може да инициира връзката с базата данни по адрес: " & FullPathToMDBfile & ". Вероятно адресът, на който се намира БД трябва да се поднови в лист настройки."
    Case 250: ErrMsg = "Програмата се свързва с базата данни по адрес: " & FullPathToMDBfile & ". Но пробната заявка се връща с грешка."
    Case Else: ErrMsg = "Грешка при опит за свързване с базата данни по адрес: " & FullPathToMDBfile
End Select
Call EmergencyExit(ErrMsg)
End Sub
Sub CloseTheConnToAccess(ByRef conn As ADODB.Connection)
If conn Is Nothing Then
    Debug.Print "Връзката не е била отворена, че да се затвори"
    Else: Set conn = Nothing
End If
End Sub

Public Function getConnectionFromPull(Path As String) As Object
    Dim Tempconn As ADODB.Connection
    On Error GoTo ErrHandler
        If Not ObjectIsInCollection(DC, Path) Then
            Set Tempconn = GetNewConnToAccess(Path, False)
            DC.Add Tempconn, Path
        End If
    Set getConnectionFromPull = DC.Item(Path)
Exit Function

ErrHandler:
Call EmergencyExit("Проблем в модул getConnectionFromPull")
End Function

Public Function GetFirstRecordFromRSData(ByRef conn As ADODB.Connection, ByVal sqlCmd As String, Optional ConnCloseAfter As Boolean) As Variant
Dim r As Variant
    r = GetRSData(conn, sqlCmd, ConnCloseAfter)
    If IsArray(r) Then GetFirstRecordFromRSData = r(LBound(r, 1), LBound(r, 2))
End Function

Public Function GetRSData(ByRef conn As ADODB.Connection, ByVal sqlCmd As String, Optional ConnCloseAfter As Boolean, Optional Parameters As Variant = "", Optional cmdType = adCmdText) As Variant
Dim RS As ADODB.Recordset
Dim Cmd As ADODB.Command
On Error GoTo ErrHandler
    If conn.State = 0 Then conn.Open
    Set Cmd = GetNewADODBCommand(conn, sqlCmd, Parameters, cmdType)
        Set RS = Cmd.Execute
            If Not RS.EOF And Not RS.BOF Then
                GetRSData = RS.GetRows
            Else:
                GetRSData = Empty
            End If
        RS.Close
    Set RS = Nothing
    If ConnCloseAfter Then Call CloseTheConnToAccess(conn)
Exit Function
ErrHandler:
    Debug.Print err.Number
    GetRSData = Empty
    Set RS = Nothing
End Function

Public Sub ExecuteStoredProcedure(ByVal conn As ADODB.Connection, ByVal StorProcName As String, Optional ByVal Parameters As Variant)
    Dim TemCmd As New ADODB.Command
On Error GoTo ErrHandler
Set TemCmd = GetNewADODBCommand(conn, StorProcName, Parameters, adCmdStoredProc)
Call TemCmd.Execute
Exit Sub

ErrHandler:
Call EmergencyExit("Модул ExecuteStoredProcedure")
End Sub
    

    
    Private Function GetNewADODBCommand(ByRef conn As ADODB.Connection, ByVal CommandText As String, Optional ByVal Parameters As Variant, Optional cmdType = adCmdText) As ADODB.Command
    Dim i As Long
    Dim TekCmnd As New ADODB.Command
    Dim ParametersArray As Variant
        Set TekCmnd.ActiveConnection = conn
        TekCmnd.CommandText = CommandText
        TekCmnd.CommandType = cmdType
        TekCmnd.CommandTimeout = 15
    
    If IsMissing(Parameters) Then Parameters = vbNullString
  'ADO does not correctly retrieve named parameters, so the names are ignored. Thus, the order of parameters in the ParametersArray must be the same ас in the stored procedure!'ADO does not correctly retrieve named parameters, so the names are ignored. Thus, the order of parameters in the ParametersArray must be the same ас in the stored procedure!

        ParametersArray = GetParametersAsArray(Parameters)
        
        For i = LBound(ParametersArray) To UBound(ParametersArray)
            Select Case ParametersArray(i)
            Case vbNullString
                Call TekCmnd.Parameters.Append(CreateCommandParameter("none", Null))
            Case Else
                Call TekCmnd.Parameters.Append(CreateCommandParameter(ParametersArray(i), ParametersArray(i)))
            End Select
        Next i
        
        Set GetNewADODBCommand = TekCmnd
        Exit Function
ErrHandler:
        Call EmergencyExit("Function GetNewADODBCommand")
    End Function
    
    Private Function CreateCommandParameter(ByVal name As String, ByVal value As Variant, Optional ByVal numPrecision As Integer = 4, Optional ByVal numScale As Integer = 4) As ADODB.Parameter
    ' based on https://codereview.stackexchange.com/questions/144063/passing-multiple-parameters-to-an-sql-Query
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
    
            Case VbVarType.vbString, VbVarType.vbNull
                result.Type = adVarChar
                If VarType(value) = vbNull Then result.Size = 1 Else result.Size = Len(value)
                If value = vbNullString Then result.Size = 1
        
            Case Else
                err.Raise 5, Description:="Data type not supported"
        End Select

        Set CreateCommandParameter = result
    End Function

Private Function GetParametersAsArray(ByRef Parameters As Variant) As Variant
On Error GoTo ErrHandler
If IsArray(Parameters) Then GetParametersAsArray = Parameters: Exit Function
Select Case VarType(Parameters)
    Case VbVarType.vbArray
        GetParametersAsArray = Parameters
    Case VbVarType.vbObject
        GetParametersAsArray = getParamArrayFromColl(Parameters)
    Case Else
        GetParametersAsArray = Array(Parameters)
End Select
Exit Function

ErrHandler:
        Call EmergencyExit("Function GetParametersAsArray")
End Function
Private Function getParamArrayFromColl(ByVal inColl As Collection) As Variant
    Dim v As Variant
    Dim result As String
On Error GoTo ErrHandler
    For Each v In inColl
        result = result & v & ";"
    Next v
    If Len(result) > 0 Then result = Left(result, Len(result) - 1)
    getParamArrayFromColl = Split(result, ";")
Exit Function
ErrHandler:
    Call EmergencyExit("Функция getParamArrayFromColl")
End Function
Public Function AdaptedQuery(ByVal SQLQuery As String, Optional ByVal Criteria As String = vbNullString) As String
AdaptedQuery = Replace(SQLQuery, "@Criteria", "'%" & Criteria & "%'")
End Function




