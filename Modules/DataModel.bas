Option Explicit
Public Sub AddDatasource(ByRef Datasources As Collection, Key As String, DataIn As Variant, Optional Overwrite As Boolean = True)
If Datasources Is Nothing Then Set Datasources = New Collection
    If ObjectIsInCollection(Datasources, Key) Then
        If Overwrite Then Datasources.Remove (Key) Else: Exit Sub
    End If
    If IsObject(DataIn) Then
        Datasources.Add DataIn, Key
    Else: Datasources.Add getNewBean(Key, DataIn), Key
    End If
Exit Sub
End Sub



Public Function GetPresetCollection(Context As String, Optional ForceReset As Boolean = False, Optional AddToDS As Boolean = True) As Collection
    Dim Itm As Collection, TempCol As Collection, TempItm As Collection
    Dim MarkNastr As Collection
    Dim LocalData As Variant
    Dim conn As ADODB.Connection
    Dim i As Long
    Dim parameters As Variant
    Dim queryName As String, queryType As Variant, KeyString As String
    Dim WS As Worksheet
    Dim StartRow As Long, LastCol As Long, LastRow As Long
    Dim TempBool As Boolean
On Error GoTo ErrHandler
If Not IsInitialized Then Call Inicial_Main

'First check whether the datasource is already in cache
If Not ForceReset Then
    If ObjectIsInCollection(DS, Context) Then
        Set GetPresetCollection = DS(Context)
        Exit Function
    End If
End If
Set GetPresetCollection = New Collection
'====== Settings related to the context ======
Select Case Context

End Select
'=============================================
Select Case Context
    Case "SomeDataInCollection"
        LocalData = getPresetData("SomeDBdata")
        For i = LBound(LocalData, 2) To UBound(LocalData, 2)
            If Not ValueIsInCollection(GetPresetCollection, CStr(LocalData(0, i))) Then GetPresetCollection.Add CStr(LocalData(0, i)), CStr(LocalData(0, i))
        Next i
    
    
    Case Else
        Call EmergencyExit("Неразпознат контекст '" & Context & "' в процедура GetPresetCollection")
End Select

If AddToDS Then Call AddDatasource(DS, Context, GetPresetCollection, ForceReset)
Exit Function

ErrHandler:
Call EmergencyExit("Function GetPresetCollection, Context:" & Context)
End Function

Private Function GetFilteredCollection(ByVal InData As Variant, Context As String) As Collection
    Dim Itm As Object, i As Long
    Dim KeyString As String
    Set GetFilteredCollection = New Collection
    Select Case Context
        Case vbNullString
            Set GetFilteredCollection = InData: Exit Function
    
    End Select

End Function

Public Function getPresetData(Context As String, Optional ForceReset As Boolean = False, Optional AddToDS As Boolean = True) As Variant
    Dim WS As Worksheet
    Dim StartRow As Long, LastRow As Long, LastCol As Long
    Dim LocalData As Variant, TempVar As Variant
    Dim conn As ADODB.Connection
    Dim i As Long, b As Long
    Dim parameters As Variant
    Dim queryName As String, queryType As Variant, KeyString As String
    Dim SetColl As Collection, TempCol As Collection
    Dim TempStr As String
On Error GoTo ErrHandler
If Not IsInitialized Then Call Inicial_Main

If Not ForceReset Then 'First check whether the datasource is already in cache
    Select Case Context
        Case "EXAMPLE_always_reset"  'Some datasource that always should be reseted
        Case Else                      'datasources that allowed to be taken from cache
            If ObjectIsInCollection(DS, Context) Then getPresetData = DS(Context).Val: Exit Function
    End Select
End If
'====== Settings related to the context ======
Select Case Context
    Case "SampleWSData", "OtherSampleWSData"
        Set WS = ERPSheet
        Set SetColl = Nastr("ERPMark")
    Case "SomeDBdata"
        Set conn = GetNewConnToAccess(Nastr("Datasources").Item("DB_Expedition").Val, True)
        queryName = "sample_db_stored_procedure_name"
        queryType = adCmdStoredProc
        parameters = Empty
    
        
End Select


'====== Instructions for retrieving certain data =========
Select Case Context
    Case "SomeDBdata"
        getPresetData = GetRSData(conn, queryName, True, parameters, queryType)
    Case "SomeSpecificDBdata"
        LocalData = getPresetData("SomeDBdata", ForceReset)
            If IsEmpty(LocalData) Then Call EmergencyExit("Не откривам качени заявки за НТ номер / дата " & DS("NTNum").Val & "/" & DS("NTDate").Val)
            ReDim TempVar(1 To SetColl("LastCol").Val, 0 To 0)
            b = LBound(LocalData, 1)
            For i = LBound(LocalData, 2) To UBound(LocalData, 2)
                TempVar(SetColl("SampleCol").Val, UBound(TempVar, 2)) = LocalData(b + 6, i) & "/" & LocalData(b + 7, i)
                TempVar(SetColl("SomeOtherColumn").Val, UBound(TempVar, 2)) = vbNullString
                ReDim Preserve TempVar(1 To SetColl("LastCol").Val, 0 To UBound(TempVar, 2) + 1)
            Next i
            ReDim Preserve TempVar(1 To SetColl("LastCol").Val, 0 To UBound(TempVar, 2) - 1)
            getPresetData = TransposeArray(TempVar, , -1)

    Case Else
         Call EmergencyExit("Неразпознат контекст '" & Context & "' в процедура getPresetData")
End Select

If AddToDS Then Call AddDatasource(DS, Context, getPresetData, ForceReset)
Exit Function

ErrHandler:
Call EmergencyExit("Function getPresetData, Context:" & Context)
End Function



    Private Sub testgetPresetData()
        Dim t As Variant
        t = getPresetData("SomeDBdata")
        Stop
    End Sub
    Private Sub testGetPresetCollection()
        Dim t As Collection
        Set t = GetPresetCollection("SomeDataInCollection")
        Stop
    End Sub
