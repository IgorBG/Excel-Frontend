Attribute VB_Name = "DataModel"
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

Public Function getDBRequest(Key As String, ByRef DataReqCol As Collection) As Object
    Dim DBRequest As CDBDataRequest
If DataReqCol Is Nothing Then Set DataReqCol = New Collection
If ObjectIsInCollection(DataReqCol, Key) Then Set getDBRequest = DataReqCol.Item(Key): Exit Function

Set DBRequest = New CDBDataRequest
With DBRequest
    
    Select Case Key
        Case "TovFltr"
                .QuerryCommand = "SELECT DISTINCT "
                Call .DisplayAtributesAdd("NarTovDate")
                Call .SourceTablesAdd("ZaqvkiTovarene")
                .PreparedConnection = GetNewConnToAccess(Nastr.Item("Datasources").Item("DB_Expedition").Val, False)

        Case "ZaqvkiTovarene"
                .QuerryCommand = "SELECT "
                Call .DisplayAtributesAdd("0,Null, OrdNum, SpedData, NomIme, Null, Null, Sdelka, Null, BrZaqveno, Null, Null, CenaObsht, Sklad, Null,null,null,null,null,null, NarTovNom & '/' & NarTovDate,null,NomNom,null,null,null,null,null, Klient, Grad, DostObekt, DostGrad")
                Call .SourceTablesAdd("ZaqvkiTovarene")
                .PreparedConnection = GetNewConnToAccess(Nastr.Item("Datasources").Item("DB_Expedition").Val, False)
    End Select
End With
DataReqCol.Add DBRequest, Key
Set getDBRequest = DBRequest
Exit Function
End Function


Public Function GetPresetCollection(Context As String, Optional ForceReset As Boolean = False, Optional AddToDS As Boolean = True) As Collection
    Dim itm As Collection
    Dim MarkNastr As Collection
    Dim LocalData As Variant
    Dim conn As ADODB.Connection
    Dim i As Long
    Dim Parameters As Variant
    Dim queryName As String, queryType As Variant, KeyString As String
    Dim WS As Worksheet
    Dim StartRow As Long, LastCol As Long, LastRow As Long
On Error GoTo ErrHandler
If Not IsInitialized Then Call Inicial_Main

'First check whether the datasource is already in cache
If Not ForceReset Then
    If ObjectIsInCollection(DS, Context) Then
        Set GetPresetCollection = DS(Context)
        Exit Function
    End If
End If
'====== Settings related to the context ======
Select Case Context
    Case "DBSavedStops"
        Set conn = GetNewConnToAccess(Nastr("Datasources").Item("DB_Expedition").Val)
        Set GetPresetCollection = New Collection                                                                      'Parameters = ([@SpeditionDate] [@RouteListType] [@RouteListID])
        Parameters = Array(CDbl(CDate(CheckedValue("date", False, DS("TempValues").Item("SpeditionDate").Val))), 1, 0)
        queryName = "selectStopsForSpecificDate"
        queryType = adCmdStoredProc
    Case "WS_Order_Clients"
        Set MarkNastr = Nastr("ERPMark")
'    Case "IgnoreList"  'just for example
'        Set GetPresetCollection = New Collection
'        Set WS = ExSheet
'        StartRow = 2
'        LastCol = 1
'        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
End Select
'=============================================
Select Case Context
    Case "DBSavedStops"
        'Scanning DB for saved unloading scheme
        LocalData = GetRSData(conn, queryName, True, Parameters, queryType)
        For i = LBound(LocalData, 2) To UBound(LocalData, 2)                                    'Mapping: ObektID=0, StopType=1, RouteNum=2, StopNum=3, ClientName=4, ClientCity=5, UnloadName=6, UnloadCity=7, DateDeliveryPlan=8, TimeDeliveryPlan=9
                                                                                            'Keystring = ClientName / ClientCity || UnloadName / UnloadCity
            KeyString = CStr(LocalData(4, i)) & " / " & CStr(LocalData(5, i)) & " || " & CStr(LocalData(6, i)) & " / " & CStr(LocalData(7, i))
            If Not ObjectIsInCollection(GetPresetCollection, KeyString) Then
                Set itm = New Collection
                    Call itm.Add(getNewBean("RouteNum", CStr(LocalData(2, i))), "RouteNum")
                    Call itm.Add(getNewBean("StopNum", CStr(LocalData(3, i))), "StopNum")
                    Call itm.Add(getNewBean("RouteNum", CStr(LocalData(4, i))), "ClientName")
                    Call itm.Add(getNewBean("StopNum", CStr(LocalData(5, i))), "ClientCity")
                    Call itm.Add(getNewBean("UnloadPlace", CStr(LocalData(6, i))), "UnloadPlace")
                    Call itm.Add(getNewBean("UnloadCity", CStr(LocalData(7, i))), "UnloadCity")
                    Call itm.Add(getNewBean("DateDelivery", CStr(LocalData(8, i))), "DateDelivery")
                    Call itm.Add(getNewBean("TimeDelivery", CStr(LocalData(9, i))), "TimeDelivery")
                GetPresetCollection.Add itm, KeyString
                Set itm = Nothing
            End If
        Next i
    Case "WS_Order_Clients"
        LocalData = getPresetData("OrderList")
            For i = LBound(LocalData) To UBound(LocalData)
                If Not Len(LocalData(i, MarkNastr("KlientCol").Val)) = 0 Then
                    KeyString = CStr(LocalData(i, MarkNastr("KlientCol").Val)) & " / " & CStr(LocalData(i, MarkNastr("GradCol").Val))
                    If Not ObjectIsInCollection(GetPresetCollection, KeyString) Then
                        Set itm = New Collection
                            Call itm.Add(getNewBean("ClientName", CStr(LocalData(i, MarkNastr("KlientCol").Val))), "ClientName")
                            Call itm.Add(getNewBean("ClientCity", CStr(LocalData(i, MarkNastr("GradCol").Val))), "ClientCity")
                            Call itm.Add(getNewBean("ClntNum", CStr(GetPresetCollection.Count + 1)), "ClntNum")
                        GetPresetCollection.Add itm, KeyString
                        Set itm = Nothing
                    End If
                End If
            Next i
'    Case "IgnoreList"
'        LocalData = WS.Range(WS.Cells(StartRow, 1), WS.Cells(LastRow, LastCol))
'        For i = LBound(LocalData) To UBound(LocalData)
'            KeyString = CStr(LocalData(i, 1))
'            If Not ValueIsInCollection(GetPresetCollection, KeyString) Then
'                GetPresetCollection.Add KeyString, KeyString
'            End If
'        Next i
    Case "DBSavedStops_FilteredUnloadPlaces"
        Set GetPresetCollection = GetFilteredCollection(GetPresetCollection("DBSavedStops"), Context)    'Recursive Filtering a new collection: Unloading places only
    Case "DBSavedStops_FilteredClients"
        Set GetPresetCollection = GetFilteredCollection(GetPresetCollection("DBSavedStops"), Context)    'Recursive Filtering a new collection: Clients only
    Case "WS_Orders_FilteredClients"
        Set GetPresetCollection = GetFilteredCollection(getPresetData("OrderList"), Context)

End Select

If AddToDS Then Call AddDatasource(DS, Context, GetPresetCollection, ForceReset)
Exit Function

ErrHandler:
Call EmergencyExit("Function GetPresetCollection, Context:" & Context)
End Function

Private Function GetFilteredCollection(ByVal inData As Variant, Context As String) As Collection
    Dim itm As Object, i As Long
    Dim KeyString As String
    Set GetFilteredCollection = New Collection
    Select Case Context
        Case vbNullString
            Set GetFilteredCollection = inData: Exit Function
        Case "DBSavedStops_FilteredUnloadPlaces"
            For Each itm In inData
                KeyString = itm("UnloadPlace").Val & " / " & itm("UnloadCity").Val
                If Not ObjectIsInCollection(GetFilteredCollection, KeyString) Then GetFilteredCollection.Add itm, KeyString
            Next itm
        Case "DBSavedStops_FilteredClients"
            For Each itm In inData
                KeyString = itm("ClientName").Val & " / " & itm("ClientCity").Val
                If Not ObjectIsInCollection(GetFilteredCollection, KeyString) Then GetFilteredCollection.Add itm, KeyString
            Next itm
        Case "WS_Orders_FilteredClients"
           
        
    
    End Select

End Function


Public Function getPresetData(Context As String, Optional ForceReset As Boolean = False, Optional AddToDS As Boolean = True, Optional ByVal InParameters as Variant) As Variant
    Dim WS As Worksheet
    Dim StartRow As Long, LastRow As Long, LastCol As Long
    Dim LocalData As Variant, TempVar1 As Variant
    Dim conn As ADODB.Connection
    Dim i As Long
    Dim Parameters As Variant
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
    Case "SampleDataFromWS"
        Set WS = SampleSheet
        Set SetColl = Nastr("SampleSheetMark")
    Case "ClientData"
        Set WS = AnotherSheet
        Set SetColl = Nastr("AnotherSheetMark")
        StartRow = SetColl("StartRow").Val
        LastCol = SetColl("LastCol").Val
        LastRow = WS.Cells(Rows.Count, SetColl("ObektCol").Val).End(xlUp).Row
    Case "SampleDataFromDBStoredProcedure"
        Set conn = GetNewConnToAccess(Nastr("Datasources").Item("Sample_DB_Path").Val, True)
        queryName = "getSomeDataQuery"
        queryType = adCmdStoredProc
        Parameters = Array(PersonName, PersonCity)
                                                                                                                                            
End Select
'=============================================
Select Case Context
    Case "OrderList"
        If StartRow = 0 Or LastRow = 0 Or LastCol = 0 Then Call MarkupOrderList(StartRow, LastRow, LastCol)
        getPresetData = WS.Range(WS.Cells(StartRow, 1), WS.Cells(LastRow, LastCol))
    Case "SampleDataFromDBStoredProcedure"
        getPresetData = GetRSData(conn, queryName, True, Parameters, queryType)                                                                                                                                            
End Select

If AddToDS Then Call AddDatasource(DS, Context, getPresetData, ForceReset)
Exit Function

ErrHandler:
Call EmergencyExit("Function getPresetData, Context:" & Context)

End Function

Private Sub MarkupOrderList(Optional ByRef StartRowInfo As Long, Optional ByRef LastRowInfo As Long, Optional ByRef LastColInfo As Long)
If Not IsInitialized Then Call Inicial_Main
    Dim i As Long, LastRow As Long, StartRow As Long, LastCol As Long
On Error GoTo Err_Handler
If StartRowInfo = 0 Then
    For i = 1 To 256
        If ERPSheet.Cells(i, 5).Value <> vbNullString Then StartRow = i + 1: Exit For
    Next i
    StartRowInfo = StartRow
End If
    If Not ObjectIsInCollection(DS("Mark"), "OrderStartRow") Then
        DS("Mark").Add getNewBean("OrderStartRow", StartRowInfo), "OrderStartRow"
        Else: DS("Mark").Item("OrderStartRow").Val = StartRowInfo
    End If
If LastRowInfo = 0 Then
    LastRow = ERPSheet.Cells(Rows.Count, 1).End(xlUp).Row
    If LastRow < StartRow Then Call EmergencyExit("Програмата не намира изделия на листа със заявките (От колона A)")
    LastRowInfo = LastRow
End If
    If Not ObjectIsInCollection(DS("Mark"), "OrderLastRow") Then
        DS("Mark").Add getNewBean("OrderLastRow", LastRowInfo), "OrderLastRow"
        Else: DS("Mark").Item("OrderLastRow").Val = LastRowInfo
    End If
If LastColInfo = 0 Then
    LastCol = Nastr.Item("ERPMark").Item("LastCol").Val
    If LastCol <= 1 Then Call EmergencyExit("Програмата не намира правилните настройки за маркировка на листа с поръчки. Невалиден параметър за последната колона")
    LastColInfo = LastCol
End If
    If Not ObjectIsInCollection(DS("Mark"), "OrderLastCol") Then
        DS("Mark").Add getNewBean("OrderLastCol", LastColInfo), "OrderLastCol"
        Else: DS("Mark").Item("OrderLastCol").Val = LastColInfo
    End If

Exit Sub

Err_Handler:
Call EmergencyExit("Програмата не намира правилните настройки за маркировка на листа с поръчки.")
End Sub


