Attribute VB_Name = "DataModel"
Public Sub AddDatasource(Key As String, Datasources As Collection, DSObject As CDataStorage, Optional Replace As Boolean = True)
If Datasources Is Nothing Then Set Datasources = New Collection
If ObjectIsInCollection(Datasources, Key) Then
    If Replace Then
        Datasources.Remove (Key)
        Datasources.Add DSObject, Key
    End If
Else: Datasources.Add DSObject, Key
End If
Exit Sub

End Sub

Public Function getDBRequest(Key As String, DataReqCol As Collection) As Object
    Dim DBRequest As CDBDataRequest
If DataReqCol Is Nothing Then Set DataReqCol = New Collection
If ObjectIsInCollection(DataReqCol, Key) Then
    Set getDBRequest = DataReqCol.Item(Key): Exit Function
End If
    Select Case Key
        Case "ImagesFromDB"
            Set DBRequest = New CDBDataRequest
            With DBRequest
                .QuerryCommand = "SELECT"
                Call .DisplayAtributesAdd("FileName, FilePath")
                Call .SourceTablesAdd("Images")
                Call .SortingAtributesAdd("FileName")
                .PreparedConnection = GetNewConnToAccess(PaketiDBPath, False)
            End With
    
    End Select
DataReqCol.Add DBRequest, Key
Set getDBRequest = DBRequest
Exit Function
End Function


Public Sub RefreshDatasourceView(Context, Datasource As CDataStorage, ByRef TekForm As Object, Optional Limit As Long)
    Dim i As Long
    Dim LocalData As Variant

LocalData = Datasource.getContent
Select Case Context
    Case "ImagesFromDB"
        TekForm.ListSearchResult.Clear
        If IsArray(LocalData) Then
            If Limit = 0 Then Limit = UBound(LocalData, 2)
            For i = LBound(LocalData, 2) To Limit
                TekForm.ListSearchResult.AddItem CStr(LocalData(Datasource.getClmnIndx("FileName"), i)), i
            Next i
        End If
End Select
End Sub
