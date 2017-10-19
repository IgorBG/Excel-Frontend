VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDBDataRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' For now Access 2003 db engine only, but in case of other DB Engines there should be added Select Case


Private pConn As Object                   'DB connection object
Private pCommand As String              'Example: SELECT
Private pAtribNames As Collection       'Example: FileName, FilePath
Private pTables As Collection           'Example: Images (From clause)
Private pFilters As Collection          'Example: WHERE FileName = "cat.jpg"
Private pFltrUnion As String            'Example: AND
Private pSortBy As Collection           'Example: FileName

Private Const pDEF_COMMAND As String = "SELECT"
Private Const pDEF_FILTR_UNION As String = "AND"



Public Property Let FilterUnion(inStrng As String)
    pFltrUnion = inStrng
End Property
Public Property Let QuerryCommand(inStrng As String)
    pCommand = inStrng
End Property
Public Property Let PreparedConnection(inConn As Object)
    Set pConn = inConn
End Property

' Work with sorting
    Public Sub SortingAtributesAdd(AtributeName As String)
        Call AddAtributes(pSortBy, AtributeName)
    End Sub
    Public Sub SortingAtributesReset()
        Set pSortBy = New Collection
    End Sub
' Work with displaying atributes
    Public Sub DisplayAtributesAdd(AtributeName As String)
        Call AddAtributes(pAtribNames, AtributeName)
    End Sub
    Public Sub DisplayAtributesReset()
        Set pAtribNames = New Collection
    End Sub
' Work with included tables
    Public Sub SourceTablesAdd(TableName As String)
        Call AddAtributes(pTables, TableName)
    End Sub
    Public Sub SourceTablesReset()
        Set pTables = New Collection
    End Sub


Public Sub ExecuteDataRequest(Optional ByRef DS As CDataStorage) 'compiles new datasource from the avilable settings in the DataRequest object
    Dim SQLstr As String
    Dim AtrbtString As String
AtrbtString = CollectionToString(pAtribNames)
SQLstr = pCommand & " " & AtrbtString & " " & "FROM" & " " & CollectionToString(pTables)
If Not pFilters.Count = 0 Then SQLstr = SQLstr & " WHERE " & FiltersToString
If Not pSortBy.Count = 0 Then SQLstr = SQLstr & " ORDER BY " & CollectionToString(pSortBy)
    If DS Is Nothing Then Set DS = New CDataStorage
    pConn.Open
    DS.letContent = GetRSData(pConn, SQLstr, False)
    pConn.Close
    DS.letClmnNames = AtrbtString
    Exit Sub
ErrHandler:
End Sub

Public Property Get getDatasource() As CDataStorage
    Dim TekDS As New CDataStorage
    Call Me.ExecuteDataRequest(TekDS)
    Set getDatasource = TekDS
End Property


Private Sub Class_Initialize()
    If pFilters Is Nothing Then Set pFilters = New Collection
    If pSortBy Is Nothing Then Set pSortBy = New Collection
    If pTables Is Nothing Then Set pTables = New Collection
    If pAtribNames Is Nothing Then Set pAtribNames = New Collection
    If pFltrUnion = vbNullString Then pFltrUnion = pDEF_FILTR_UNION
    If pCommand = vbNullString Then pCommand = pDEF_COMMAND
End Sub

' Working with filters
    Private Sub FilterAddAsObject(ByRef inFilter As CFilter)
    If pFilters Is Nothing Then Set pFilters = New Collection
        pFilters.Add inFilter
    End Sub
    Public Sub FilterReset()
        Set pFilters = New Collection
    End Sub
    Public Sub FilterAddByValues(AtributeName As String, CompareMode As String, CondValue As Variant, ConvertToType As String)
        Dim oFilter As New CFilter
    On Error GoTo ErrHandler
            oFilter.AtributeName = AtributeName
            oFilter.CompareMode = CompareMode
            oFilter.AddCondContent CondValue
            oFilter.ConvertToType = ConvertToType
            Call FilterAddAsObject(oFilter)
        Set oFilter = Nothing
    Exit Sub
ErrHandler:
    EmergencyExit ("Неуспешен опит за създаване на филтър. Процедура AddFilterSingleCondValue. Модул CDataStorage")
    End Sub
        Private Function FiltersToString() As String
            Dim F As CFilter
        On Error GoTo ErrHandler
        If pFilters.Count = 0 Then Exit Function
        For Each F In pFilters
            FiltersToString = FiltersToString & F.toString & pFltrUnion
        Next F
        FiltersToString = Left(FiltersToString, Len(FiltersToString) - Len(pFltrUnion))
        Exit Function
ErrHandler: EmergencyExit ("Неуспешен опит за прилагане на филтрите към запитването. Функция FiltersToString. Модул CDataStorage")
        End Function

Private Function CollectionToString(InColl As Collection, Optional Delim As String = ",", Optional BeginStr As String, Optional EndStr As String) As String
    Dim v As Variant
    For Each v In InColl
        CollectionToString = CollectionToString & CStr(v) & Delim
    Next v
If Len(CollectionToString) = 0 Then Exit Function
CollectionToString = Left(CollectionToString, Len(CollectionToString) - Len(Delim))
CollectionToString = BeginStr & CollectionToString & EndStr
End Function

Private Sub AddAtributes(ByRef Coll As Collection, ByRef inValue As String, Optional ByVal Delim As String = ",")
    Dim v As Variant, i As Long
    v = Split(inValue, Delim)
    For i = LBound(v) To UBound(v)
        Coll.Add (v(i))
    Next i
End Sub



