Option Explicit

Private pFilterContent As Collection
Private UnionType As String

Private Sub Class_Initialize()
    Set pFilterContent = New Collection
    If UnionType = vbNullString Then UnionType = "AND"
End Sub
Public Sub AddFilterObject(ByVal inValue As CFilter)
If pFilterContent Is Nothing Then Set pFilterContent = New Collection
    pFilterContent.Add inValue
End Sub
Private Sub ClearCondContent()
Set pFilterContent = Nothing
End Sub


Public Property Get toString() As String '(Optional ByVal dbEngine As String) As String
'Converts Collection of Filters to string so it can be injected to sql querry
' For now Access 2003 db engine only, but in case of other DB Engines there should be added Select Case
    Dim F As CFilter
On Error GoTo ErrHandler
If pFilterContent.Count = 0 Then Exit Property
For Each F In pFilterContent
    toString = toString & F.toString & UnionType
Next F
toString = Left(toString, Len(toString) - Len(UnionType))
Exit Property

ErrHandler:
EmergencyExit ("Неуспешен опит за прилагане на филтрите към запитването. Модул CFilterCollection")
End Property


Public Sub AddFilterSingleCondValue(AtributeName As String, CompareMode As String, CondValue As Variant, ConvertToType As String)
    Dim oFilter As CFilter
On Error GoTo ErrHandler
If pFilterContent Is Nothing Then Set pFilterContent = New Collection
    Set oFilter = New CFilter
        oFilter.AtributeName = AtributeName
        oFilter.CompareMode = CompareMode
        oFilter.AddCondContent CondValue
        oFilter.ConvertToType = ConvertToType
        Me.AddFilterObject oFilter
    Set oFilter = Nothing
Exit Sub
ErrHandler:
EmergencyExit ("Неуспешен опит за създаване на филтър. Процедура AddFilterSingleCondValue. Модул CFilterCollection")
End Sub
