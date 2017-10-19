VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDataContent As Variant         'Actual Array with data
Private pClmnNames As Collection        'Column Name Indexes
Private pRowNames As Collection

Public Property Get getContent() As Variant
    getContent = pDataContent
End Property
Public Property Let letContent(ValueIn As Variant)
    pDataContent = ValueIn
End Property

Public Property Get getLbound() As Long
    getLbound = LBound(pDataContent)
End Property
Public Property Get getUbound() As Long
    getUbound = UBound(pDataContent)
End Property


Public Property Get getLbound2() As Long
    getLbound2 = LBound(pDataContent, 2)
End Property
Public Property Get getUbound2() As Long
    getUbound2 = UBound(pDataContent, 2)
End Property

Public Property Get getClmnIndx(ByVal ClmnName As String) As Long
    getClmnIndx = pClmnNames(ClmnName).Stojnost
End Property

Private Sub Class_Initialize()
If pClmnNames Is Nothing Then Set pClmnNames = New Collection
End Sub
Public Sub ResetMap()
    Set pClmnNames = New Collection
End Sub


Public Property Get getValue(ByVal rowIndex As String, ByVal clmnIndex As String) As Variant
    rowIndex = TestedIndex(rowIndex, pRowNames)
    clmnIndex = TestedIndex(clmnIndex, pClmnNames)
    getValue = pDataContent(CLng(clmnIndex), CLng(rowIndex))
End Property
Private Function TestedIndex(ByVal indx As String, ByRef col As Collection) As Variant
    If Not ObjectIsInCollection(col, indx) Then
        If Not IsNumeric(indx) Then Call EmergencyExit("Грешка при сканиране на datasource за наличие на колона с индекс: " & indx & ". Модул CDataStorage")
        TestedIndex = indx: Exit Function
    End If
    TestedIndex = getClmnIndx(indx)
End Function



Public Property Let letClmnNames(ByVal CommaSeparNames As String)
    Dim v As Variant, i As Long
    Dim StartPoint As Integer
    
    CommaSeparNames = Replace(CommaSeparNames, " ", vbNullString)
    If CommaSeparNames = vbNullString Then Exit Property
    Call ResetMap
    
    If Not IsArray(pDataContent) Then ReDim pDataContent(0 To 0, 0 To 0)
    v = Split(CommaSeparNames, ",")
    For i = LBound(v) To UBound(v)
        Call addClmnAtribt(v(i), i)
    Next i
End Property
    Private Sub addClmnAtribt(ByVal AtribtName As String, ClmnNum As Long)
        Dim TekAtribt As New CNastrojka
        With TekAtribt
            TekAtribt.Ime = AtribtName
            TekAtribt.Stojnost = ClmnNum
            pClmnNames.Add TekAtribt, AtribtName
        End With
        Set TekAtribt = Nothing
    End Sub
