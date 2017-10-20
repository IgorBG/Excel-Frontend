VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private pAtributeName As String
Private pCompareMode As String
Private pCondContent As Collection
Private pConvertToType As String
Private pDelim As String
Private pApostropheIfString As String
Private typeRecognized As Boolean

    Public Property Let AtributeName(ByVal ValueIn As String)
        pAtributeName = ValueIn
    End Property
    Public Property Get AtributeName() As String
        AtributeName = pAtributeName
    End Property
    Public Property Let CompareMode(ByVal ValueIn As String)
        pCompareMode = ValueIn
    End Property
    Public Property Get CompareMode() As String
        CompareMode = pCompareMode
    End Property
    Public Property Let ConvertToType(ByVal ValueIn As String)
        pConvertToType = ValueIn
    End Property
    Public Property Get ConvertToType() As String
        ConvertToType = pConvertToType
    End Property
    
Public Sub AddCondContent(ByVal inValue As Variant)
If pCondContent Is Nothing Then Set pCondContent = New Collection
    pCondContent.Add inValue
End Sub
Public Sub ClearCondContent()
Set pCondContent = Nothing
End Sub


Public Property Get toString() '(Optional ByVal dbEngine As String) As String
'Converts Filter object to string so it can be injected to sql querry
' For now Access 2003 db engine only, but in case of other DB Engines there should be added Select Case
Dim condBegin As String, condEnd As String
Dim breakPoint As Long
Dim v As Variant
On Error GoTo ErrHandler
    breakPoint = 50
    Call setQuerryWrap(condBegin, condEnd, pCompareMode)
    breakPoint = 200
    toString = " " & pAtributeName & condBegin & CondContentString & condEnd & " "
Exit Property

ErrHandler:
Select Case breakPoint
    Case 50: EmergencyExit ("Неуспешен опит за прилагане на филтрите към запитването. Модул Филтри")
    Case 200: EmergencyExit ("Неуспешен опит за преконвертиране на стойността. Модул Филтри")
    Case 250: EmergencyExit ("Програмата не може да преконвертира данните в тип " & pConvertToType)
    Case Else: EmergencyExit ("Неуспешен опит за прилагане на филтрите към запитването. Модул Филтри")
End Select
End Property


Private Sub typeRecognition()
    If pConvertToType = vbNullString Then
        If IsDate(pCondContent.Item(1)) Then pConvertToType = "long"
        If IsNumeric(pCondContent.Item(1)) = False Then pConvertToType = "string"
    End If
    
    Select Case LCase(pConvertToType)
        Case "string", "text"
            pApostropheIfString = "'"
    End Select
Delim = pApostropheIfString & "," & pApostropheIfString
typeRecognized = True
End Sub

Private Sub setQuerryWrap(ByRef condBegin As String, ByRef condEnd As String, ByVal CmprSign)
If Not typeRecognized Then Call typeRecognition
On Error GoTo ErrHandler
CmprSign = Replace(pCompareMode, " ", vbNullString)
    Select Case LCase(CmprSign)
        Case "=", ">=", "<=", ">", "<"
            If pCondContent.Count > 1 Then breakPoint = 100: GoTo ErrHandler
            condBegin = pCompareMode & pApostropheIfString
            condEnd = pApostropheIfString
        Case "in", "notin"
            condBegin = " " & pCompareMode & " (" & pApostropheIfString
            condEnd = pApostropheIfString & ")"
        Case "like", "notlike"
            If pCondContent.Count > 1 Then breakPoint = 100: GoTo ErrHandler
            condBegin = " " & pCompareMode & " '%"
            condEnd = "%'"
        Case Else: breakPoint = 150: GoTo ErrHandler
    End Select
Exit Sub
ErrHandler:
Select Case breakPoint
    Case 100: EmergencyExit ("Филтърът със посочения знак не може да поема множественни условия. Процедура setQuerryWrap, модул CFilter")
    Case 150: EmergencyExit ("Непознат вид на филтъра на данните. Процедура setQuerryWrap, модул CFilter")
    Case Else: EmergencyExit ("Не мога да съставя условието за филтъра. Процедура setQuerryWrap, модул CFilter")
End Select
End Sub



Private Function CondContentString() As String
    Dim v As Variant
    For Each v In pCondContent
        CondContentString = CondContentString & CnvrtDataType(v, pConvertToType) & pDelim
    Next v
CondContentString = Left(CondContentString, Len(CondContentString) - Len(pDelim))
End Function
Private Function CnvrtDataType(ByRef inValue As Variant, DataType As String) As Variant
        Select Case LCase(pConvertToType)
            Case "long": CnvrtDataType = CLng(inValue): Exit Function
            Case "bool": CnvrtDataType = CBool(inValue): Exit Function
            Case "string", "text": CnvrtDataType = CStr(inValue): Exit Function
            Case vbNullString: CnvrtDataType = inValue: Exit Function
        End Select
End Function



