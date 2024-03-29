VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CZRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Private pQuotDate As Date
    Private pQuotID As String
    Private pOrderNumber As Long
    Private pLoadDate As Date
    Private pArticleName As String
    Private pColor1 As String
    Private pColor2 As String
    Private pProjectRef As String
    Private pQuotedQnty As Double
    Private pUnitPrice As Double
    Private pDiscount As Double
    Private pTotalPrice As Double
    Private pWHCode As Long
    Private pReservedQnty As Double
    Private pSoldQnty As Double
    Private pLeftQnty As Double
    Private pToProduceQnty As Double
    Private pQuotApproved As Boolean
    Private pMONumber As Long
    Private pMODate As Date
    Private pDONumber As Long
    Private pDODate As Date
    Private pTransferedQnty As Double
    Private pArticleID As String
'    Private pInfo1 As String
'    Private pDueDate As Date
'    Private pWHDoc As String
    Private pInfo2 As String
'    Private pNote As String
    Private pClientName As String
    Private pClientLoc As String
    Private pDeliveryPlaceName As String
    Private pDeliveryPlaceLoc As String
    Private phasTransport As Boolean
    Private pParentArticle As String
    Private pProdLevel As Long
    
    
    Private pErrMsg As String

Public Property Get ErrMsg() As String
    ErrMsg = pErrMsg
End Property
Public Property Let ErrMsg(ByVal ValueIn As String)
    Const Limit As Long = 32
    pErrMsg = pErrMsg & "; " & Left(ValueIn, Limit)
End Property


Public Property Get QuotDate() As Date
    QuotDate = pQuotDate
End Property
Public Property Let QuotDate(ByVal ValueIn As Date)
    'If Not isValidValue("date", True, ValueIn) Then ErrMsg = "���� (������)" Else
    pQuotDate = CDate(ValueIn)
End Property

Public Property Get QuotID() As String
    QuotID = pQuotID
End Property
Public Property Let QuotID(ByVal ValueIn As String)
    Const Limit As Long = 20
    'If Not isValidValue("string", True, ValueIn) Then ErrMsg = "����� (������)"
    'Else:
        If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
        pQuotID = ValueIn
    'End If
End Property

Public Property Get OrderNumber() As Long
    OrderNumber = pOrderNumber
End Property
Public Property Let OrderNumber(ByVal ValueIn As Long)
    'If Not isValidValue("long", True, ValueIn) Then ErrMsg = "���. �����"
    'Else:
        'If IsNull(pOrderNumber) Then pOrderNumber = 0 Else
        pOrderNumber = CLng(ValueIn)
    'End If
End Property

Public Property Get LoadDate() As Date
    LoadDate = pLoadDate
End Property
Public Property Let LoadDate(ByVal ValueIn As Date)
    'If Not isValidValue("date", True, ValueIn) Then ErrMsg = "���� �� ��������" Else
    pLoadDate = CDate(ValueIn)
End Property

Public Property Get ArticleName() As String
    ArticleName = pArticleName
End Property
Public Property Let ArticleName(ByVal ValueIn As String)
    Const Limit As Long = 128
    'If Not isValidValue("string", False, ValueIn) Then ErrMsg = "�������"
    'Else:
        ValueIn = Application.WorksheetFunction.Trim(ValueIn)
        If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
        pArticleName = ValueIn
    'End If
End Property

Public Property Get Color1() As String
    Color1 = pColor1
End Property
Public Property Let Color1(ByVal ValueIn As String)
    Const Limit As Long = 32
    'If Not isValidValue("string", True, ValueIn) Then ErrMsg = "����� 1"
    'Else:
        If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
        pColor1 = ValueIn
    'End If
End Property

Public Property Get Color2() As String
    Color2 = pColor2
End Property
Public Property Let Color2(ByVal ValueIn As String)
    Const Limit As Long = 32
    'If Not isValidValue("string", True, ValueIn) Then ErrMsg = "����� 2"
    'Else:
        If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
        pColor2 = ValueIn
    'End If
End Property

Public Property Get ProjectRef() As String
    ProjectRef = pProjectRef
End Property
Public Property Let ProjectRef(ByVal ValueIn As String)
    Const Limit As Long = 16
    'If Not isValidValue("string", True, ValueIn) Then ErrMsg = "�� ������"
    'Else:
        If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
        pProjectRef = ValueIn
    'End If
End Property

Public Property Get QuotedQnty() As Double
    QuotedQnty = pQuotedQnty
End Property
Public Property Let QuotedQnty(ByVal ValueIn As Double)
    'If Not isValidValue("double", False, ValueIn) Then ErrMsg = "�������" Else:
    pQuotedQnty = CDbl(ValueIn)
End Property

Public Property Get UnitPrice() As Double
    UnitPrice = pUnitPrice
End Property
Public Property Let UnitPrice(ByVal ValueIn As Double)
     'If Not isValidValue("double", True, ValueIn) Then ErrMsg = "���� � ���" Else:
     pUnitPrice = CDbl(ValueIn)
End Property

Public Property Get Discount() As Double
    Discount = pDiscount
End Property
Public Property Let Discount(ByVal ValueIn As Double)
    'If Not isValidValue("double", True, ValueIn) Then ErrMsg = "�� � �� ������" Else:
    pDiscount = CDbl(ValueIn)
End Property

Public Property Get TotalPrice() As Double
    TotalPrice = pTotalPrice
End Property
Public Property Let TotalPrice(ByVal ValueIn As Double)
    'If Not isValidValue("double", True, ValueIn) Then ErrMsg = "��-�� � ���" Else:
    pTotalPrice = CDbl(ValueIn)
End Property

Public Property Get WHCode() As Long
    WHCode = pWHCode
End Property
Public Property Let WHCode(ByVal ValueIn As Long)
     'If Not isValidValue("long", False, ValueIn) Then ErrMsg = "����� ���" Else:
     pWHCode = CLng(ValueIn)
End Property

Public Property Get ReservedQnty() As Double
    ReservedQnty = pReservedQnty
End Property
Public Property Let ReservedQnty(ByVal ValueIn As Double)
    'If Not isValidValue("double", True, ValueIn) Then ErrMsg = "�e���������" Else:
    pReservedQnty = CDbl(ValueIn)
End Property

Public Property Get SoldQnty() As Double
    SoldQnty = pSoldQnty
End Property
Public Property Let SoldQnty(ByVal ValueIn As Double)
    'If Not isValidValue("double", True, ValueIn) Then ErrMsg = "���������" Else:
    pSoldQnty = CDbl(ValueIn)
End Property

Public Property Get LeftQnty() As Double
    LeftQnty = pLeftQnty
End Property
Public Property Let LeftQnty(ByVal ValueIn As Double)
    'If Not isValidValue("double", True, ValueIn) Then ErrMsg = "�������" Else:
    pLeftQnty = CDbl(ValueIn)
End Property

Public Property Get ToProduceQnty() As Double
    ToProduceQnty = pToProduceQnty
End Property
Public Property Let ToProduceQnty(ByVal ValueIn As Double)
    'If Not isValidValue("double", True, ValueIn) Then ErrMsg = "� ���. �� ������." Else:
    pToProduceQnty = CDbl(ValueIn)
End Property

'Public Property Get QuotApproved() As Boolean
'    QuotApproved = pQuotApproved
'End Property
'Public Property Let QuotApproved(ByVal ValueIn As String)
'    If ValueIn = "��" Or LCase(ValueIn) = "true" Then pQuotApproved = True Else: pQuotApproved = False
'End Property

Public Property Get MONumber() As Long
    MONumber = pMONumber
End Property
Public Property Let MONumber(ByVal ValueIn As Long)
    'If Not isValidValue("long", True, ValueIn) Then ErrMsg = "�� �����"
    'Else:
    'If IsNull(ValueIn) Then pMONumber = 0 Else:
    pMONumber = CLng(ValueIn)
    'End If
End Property

Public Property Get MOdate() As Date
    MOdate = pMODate
End Property
Public Property Let MOdate(ByVal ValueIn As Date)
    'If Not isValidValue("date", True, ValueIn) Then ErrMsg = "�� ����"
    'Else:
    'If IsNull(ValueIn) Then pMODate = 0 Else:
    pMODate = CDate(ValueIn)
End Property

Public Property Get DONumber() As Long
    DONumber = pDONumber
End Property
Public Property Let DONumber(ByVal ValueIn As Long)
    'If Not isValidValue("long", True, ValueIn) Then ErrMsg = "�T �����" Else:
    pDONumber = ValueIn
End Property

Public Property Get DODate() As Date
    DODate = pDODate
End Property
Public Property Let DODate(ByVal ValueIn As Date)
    'If Not isValidValue("date", True, ValueIn) Then ErrMsg = "�T ����" Else:
    pDODate = ValueIn
End Property

Public Property Get TransferedQnty() As Double
    TransferedQnty = pTransferedQnty
End Property
Public Property Let TransferedQnty(ByVal ValueIn As Double)
    'If Not isValidValue("double", True, ValueIn) Then ErrMsg = "���������" Else:
    pTransferedQnty = CDbl(ValueIn)
End Property

Public Property Get ArticleID() As String
    ArticleID = pArticleID
End Property
Public Property Let ArticleID(ByVal ValueIn As String)
    Const Limit As Long = 20
    'If Not isValidValue("string", False, ValueIn) Then ErrMsg = "��� �����"
    'Else:
    If Limit > 0 And Len(ValueIn) > Limit Then
        ErrMsg = "��� �����: ��������� ����������� ������� �� ������ - " & Limit & " �������"
        Else
        pArticleID = ValueIn
    End If
End Property

'Public Property Get Info1() As String
'    Info1 = pInfo1
'End Property
'Public Property Let Info1(ByVal ValueIn As String)
'    Const Limit As Long = 24
'    'If Not isValidValue("string", True, ValueIn) Then ErrMsg = "����"
'    'Else:
'    If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
'    pInfo1 = ValueIn
'    'End If
'End Property

'Public Property Get DueDate() As Date
'    DueDate = pDueDate
'End Property
'Public Property Let DueDate(ByVal ValueIn As Date)
'    If Not isValidValue("date", True, ValueIn) Then Errmsg = "���� ��������" Else: pDueDate = CDate(ValueIn)
'End Property

'Public Property Get WHDoc() As String
'    WHDoc = pWHDoc
'End Property
'Public Property Let WHDoc(ByVal ValueIn As String)
'    pWHDoc = ValueIn
'End Property

Public Property Get Info2() As String
    Info2 = pInfo2
End Property
Public Property Let Info2(ByVal ValueIn As String)
    Const Limit As Long = 24
    'If Not isValidValue("string", True, ValueIn) Then ErrMsg = "���. ���� � �� ������"
    'Else:
    If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
    pInfo2 = ValueIn
    'End If
End Property

'Public Property Get Note() As String
'    Note = pNote
'End Property
'Public Property Let Note(ByVal ValueIn As String)
'    Const Limit As Long = 24
'    If Not isValidValue("string", True, ValueIn) Then ErrMsg = "���������"
'    Else: If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
'       pNote = ValueIn
'    End If
'End Property
Public Property Get ClientName() As String
    ClientName = pClientName
End Property
Public Property Let ClientName(ByVal ValueIn As String)
    Const Limit As Long = 50
    If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
    pClientName = ValueIn
End Property

Public Property Get ClientLoc() As String
    ClientLoc = pClientLoc
End Property
Public Property Let ClientLoc(ByVal ValueIn As String)
    Const Limit As Long = 24
    If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
    pClientLoc = ValueIn
End Property

Public Property Get DeliveryPlaceName() As String
    DeliveryPlaceName = pDeliveryPlaceName
End Property
Public Property Let DeliveryPlaceName(ByVal ValueIn As String)
    Const Limit As Long = 50
    If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
    pDeliveryPlaceName = ValueIn
End Property

Public Property Get DeliveryPlaceLoc() As String
    DeliveryPlaceLoc = pDeliveryPlaceLoc
End Property
Public Property Let DeliveryPlaceLoc(ByVal ValueIn As String)
    Const Limit As Long = 24
    If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
    pDeliveryPlaceLoc = ValueIn
End Property

Public Property Get hasTransport() As Boolean
    hasTransport = phasTransport
End Property
Public Property Let hasTransport(ByVal ValueIn As Boolean)
    phasTransport = ValueIn
End Property

Public Property Get ParentArticle() As String
    ParentArticle = pParentArticle
End Property
Public Property Let ParentArticle(ByVal ValueIn As String)
    Const Limit As Long = 16
    If Limit > 0 And Len(ValueIn) > Limit Then ValueIn = Left(ValueIn, Limit)
    pParentArticle = ValueIn
End Property
Public Property Get ProdLevel() As Long
    ProdLevel = pProdLevel
End Property
Public Property Let ProdLevel(ByVal ValueIn As Long)
    pProdLevel = ValueIn
End Property


Public Sub PopulateZRowViaInterface(ByRef arr As Variant, ByVal mCol As Collection, Optional ByRef s As Long, Optional outMsg As String)
    Dim p As Long
On Error GoTo ErrHandler
s = 1
p = 1: QuotDate = arr(mCol("ZaqDataCol").Val)
    If IsNull(QuotDate) Then QuotDate = vbNullString
    If IsNumeric(QuotDate) Or IsDate(QuotDate) Then QuotDate = Application.WorksheetFunction.RoundDown(QuotDate, 0)
p = 2: QuotID = arr(mCol("ZaqNomCol").Val)
p = 3: If CStr((arr(mCol("PorNomCol").Val))) = vbNullString Then arr(mCol("PorNomCol").Val) = 0
    OrderNumber = arr(mCol("PorNomCol").Val)
p = 4: LoadDate = arr(mCol("SpedDataCol").Val)
p = 5: ArticleName = arr(mCol("NomImeCol").Val)
    If ArticleName = vbNullString Then s = 0
p = 6: Color1 = arr(mCol("Color1Col").Val)
p = 7: Color2 = arr(mCol("Color2Col").Val)
p = 8: ProjectRef = arr(mCol("SdelkaCol").Val)
p = 9: QuotedQnty = arr(mCol("BrZaqvenoCol").Val)
    If QuotedQnty = 0 Then s = 0
p = 10: UnitPrice = arr(mCol("CenaCol").Val)
p = 11: Discount = arr(mCol("TOtstCol").Val)
p = 12: TotalPrice = arr(mCol("ObshtCenaCol").Val)
p = 13 'WHCode = arr(mCol("").Val) '������� �� �������
p = 14: ReservedQnty = arr(mCol("RezervCol").Val)
p = 15: SoldQnty = arr(mCol("ProdadCol").Val)
p = 16: LeftQnty = arr(mCol("OstatykCol").Val)
p = 17: ToProduceQnty = arr(mCol("ZaProizvCol").Val)
p = 18 'QuotApproved
p = 19 'MONumber   '������� �� �������
p = 20 'MODate     '������� �� �������
p = 21 'DONumber   '������� �� �������
p = 22 'DODate     '������� �� �������
p = 23: TransferedQnty = arr(mCol("PredadenCol").Val)
p = 24: ArticleID = arr(mCol("NomNomCol").Val)
    If ArticleID = vbNullString Then s = 0
p = 25 'Info1
p = 26 'DueDate
p = 27 'WHDoc
p = 28: Info2 = arr(mCol("DopInfoCol").Val)
p = 29 'Note
p = 30: ProdLevel = IIf(Left(ArticleName, 4) = "### ", 2, 1)
Exit Sub

ErrHandler:
s = 0
outMsg = getPhase(p)
End Sub

Private Function getPhase(ByVal p As Long) As String
    Select Case p
        Case 1: getPhase = "���� (������)"
        Case 2: getPhase = "����� (������)"
        Case 3: getPhase = "���. �����"
        Case 4: getPhase = "���� �� ��������"
        Case 5: getPhase = "�������"
        Case 6: getPhase = "����� 1"
        Case 7: getPhase = "����� 2"
        Case 8: getPhase = "�� ������"
        Case 9: getPhase = "�������"
        Case 10: getPhase = "���� � ���"
        Case 11: getPhase = "�� � �� ������"
        Case 12: getPhase = "��-�� � ���"
        Case 13: getPhase = "����� ���"
        Case 14: getPhase = "�����������"
        Case 15: getPhase = "���������"
        Case 16: getPhase = "�������"
        Case 17: getPhase = "� ���. �� ������."
        Case 23: getPhase = "���������"
        Case 24: getPhase = "��� �����"
        Case 28: getPhase = "���. ���� � �� ������"
        Case Else: getPhase = "�����"
    End Select
End Function
