VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Use for artcle order and production lists. Last update: 20220114
Private pOrdNum As Long
Private pDate1 As Date
Private pCompany1 As String
Private pProdDir As String
Private pERPRef As String
Private pArtNum As String
Private pArtName As String
Private pQnty As Double
Private pNote As String
Private pQuoteNum As String 'quatation number
Private pDBID As Long
Private pPaketazh As String
Private pGrad As String
Private pIsApproved As Boolean
Private pDoClearSizes As Boolean
Private pItems As Collection 'packs
Private pAdditInfo As Collection 'additional characteristics
Private pBrPak As Long


Public Property Get DoClearSizes() As Boolean
    DoClearSizes = pDoClearSizes
End Property
Public Property Let DoClearSizes(ByVal ValueIn As Boolean)
    pDoClearSizes = ValueIn
End Property
Public Property Get isApproved() As Boolean
    isApproved = pIsApproved
End Property
Public Property Let isApproved(ByVal ValueIn As Boolean)
    pIsApproved = ValueIn
End Property
Public Property Get Grad() As String
    Grad = pGrad
End Property
Public Property Let Grad(ByVal ValueIn As String)
    pGrad = ValueIn
End Property
Public Property Get Paketazh() As String
    Paketazh = pPaketazh
End Property
Public Property Let Paketazh(ByVal ValueIn As String)
    pPaketazh = ValueIn
End Property
Public Property Get OrdNum() As String
    OrdNum = pOrdNum
End Property
Public Property Let OrdNum(ByVal ValueIn As String)
    If IsNumeric(ValueIn) Then pOrdNum = ValueIn
End Property
Public Property Get Date1() As Date
    Date1 = pDate1
End Property
Public Property Let Date1(ByVal ValueIn As Date)
    pDate1 = ValueIn
End Property
Public Property Get ClientCompany() As String
    ClientCompany = pCompany1
End Property
Public Property Let ClientCompany(ByVal ValueIn As String)
    pCompany1 = ValueIn
End Property
Public Property Get ERPRef() As String
    ERPRef = pERPRef
End Property
Public Property Let ERPRef(ByVal ValueIn As String)
    pERPRef = ValueIn
End Property
Public Property Get artNum() As String
    artNum = pArtNum
End Property
Public Property Let artNum(ByVal ValueIn As String)
    pArtNum = ValueIn
End Property
Public Property Get artName() As String
    artName = pArtName
End Property
Public Property Let artName(ByVal ValueIn As String)
    pArtName = ValueIn
End Property
Public Property Get Qnty() As Double
    Qnty = pQnty
End Property
Public Property Let Qnty(ByVal ValueIn As Double)
    pQnty = ValueIn
End Property
Public Property Get ProdDir() As String
    ProdDir = pProdDir
End Property
Public Property Let ProdDir(ByVal ValueIn As String)
    pProdDir = ValueIn
End Property
Public Property Get Note() As String
    Note = pNote
End Property
Public Property Let Note(ByVal ValueIn As String)
    pNote = ValueIn
End Property
Public Property Get quoteNum() As String
    quoteNum = pQuoteNum
End Property
Public Property Let quoteNum(ByVal ValueIn As String)
    pQuoteNum = ValueIn
End Property
Public Property Get DBID() As Long
    DBID = pDBID
End Property
Public Property Let DBID(ByVal ValueIn As Long)
    pDBID = ValueIn
End Property

Public Sub AddItem(ByVal obj As Object)
If pItems Is Nothing Then Set pItems = New Collection
    pItems.Add obj
End Sub
Public Property Get Items() As Collection
    Set Items = pItems
End Property
Public Property Get BrPak() As Long
    BrPak = pBrPak
End Property
Public Property Let BrPak(ByVal ValueIn As Long)
    pBrPak = ValueIn
End Property
Public Sub AddAdditInfo(ByVal obj As Object)
If pAdditInfo Is Nothing Then Set pAdditInfo = New Collection
    pAdditInfo.Add obj
End Sub
Public Property Get AdditInfo() As Collection
    Set AdditInfo = pAdditInfo
End Property



