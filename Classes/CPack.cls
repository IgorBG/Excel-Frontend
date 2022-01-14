VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Use for packages, parts in packages
Private pID As Long        ' пореден номер
Private pHandler As Long        ' пореден номер на пакета в пакетажа
Private pName As String
Private pInfo1 As String
Private pInfo2 As String
Private pInfoContent As String  'Модул
Private pImgFile As String
Private pQnty As Double
Private pTeg As Double
Private pDyl As Double
Private pShi As Double
Private pDeb As Double
Private pNote As String
Private pItems As Collection


Public Property Get ID() As Long
    ID = pID
End Property
Public Property Let ID(ByVal ValueIn As Long)
    pID = ValueIn
End Property
Public Property Get Handler() As Long
    Handler = pHandler
End Property
Public Property Let Handler(ByVal ValueIn As Long)
    pHandler = ValueIn
End Property
Public Property Get name() As String
    name = pName
End Property
Public Property Let name(ByVal ValueIn As String)
    pName = ValueIn
End Property
Public Property Get info1() As String
    info1 = pInfo1
End Property
Public Property Let info1(ByVal ValueIn As String)
    pInfo1 = ValueIn
End Property
Public Property Get Info2() As String
    Info2 = pInfo2
End Property
Public Property Let Info2(ByVal ValueIn As String)
    pInfo2 = ValueIn
End Property
Public Property Get InfoContent() As String
    InfoContent = pInfoContent
End Property
Public Property Let InfoContent(ByVal ValueIn As String)
    pInfoContent = ValueIn
End Property
Public Property Get ImgFile() As String
    ImgFile = pImgFile
End Property
Public Property Let ImgFile(ByVal ValueIn As String)
    pImgFile = ValueIn
End Property
Public Property Get Qnty() As Double
    Qnty = pQnty
End Property
Public Property Let Qnty(ByVal ValueIn As Double)
    pQnty = ValueIn
End Property
Public Property Get Teg() As Double
    Teg = pTeg
End Property
Public Property Let Teg(ByVal ValueIn As Double)
    pTeg = ValueIn
End Property
Public Property Get Dyl() As Double
    Dyl = pDyl
End Property
Public Property Let Dyl(ByVal ValueIn As Double)
    pDyl = ValueIn
End Property
Public Property Get Shi() As Double
    Shi = pShi
End Property
Public Property Let Shi(ByVal ValueIn As Double)
    pShi = ValueIn
End Property
Public Property Get Deb() As Double
    Deb = pDeb
End Property
Public Property Let Deb(ByVal ValueIn As Double)
    pDeb = ValueIn
End Property
Public Property Get Note() As String
    Note = pNote
End Property
Public Property Let Note(ByVal ValueIn As String)
    pNote = ValueIn
End Property


Public Sub AddItem(ByVal obj As Object)
If pItems Is Nothing Then Set pItems = New Collection
    pItems.Add obj
End Sub
Public Property Get Items() As Collection
    Set Items = pItems
End Property


