VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNastrojka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Private pIme As String
    Private pStojnost As Variant
Public Property Get Ime() As String
    Ime = pIme
End Property
Public Property Let Ime(ByVal ValueIn As String)
    pIme = ValueIn
End Property
Public Property Get Stojnost() As Variant
    Stojnost = pStojnost
End Property
Public Property Let Stojnost(ByVal ValueIn As Variant)
    pStojnost = ValueIn
End Property

