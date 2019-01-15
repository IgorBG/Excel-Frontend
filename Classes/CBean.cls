VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Private pProp As String 'Property Name
    Private pVal As Variant 'Value
Public Property Get Prop() As String
    Prop = pProp
End Property
Public Property Let Prop(ByVal ValueIn As String)
    pProp = ValueIn
End Property
Public Property Get Val() As Variant
    Val = pVal
End Property
Public Property Let Val(ByVal ValueIn As Variant)
    pVal = ValueIn
End Property
