VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Private pEventName As String
    Public Event Change(ByVal ValueIn As String)
    
Public Property Get EventName() As String
    EventName = pEventName
End Property
Public Property Let EventName(ByVal ValueIn As String)
    pEventName = ValueIn
    RaiseEvent Change(pEventName)
End Property
        
Public Sub ResetLog()
    pEventName = vbNullString
End Sub
