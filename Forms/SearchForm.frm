VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchForm 
Caption         =   "Search"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   OleObjectBlob   =   "SearchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

'Private pEventLog As String
Private pSearchQuery As String
Private pQueryType As Variant
Private pResults As Variant
Private pResultObj As Collection
Private pContext As String
Private Const defWidth As Double = 450
Private Const defHeight As Double = 200
Private Const frmBord As Double = 4
Private Const defWindowCaption As String = "Прозорец за търсене"
Private Const defHelpCaption As String = "Въведете името по което желаете да търсите елементите в Базата данни"
    
Private WithEvents EventLogger As CListener
    
    
    Public Property Let SearchQuery(ByVal ValueIn As String)
        pSearchQuery = ValueIn
    End Property
    Public Property Get SearchQuery() As String
        SearchQuery = pSearchQuery
    End Property
    Public Property Let queryType(ByVal ValueIn As Variant)
        pQueryType = ValueIn
    End Property
    Public Property Get queryType() As Variant
        queryType = pQueryType
    End Property
    Public Property Let Results(ByVal DataIn As Variant)
        pResults = DataIn
    End Property
    Public Property Get Results() As Variant
        Results = pResults
    End Property
    Public Property Get Context() As String
        Context = pContext
    End Property
    Public Property Set ResultObject(ByVal DataIn As Variant)
        Set pResultObj = DataIn
    End Property
    Public Property Get ResultObject() As Variant
        Set ResultObject = pResultObj
    End Property

Private Sub EventLogger_Change(ByVal name As String)
If Not name = vbNullString Then Call ProcessFormEvent(Me, name, Me.Context)
EventLogger.ResetLog
End Sub

Private Sub UserForm_Initialize()
    Set EventLogger = New CListener
    Set pResultObj = New Collection
    Me.btnOption1.Visible = False
    Me.LBResultList.MultiSelect = 0
    Me.TBSearchText.SetFocus
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
    EventLogger.EventName = "Cancel"
    End If
End Sub


Public Sub Constructor(Optional frmWidth As Double, Optional frmHeight As Double, Optional frmCaption As String, Optional helpCaption As String, Optional clmnCount As Long, Optional Context As String)
'Конструкторът на прозореца за търсене. Подрежда полета и бутоните според размера на прозореца
    Dim TempDbl As Double
'Defaults
If frmWidth = 0 Then frmWidth = defWidth
If frmHeight = 0 Then frmHeight = defHeight
If clmnCount = 0 Then clmnCount = 1
If frmCaption = vbNullString Then frmCaption = defWindowCaption
If helpCaption = vbNullString Then helpCaption = defHelpCaption
With Me
'Positions
    .Width = frmWidth
    .Height = frmHeight
    .TBSearchText.Left = frmBord: .TBSearchText.Width = Width - frmBord * 2
    .HeadText.Left = frmBord: .HeadText.Width = Width - frmBord * 2
    .LBResultList.Left = frmBord: .LBResultList.Width = Width - frmBord * 2
    .btnOk.Top = frmHeight - .btnOk.Height * 2: .btnOk.Left = Width - .btnOk.Width * 2 - (frmBord * 4)
    .btnCancel.Top = frmHeight - .btnCancel.Height * 2: .btnCancel.Left = Width - .btnCancel.Width - (frmBord * 2)
    .btnOption1.Top = frmHeight - .btnOption1.Height * 2: .btnOption1.Left = frmBord
    .LBResultList.Height = frmHeight - .LBResultList.Top - (frmHeight - .btnOk.Top) - frmBord
'Values
    .Caption = frmCaption
    .HeadText.Caption = helpCaption
'Settings
    .LBResultList.ColumnCount = clmnCount
    pContext = Context
End With
End Sub


Private Sub OKProcess()
EventLogger.EventName = "Ok"
End Sub

Private Sub CancelProcess()
EventLogger.EventName = "Cancel"
End Sub

Private Sub btnOk_Click()
 Call OKProcess
End Sub

Private Sub LBResultList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 Call OKProcess
End Sub
Private Sub LBResultList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 13: Call OKProcess
    End Select
End Sub
Private Sub TBSearchText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 13: EventLogger.EventName = "tbxSearch_Update"
    End Select
End Sub


Private Sub btnCancel_Click()
 Call CancelProcess
End Sub

Private Sub btnOption1_Click()
    EventLogger.EventName = "OB1_Click"
End Sub

Private Sub TBSearchText_AfterUpdate()
    EventLogger.EventName = "tbxSearch_Update"
End Sub
