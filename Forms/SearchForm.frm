VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchForm 
   Caption         =   "Òúðñåíå"
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
Private pSearchQuerry As String
Private pResults As Variant
Private pContext As String
Private Const defWidth As Double = 450
Private Const defHeight As Double = 200
Private Const frmBord As Double = 4
Private Const defWindowCaption As String = "Ïðîçîðåö çà òúðñåíå"
Private Const defHelpCaption As String = "Âúâåäåòå èìåòî ïî êîåòî æåëàåòå äà òúðñèòå åëåìåíòèòå â Áàçàòà äàííè"
    
Private WithEvents EventLogger As CListener
Attribute EventLogger.VB_VarHelpID = -1
    
    
    Public Property Let SearchQuerry(ByVal ValueIn As String)
        pSearchQuerry = ValueIn
    End Property
    Public Property Get SearchQuerry() As String
        SearchQuerry = pSearchQuerry
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

Private Sub EventLogger_Change(ByVal Name As String)
If Not Name = vbNullString Then Call ProcessFormEvent(Me, Name, Me.Context)
EventLogger.ResetLog
End Sub

Private Sub UserForm_Initialize()
    Set EventLogger = New CListener
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
'Êîíñòðóêòîðúò íà ïðîçîðåöà çà òúðñåíå. Ïîäðåæäà ïîëåòà è áóòîíèòå ñïîðåä ðàçìåðà íà ïðîçîðåöà
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



