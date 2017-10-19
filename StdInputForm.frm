VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StdInputForm 
   Caption         =   "Въвеждане в база данни"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   OleObjectBlob   =   "StdInputForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StdInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private pEventLog As String
Private pSearchQuerry As String
Private pResults As Variant
Private pContext As String
Private Const frmBord As Double = 4
Private Const defWindowCaption As String = "Прозорец за въвеждане"
Private Const defHelpCaption As String = "Въведете данните, които желаете да въедете в Базата данни"
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



Private Sub UserForm_Initialize()
    If Me.txtInput1Main.Enabled Then Me.txtInput1Main.SetFocus
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        pEventLog = "Cancel"
        Me.Hide
    End If
End Sub
Public Sub Constructor(Optional frmWidth As Double, Optional frmHeight As Double, Optional frmCaption As String, Optional helpCaption As String, Optional clmnCount As Long, Optional Context As String)
'Конструкторът на прозореца за търсене. Подрежда полета и бутоните според размера на прозореца
    Dim TempDbl As Double
'Defaults
If frmWidth = 0 Then frmWidth = Me.Width
If frmHeight = 0 Then frmHeight = Me.Height
If frmCaption = vbNullString Then frmCaption = defWindowCaption
If helpCaption = vbNullString Then helpCaption = defHelpCaption
With Me
'Positions
    .Width = frmWidth
    .Height = frmHeight
    .txtInput1Main.Left = frmBord: .txtInput1Main.Width = Width - frmBord * 2
    .lblHelpHeader.Left = frmBord: .lblHelpHeader.Width = Width - frmBord * 2
    .btnOk.Top = frmHeight - .btnOk.Height * 2: .btnOk.Left = .btnOk.Width * 2 - frmBord
    .btnCancel.Top = frmHeight - .btnCancel.Height * 2: .btnCancel.Left = Width - .btnCancel.Width - frmBord * 2
    .btnOption1.Top = frmHeight - .btnOption1.Height * 2: .btnOption1.Left = frmBord
'Values
    .Caption = frmCaption
    .lblHelpHeader.Caption = helpCaption
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
Private Sub btnCancel_Click()
 Call CancelProcess
End Sub

Private Sub btnOption1_Click()
    EventLogger.EventName = "OB1_Click"
End Sub
'option element events on main page
Private Sub btnOB1Main_Click()
    EventLogger.EventName = "btnOB1Main_Click"
End Sub
Private Sub btnOB2Main_Click()
    EventLogger.EventName = "btnOB2Main_Click"
End Sub




