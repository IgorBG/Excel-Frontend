Attribute VB_Name = "BeansHandler"
Option Explicit

Public Sub Inicial_Add(ByRef Parent As Collection, KeyText As String, Optional Mode As String, Optional ByVal NameCol As String = 1, Optional ByVal ValCol As String = 3)
    Dim StartRowCol As Collection, StartRow As Long
If Not isInitialized Then Call Inicial_Main
On Error GoTo ErrHandler

Set StartRowCol = Collection_Nastrojki_Edna_Kolona("#Content", "#ContentEnd", "LastCol", 2, nastrSheet, 1)
If ObjectIsInCollection(StartRowCol, KeyText) Then
    StartRow = StartRowCol.Item(KeyText).Val
    Else: StartRow = 0
End If

If Parent Is Nothing Then Set Parent = New Collection

If Mode = vbNullString Then Mode = "SINGLE"
Select Case Mode
    Case "SINGLE"
    Parent.Add Collection_Nastrojki_Edna_Kolona(CStr("#" & KeyText), CStr("#" & KeyText & "End"), "LastCol", 2, NastrSheet, StartRowCol.Item(KeyText).Val, NameCol, ValCol), KeyText
    Case "MULTI"
    Parent.Add Collection_Nastrojki_Mnogo_Koloni(CStr("#" & KeyText), CStr("#" & KeyText & "End"), "LastCol", 2, NastrSheet, 1, StartRowCol.Item(KeyText).Val, NameCol), KeyText
End Select
Exit Sub

ErrHandler:
EmergencyExit "Модул Inicial_Add. Не откривам необходимата маркировка за таблица " & KeyText
End Sub

Public Sub Inicial_Main()
    Dim StartRowCol As Collection
Call InicialWS(oSheet, "Porychki")
Call InicialWS(pakSheet, "Paketi")
Call InicialWS(printSheet, "Print")
Call InicialWS(NPSheet, "NarProizv")
Call InicialWS(nastrSheet, "Nastrojki")
Call InicialWS(impSheet, "import")

eventImportSheetChange = True

Set StartRowCol = Collection_Nastrojki_Edna_Kolona("#Content", "#ContentEnd", "LastCol", 2, nastrSheet, 1)
Set Nastr = New Collection
Set DS = New Collection
Set DC = New Collection
Nastr.Add Collection_Nastrojki_Edna_Kolona("#OrdMark", "#OrdMarkEnd", "LastCol", 2, nastrSheet, StartRowCol.Item("OrdMark").Val), "OrdMark"
Nastr.Add Collection_Nastrojki_Edna_Kolona("#PakMark", "#PakMarkEnd", "LastCol", 2, nastrSheet, StartRowCol.Item("PakMark").Val), "PakMark"
Nastr.Add Collection_Nastrojki_Edna_Kolona("#DetMark", "#DetMarkEnd", "LastCol", 2, nastrSheet, StartRowCol.Item("DetMark").Val), "DetMark"
Nastr.Add Collection_Nastrojki_Edna_Kolona("#Datasources", "#DatasourcesEnd", "LastCol", 2, nastrSheet, StartRowCol.Item("Datasources").Val), "Datasources"
Nastr.Add Collection_Nastrojki_Edna_Kolona("#NPMark", "#NPMarkEnd", "LastCol", 2, nastrSheet, StartRowCol.Item("NPMark").Val), "NPMark"
Nastr.Add Collection_Nastrojki_Mnogo_Koloni("#NPFltr", "#NPFltrEnd", "LastCol", 2, nastrSheet, 1, StartRowCol.Item("NPFltr").Val), "NPFltr"
Nastr.Add Collection_Nastrojki_Mnogo_Koloni("#OrdFltr", "#OrdFltrEnd", "LastCol", 2, nastrSheet, 1, StartRowCol.Item("OrdFltr").Val), "OrdFltr"
Nastr.Add Collection_Nastrojki_Edna_Kolona("#Print", "#PrintEnd", "LastCol", 2, nastrSheet, StartRowCol.Item("Print").Val), "Print"

Call InicialDC

isInitialized = True
End Sub

    Private Sub InicialWS(ByRef WSObject As Worksheet, ByVal WSName As String)
        On Error GoTo ErrHandler
        Set WSObject = ThisWorkbook.Worksheets(WSName): Exit Sub
ErrHandler:
        MsgBox "Програмата не може да открие лист с име '" & WSName & "'. Вероятно той е бил изтрит или преименуван." & _
                "Моля, възстановете го или ползвайте последната работеща версия на програмата": End
End Sub
    Private Sub InicialDC() 'Singletone inicialisation of all data connections
       ' On Error GoTo ErrHandler
        Dim b As CBean
        Dim Path As String
            
            For Each b In Nastr.Item("Datasources")
                Path = b.Val
                If Not ObjectIsInCollection(DC, Path) Then DC.Add GetNewConnToAccess(Path, False), Path
            Next b
    Exit Sub
ErrHandler:
    Call EmergencyExit("Програмата не може да открие файл с име '" & Path & "'. Вероятно той е бил изтрит или преименуван." & _
                "Моля, възстановете го или обърнете се за помощ при администратора на програмата")
    End Sub
            Private Function Collection_Nastrojki_Edna_Kolona(ByRef NachZapis As String, ByRef KrajZapis As String, ByRef LastKolZapis As String, IgnoreRows As Integer, SourceSheet As Worksheet, Optional StartRow As Long) As Collection
            ' For Single column settings
            Dim Col As New Collection
            Dim Data As Variant
            Dim i As Integer
            Dim Tek_Atr As CBean
            If StartRow = 0 Then StartRow = 1
            Data = Data_Nastrojki(NachZapis, KrajZapis, LastKolZapis, IgnoreRows, SourceSheet, StartRow)
                For i = 1 To UBound(Data)
                    Set Tek_Atr = New CBean
                    Tek_Atr.Prop = Data(i, 1)
                    Tek_Atr.Val = Data(i, 3)
                    On Error Resume Next
                    Col.Add Tek_Atr, Tek_Atr.Prop
                    Set Tek_Atr = Nothing
                Next i
            Set Data = Nothing
            Set Collection_Nastrojki_Edna_Kolona = Col
            Set Col = Nothing
            End Function
            
            Private Function Collection_Nastrojki_Mnogo_Koloni(ByRef NachZapis As String, ByRef KrajZapis As String, ByRef LastKolZapis As String, IgnoreRows As Integer, SourceSheet As Worksheet, KeyRow As Integer, Optional StartRow As Long) As Collection
            Dim Col As New Collection
            Dim ColonaNastr As Collection
            Dim Data As Variant
            Dim i As Integer, j As Integer
            Dim Tek_Atr As CBean
            If StartRow = 0 Then StartRow = 1
                Data = Data_Nastrojki(NachZapis, KrajZapis, LastKolZapis, IgnoreRows, SourceSheet, StartRow)
            For j = LBound(Data, 2) + 2 To UBound(Data, 2)   'From the each colunm with attributes in the marked range
                Set ColonaNastr = New Collection
                For i = 1 To UBound(Data)
                    Set Tek_Atr = New CBean
                    Tek_Atr.Prop = Data(i, 1)
                    Tek_Atr.Val = Data(i, j)
                    On Error Resume Next
                    ColonaNastr.Add Tek_Atr, Tek_Atr.Prop
                    Set Tek_Atr = Nothing
                Next i
            
            If Col Is Nothing Then Set Col = New Collection
            Col.Add ColonaNastr, CStr(Data(KeyRow, j))
            Set ColonaNastr = Nothing
            Next j
            
            Set Data = Nothing
            Set Collection_Nastrojki_Mnogo_Koloni = Col
            Set Col = Nothing
            End Function
             
      Private Function Data_Nastrojki(ByRef NachZapis As String, ByRef KrajZapis As String, ByRef LastKolZapis As String, IgnoreRows As Integer, SourceSheet As Worksheet, StartRow As Long) As Variant
            Dim SourceSheetLastRow  As Integer, SourceSheetLastCol  As Integer
            Dim FirstRow As Integer, LastRow As Integer, LastCol As Integer
            Dim i As Integer
            On Error GoTo ErrHandler
             SourceSheetLastRow = SourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
             'SourceSheetLastCol = SourceSheet.Cells(1, Columns.Count).End(xlToLeft).Column
             SourceSheetLastCol = 255
            ' търсим настройките на страницата
            For i = StartRow To SourceSheetLastRow ' търсим реда с кодовото име NachZapis
                If SourceSheet.Cells(i, 1).value = NachZapis Then FirstRow = i: Exit For
            Next i
            For i = FirstRow To SourceSheetLastRow ' търсим реда с кодовото име KrajZapis
                If SourceSheet.Cells(i, 1).value = KrajZapis Then LastRow = i - 1: Exit For
            Next i
            For i = 1 To SourceSheetLastCol ' търсим колоната с кодовото име LastKolZapis
                If SourceSheet.Cells(FirstRow, i).value = LastKolZapis Then LastCol = i - 1: Exit For
            Next i
            Data_Nastrojki = SourceSheet.Range(SourceSheet.Cells(FirstRow + IgnoreRows, 1), SourceSheet.Cells(LastRow, LastCol))
            Exit Function

ErrHandler:
Call EmergencyExit("Не откривам необходимата маркировка за таблица " & NachZapis & vbCrLf & _
                    "Откритите параметри:" & vbCrLf & _
                    "Първия ред: " & FirstRow & " - Ключ:" & NachZapis & vbCrLf & _
                    "Последната колона: " & LastCol & " - Ключ:" & LastKolZapis & vbCrLf & _
                    "Последния ред: " & LastRow & " - Ключ:" & KrajZapis)
End Function

Public Function getNewBean(ByVal PropertyName As String, ByVal value As Variant) As CBean
'On Error GoTo ErrHandler
Set getNewBean = New CBean
    getNewBean.Prop = PropertyName
    getNewBean.Val = value
Exit Function
ErrHandler:
Call EmergencyExit("Функция getNewBean")
End Function


Public Sub DeleteUserMenu()
    Dim CntxMenu As CommandBar
    Dim UserMenu As CommandBarControl
Set CntxMenu = Application.CommandBars("Cell")
For Each UserMenu In CntxMenu.Controls
    If UserMenu.Tag = "AddedByUser" Then UserMenu.Delete    ' Delete the menu added by user.
Next UserMenu
End Sub
