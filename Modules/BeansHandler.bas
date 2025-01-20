Option Explicit

Public Sub Inicial_Add(ByRef Parent As Collection, KeyText As String, Optional Mode As String, Optional ByVal NameCol As String = 1, Optional ByVal ValCol As String = 3)
    Dim StartRowCol As Collection
If Not IsInitialized Then Call Inicial_Main
Set StartRowCol = Collection_Nastrojki_Edna_Kolona("#Content", "#ContentEnd", "LastCol", 2, NastrSheet, 1)
If Parent Is Nothing Then Set Parent = New Collection

If Mode = vbNullString Then Mode = "SINGLE"
Select Case Mode
    Case "SINGLE"
    Parent.Add Collection_Nastrojki_Edna_Kolona(CStr("#" & KeyText), CStr("#" & KeyText & "End"), "LastCol", 2, NastrSheet, StartRowCol.Item(KeyText).Val, NameCol, ValCol), KeyText
    Case "MULTI"
    Parent.Add Collection_Nastrojki_Mnogo_Koloni(CStr("#" & KeyText), CStr("#" & KeyText & "End"), "LastCol", 2, NastrSheet, 1, StartRowCol.Item(KeyText).Val, NameCol), KeyText
End Select
End Sub

Public Sub Inicial_Main()
On Error GoTo ErrHandler
    Dim StartRowCol As Collection
Call InicialWS(ERPSheet, "Tovarene")
Call InicialWS(KOMSkladSheet, "KomSklad")
Call InicialWS(KOMPorchSheet, "KomPorych")
Call InicialWS(ArtLoadSheet, "Obshto")
Call InicialWS(MatrLoadSheet, "Matraci")
Call InicialWS(NastrSheet, "Nastrojki")

Set StartRowCol = Collection_Nastrojki_Edna_Kolona("#Content", "#ContentEnd", "LastCol", 2, NastrSheet, 1)
Set Nastr = New Collection
Set DS = New Collection

Nastr.Add Collection_Nastrojki_Edna_Kolona("#ERPMark", "#ERPMarkEnd", "LastCol", 2, NastrSheet, StartRowCol.Item("ERPMark").Val), "ERPMark"
Nastr.Add Collection_Nastrojki_Edna_Kolona("#Datasources", "#DatasourcesEnd", "LastCol", 2, NastrSheet, StartRowCol.Item("Datasources").Val), "Datasources"
Nastr.Add Collection_Nastrojki_Edna_Kolona("#KomMark", "#KomMarkEnd", "LastCol", 2, NastrSheet, StartRowCol.Item("KomMark").Val), "KomMark"
Nastr.Add Collection_Nastrojki_Edna_Kolona("#ArtLoadMark", "#ArtLoadMarkEnd", "LastCol", 2, NastrSheet, StartRowCol.Item("ArtLoadMark").Val), "ArtLoadMark"

IsInitialized = True
Exit Sub

ErrHandler:
Call EmergencyExit("Module Inicial_Main")
End Sub

Private Sub InicialWS(ByRef WSObject As Worksheet, ByVal WSName As String)
    On Error GoTo ErrHandler
    Set WSObject = ThisWorkbook.Worksheets(WSName): Exit Sub
ErrHandler:
    MsgBox "Програмата не може да открие лист с име '" & WSName & "'. Вероятно той е бил изтрит или преименуван." & _
            "Моля, възстановете го или ползвайте последната работеща версия на програмата": End
End Sub

            Private Function Collection_Nastrojki_Edna_Kolona(ByRef NachZapis As String, ByRef KrajZapis As String, ByRef LastKolZapis As String, _
                                                            IgnoreRows As Integer, SourceSheet As Worksheet, Optional StartRow As Long, _
                                                            Optional ByVal NameCol As Long = 1, Optional ByVal ValCol As Long = 3) As Collection
            ' For Single column settings
            Dim col As New Collection
            Dim Data As Variant
            Dim i As Integer
            Dim Tek_Atr As CBean
            If StartRow = 0 Then StartRow = 1
            Data = Data_Nastrojki(NachZapis, KrajZapis, LastKolZapis, IgnoreRows, SourceSheet, StartRow)
                For i = 1 To UBound(Data)
                    Set Tek_Atr = New CBean
                    Tek_Atr.Prop = Data(i, NameCol)
                    Tek_Atr.Val = Data(i, ValCol)
                    On Error Resume Next
                    col.Add Tek_Atr, Tek_Atr.Prop
                    Set Tek_Atr = Nothing
                Next i
            Set Data = Nothing
            Set Collection_Nastrojki_Edna_Kolona = col
            Set col = Nothing
            End Function
            
            Private Function Collection_Nastrojki_Mnogo_Koloni(ByRef NachZapis As String, ByRef KrajZapis As String, ByRef LastKolZapis As String, IgnoreRows As Integer, SourceSheet As Worksheet, KeyRow As Integer, Optional StartRow As Long, Optional ByVal NameCol As Long = 1) As Collection
            Dim col As New Collection
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
                    Tek_Atr.Prop = Data(i, NameCol)
                    Tek_Atr.Val = Data(i, j)
                    On Error Resume Next
                    ColonaNastr.Add Tek_Atr, Tek_Atr.Prop
                    Set Tek_Atr = Nothing
                Next i
            
            If col Is Nothing Then Set col = New Collection
            col.Add ColonaNastr, CStr(Data(KeyRow, j))
            Set ColonaNastr = Nothing
            Next j
            
            Set Data = Nothing
            Set Collection_Nastrojki_Mnogo_Koloni = col
            Set col = Nothing
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
            MsgBox ("Не откривам необходимата маркировка за таблица " & NachZapis & vbCrLf & _
                    "Откритите параметри:" & vbCrLf & _
                    "Първия ред: " & FirstRow & " - Ключ:" & NachZapis & vbCrLf & _
                    "Последната колона: " & LastCol & " - Ключ:" & LastKolZapis & vbCrLf & _
                    "Последния ред: " & LastRow & " - Ключ:" & KrajZapis)
            End
End Function
