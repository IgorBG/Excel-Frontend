Attribute VB_Name = "WSWorker"
Public Function CheckTemplate() As String
If Not IsInitialized Then Call Inicial_Main
 If Not Len(PrintTemplateString) = 0 Then CheckTemplate = PrintTemplateString: Exit Function
 CheckTemplate = OSheet.Range(Nastr.Item("OrdMark").Item("TmpltCell").Stojnost).Value
 If Len(CheckTemplate) = 0 Then Call EmergencyExit("Не е избран шаблон за печат. Моля, първо изберете един от достъпните шаблони за печат на страницата с поръчките")
End Function

Public Sub PutTemplateInfo(ByVal TemplateStr As String)
If Not IsInitialized Then Call Inicial_Main
    PrintTemplateString = TemplateStr
    OSheet.Range(Nastr.Item("OrdMark").Item("TmpltCell").Stojnost).Value = TemplateStr
End Sub

Public Sub ClearPages()
    Dim i As Long, LastRow As Long
    Dim Rng As Range
On Error GoTo ErrHandler
If Not IsInitialized Then Call Inicial_Main
Call Optimization_ON
    i = Nastr.Item("OrdMark").Item("StartRow").Stojnost 'Изчиства вписаната информация в лист Porychki
    LastRow = 65536
    OSheet.Rows(i & ":" & LastRow).ClearContents
    OSheet.Rows(i & ":" & LastRow).ClearFormats
    i = Nastr.Item("PakMark").Item("StartRow").Stojnost 'Изчиства вписаната информация в лист Paketi
    LastRow = 65536
    PakSheet.Rows(i & ":" & LastRow).ClearContents
    PakSheet.Rows(i & ":" & LastRow).ClearFormats
    PrintSheet.Cells.ClearContents 'Изчиства вписаната информация в шаблона на етикета
        Dim shp As Shape
        For Each shp In PrintSheet.Shapes
         If shp.Name Like "*Picture*" Then shp.Delete
        Next shp
Call PutTemplateInfo(vbNullString)
Call Optimization_OFF

Exit Sub


ErrHandler:
Call EmergencyExit("Програмата не успя да изчисти съдържанието на страниците")
End Sub


Public Sub ApproveOrderPaketazhFromMenu()
If Not IsInitialized Then Call Inicial_Main
    Dim OrdersCol As New Collection
    Dim cell As Range
    Dim TmpStr As String
'Проверява маркираното от потребителя. Ако е маркирал няколко поръчки наведнъж го спира (Щом е забранено пакетното потвърждаване на поръчките).
For Each cell In Selection.Rows
    TmpStr = PakSheet.Cells(cell.Row, Nastr.Item("PakMark").Item("OrdNumCol").Stojnost).Value
    If Not ValueIsInCollection(OrdersCol, TmpStr) Then OrdersCol.Add TmpStr, TmpStr
    If OrdersCol.Count > 1 Then
        Call EmergencyExit("Маркирали сте повече от един номер поръчка. Поръчките трябва да се потвърждават една по една.")
    End If
Next cell

For Each cell In Selection.Rows
    PakSheet.Cells(cell.Row, Nastr.Item("PakMark").Item("PktzhCol").Stojnost).Font.ColorIndex = 1
    PakSheet.Cells(cell.Row, Nastr.Item("PakMark").Item("ApprvdPzh").Stojnost).Value = "TRUE"
Next cell
Exit Sub

Err_Handler:
    Call EmergencyExit("Избранният пакетаж не може да се потвърди. Ако нищо не помага, ще се наложи той да бъде изтрит от списъка.")

End Sub

Public Sub ExportCSV()
    Dim sOutput As String
    Dim fName As String, lFnum As Long
    Dim cPak As Collection
    Dim LocalData As Variant
    Dim i As Long, j As Long, b As Long
    Dim LastRow As Long
    Const DELIMITER = ";"
    Const FILEFLTR = "CSV Files (*.csv), *.csv"
If Not IsInitialized Then Call Inicial_Main
b = 1 'Base for MS Office
On Error GoTo ErrHandler
Set cPak = Nastr.Item("PakMark")
LastRow = PakSheet.Cells(Rows.Count, cPak("QntyEtikCol").Stojnost).End(xlUp).Row
If LastRow = cPak("StartRow").Stojnost Then GoTo EmptyListExit
LocalData = PakSheet.Range(PakSheet.Cells(cPak("StartRow").Stojnost, 1), PakSheet.Cells(LastRow, cPak("LastCol").Stojnost))

    lFnum = FreeFile
    fName = Application.GetSaveAsFilename(FileFilter:=FILEFLTR)
    If Not fName = "False" Then
        Open fName For Output As lFnum
            For i = LBound(LocalData, b) To UBound(LocalData, b)
                For j = LBound(LocalData, b + 1) To UBound(LocalData, b + 1)
                    sOutput = sOutput & CStr(LocalData(i, j)) & DELIMITER
                Next j
                sOutput = Left(sOutput, Len(sOutput) - Len(DELIMITER))
                Print #lFnum, sOutput
                sOutput = vbNullString
            Next i
        Close lFnum
    End If
Exit Sub

EmptyListExit:
Call EmergencyExit("Програмата не вижда да има данни на лист Paketi. Моля запазете ги ръчно. Ако проблемът се повтаря - докладвайте за проблема.")
Exit Sub
ErrHandler:
Call EmergencyExit("Проблем при запазване на данните в CSV файл. Моля запазете данните ръчно. Ако проблемът се повтаря - докладвайте за проблема.")
Exit Sub
End Sub


Public Sub DeleteUserMenu()
    Dim CntxMenu As CommandBar
    Dim UserMenu As CommandBarControl
Set CntxMenu = Application.CommandBars("Cell")
For Each UserMenu In CntxMenu.Controls
    If UserMenu.Tag = "AddedByUser" Then UserMenu.Delete    ' Delete the menu added by user.
Next UserMenu
End Sub

Function GetPartsArray() As Variant ' Collect parts from WS and convert it into 2D-array
If Not IsInitialized Then Inicial_Main
Dim LastRow As Long
LastRow = SpecSheet.Cells(Rows.Count, ERPNastrCol("Det_kr_ime").Stojnost).End(xlUp).Row
Select Case LastRow
    Case 0, 1, 65536
        Set GetPartsArray = Nothing: Exit Function
    Case Else
        GetPartsArray = SpecSheet.Range(SpecSheet.Cells(ERPNastrCol("StartRow").Stojnost, ERPNastrCol("FirstCol").Stojnost), SpecSheet.Cells(LastRow, ERPNastrCol("LastCol").Stojnost))
        Exit Function
End Select
End Function

Public Sub PrintArray(v As Variant, Rng As Range, Optional Transpose As Boolean, Optional addBase As Long)
    If Not IsArray(v) Then Exit Sub
        Select Case Transpose
            Case True: Rng.Resize(UBound(v, 2) + addBase, UBound(v, 1) + addBase) = WorksheetFunction.Transpose(v)
            Case False: Rng.Resize(UBound(v, 1) + addBase, UBound(v, 1) + addBase) = v
        End Select
End Sub

