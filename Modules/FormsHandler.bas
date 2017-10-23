Attribute VB_Name = "FormsHandler"
Option Explicit
Public Sub ChoozeObektClient()
'Прозорец за избор на търговския обект на клиента
    Dim SF As Object
    Dim i As Long
    Dim ERPName As String
If Not IsInitialized Then Call Inicial_Main
    
    ERPName = OSheet.Cells(RowClicked, Nastr.Item("OrdMark").Item("ERPKlientCol").Stojnost).Value
On Error GoTo ErrHandler
Set SF = New SearchForm
    Call SF.Constructor(250, 200, "Избор търговски обект", "Изберете търговския обект на контрагент " & ERPName, 2)
        SF.TBSearchText.Visible = False
        SF.HeadText.Height = SF.HeadText.Height + SF.TBSearchText.Height
        SF.LBResultList.ColumnWidths = "120;100"
        SF.btnOption1.Visible = True: SF.btnOption1.Caption = "New"
BackToForm:
        If SF.LBResultList.ListCount = 0 Then Call SF.PopulateListBoxWithFromAccessDB(SF.LBResultList, Nastr.Item("PaketiDB").Item("PaketiDBPath").Stojnost, "SELECT ObektIme, ObektGrad FROM KlientiObekti WHERE KlientERPIme = '" & ERPName & "' AND Aktiven =TRUE ORDER BY ObektGrad;", SF.LBResultList.ColumnCount)
        SF.LBResultList.SetFocus
        SF.Show
            If IsError(SF.EventLog) Then GoTo EmptyExit
            Select Case SF.EventLog
                Case "OK"
                    If SF.LBResultList.ListCount = 0 Then GoTo EmptyExit
                        For i = LBound(SF.LBResultList.List) To UBound(SF.LBResultList.List)
                            If SF.LBResultList.Selected(i) Then
                                OSheet.Cells(RowClicked, Nastr.Item("OrdMark").Item("EtiketKlntCol").Stojnost).Value = SF.LBResultList.List(SF.LBResultList.ListIndex)
                                OSheet.Cells(RowClicked, Nastr.Item("OrdMark").Item("GradCol").Stojnost).Value = SF.LBResultList.List(SF.LBResultList.ListIndex, 1)
                                Exit For
                            End If
                        Next i
                Case "OB1_Click"
                    SF.Hide
                    NewObektGradForm.Show  'TODO: Call macro that shows new client/object input
                        If Not SF.LBResultList.ListCount = 0 Then SF.LBResultList.Clear
                    GoTo BackToForm
            End Select
        Set SF = Nothing
Exit Sub
EmptyExit:
    Set SF = Nothing
   Call UserExit: Exit Sub
ErrHandler:
    Set SF = Nothing
    Call EmergencyExit("Грешка при опит за избор на търговския обект от прозореца")
End Sub

Private Sub InputArticleWindow()
Dim InputForm As StdInputForm
On Error GoTo Err_Handler
Set oSpec = New CSpec
    Set InputForm = New StdInputForm
    With InputForm
        .btnOB1Main.Visible = False
        .lblTxtInput1Main.Visible = True: .lblTxtInput1Main.Caption = "Име на спецификация *"
        .txtInput1Main.Value = ERP_ImeIzdelie(SpecSheet.Cells(ERPNastrCol.Item("StartRow").Stojnost, ERPNastrCol.Item("Det_pyl_ime").Stojnost).Value)
        .lblTxtInput2Main.Visible = True: .lblTxtInput2Main.Caption = "Кратко пояснение"
        .txtInput2Main.Visible = True
        .lblTxtInput3Main.Visible = True: .lblTxtInput3Main.Caption = "Отнася се към номенклатури:"
        .txtInput3Main.Visible = True
        .check1Main.Caption = "Краен артикул": .check1Main.Visible = True
        .Show 'Opens the pop-up window. Declaring the names and other details on the specification
    End With
        Select Case InputForm.EventLog
            Case "OK"
                oSpec.Name = InputForm.txtInput1Main.Value
                oSpec.Descr = InputForm.txtInput2Main.Value
                oSpec.isCompleteArt = InputForm.check1Main.Value
            Case "Cancel"
                Call UserExit
        End Select
Exit Sub
Err_Handler:
    Call EmergencyExit("Проблем при обработка на прозореца за въвеждане")
Exit Sub
ForceExit:
    Call UserExit
    End
End Sub
