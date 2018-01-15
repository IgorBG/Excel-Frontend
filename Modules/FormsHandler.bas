Attribute VB_Name = "FormsHandler"
Option Explicit

Private Sub PopulateListControlWithArray(ByVal DataIn As Variant, ByRef ListControl As Object)
If Not IsEmpty(DataIn) Then
    ListControl.Clear
    ListControl.ColumnCount = UBound(DataIn, 2) - LBound(DataIn, 2) + 1
    ListControl.List = DataIn
    ListControl.SetFocus
End If
End Sub
Private Sub PopulateListControlFromRequest(ByVal DBPath, SQLQuery As String, ByRef ListControl As Object, Optional Parameters As Variant, Optional queryType As Variant = adCmdText)
    Dim LocalData As Variant
    LocalData = TransposeArray(GetRSData(GetNewConnToAccess(DBPath), SQLQuery, True, Parameters, queryType))
    If IsArray(LocalData) Then Call PopulateListControlWithArray(LocalData, ListControl)
End Sub
Public Sub LoadPictureToForm(ByRef ImageCntrl As Image, FilePath As String)
    ImageCntrl.PictureSizeMode = fmPictureSizeModeZoom
    ImageCntrl.Picture = LoadPicture(FilePath)
End Sub
    Function GetStringResultViaForm(ByVal Context As String) As String
        Dim TekForm As Object
    If Not IsInitialized Then Call Inicial_Main
    Set TekForm = GetFormByContext(Context)
    TekForm.Show
    On Error Resume Next
    If IsEmpty(TekForm.Results) Then Exit Function
    If TekForm.Results = vbNullString Then Exit Function
    GetStringResultViaForm = TekForm.Results
    End Function

    Function GetResultCollectionViaForm(ByVal Context As String) As Collection
        Dim TekForm As Object
    If Not IsInitialized Then Call Inicial_Main
    Set TekForm = GetFormByContext(Context)
    TekForm.Show
    On Error Resume Next
    If TekForm.ResultsCollection Is Nothing Then Exit Function
    Set GetResultCollectionViaForm = TekForm.ResultsCollection
    End Function
    

Public Function GetFormByContext(Context As String) As Object
    Dim TempForm As Object
    Dim LocalData As Variant
    Dim TempCol As Collection
    Dim i As Long
    Dim TempStr1 As String, TempStr2 As String
    Dim ProdDirID As Long
    Dim strObekt As String, strGrad As String
    Dim DSCol As Collection
    Dim FormWidth As Double, Caption As String, HeadText As String, ColumnWidths As String
    Dim DBPath As String, Query As String, queryType As Variant

    Dim CertainClientMode As Boolean
    
If Not IsInitialized Then Call Inicial_Main
Set DSCol = Nastr.Item("Datasources")
Select Case Context
    
    Case "SearchTruck"
        FormWidth = 400
        Caption = "Списък МПС"
        HeadText = "Изберете МПС от списъка (рег.ном, модел, капацитет)"
        ColumnWidths = "0;80;200;64"
        DBPath = DSCol("DB_Transport_Path").Val
        Query = "searchTrucksByKeyword"
        queryType = adCmdStoredProc
    
    Case "SearchDriver"
        FormWidth = 350
        Caption = "Списък шофьори"
        HeadText = "Изберете шофьор от списъка"
        ColumnWidths = "180;100"
        DBPath = DSCol("DB_Transport_Path").Val
        Query = "searchDriversByPartOfName"
        queryType = adCmdStoredProc
    
    Case "SearchLine"
        FormWidth = 420
        Caption = "Списък запазени маршрути"
        HeadText = "Изберете шаблонен маршрут от списъка"
        ColumnWidths = "80;320"
        DBPath = DSCol("DB_Transport_Path").Val
        Query = "searchTransportLineByKeyword"
        queryType = adCmdStoredProc
    
    Case "ListRedirectedClients"
        FormWidth = 350
        Caption = "Списък търговски обекти"
        HeadText = "Маркираното посещение включва разтоварване на следните търговски обекти:"
        ColumnWidths = "180;120"
    
    Case "ViewUnrecognizedArticles"
        FormWidth = 450
        Caption = "Списък неразпознати изделия"
        HeadText = "Долу са изредени изделия, за които програмата не разполага с данни за габарити, тегло, друго"
        ColumnWidths = "24;400"
    
    Case "ViewMekaArticles"
        FormWidth = 450
        Caption = "Списък изделия от Мека Мебел"
        HeadText = "Долу са изредени изделия от категория мека мебел за избрания контрагент"
        ColumnWidths = "24;400"
End Select




Select Case Context
    Case "SearchLine", "SearchTruck", "SearchDriver"
        Set TempForm = New SearchForm
        Call TempForm.Constructor(FormWidth, , Caption, HeadText, , Context)
            With TempForm
                .LBResultList.ColumnCount = UBound(Split(ColumnWidths, ";")) - LBound(Split(ColumnWidths, ";")) + 1: .LBResultList.ColumnWidths = ColumnWidths
                .LBResultList.Font.Size = 12
                .SearchQuery = Query: .queryType = queryType
                Call PopulateListControlWithArray(TransposeArray(GetRSData(GetNewConnToAccess(DBPath), Query, True, TempForm.TBSearchText.Value, queryType)), TempForm.LBResultList)
                .TBSearchText.Font.Size = 12: .TBSearchText.SetFocus
            End With
    
    Case "ViewUnrecognizedArticles"
        If RowClicked >= Nastr("RazpredKlientiMark").Item("StartRow").Val Then CertainClientMode = True
        Set TempForm = New SearchForm
        Call TempForm.Constructor(FormWidth, , Caption, HeadText, , Context)
            With TempForm
            .TBSearchText.Visible = False:  .LBResultList.ColumnCount = 2:   .LBResultList.ColumnWidths = ColumnWidths
            LocalData = getPresetData("OrderList")  'GetOData
                If Not IsEmpty(LocalData) Then
                    TempForm.LBResultList.Clear
                    If CertainClientMode Then
                        strObekt = RAZPSheet.Cells(RowClicked, Nastr("RazpredKlientiMark").Item("ObektCol").Val).Value
                        strGrad = RAZPSheet.Cells(RowClicked, Nastr("RazpredKlientiMark").Item("GradCol").Val).Value
                    End If
                    For i = LBound(LocalData, 1) To UBound(LocalData, 1)
                        ProdDirID = GetProdDir(LocalData(i, Nastr("ERPMark").Item("SkladCol").Val))
                        If Not ProdDirID = 0 Then
                        If LocalData(i, Nastr("ERPMark").Item("PktzhCol").Val) = vbNullString Then
                        If CertainClientMode = False Or CStr(LocalData(i, Nastr("ERPMark").Item("RaztObekt").Val)) = strObekt Then
                        If CertainClientMode = False Or CStr(LocalData(i, Nastr("ERPMark").Item("RaztGrad").Val)) = strGrad Then
                        Select Case ProdDirID
                            Case 1, 2, 3, 4, 5, 6
                            If LocalData(i, Nastr("ERPMark").Item("ObshtCenaCol").Val) > 0 Then
                                TempForm.LBResultList.addItem CStr(LocalData(i, Nastr("ERPMark").Item("BrZaqvenoCol").Val))
                                TempForm.LBResultList.List(TempForm.LBResultList.ListCount - 1, 1) = CStr(LocalData(i, Nastr("ERPMark").Item("NomImeCol").Val))
                            End If
                            Case 7, 9, 10
                                TempForm.LBResultList.addItem CStr(LocalData(i, Nastr("ERPMark").Item("BrZaqvenoCol").Val))
                                TempForm.LBResultList.List(TempForm.LBResultList.ListCount - 1, 1) = CStr(LocalData(i, Nastr("ERPMark").Item("NomImeCol").Val))
                        End Select
                        End If
                        End If
                        End If
                        End If
                    Next i
                    TempForm.LBResultList.SetFocus
                End If
            End With
    Case "ViewMekaArticles"
        Set TempForm = New SearchForm
        Call TempForm.Constructor(FormWidth, , Caption, HeadText, , Context)
            With TempForm
            .TBSearchText.Visible = False:  .LBResultList.ColumnCount = 2:   .LBResultList.ColumnWidths = ColumnWidths
            LocalData = getPresetData("OrderList")  'GetOData
                If Not IsEmpty(LocalData) Then
                    TempForm.LBResultList.Clear
                    strObekt = RAZPSheet.Cells(RowClicked, Nastr("RazpredKlientiMark").Item("ObektCol").Val).Value
                    strGrad = RAZPSheet.Cells(RowClicked, Nastr("RazpredKlientiMark").Item("GradCol").Val).Value
                    For i = LBound(LocalData, 1) To UBound(LocalData, 1)
                        ProdDirID = GetProdDir(LocalData(i, Nastr("ERPMark").Item("SkladCol").Val))
                        If ProdDirID = 6 Then
                        If CStr(LocalData(i, Nastr("ERPMark").Item("RaztObekt").Val)) = strObekt Then
                        If CStr(LocalData(i, Nastr("ERPMark").Item("RaztGrad").Val)) = strGrad Then
                        If Not Left(CStr(LocalData(i, Nastr("ERPMark").Item("NomImeCol").Val)), 3) = "###" Then
                                TempForm.LBResultList.addItem CStr(LocalData(i, Nastr("ERPMark").Item("BrZaqvenoCol").Val))
                                TempForm.LBResultList.List(TempForm.LBResultList.ListCount - 1, 1) = CStr(LocalData(i, Nastr("ERPMark").Item("NomImeCol").Val))
                        End If
                        End If
                        End If
                        End If
                    Next i
                    TempForm.LBResultList.SetFocus
                End If
            End With
    Case "ListRedirectedClients"
        Set TempForm = New SearchForm
        Call TempForm.Constructor(FormWidth, , , , , Context)
            With TempForm
            .Caption = Caption: .HeadText.Caption = HeadText
            .TBSearchText.Visible = False:  .LBResultList.ColumnCount = 2:   .LBResultList.ColumnWidths = ColumnWidths: .LBResultList.Font.Size = 12
            LocalData = getPresetData("OrderList")  'GetOData
                If Not IsEmpty(LocalData) Then
                    TempForm.LBResultList.Clear
                    Set TempCol = New Collection
                    strObekt = RAZPSheet.Cells(RowClicked, Nastr("RazpredKlientiMark").Item("ObektCol").Val).Value
                    strGrad = RAZPSheet.Cells(RowClicked, Nastr("RazpredKlientiMark").Item("GradCol").Val).Value
                    For i = LBound(LocalData, 1) To UBound(LocalData, 1)
                        If CStr(LocalData(i, Nastr("ERPMark").Item("RaztObekt").Val)) = strObekt Then
                        If CStr(LocalData(i, Nastr("ERPMark").Item("RaztGrad").Val)) = strGrad Then
                        TempStr1 = LocalData(i, Nastr("ERPMark").Item("KlientCol").Val)
                        TempStr2 = LocalData(i, Nastr("ERPMark").Item("GradCol").Val)
                        If Not ValueIsInCollection(TempCol, TempStr1 & " / " & TempStr2) Then
                                TempForm.LBResultList.addItem TempStr1
                                TempForm.LBResultList.List(TempForm.LBResultList.ListCount - 1, 1) = TempStr2
                                TempCol.Add TempStr1 & " / " & TempStr2, TempStr1 & " / " & TempStr2
                        End If
                        End If
                        End If
                    Next i
                    TempForm.LBResultList.SetFocus
                End If
            End With
    
    Case Else: Call EmergencyExit("Непознато действие във прозореца")
End Select
Set GetFormByContext = TempForm
End Function


Public Sub ProcessFormEvent(ByRef TekForm As Object, Optional ByRef EventMsg As String, Optional Context As String, Optional arg1 As Variant)
    Dim myFile As Variant, targetFldr As String, SourcePath As String
    Dim SQLstr As String
    Dim LocalData As Variant
    Dim i As Long
    Dim InnerForm As Object
    Dim SuccessBool As Boolean
If Not IsInitialized Then Call Inicial_Main
If EventMsg = vbNullString Then Exit Sub
Select Case Context
    Case "SearchTruck", "SearchDriver", "SearchLine"
        SourcePath = Nastr.Item("Datasources").Item("DB_Transport_Path").Val
End Select

Select Case Context
    Case "SearchTruck", "SearchDriver", "SearchLine", "ViewUnrecognizedArticles", "ViewMekaArticles", "ListRedirectedClients"
        Select Case EventMsg
            Case "Ok"
                TekForm.Hide
                If TekForm.LBResultList.ListCount = 0 Then Exit Sub
                For i = LBound(TekForm.LBResultList.List) To UBound(TekForm.LBResultList.List)
                    If TekForm.LBResultList.Selected(i) Then LocalData = LocalData & TekForm.LBResultList.List(i) & ",": Exit For
                Next i
                If IsEmpty(LocalData) Then Exit Sub
                LocalData = Left(LocalData, Len(LocalData) - 1)
                TekForm.Results = LocalData
            
            Case "Cancel"
                TekForm.Hide
                TekForm.Results = vbNullString
            
            Case "tbxSearch_Update"
                Call PopulateListControlWithArray(TransposeArray(GetRSData(GetNewConnToAccess(SourcePath), TekForm.SearchQuery, True, TekForm.TBSearchText.Value, TekForm.queryType)), TekForm.LBResultList)

            Case Else: Call EmergencyExit("Непознато действие във прозореца")
        End Select
End Select
ByPass:
Exit Sub

ErrHandler:
Call EmergencyExit("Неуспешен опит при създаване на потребителски прозорец")
End Sub
