Option Explicit

' Эта рабочая книга
Public ControlWB    As Workbook
Public ControlWS    As Worksheet

' Имена, котрорые заданы в документе в "Диспетчере имен". Так проще работать с диапазонами
' плюс ссылки на основные страницы на zakupki.gov.ru
Const SUPPLIER_RESULTS As String = "https://zakupki.gov.ru/epz/order/notice/ea44/view/supplier-results.html?regNumber="
Const CONTROL_NAME  As String = ""
Const ZAK_COMMON    As String = "https://zakupki.gov.ru/epz/order/notice/ea44/view/common-info.html?regNumber="
Const CONTRACT_COMMON As String = "https://zakupki.gov.ru/epz/order/notice/rpec/common-info.html?regNumber="
Const CONTRACT_EVENT_JOURNAL As String = "https://zakupki.gov.ru/epz/order/notice/rpec/event-journal.html?regNumber="
Public Const SIGNING As String = "Подписание"
Public Const CONCLUDED As String = "Заключен"
Public Const RNG_NUMBER_NAME As String = "Номер"
Public Const RNG_STATUS As String = "Статус"
Public Const RNG_PROJECT_NAME As String = "Проект"
Public Const RNG_REVISION_NAME As String = "Доработ"
Public Const RNG_DISAGREE_NAME As String = "ПротоколРазногласий"
Public Const RNG_RESULT_PROT_NAME As String = "ПротоколИтогов"
Public Const RNG_CUSTOMER_INN As String = "ЗаказчикИНН"
Public Const RNG_CUSTOMER_NAME As String = "ЗаказчикНаим"
Public Const RNG_PROC_NUM_NAME As String = "НомерПроцедурыЗак"
Public Const RNG_CONTRACT_NUM_NAME As String = "РеестровыйНомер"
Public Const RNG_CONTRACT_OFF_NYM_NAME As String = "Офиц_Номер"
Public Const RNG_CONTRACT_DATE_NAME As String = "Дата_Заключения"
Public Const RNG_NOTE As String = "Примечание"
Public Const RNG_ORG As String = "Организация"
Public Const RNG_REGION As String = "region"
Public Const RNG_SID As String = "SID"
Public Const RNG_SHORT_ABOUT As String = "Предмет"
Public Const RNG_LVL_NOTIF_NAME As String = "Уровень_оповещения"

' глобальный флаг для работы с автофильтром документа
' используется для запуска общей процедуры проверки,
' чтобы постоянно не дергать ручку автофильтра в каждой отдельной процедуре (это долго)
Private FILTER_MODE_OFF As Boolean

' Пути где лежат реестры
Const REGISTRY_NAME As String = ""
Const EDIN_REGISTRY_NAME As String = ""
Const REGISRTY_PATH As String = "" + REGISTRY_NAME
Const EDIN_REGISRTY_PATH As String = "" + EDIN_REGISTRY_NAME
Sub Общая_Проверка()
    Dim cw_is_set   As Boolean
    
    FILTER_MODE_OFF = TRUE
    
    Call toggle_screen_upd
    
    cw_is_set = set_control_workbook(ControlWB, ControlWS)
    
    With ControlWB.Worksheets(ControlWS.Name)
        
        .AutoFilterMode = FALSE
        .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=SIGNING
        
        .Range("A1").AutoFilter Field:=.Range(RNG_RESULT_PROT_NAME)(1).Column, Criteria1:="="
        
        Call Протоколы_Итогов22
        
        .Range("A1").AutoFilter Field:=.Range(RNG_RESULT_PROT_NAME)(1).Column
        
        Call Проверка_Контрактов
        
        .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=CONCLUDED
        .Range("A1").AutoFilter Field:=.Range(RNG_CONTRACT_OFF_NYM_NAME)(1).Column, Criteria1:="="
        
        Call Номера_Контрактов
        
        .Range("A1").AutoFilter Field:=.Range(RNG_CONTRACT_OFF_NYM_NAME)(1).Column
        
        .Range("A1").AutoFilter Field:=.Range(RNG_CONTRACT_DATE_NAME)(1).Column, Criteria1:="="
        
        Call Подписание_Заказчиком
        
        .Range("A1").AutoFilter Field:=.Range(RNG_CONTRACT_DATE_NAME)(1).Column
        .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=SIGNING
        
    End With
    
    Call toggle_screen_upd
    FILTER_MODE_OFF = FALSE
End Sub
Private Sub Проверка_Контрактов()
    Dim http        As New MSXML2.XMLHTTP60
    Dim doc         As HTMLDocument
    Dim isSet       As Boolean, match As Boolean, SomeSet As Boolean, cw_is_set As Boolean
    Dim lastrow     As Integer, firstRow As Integer, suffixNumber As Integer, i As Integer, v As Long
    Dim contractNumber As String, zakNumber As String, serchS As String, resp As String, Text As String
    Dim span        As IHTMLElementCollection
    Dim elements    As IHTMLDOMChildrenCollection
    Dim dict        As New Scripting.Dictionary
    Dim visible_range As Range, num As Range, total_rows As Integer, counter As Integer
    Const label     As String = "Контроль Проверка Проектов Контрактов "
    
    cw_is_set = set_control_workbook(ControlWB, ControlWS)
    
    With ControlWB.Worksheets(ControlWS.Name)
        
        If FILTER_MODE_OFF = FALSE Then
            Call toggle_screen_upd
            .AutoFilterMode = FALSE
            .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=SIGNING
        End If
        
        lastrow = .Range(RNG_NUMBER_NAME).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
        
        If lastrow = 1 Then
            If FILTER_MODE_OFF = FALSE Then Call toggle_screen_upd
            Exit Sub
        End If
        
        Set visible_range = .Range(.Range(RNG_NUMBER_NAME)(2).Address & ":" & .Range(RNG_NUMBER_NAME)(lastrow).Address).SpecialCells(xlCellTypeVisible)
        total_rows = visible_range.Count
        counter = 1
        
        For Each num In visible_range
            
            i = num.Row
            
            If IsEmpty(.Range(RNG_PROJECT_NAME)(i)) Or IsEmpty(.Range(RNG_REVISION_NAME)(i)) Or IsEmpty(.Range(RNG_DISAGREE_NAME)(i)) Then
                match = FALSE
                zakNumber = Trim(Replace(.Range(RNG_NUMBER_NAME)(i).Value, "№", ""))
                suffixNumber = 1
                If IsEmpty(.Range(RNG_PROC_NUM_NAME)(i).Value) Then
                    contractNumber = zakNumber & "000" & suffixNumber
                    Do While dict.Exists(contractNumber) = TRUE
                        If suffixNumber < 10 Then
                            contractNumber = zakNumber & "000" & suffixNumber
                        Else
                            contractNumber = zakNumber & "00" & suffixNumber
                        End If
                        suffixNumber = suffixNumber + 1
                    Loop
                Else
                    contractNumber = .Range(RNG_PROC_NUM_NAME)(i).Value
                End If
                If Not IsEmpty(.Range(RNG_PROJECT_NAME)(i).Value) And IsEmpty(.Range(RNG_PROC_NUM_NAME)(i).Value) And Not IsEmpty(.Range(RNG_CUSTOMER_INN)(i).Value) Then
                    suffixNumber = 1
                    Do While match = FALSE
                        If suffixNumber < 10 Then
                            contractNumber = zakNumber & "000" & suffixNumber
                        Else
                            contractNumber = zakNumber & "00" & suffixNumber
                        End If
                        Set doc = HTMLDoc(CONTRACT_COMMON & contractNumber)
                        http.send
                        Set span = doc.getElementsByTagName("span")
                        isSet = FALSE
                        For v = span.Length - 1 To 0 Step -1
                            If isSet Then Exit For
                            If UCase(span(v).innerText) Like Trim(UCase("*Наименование*заказчика*")) Then
                                If .Range(RNG_CUSTOMER_INN)(i).Value = span(v + 5).innerText Then
                                    match = TRUE
                                End If
                                isSet = TRUE
                            End If
                        Next v
                        suffixNumber = suffixNumber + 1
                    Loop
                Else
                    match = TRUE
                End If
                
                If match Then
                    
                    If IsEmpty(.Range(RNG_SID)(i).Value) Then
                        resp = HttpGetEventJournalSid(CONTRACT_EVENT_JOURNAL & contractNumber)
                        If resp <> "HttpGet Error" Then .Range(RNG_SID)(i).Value = resp
                    End If
                    
                    Set doc = HTMLDoc("https://zakupki.gov.ru/epz/order/notice/card/event/journal/list.html?number=&sid=" & .Range(RNG_SID)(i).Value & "&page=1&pageSize=50&qualifier=rpecJournalEventService")
                    SomeSet = FALSE
                    Set elements = doc.querySelectorAll("table td:nth-child(2)")
                    
                    For v = 0 To elements.Length - 1
                        serchS = ""
                        Text = UCase(elements.Item(v).innerText)
                        
                        If Text Like "*РАЗМЕЩЕН ДОКУМЕНТ*«ПРОЕКТ КОНТРАКТА»*" Then
                            
                            If IsEmpty(.Range(RNG_PROJECT_NAME)(i)) Then
                                If Mid(text, InStr(1, text, "«ПРОЕКТ КОНТРАКТА» ОТ", vbTextCompare) + 22, 10) Like "##.##.####" Then
                                    serchS = Mid(text, InStr(1, text, "«ПРОЕКТ КОНТРАКТА» ОТ", vbTextCompare) + 22, 10)
                                Else
                                    serchS = Mid(text, InStr(1, text, "«ПРОЕКТ КОНТРАКТА» РЕД.", vbTextCompare) + 29, 10)
                                End If
                                .Range(RNG_PROJECT_NAME)(i).Value = CDate(serchS)
                            End If
                            
                            SomeSet = TRUE
                        End If
                        
                        If IsEmpty(.Range(RNG_DISAGREE_NAME)(i)) And Text Like "*ПОЛУЧЕН ДОКУМЕНТ*«ПРОТОКОЛ РАЗНОГЛАСИЙ»*" Then
                            .Range(RNG_DISAGREE_NAME)(i).Value = CDate(Mid(text, InStr(1, text, "«ПРОТОКОЛ РАЗНОГЛАСИЙ» ОТ", vbTextCompare) + 26, 10))
                        End If
                        
                        If IsEmpty(.Range(RNG_REVISION_NAME)(i)) And Text Like Trim(UCase("*РАЗМЕЩЕН ДОКУМЕНТ*«ДОРАБОТАННЫЙ ПРОЕКТ КОНТРАКТА»*")) Then
                            If Mid(text, InStr(1, text, "«ДОРАБОТАННЫЙ ПРОЕКТ КОНТРАКТА» ОТ", vbTextCompare) + 35, 10) Like "##.##.####" Then
                                serchS = Mid(text, InStr(1, text, "«ДОРАБОТАННЫЙ ПРОЕКТ КОНТРАКТА» ОТ", vbTextCompare) + 35, 10)
                            Else
                                serchS = Mid(text, InStr(1, text, "«ДОРАБОТАННЫЙ ПРОЕКТ КОНТРАКТА» РЕД.", vbTextCompare) + 42, 10)
                            End If
                            .Range(RNG_REVISION_NAME)(i).Value = CDate(serchS)
                        End If
                        
                    Next v
                    
                    If SomeSet Then .Range(RNG_PROC_NUM_NAME)(i).Value = contractNumber
                    
                    If Not IsEmpty(.Range(RNG_PROJECT_NAME)(i)) And .Range(RNG_CUSTOMER_INN)(i).Value = "" Then
                        
                        Set doc = HTMLDoc(CONTRACT_COMMON & contractNumber)
                        isSet = FALSE
                        Set span = doc.getElementsByTagName("span")
                        For v = span.Length - 1 To 0 Step -1
                            If isSet Then Exit For
                            
                            If UCase(span(v).innerText) Like Trim(UCase("*Наименование*заказчика*")) Then
                                .Range(RNG_CUSTOMER_NAME)(i).Value = Trim(span(v + 1).innerText)
                                .Range(RNG_CUSTOMER_INN)(i).Value = Trim(span(v + 5).innerText)
                                isSet = TRUE
                            End If
                            
                        Next v
                        
                    End If
                    
                End If
                If SomeSet Or Not IsEmpty(.Range(RNG_PROC_NUM_NAME)(i)) Then
                    If dict.Exists(contractNumber) Then
                        dict.Item(contractNumber) = TRUE
                    Else
                        dict.Add contractNumber, TRUE
                    End If
                End If
            Else
                dict.Add .Range(RNG_PROC_NUM_NAME)(i).Value, TRUE
            End If
            
            Call show_status(counter, total_rows, label)
            counter = counter + 1
            
        Next num
        
        Set doc = Nothing
        Set http = Nothing
        Set dict = Nothing
        
    End With
    
    If FILTER_MODE_OFF = FALSE Then Call toggle_screen_upd
    
End Sub
Private Sub Протоколы_Итогов()
    Dim doc         As HTMLDocument
    Dim lastrow     As Integer, i As Integer, v As Long, counter As Integer, total_rows As Integer
    Dim Text        As String
    Dim protocol_date As IHTMLElement, a As IHTMLElement
    Dim OrgIsSet    As Boolean, cw_is_set As Boolean
    Dim finalPrice  As Double, max_price As Double
    Dim elements    As IHTMLDOMChildrenCollection
    Dim visible_range As Range, num As Range
    Const label     As String = "Контроль Протоколы Итогов "
    
    cw_is_set = set_control_workbook(ControlWB, ControlWS)
    
    With ControlWB.Worksheets(ControlWS.Name)
        
        If FILTER_MODE_OFF = FALSE Then
            Call toggle_screen_upd
            .AutoFilterMode = FALSE
            .Range("A1").AutoFilter Field:=11, Criteria1:=SIGNING
            .Range("A1").AutoFilter Field:=.Range(RNG_RESULT_PROT_NAME)(1).Column, Criteria1:="="
        End If
        
        lastrow = .Range(RNG_NUMBER_NAME).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
        
        If lastrow = 1 Then
            If FILTER_MODE_OFF = FALSE Then
                .Range("A1").AutoFilter Field:=.Range(RNG_RESULT_PROT_NAME)(1).Column
                Call toggle_screen_upd
            End If
            Exit Sub
        End If
        
        Set visible_range = .Range(.Range(RNG_NUMBER_NAME)(2).Address & ":" & .Range(RNG_NUMBER_NAME)(lastrow).Address).SpecialCells(xlCellTypeVisible)
        total_rows = visible_range.Count
        counter = 1
        
        For Each num In visible_range
            i = num.Row
            
            Set doc = HTMLDoc(SUPPLIER_RESULTS & Replace(.Range(RNG_NUMBER_NAME)(i).Value, "№", ""))
            OrgIsSet = FALSE
            finalPrice = 0
            Set protocol_date = doc.querySelector("section.blockInfo__section section.blockInfo__section:last-child span:last-child")
            If Not protocol_date Is Nothing Then
                .Range(RNG_RESULT_PROT_NAME)(i).Value = CDate(Left(Trim(protocol_date.innerText), 10))
                Set elements = doc.querySelectorAll("td.tableBlock__col")
                
                For v = 0 To elements.Length - 1
                    
                    If OrgIsSet Then Exit For
                    Text = UCase(elements.Item(v).innerText)
                    
                    If is_winner(text) Then
                        .Range(RNG_ORG)(i).Value = ControlOrgStr(Trim(elements.Item(v - 1).innerText))
                        finalPrice = Val(Replace(elements.Item(v + 1).innerText, ",", "."))
                        OrgIsSet = TRUE
                    End If
                    
                    If is_sole_participant(text) Then
                        
                        .Range(RNG_ORG)(i).Value = ControlOrgStr(Trim(elements.Item(v - 1).innerText))
                        OrgIsSet = TRUE
                        finalPrice = Val(Replace(elements.Item(v + 1).innerText, ",", "."))
                        
                    End If
                    
                Next v
                
                If finalPrice <> 0 Then
                    max_price = Val(Replace(doc.querySelector(".cardMainInfo__content.cost").innerText, ",", "."))
                    
                    If IsLess25(max_price, finalPrice) Then
                        .Range("_25")(i).Value = ">"
                        .Range(RNG_NOTE)(i).Value = "Добросовестность!!!"
                        If IsOneAndHalfGuarantee(max_price) Then
                            .Range(RNG_NOTE)(i).Value = "Полуторная!!!"
                        End If
                    Else
                        .Range("_25")(i).Value = "-"
                    End If
                    
                ElseIf Not IsEmpty(.Range(RNG_RESULT_PROT_NAME)(i).Value) Then
                    .Range("_25")(i).Value = "-"
                End If
                
                If IsEmpty(.Range(RNG_SHORT_ABOUT)(i)) Then
                    .Range(RNG_SHORT_ABOUT)(i).Value = Trim(doc.querySelector(".cardMainInfo__content").innerText)
                End If
                
                If IsEmpty(.Range(RNG_REGION)(i)) Then
                    Set a = doc.querySelector(".cardMainInfo__content > a")
                    Set doc = HTMLDoc(a.href)
                    .Range(RNG_REGION)(i).Value = RegionStr(doc.querySelector("div.registry-entry__body-value").innerText)
                End If
                
            End If
            Call show_status(counter, total_rows, label)
            counter = counter + 1
        Next num
        
        If FILTER_MODE_OFF = FALSE Then
            .Range("A1").AutoFilter Field:=.Range(RNG_RESULT_PROT_NAME)(1).Column
            Call toggle_screen_upd
        End If
        
    End With
    
End Sub
Private Sub Протоколы_Итогов22()
    Dim doc         As HTMLDocument
    Dim lastrow     As Integer, i As Integer, v As Long, counter As Integer, total_rows As Integer
    Dim Text        As String
    Dim protocol_date As IHTMLElement, a As IHTMLElement
    Dim final_price_is_set As Boolean, cw_is_set As Boolean
    Dim finalPrice  As Double, max_price As Double
    Dim elements    As IHTMLDOMChildrenCollection
    Dim visible_range As Range, num As Range
    Const label     As String = "Контроль Протоколы Итогов "
    
    cw_is_set = set_control_workbook(ControlWB, ControlWS)
    
    With ControlWB.Worksheets(ControlWS.Name)
        
        If FILTER_MODE_OFF = FALSE Then
            Call toggle_screen_upd
            .AutoFilterMode = FALSE
            .Range("A1").AutoFilter Field:=11, Criteria1:=SIGNING
            .Range("A1").AutoFilter Field:=.Range(RNG_RESULT_PROT_NAME)(1).Column, Criteria1:="="
        End If
        
        lastrow = .Range(RNG_NUMBER_NAME).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
        
        If lastrow = 1 Then
            If FILTER_MODE_OFF = FALSE Then
                .Range("A1").AutoFilter Field:=.Range(RNG_RESULT_PROT_NAME)(1).Column
                Call toggle_screen_upd
            End If
            Exit Sub
        End If
        
        Set visible_range = .Range(.Range(RNG_NUMBER_NAME)(2).Address & ":" & .Range(RNG_NUMBER_NAME)(lastrow).Address).SpecialCells(xlCellTypeVisible)
        total_rows = visible_range.Count
        counter = 1
        
        For Each num In visible_range
            i = num.Row
            
            Set doc = HTMLDoc(SUPPLIER_RESULTS & Replace(.Range(RNG_NUMBER_NAME)(i).Value, "№", ""))
            final_price_is_set = FALSE
            finalPrice = 0
            Set protocol_date = doc.querySelector("section.blockInfo__section section.blockInfo__section:last-child span:last-child")
            If Not protocol_date Is Nothing Then
                .Range(RNG_RESULT_PROT_NAME)(i).Value = CDate(Left(Trim(protocol_date.innerText), 10))
                Set elements = doc.querySelectorAll("td.tableBlock__col")
                
                For v = 0 To elements.Length - 1
                    
                    If final_price_is_set Then Exit For
                    Text = UCase(elements.Item(v).innerText)
                    
                    If is_winner(text) Then
                        finalPrice = Val(Replace(elements.Item(v + 1).innerText, ",", "."))
                        final_price_is_set = TRUE
                    End If
                    
                Next v
                
                If finalPrice <> 0 Then
                    max_price = Val(Replace(doc.querySelector(".cardMainInfo__content.cost").innerText, ",", "."))
                    
                    If IsLess25(max_price, finalPrice) Then
                        .Range("_25")(i).Value = ">"
                        .Range(RNG_NOTE)(i).Value = "Добросовестность!!!"
                        If IsOneAndHalfGuarantee(max_price) Then
                            .Range(RNG_NOTE)(i).Value = "Полуторная!!!"
                        End If
                    Else
                        .Range("_25")(i).Value = "-"
                    End If
                    
                ElseIf Not IsEmpty(.Range(RNG_RESULT_PROT_NAME)(i).Value) Then
                    .Range("_25")(i).Value = "-"
                End If
                
                If IsEmpty(.Range(RNG_SHORT_ABOUT)(i)) Then
                    .Range(RNG_SHORT_ABOUT)(i).Value = Trim(doc.querySelector(".cardMainInfo__content").innerText)
                End If
                
                If IsEmpty(.Range(RNG_REGION)(i)) Then
                    Set a = doc.querySelector(".cardMainInfo__content > a")
                    Set doc = HTMLDoc(a.href)
                    .Range(RNG_REGION)(i).Value = RegionStr(doc.querySelector("div.registry-entry__body-value").innerText)
                End If
                
            End If
            Call show_status(counter, total_rows, label)
            counter = counter + 1
        Next num
        
        If FILTER_MODE_OFF = FALSE Then
            .Range("A1").AutoFilter Field:=.Range(RNG_RESULT_PROT_NAME)(1).Column
            Call toggle_screen_upd
        End If
        
    End With
    
End Sub
Private Sub Номера_Контрактов()
    Dim doc         As HTMLDocument
    Dim v           As Long, i As Integer, lastrow As Integer
    Dim total_rows  As Integer, counter As Integer
    Dim zakNumber   As String, cw_is_set As Boolean
    Dim elements    As IHTMLDOMChildrenCollection
    Dim visible_range As Range, num As Range
    Dim new_contracts As Scripting.Dictionary
    Const label     As String = "Контроль Номера Контрактов "
    
    cw_is_set = set_control_workbook(ControlWB, ControlWS)
    
    With ControlWB.Worksheets(ControlWS.Name)
        
        If FILTER_MODE_OFF = FALSE Then
            Call toggle_screen_upd
            .AutoFilterMode = FALSE
            .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=CONCLUDED
            .Range("A1").AutoFilter Field:=.Range(RNG_CONTRACT_OFF_NYM_NAME)(1).Column, Criteria1:="="
        End If
        
        lastrow = .Range(RNG_NUMBER_NAME).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
        Set visible_range = .Range(.Range(RNG_NUMBER_NAME)(2).Address & ":" & .Range(RNG_NUMBER_NAME)(lastrow).Address).SpecialCells(xlCellTypeVisible)
        total_rows = visible_range.Count
        counter = 1
        
        Set new_contracts = New Scripting.Dictionary
        
        For Each num In visible_range
            i = num.Row
            zakNumber = Trim(Replace(.Range(RNG_NUMBER_NAME)(i).Value, "№", ""))
            
            If IsEmpty(.Range(RNG_CONTRACT_OFF_NYM_NAME)(i)) And .Range(RNG_STATUS)(i).Value = CONCLUDED Then
                Set doc = HTMLDoc("https://zakupki.gov.ru/epz/contract/search/results.html?orderNumber=" & zakNumber)
                Set elements = doc.querySelectorAll("div.registry-entry__form")
                For v = 0 To elements.Length - 1
                    If .Range(RNG_CUSTOMER_NAME)(i).Value = Trim(elements.Item(v).querySelector(".registry-entry__body-href > a").innerText) Then
                        .Range(RNG_CONTRACT_NUM_NAME)(i).Value = Trim(Replace(elements.Item(v).querySelector(".registry-entry__header-mid__number > a").innerText, "№", ""))
                        .Range(RNG_CONTRACT_OFF_NYM_NAME)(i).Value = Trim(Replace(elements.Item(v).querySelector(".registry-entry__body-value").innerText, "№", ""))
                        .Range(RNG_CONTRACT_DATE_NAME)(i).Value = CDate(Trim(elements.Item(v).querySelector(".data-block__value").innerText))
                        
                    End If
                Next v
            End If
            
            Call show_status(counter, total_rows, label)
            counter = counter + 1
        Next num
        
        If new_contracts.Count > 0 Then Call handle_new_signings(new_contracts, True)
        
        If FILTER_MODE_OFF = FALSE Then
            .Range("A1").AutoFilter Field:=.Range(RNG_CONTRACT_OFF_NYM_NAME)(1).Column
            .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=SIGNING
            Call toggle_screen_upd
        End If
        
    End With
    
End Sub
Private Sub Подписание_Заказчиком()
    Dim doc         As HTMLDocument
    Dim v           As Long, i As Integer, lastrow As Integer, str_len As Integer
    Dim serchS      As String, Text As String
    Dim total_rows  As Integer, counter As Integer
    Dim zakNumber   As String, cw_is_set As Boolean, finded As Boolean
    Dim elements    As IHTMLDOMChildrenCollection
    Dim visible_range As Range, num As Range
    Dim signed      As Scripting.Dictionary
    Const label     As String = "Контроль Проверка подписания Заказчиком "
    
    cw_is_set = set_control_workbook(ControlWB, ControlWS)
    
    With ControlWB.Worksheets(ControlWS.Name)
        
        If FILTER_MODE_OFF = FALSE Then
            Call toggle_screen_upd
            .AutoFilterMode = FALSE
            .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=CONCLUDED
            .Range("A1").AutoFilter Field:=.Range(RNG_CONTRACT_DATE_NAME)(1).Column, Criteria1:="="
        End If
        
        lastrow = .Range(RNG_NUMBER_NAME).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
        Set visible_range = .Range(.Range(RNG_NUMBER_NAME)(2).Address & ":" & .Range(RNG_NUMBER_NAME)(lastrow).Address).SpecialCells(xlCellTypeVisible)
        total_rows = visible_range.Count
        counter = 1
        
        Set signed = New Scripting.Dictionary
        
        For Each num In visible_range
            i = num.Row
            zakNumber = Trim(Replace(.Range(RNG_NUMBER_NAME)(i).Value, "№", ""))
            finded = FALSE
            
            Set doc = HTMLDoc("https://zakupki.gov.ru/epz/order/notice/card/event/journal/list.html?number=&sid=" & .Range(RNG_SID)(i).Value & "&page=1&pageSize=50&qualifier=rpecJournalEventService")
            Set elements = doc.querySelectorAll("table td:nth-child(2)")
            
            For v = 0 To elements.Length - 1
                If finded Then Exit For
                serchS = ""
                Text = UCase(elements.Item(v).innerText)
                
                If Text Like "*ПЕРЕДАН ДОКУМЕНТ*«ПОДПИСАННЫЙ КОНТРАКТ»*" Then
                    str_len = Len("«ПОДПИСАННЫЙ КОНТРАКТ» ОТ") + 1
                    If Mid(text, InStr(1, text, "«ПОДПИСАННЫЙ КОНТРАКТ» ОТ", vbTextCompare) + str_len, 10) Like "##.##.####" Then
                        serchS = Mid(text, InStr(1, text, "«ПОДПИСАННЫЙ КОНТРАКТ» ОТ", vbTextCompare) + str_len, 10)
                    Else
                        serchS = Mid(text, InStr(1, text, "«ПОДПИСАННЫЙ КОНТРАКТ» РЕД.", vbTextCompare) + Len("«ПОДПИСАННЫЙ КОНТРАКТ» РЕД.") + 1, 10)
                    End If
                    
                    .Range(RNG_CONTRACT_DATE_NAME)(i).Value = CDate(serchS)
                    
                    finded = TRUE
                    
                End If
                
            Next v
            
            Call show_status(counter, total_rows, label)
            counter = counter + 1
        Next num
        
        If signed.Count > 0 Then Call handle_new_signings(signed, False)
        
        If FILTER_MODE_OFF = FALSE Then
            .Range("A1").AutoFilter Field:=.Range(RNG_CONTRACT_DATE_NAME)(1).Column
            .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=SIGNING
            Call toggle_screen_upd
        End If
        
    End With
    
End Sub
Private Sub Проверка_Новых()
    Dim lastrow     As Integer
    Dim v           As Long, i As Long, time_now As Long, weak_ago As Long, customers_count As Integer
    Dim doc         As HTMLDocument, protocol_date As IHTMLElement
    Dim elements    As IHTMLDOMChildrenCollection
    Dim OrgIsSet    As Boolean, cw_is_set As Boolean, wk_is_open As Boolean
    Dim Text        As String, k As Variant
    Dim notifications_to_add As Scripting.Dictionary
    Dim finded      As Variant
    Dim wk          As Workbook
    Dim r           As Range
    Const label     As String = "Контроль проверка не внесенных закупок "
    
    Call toggle_screen_upd
    
    Set notifications_to_add = New Scripting.Dictionary
    
    wk_is_open = open_registry_wb()
    
    For Each wk In Application.Workbooks
        
        If wk.Name = REGISTRY_NAME Or wk.Name = EDIN_REGISTRY_NAME Then
            time_now = Now()
            weak_ago = DateAdd("ww", -3, Now())
            wk.Sheets(1).AutoFilterMode = FALSE
            
            With wk.Sheets(1).Range("A1")
                
                .AutoFilter Field:=.Range("A1").Range(RNG_STATUS)(1).Column, Criteria1:=Array("допущены", "заявлены", "выиграли"), Operator:=xlFilterValues
                
                .AutoFilter Field:=.Range("Дата_проведения_аукциона_конкурса")(1).Column, Criteria1:=">=" & weak_ago, _
                            Operator:=xlAnd, Criteria2:="<=" & time_now
                
                For Each r In .Range(RNG_NUMBER_NAME).Rows.SpecialCells(xlCellTypeVisible, xlTextValues)
                    OrgIsSet = FALSE
                    If r.Value = "" Then Exit For
                    Set doc = HTMLDoc(SUPPLIER_RESULTS & Replace(r.Value, "№", ""))
                    Set protocol_date = doc.querySelector("section.blockInfo__section section.blockInfo__section:last-child span:last-child")
                    
                    If Not protocol_date Is Nothing Then
                        
                        customers_count = doc.querySelectorAll("table.blockInfo__table:nth-child(1) > tbody > tr").Length
                        Set elements = doc.querySelectorAll("td.tableBlock__col")
                        
                        For v = 0 To elements.Length - 1
                            
                            If OrgIsSet Then Exit For
                            Text = UCase(elements.Item(v).innerText)
                            
                            If is_sole_participant(text) Then
                                
                                OrgIsSet = TRUE
                            End If
                            
                        Next v
                        
                    End If
                    
                Next r
                
            End With
            
            wk.Close SaveChanges:=False
            
        End If
        
    Next wk
    
    If notifications_to_add.Count = 0 Then Exit Sub
    
    cw_is_set = set_control_workbook(ControlWB, ControlWS)
    
    With ControlWB.Worksheets(ControlWS.Name)
        
        .AutoFilterMode = FALSE
        .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=Array(SIGNING, CONCLUDED), Operator:=xlFilterValues
        lastrow = .Range(RNG_NUMBER_NAME).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1
        
        For Each k In notifications_to_add.Keys
            Set finded = .Range(RNG_NUMBER_NAME).Find(What:=notifications_to_add(k), LookIn:=xlValues)
            
            If finded Is Nothing Then
                customers_count = Val(k)
                If customers_count = 1 Then
                    .Range(RNG_NUMBER_NAME)(lastrow).Value = notifications_to_add(k)
                    lastrow = lastrow + 1
                Else
                    For v = 0 To customers_count - 1
                        .Range(RNG_NUMBER_NAME)(lastrow).Value = notifications_to_add(k)
                        lastrow = lastrow + 1
                    Next v
                End If
            End If
            
        Next k
        
        .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=SIGNING
        
    End With
    
    Call toggle_screen_upd
    
End Sub
Sub Проверка_Новых22()
    Dim lastrow     As Integer
    Dim v           As Long, i As Long, time_now As Long, two_weaks_ago As Long
    Dim cw_is_set   As Boolean, wk_is_open As Boolean
    Dim k           As Variant
    Dim notifications_to_add As Scripting.Dictionary
    Dim finded      As Variant
    Dim wk          As Workbook
    Dim r           As Range
    Const label     As String = "Контроль проверка не внесенных закупок "
    
    Call toggle_screen_upd
    
    Set notifications_to_add = New Scripting.Dictionary
    
    wk_is_open = open_registry_wb()
    
    For Each wk In Application.Workbooks
        
        If wk.Name = REGISTRY_NAME Or wk.Name = EDIN_REGISTRY_NAME Then
            time_now = Now()
            two_weaks_ago = DateAdd("ww", -2, Now())
            wk.Sheets(1).AutoFilterMode = FALSE
            
            With wk.Sheets(1).Range("A1")
                
                .AutoFilter Field:=.Range("A1").Range(RNG_STATUS)(1).Column, Criteria1:="выиграли", Operator:=xlFilterValues
                
                .AutoFilter Field:=.Range("Дата_проведения_аукциона_конкурса")(1).Column, Criteria1:=">=" & two_weaks_ago, _
                            Operator:=xlAnd, Criteria2:="<=" & time_now
                
                For Each r In .Range(RNG_NUMBER_NAME).Rows.SpecialCells(xlCellTypeVisible, xlTextValues)
                    
                    If r.Value = "" Then Exit For
                    If r.Row <> 1 Then notifications_to_add.Add r.Value, r.Value
                    
                Next r
                
            End With
            
            wk.Close SaveChanges:=False
            
        End If
        
    Next wk
    
    If notifications_to_add.Count = 0 Then Exit Sub
    
    cw_is_set = set_control_workbook(ControlWB, ControlWS)
    
    With ControlWB.Worksheets(ControlWS.Name)
        
        .AutoFilterMode = FALSE
        .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=Array(SIGNING, CONCLUDED), Operator:=xlFilterValues
        lastrow = .Range(RNG_NUMBER_NAME).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1
        
        For Each k In notifications_to_add.Keys
            Set finded = .Range(RNG_NUMBER_NAME).Find(What:=notifications_to_add(k), LookIn:=xlValues)
            
            If finded Is Nothing Then
                .Range(RNG_NUMBER_NAME)(lastrow).Value = notifications_to_add(k)
                lastrow = lastrow + 1
            End If
            
        Next k
        
        .Range("A1").AutoFilter Field:=.Range(RNG_STATUS)(1).Column, Criteria1:=SIGNING
        
    End With
    
    Call toggle_screen_upd
    
End Sub

Private Function RegionStr(s As String) As String
    RegionStr = UCase(s)
    If RegionStr Like "*КАРЕЛ*РЕСП*" Then RegionStr = "Карелия": Exit Function
    If RegionStr Like "*ПЕТРОЗАВОДСК,*" Then RegionStr = "Карелия": Exit Function
    If RegionStr Like "*КОМИ*РЕСП*" Then RegionStr = "Коми": Exit Function
    If RegionStr Like "*СЫКТЫВКАР,*" Then RegionStr = "Коми": Exit Function
    If RegionStr Like "*ЧЕЧЕН*РЕСП*" Then RegionStr = "Чечня": Exit Function
    If RegionStr Like "*ГРОЗНЫЙ,*" Then RegionStr = "Чечня": Exit Function
    If RegionStr Like "*ЧУВАШ*РЕСП*" Then RegionStr = "Чувашия": Exit Function
    If RegionStr Like "*ЭЛИСТА,*" Then RegionStr = "Чувашия": Exit Function
    If RegionStr Like "*ЧУКОТСКИЙ*" Then RegionStr = "Чукотка": Exit Function
    If RegionStr Like "*АНЫДЫРЬ,*" Then RegionStr = "Чукотка": Exit Function
    If RegionStr Like "*УДМУРТ*РЕСП*" Then RegionStr = "Удмуртия": Exit Function
    If RegionStr Like "*ИЖЕВСК,*" Then RegionStr = "Удмуртия": Exit Function
    If RegionStr Like "*ИНГУШ*РЕСП*" Then RegionStr = "Ингушетия": Exit Function
    If RegionStr Like "*КЕМЕРОВ*ОБЛ*" Then RegionStr = "Кемерово": Exit Function
    If RegionStr Like "*КЕМЕРОВО,*" Then RegionStr = "Кемерово": Exit Function
    If RegionStr Like "*ДАГЕСТ*РЕСП*" Then RegionStr = "Дагестан": Exit Function
    If RegionStr Like "*МАХАЧКАЛА,*" Then RegionStr = "Дагестан": Exit Function
    If RegionStr Like "*КРЫМ*РЕСП*" Then RegionStr = "Крым": Exit Function
    If RegionStr Like "*СИМФЕРОПОЛЬ,*" Then RegionStr = "Крым": Exit Function
    If RegionStr Like "*САХА*ЯКУТИ*" Then RegionStr = "Якутия": Exit Function
    If RegionStr Like "*ЯКУТСК,*" Then RegionStr = "Якутия": Exit Function
    If RegionStr Like "*ХАКАС*" Then RegionStr = "Хакасия": Exit Function
    If RegionStr Like "*АБАКАН,*" Then RegionStr = "Хакасия": Exit Function
    If RegionStr Like "*ХАНТЫ*МАН*" Then RegionStr = "ХМАО": Exit Function
    If RegionStr Like "*ХАНТЫ-МАНСИЙСК,*" Then RegionStr = "ХМАО": Exit Function
    If RegionStr Like "*БАШКОРТ*РЕСП*" Then RegionStr = "Башкирия": Exit Function
    If RegionStr Like "*УФА,*" Then RegionStr = "Башкирия": Exit Function
    If RegionStr Like "*САНКТ*ПЕТ*" Then RegionStr = "Санкт-Петербург": Exit Function
    If RegionStr Like "*САНКТ-ПЕТЕРБУРГ,*" Then RegionStr = "Санкт-Петербург": Exit Function
    If RegionStr Like "*ЯМАЛО*НЕН*" Then RegionStr = "Ямало-Ненецкий": Exit Function
    If RegionStr Like "*САЛЕХАРД,*" Then RegionStr = "Ямало-Ненецкий": Exit Function
    If RegionStr Like "*ТАТАРСТАН*" Then RegionStr = "Татарстан": Exit Function
    If RegionStr Like "*КАЗАНЬ,*" Then RegionStr = "Татарстан": Exit Function
    If RegionStr Like "*КРАСНОДАР*КРАЙ*" Then RegionStr = "Краснодар": Exit Function
    If RegionStr Like "*КРАСНОДАР,*" Then RegionStr = "Краснодар": Exit Function
    If RegionStr Like "*Г.КРАСНОДАР*" Then RegionStr = "Краснодар": Exit Function
    If RegionStr Like "*ЧЕЛЯБИНСК*ОБЛ*" Then RegionStr = "Челябинск": Exit Function
    If RegionStr Like "*ЧЕЛЯБИНСК,*" Then RegionStr = "Челябинск": Exit Function
    If RegionStr Like "*ОСЕТИ*РЕСП*" Then RegionStr = "Владикавказ": Exit Function
    If RegionStr Like "*ВЛАДИКАВКАЗ,*" Then RegionStr = "Владикавказ": Exit Function
    If RegionStr Like "*СТАВРОПОЛ*КРАЙ*" Then RegionStr = "Ставрополь": Exit Function
    If RegionStr Like "*СТАВРОПОЛЬ,*" Then RegionStr = "Ставрополь": Exit Function
    If RegionStr Like "*КАБАРД*РЕСП*" Then RegionStr = "КБР": Exit Function
    If RegionStr Like "*НАЛЬЧИК,*" Then RegionStr = "КБР": Exit Function
    If RegionStr Like "*АСТРАХАН*ОБЛ*" Then RegionStr = "Астрахань": Exit Function
    If RegionStr Like "*АСТРАХАНЬ,*" Then RegionStr = "Астрахань": Exit Function
    If RegionStr Like "*АДЫГ*РЕСП*" Then RegionStr = "Адыгея": Exit Function
    If RegionStr Like "*МАЙКОП,*" Then RegionStr = "Адыгея": Exit Function
    If RegionStr Like "*КАРАЧ*РЕСП*" Then RegionStr = "КЧР": Exit Function
    If RegionStr Like "*ЧЕРКЕССК,*" Then RegionStr = "КЧР": Exit Function
    If RegionStr Like "*МОСКВ*" Then RegionStr = "Москва": Exit Function
    If RegionStr Like "*МОСКВА,*" Then RegionStr = "Москва": Exit Function
    If RegionStr Like "*РОСТОВ*ОБЛ*" Then RegionStr = "Ростов": Exit Function
    If RegionStr Like "*РОСТОВ-НА-ДОНУ,*" Then RegionStr = "Ростов": Exit Function
    If RegionStr Like "*АЛТАЙ*РЕСП*" Then RegionStr = "Алтай": Exit Function
    If RegionStr Like "*ГОРНО-АЛТАЙСК,*" Then RegionStr = "Алтай": Exit Function
    If RegionStr Like "*АЛТАЙСК*КРАЙ*" Then RegionStr = "Алтай": Exit Function
    If RegionStr Like "*БАРНАУЛ,*" Then RegionStr = "Алтай": Exit Function
    If RegionStr Like "*АМУРСК*ОБЛ*" Then RegionStr = "Амур": Exit Function
    If RegionStr Like "*БЛАГОВЕЩЕНСК,*" Then RegionStr = "Амур": Exit Function
    If RegionStr Like "*АРХАНГЕЛЬСК*ОБЛ*" Then RegionStr = "Архангельск": Exit Function
    If RegionStr Like "*АРХАНГЕЛЬСК,*" Then RegionStr = "Архангельск": Exit Function
    If RegionStr Like "*БРЯНСК*ОБЛ*" Then RegionStr = "Брянск": Exit Function
    If RegionStr Like "*БРЯНСК,*" Then RegionStr = "Брянск": Exit Function
    If RegionStr Like "*БУРЯТ*РЕСП*" Then RegionStr = "Бурятия": Exit Function
    If RegionStr Like "*УЛАН-УДЭ,*" Then RegionStr = "Бурятия": Exit Function
    If RegionStr Like "*ВЛАДИМИР*ОБЛ*" Then RegionStr = "Владимир": Exit Function
    If RegionStr Like "*ВЛАДИМИР*ОБЛ*" Then RegionStr = "Владимир": Exit Function
    If RegionStr Like "*ВЛАДИМИР,*" Then RegionStr = "Владимир": Exit Function
    If RegionStr Like "*ВОЛГОГРАД*ОБЛ*" Then RegionStr = "Волгоград": Exit Function
    If RegionStr Like "*ВОЛГОГРАД,*" Then RegionStr = "Волгоград": Exit Function
    If RegionStr Like "*ВОЛОГОДСК*ОБЛ*" Then RegionStr = "Вологда": Exit Function
    If RegionStr Like "*ВОЛОГДА,*" Then RegionStr = "Вологда": Exit Function
    If RegionStr Like "*ВОРОНЕЖ*ОБЛ*" Then RegionStr = "Воронеж": Exit Function
    If RegionStr Like "*ВОРОНЕЖ,*" Then RegionStr = "Воронеж": Exit Function
    If RegionStr Like "*ЕВРЕЙСК*" Then RegionStr = "Еврейская АО": Exit Function
    If RegionStr Like "*ЗАБАЙКАЛЬСК*КРАЙ*" Then RegionStr = "Чита": Exit Function
    If RegionStr Like "*ЧИТА,*" Then RegionStr = "Чита": Exit Function
    If RegionStr Like "*ИВАНОВ*ОБЛ*" Then RegionStr = "Иваново": Exit Function
    If RegionStr Like "*ИВАНОВО,*" Then RegionStr = "Иваново": Exit Function
    If RegionStr Like "*БЕЛГОРОД*ОБЛ*" Then RegionStr = "Белгород": Exit Function
    If RegionStr Like "*БЕЛГОРОД,*" Then RegionStr = "Белгород": Exit Function
    If RegionStr Like "*ТУЛЬ*ОБЛ*" Then RegionStr = "Тула": Exit Function
    If RegionStr Like "*ТУЛА,*" Then RegionStr = "Тула": Exit Function
    If RegionStr Like "*ИРКУТСК*ОБЛ*" Then RegionStr = "Иркутск": Exit Function
    If RegionStr Like "*ИРКУТСК,*" Then RegionStr = "Иркутск": Exit Function
    If RegionStr Like "*КАЛИНИНГР*ОБЛ*" Then RegionStr = "Калининград": Exit Function
    If RegionStr Like "*КАЛИНИНГРАД,*" Then RegionStr = "Калининград": Exit Function
    If RegionStr Like "*КАЛМЫК*РЕСП*" Then RegionStr = "Калмыкия": Exit Function
    If RegionStr Like "*ЭЛИСТА,*" Then RegionStr = "Калмыкия": Exit Function
    If RegionStr Like "*КАЛУЖСК*ОБЛ*" Then RegionStr = "Калуга": Exit Function
    If RegionStr Like "*КАЛУГА,*" Then RegionStr = "Калуга": Exit Function
    If RegionStr Like "*КАМЧАТСК*КРАЙ*" Then RegionStr = "Камчатка": Exit Function
    If RegionStr Like "*КОСТРОМ*ОБЛ*" Then RegionStr = "Кострома": Exit Function
    If RegionStr Like "*КОСТРОМА,*" Then RegionStr = "Кострома": Exit Function
    If RegionStr Like "*КРАСНОЯРСК*КРАЙ*" Then RegionStr = "Красноярск": Exit Function
    If RegionStr Like "*КРАСНОЯРСК,*" Then RegionStr = "Красноярск": Exit Function
    If RegionStr Like "*КУРГАН*ОБЛ*" Then RegionStr = "Курган": Exit Function
    If RegionStr Like "*КУРГАН,*" Then RegionStr = "Курган": Exit Function
    If RegionStr Like "*КУРСК*ОБЛ*" Then RegionStr = "Курск": Exit Function
    If RegionStr Like "*КУРСК,*" Then RegionStr = "Курск": Exit Function
    If RegionStr Like "*ЛЕНИНГРАД*ОБЛ*" Then RegionStr = "Ленинградская": Exit Function
    If RegionStr Like "*ЛИПЕЦК*ОБЛ*" Then RegionStr = "Липецк": Exit Function
    If RegionStr Like "*ЛИПЕЦК,*" Then RegionStr = "Липецк": Exit Function
    If RegionStr Like "*МАГАДАНСК*ОБЛ*" Then RegionStr = "Магадан": Exit Function
    If RegionStr Like "*МАГАДАН,*" Then RegionStr = "Магадан": Exit Function
    If RegionStr Like "*МАРИЙ*" Then RegionStr = "Марий Эл": Exit Function
    If RegionStr Like "*ЙОШКАР-ОЛА,*" Then RegionStr = "Марий Эл": Exit Function
    If RegionStr Like "*МОРДОВ*РЕСП*" Then RegionStr = "Мордовия": Exit Function
    If RegionStr Like "*САРАНСК,*" Then RegionStr = "Мордовия": Exit Function
    If RegionStr Like "*МОСКОВСК*ОБЛ*" Then RegionStr = "Московская область": Exit Function
    If RegionStr Like "*МУРМАНСК*ОБЛ*" Then RegionStr = "Мурманск": Exit Function
    If RegionStr Like "*МУРМАНСК,*" Then RegionStr = "Мурманск": Exit Function
    If RegionStr Like "*НИЖЕГОРОД*ОБЛ*" Then RegionStr = "Нижний": Exit Function
    If RegionStr Like "*НИЖНИЙ НОВГОРОД,*" Then RegionStr = "Нижний": Exit Function
    If RegionStr Like "*НОВГОРОДСК*ОБЛ*" Then RegionStr = "Новгород": Exit Function
    If RegionStr Like "*НОВОСИБИР*ОБЛ*" Then RegionStr = "Новосибирск": Exit Function
    If RegionStr Like "*НОВОСИБИРСК,*" Then RegionStr = "Новосибирск": Exit Function
    If RegionStr Like "*ТОМСК*ОБЛ*" Then RegionStr = "Томск": Exit Function
    If RegionStr Like "*ТОМСК,*" Then RegionStr = "Томск": Exit Function
    If RegionStr Like "*ОМСК*ОБЛ*" Then RegionStr = "Омск": Exit Function
    If RegionStr Like "*ОМСК,*" Then RegionStr = "Омск": Exit Function
    If RegionStr Like "*ОРЕНБУРГСК*ОБЛ*" Then RegionStr = "Оренбург": Exit Function
    If RegionStr Like "*ОРЕНБУРГ,*" Then RegionStr = "Оренбург": Exit Function
    If RegionStr Like "*ОРЛОВ*ОБЛ*" Then RegionStr = "Орел": Exit Function
    If RegionStr Like "*ОРЁЛ,*" Then RegionStr = "Орел": Exit Function
    If RegionStr Like "*ПЕНЗЕНСК*ОБЛ*" Then RegionStr = "Пенза": Exit Function
    If RegionStr Like "*ПЕНЗА,*" Then RegionStr = "Пенза": Exit Function
    If RegionStr Like "*ПЕРМСК*КРАЙ*" Then RegionStr = "Пермь": Exit Function
    If RegionStr Like "*ПЕРМЬ,*" Then RegionStr = "Пермь": Exit Function
    If RegionStr Like "*ПРИМОРСК*КРАЙ*" Then RegionStr = "Владивосток": Exit Function
    If RegionStr Like "*ВЛАДИВОСТОК,*" Then RegionStr = "Владивосток": Exit Function
    If RegionStr Like "*ПСКОВ*ОБЛ*" Then RegionStr = "Псков": Exit Function
    If RegionStr Like "*ПСКОВ,*" Then RegionStr = "Псков": Exit Function
    If RegionStr Like "*РЯЗАН*ОБЛ*" Then RegionStr = "Рязань": Exit Function
    If RegionStr Like "*РЯЗАНЬ,*" Then RegionStr = "Рязань": Exit Function
    If RegionStr Like "*САМАРСК*ОБЛ*" Then RegionStr = "Самара": Exit Function
    If RegionStr Like "*САМАРА,*" Then RegionStr = "Самара": Exit Function
    If RegionStr Like "*САРАТОВ*ОБЛ*" Then RegionStr = "Саратов": Exit Function
    If RegionStr Like "*САРАТОВ,*" Then RegionStr = "Саратов": Exit Function
    If RegionStr Like "*СВЕРДЛОВСК*ОБЛ*" Then RegionStr = "Екатеринбург": Exit Function
    If RegionStr Like "*ЕКАТЕРИНБУРГ,*" Then RegionStr = "Екатеринбург": Exit Function
    If RegionStr Like "*СМОЛЕНСК*ОБЛ*" Then RegionStr = "Смоленск": Exit Function
    If RegionStr Like "*СМОЛЕНСК,*" Then RegionStr = "Смоленск": Exit Function
    If RegionStr Like "*ТАМБОВ*ОБЛ*" Then RegionStr = "Тамбов": Exit Function
    If RegionStr Like "*ТАМБОВ,*" Then RegionStr = "Тамбов": Exit Function
    If RegionStr Like "*ТВЕРСК*ОБЛ*" Then RegionStr = "Тверь": Exit Function
    If RegionStr Like "*ТВЕРЬ,*" Then RegionStr = "Тверь": Exit Function
    If RegionStr Like "*ТЫВА*РЕСП*" Then RegionStr = "Тыва": Exit Function
    If RegionStr Like "*ТЮМЕНСК*ОБЛ*" Then RegionStr = "Тюмень": Exit Function
    If RegionStr Like "*ТЮМЕНЬ,*" Then RegionStr = "Тюмень": Exit Function
    If RegionStr Like "*УЛЬЯНОВСК*ОБЛ*" Then RegionStr = "Ульяновск": Exit Function
    If RegionStr Like "*УЛЬЯНОВСК,*" Then RegionStr = "Ульяновск": Exit Function
    If RegionStr Like "*ХАБАРОВСК*КРАЙ*" Then RegionStr = "Хабаровск": Exit Function
    If RegionStr Like "*ХАБАРОВСК,*" Then RegionStr = "Хабаровск": Exit Function
    If RegionStr Like "*ЯРОСЛАВСК*ОБЛ*" Then RegionStr = "Ярославль": Exit Function
    If RegionStr Like "*ЯРОСЛАВЛЬ,*" Then RegionStr = "Ярославль": Exit Function
    If RegionStr Like "*КИРОВ*ОБЛ*" Then RegionStr = "Киров": Exit Function
    If RegionStr Like "*КИРОВ,*" Then RegionStr = "Киров": Exit Function
    If RegionStr Like "*САХАЛИНСК*ОБЛ*" Then RegionStr = "Сахалин": Exit Function
    If RegionStr Like "*ЮЖНО-САХАЛИНСК,*" Then RegionStr = "Сахалин": Exit Function
    If RegionStr Like "*СЕВАСТОПОЛ*" Then RegionStr = "Севастополь": Exit Function
    If RegionStr Like "*СЕВАСТОПОЛЬ,*" Then RegionStr = "Севастополь": Exit Function
    If RegionStr Like "*ВЕЛИК*НОВГОРОД*" Then RegionStr = "Новгород": Exit Function
End Function

Function set_control_workbook(ByRef ControlWB As Workbook, ByRef ControlWS As Worksheet) As Boolean
    If ControlWB Is Nothing Then
        Set ControlWB = Workbooks(CONTROL_NAME)
        Set ControlWS = ControlWB.Sheets(1)
        set_control_workbook = TRUE
    End If
End Function

Private Function open_registry_wb() As Boolean
    Dim wb          As Workbook
    
    For Each wb In Application.Workbooks
        If wb.Name = REGISTRY_NAME Then open_registry_wb = TRUE
    Next wb
    
    If open_registry_wb = FALSE Then
        Workbooks.Open (REGISRTY_PATH), UpdateLinks:=0, ReadOnly:=True
    End If
    
    For Each wb In Application.Workbooks
        If wb.Name = EDIN_REGISTRY_NAME Then open_registry_wb = TRUE
    Next wb
    
    If open_registry_wb = FALSE Then
        Workbooks.Open (EDIN_REGISRTY_PATH), UpdateLinks:=0, ReadOnly:=True
    End If
    
    open_registry_wb = TRUE
End Function

Private Function is_winner(ByRef Text As String) As Boolean
    
    is_winner = FALSE
    
    If Text Like "*1*-*ПОБЕДИТЕЛЬ*" Then is_winner = TRUE
    
End Function

Private Function is_sole_participant(ByRef Text As String) As Boolean
    
    is_sole_participant = FALSE
    
    If Text Like "*ПОДАН*ТОЛЬКО*ОДН*ПРИЗНАНА СООТВЕТСТВУЮЩ*" Or _
       Text Like "*ПОДАН*ТОЛЬКО*ОДН*ЗАЯВКА СООТВЕТСТВУЕТ ТРЕБОВАНИЯМ*" Or _
       Text Like "*ПОДАН*ЕДИНСТВ*ПРЕДЛОЖ*О ЕЕ СООТВЕТСТВИИ*" Or _
       Text Like "*РЕШЕНИЕ*СООТВЕ*ТОЛЬКО*ОДНОЙ*" Or _
       Text Like "*ПРИЗНАН*НЕСОСТОЯВШИМ*УКАЗАНН*ПРИЧИН*НЕ НАЙДЕН*" Or _
       Text Like "*1*-*ПОБЕДИТЕЛЬ*" Then
    
    is_sole_participant = TRUE
    
End If

End Function
Sub ControlHideUnHide()
    
    If Range(Range(RNG_RESULT_PROT_NAME), Range("_25")).EntireColumn.Hidden = TRUE Then
        Range(Range(RNG_RESULT_PROT_NAME), Range("_25")).EntireColumn.Hidden = FALSE
    Else
        Range(Range(RNG_RESULT_PROT_NAME), Range("_25")).EntireColumn.Hidden = TRUE
    End If
    
End Sub
Private Function IsLess25(StartPrice, finalPrice As Double) As Boolean
    IsLess25 = (StartPrice * 0.75) > finalPrice
End Function
Private Function IsOneAndHalfGuarantee(StartPrice As Double) As Boolean
    IsOneAndHalfGuarantee = StartPrice > 15000000
End Function
Private Function HTMLDoc(url As String) As HTMLDocument
    
    Dim http        As New MSXML2.XMLHTTP60
    http.Open "GET", url, FALSE
    http.send
    Set HTMLDoc = New HTMLDocument
    HTMLDoc.body.innerHTML = http.responseText
    
End Function
Private Sub toggle_screen_upd()
    
    If Application.Calculation = xlCalculationAutomatic Then
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
    End If
    
    If Application.ScreenUpdating Then
        Application.ScreenUpdating = FALSE
    Else
        Application.ScreenUpdating = TRUE
    End If
    
    If Application.EnableEvents Then
        Application.EnableEvents = FALSE
    Else
        Application.EnableEvents = TRUE
    End If
    
    If Application.DisplayStatusBar = FALSE Then Application.DisplayStatusBar = TRUE
    
End Sub
Private Function HttpGetEventJournalSid(url As String) As String
    
    Dim v           As Long, posStart, posEnd As Integer
    Dim doc         As New HTMLDocument
    Dim scripts     As Variant
    
    With New MSXML2.XMLHTTP60
        .Open "GET", url, FALSE
        On Error Resume Next
        .send
        If .Status = 200 Then
            doc.body.innerHTML = .responseText
            Set scripts = doc.getElementsByTagName("script")
            For v = 0 To scripts.Length - 1
                posStart = InStr(scripts(v).innerHTML, "sid:        '")
                If posStart > 0 Then
                    posEnd = InStr(posStart + 6, scripts(v).innerHTML,        '")
                    HttpGetEventJournalSid = Mid(scripts(v).innerHTML, posStart + 6, posEnd - (posStart + 6))
                    Exit For
                End If
            Next v
        Else
            HttpGetEventJournalSid = "HttpGet Error"
        End If
        .abort
    End With
    
End Function
Private Sub show_status(current As Integer, total_rows As Integer, topic As String)
    
    Dim NumberOfBars As Integer, CurrentStatus As Integer
    Dim pctDone     As Integer
    
    NumberOfBars = 50
    CurrentStatus = Int((current / total_rows) * NumberOfBars)
    pctDone = Round(CurrentStatus / NumberOfBars * 100, 0)
    Application.StatusBar = topic & " [" & String(CurrentStatus, "|") & _
                            Space(NumberOfBars - CurrentStatus) & "]" & _
                            " " & pctDone & "% Завершено"
    
    If current = total_rows Then Application.StatusBar = ""
    
End Sub