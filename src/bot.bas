Option Explicit

Private Const CONTRACTS_CHAT_ID As String = ""
Private Const SEVER_CHAT_ID As String = ""
Private Const TEST_CHAT_ID As String = ""
Private Const API_URL As String = "https://api.telegram.org/bot"
Private Const TOKEN As String = ""
Private Const CUST_SIGNED As String = "Подписан Заказчиком: "
Private Const NO_NUMBER_SO_FAR As String = "Официальный номер пока не опубликован"
Private Const OF_NUMBER_PUBLISHED As String = "Опубликован номер контракта"
Private Const CR    As String = "%0A"

Sub send_message(msg As String, chat_id As String)
    Dim url         As String
    Dim resp        As String
    
    msg = EncodeUTF8noBOM(msg)
    
    url = API_URL + TOKEN + "/sendMessage?chat_id=" & chat_id & "&text=" & msg
    
    resp = post(url)
End Sub

Sub handle_new_signings(signed As Scripting.Dictionary, is_contract As Boolean)
    Dim k           As Variant
    Dim msg         As String
    Dim zak_num_last_three As String, predmet As String, region As String, chat As String
    Dim idx         As Integer, counter As Integer, total As Integer
    Dim cw_is_set   As Boolean
    Dim lvl_of_notif As Integer, lvl As Integer
    Const label     As String = "Отправка сообщений "
    
    cw_is_set = set_control_workbook(ControlWB, ControlWS)
    total = signed.Count
    counter = 0
    lvl = 0
    If is_contract Then lvl = 1
    
    With ControlWB.Worksheets(ControlWS.Name)
        .AutoFilterMode = FALSE
        
        For Each k In signed.Keys()
            
            idx = signed.Item(k)
            
            zak_num_last_three = Right(.Range(RNG_NUMBER_NAME)(idx).Value, 3)
            region = .Range(RNG_REGION)(idx).Value
            predmet = Trim(Replace(.Range(RNG_SHORT_ABOUT)(idx).Value, region, ""))
            
            lvl_of_notif = .Range(RNG_LVL_NOTIF_NAME)(idx).Value
            
            If lvl_of_notif = lvl Then
                
                If is_contract Then
                    msg = region + " " + predmet + " " + zak_num_last_three & CR & OF_NUMBER_PUBLISHED _
                        & CR & .Range(RNG_CONTRACT_OFF_NYM_NAME)(idx).Value _
                        & " от " & CStr(.Range(RNG_CONTRACT_DATE_NAME)(idx).Value) & " г."
                Else
                    msg = region + " " + predmet + " " + zak_num_last_three & CR & CUST_SIGNED & _
                          CStr(.Range(RNG_CONTRACT_DATE_NAME)(idx).Value)
                    '                    & CR & NO_NUMBER_SO_FAR
                End If
                
                If .Range(RNG_ORG)(idx).Value = "Единство" Or .Range(RNG_ORG)(idx).Value = "Опора-Север" Then
                    chat = SEVER_CHAT_ID
                Else
                    chat = CONTRACTS_CHAT_ID
                End If
                
                Call send_message(msg, chat)
                .Range(RNG_LVL_NOTIF_NAME)(idx).Value = lvl + 1
                
            End If
            
            counter = counter + 1
            Call show_status(counter, total, label)
            '        counter = counter + 1
            
        Next k
        
    End With
End Sub

Private Function post(url As String) As String
    Dim xmlhttp     As New MSXML2.XMLHTTP60
    
    xmlhttp.Open "POST", url, FALSE
    xmlhttp.send
    post = xmlhttp.responseText
    
End Function

Function EncodeUTF8noBOM(ByVal txt As String) As String
    Dim i           As Long, l As String, t As String
    For i = 1 To Len(txt)
        l = Mid(txt, i, 1)
        Select Case AscW(l)
            Case Is > 4095: t = Chr(AscW(l) \ 64 \ 64 + 224) & Chr(AscW(l) \ 64) & Chr(8 * 16 + AscW(l) Mod 64)
            Case Is > 127: t = Chr(AscW(l) \ 64 + 192) & Chr(8 * 16 + AscW(l) Mod 64)
            Case Else: t = l
        End Select
        EncodeUTF8noBOM = EncodeUTF8noBOM & t
    Next
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