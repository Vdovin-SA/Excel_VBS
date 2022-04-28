Const SRV = "https://kdvonline.ru"
Sub Кнопка1_Щелчок()
    On Error GoTo er
    Dim cat As String, url As String, action As String
    
    Set ws = Worksheets("Категории")

    For Counter = 2 To 100
        cat = ws.Cells(Counter, 1).Value
        url = ws.Cells(Counter, 2).Value
        action = ws.Cells(Counter, 3).Value
        
        If action = "Обновить" And url <> "" Then
            Call category_get_conten(cat, url)
        End If
    Next Counter
    MsgBox "Обновление данных выполнено"
    Exit Sub
er:
    MsgBox "Error: " & Err.Description & Err.Number
End Sub
' Работа с HTML
Function HTML_Get_Content(strURL)
    html = ""
    Set MyBrowser = CreateObject("MSXML2.XMLHTTP")
    
    MyBrowser.Open "GET", strURL, False
    MyBrowser.setRequestHeader "User-Agent", "Mozilla/5.0"
    MyBrowser.send
    
    If MyBrowser.Status = 200 Then
        html = MyBrowser.responseText
       ' Убираем спец символы
        html = Replace(html, Chr(9), "")
        html = Replace(html, Chr(10), "")
        html = Replace(html, Chr(13), "")
    End If
    
    HTML_Get_Content = html
    Set MyBrowser = Nothing
End Function
Function HTML_Get_Element(html, pattern)
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.pattern = pattern
 
    Set objMatches = objRegExp.Execute(html)
    If objMatches.Count = 0 Then
        MsgBox "Не найдено совпадений по шаблону\n\n" + pattern
    Else
        Set objMatch = objMatches.item(0)
    End If
 
    HTML_Get_Element = Mid(objMatch.Value, 31, objMatch.Length - 36)
End Function

Sub category_get_conten(category As String, strURL As String)
    On Error GoTo er
    Dim html As String
    
    html = HTML_Get_Content(SRV & strURL)
    If html = "" Then MsgBox "Пустой ответ: " & strURL: Exit Sub
    'html_to_file (html)
    
    ' Разбор страницы
    Dim html_doc As HTMLDocument
    Set html_doc = html_to_doc(html)
    Dim items As IHTMLElementCollection
    Set items = html_doc.getElementsByClassName("c3s8K6a5X")
    
    ' Загрузка основной страницы
    Dim i As Long, i_max As Long
    i_max = items.Length - 1
    For i = 0 To i_max
        Call category_update_item(category, items(i))
    Next i
    
    ' Поиск следующей страницы - РЕКУРСИЯ
    Set items = html_doc.getElementsByClassName("c18ybbMcB")
    i_max = items.Length - 1
    For i = 0 To i_max
        'MsgBox i & " " & items(i).innerText & " " & items(i).tagName
        If items(i).tagName = "A" And items(i).innerText = "next" Then
            url = items(i).href
            url = Right(url, Len(url) - InStr(url, "/") + 1)
            If url <> "" Then
                Call category_get_conten(category, CStr(url))
            End If
        End If
    Next i
    
Exit Sub
er:
    MsgBox "Error: caregory_get_conten [" & Erl & "] " & Err.Description
End Sub
Sub category_update_item(category As String, item As IHTMLElement)
    On Error GoTo er
    Set ws = Worksheets("Товар")
    'html_to_file (item.innerHTML)
    
    Dim url As String
    url = item.getElementsByTagName("a")(1).href
    pos = InStr(url, "/")
    url = Right(url, Len(url) - pos + 1)
    pos = InStrRev(url, "-")
    id = Right(url, Len(url) - pos)
    
    Price = ""
    is_new = "Нет"
    is_promo = "Нет"
    For Each div In item.getElementsByTagName("div")
        If div.className = "b2iP1cx1b" Then
            Price = div.innerText
            Price = Left(Price, InStr(Price, " ") - 1)
        End If
        If div.className = "b10FT7BLs a3blieLf1 l3blieLf1" Then
            is_new = "Да"
        End If
        If div.className = "d10FT7BLs a3blieLf1 m3blieLf1" Then
            is_promo = "Да"
        End If
    Next

    pos = item_pos_by_id(ws, id, 2)

    ws.Cells(pos, 1).Value = category
    ws.Cells(pos, 2).Value = id
    ws.Cells(pos, 3).Value = item.getElementsByTagName("a")(1).innerText
    ws.Cells(pos, 4).Value = Price
    'ws.Cells(pos, 5).Value = url
    Call ws.Hyperlinks.Add(ws.Cells(pos, 5), SRV + url, "", "", "Ссылка") ' Вставляем текстовую ссылку - MAX 64 000
    ws.Cells(pos, 6).Value = is_new
    ws.Cells(pos, 7).Value = is_promo
    
'    Dim cell As Range
'    Set cell = ws.Cells(pos, 8) ' выделяем колонку для изображения
'    Image_url = item.getElementsByTagName("img")(0).src ' получение прямой ссылки на изображение
'    With ws.Pictures.Insert(Image_url) ' создаем обьект изображения
'        .Left = cell.Left + 2
'        .Top = cell.Top + 2
'        .Width = cell.Width - 4 ' вписываем картинку в ячейку
'        .Height = cell.Height - 4
'    End With
    
Exit Sub
er:
    MsgBox "Error: category_update_item [" & Erl & "] " & Err.Description
End Sub
Function item_pos_by_id(ws, id, id_pos As Long) As Long
    On Error GoTo er
    Dim i As Long, i_max As Long, i_pos As Long
    
    i_max = ws.Cells(Rows.Count, id_pos).End(xlUp).Row
    
    For i = 2 To i_max
        id_cel = CStr(ws.Cells(i, id_pos).Value)
        If id = id_cel Then
            i_pos = i
            i = i_max
        End If
    Next i
    If i_pos = 0 Then i_pos = i
    
    'MsgBox "i_max = " & i_max & " i = " & i & " & i_pos = " & i_pos
    item_pos_by_id = i_pos
    
    Exit Function
er:
    MsgBox "Error: item_pos_by_id [" & Erl & "] " & Err.Description
End Function
Function html_to_doc(html As String) As HTMLDocument
    On Error GoTo er
    Dim tmp As HTMLDocument
    Set tmp = CreateObject("HTMLFile")
    tmp.body.innerHTML = html
    
    Set html_to_doc = tmp
Exit Function
er:
    MsgBox "Error: html_to_doc [" & Erl & "] " & Err.Description
End Function
Sub html_to_file(html As String)
    On Error GoTo er
    
    Open ThisWorkbook.Path & "\Output.html" For Output As #1
    Print #1, html
    Close #1
        
Exit Sub
er:
    MsgBox "Error: html_to_file [" & Erl & "] " & Err.Description
End Sub
