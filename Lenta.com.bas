Attribute VB_Name = "Module1"
Public ws_store As Worksheet
Public ws_skus As Worksheet
Public ws_price As Worksheet

Public Const SRV = "https://lenta.com/api/"
' class ApiMethods:
Const GET_CITIES = "v1/cities"
Const GET_STORES = "v1/stores"
Const GET_STORE = "v1/stores/{store_id}"                  ' GET     Получение магазина
Const GET_CITY_STORES = "v1/cities/{city_id}/stores"      ' GET     Получение списка магазинов Лента для города
Public Const STORE_SKUS = "v1/stores/{store_id}/skus"            ' POST    Поиск товара в магазине
Const STORE_SKUS_LIST = "v1/stores/{store_id}/skusList"
Const GET_STORE_SKUS = "v1/stores/{store_id}/skus/{code}" ' POST    Получение товаров магазина по иденификаторам товаров
Const GET_CATALOG = "v2/stores/{store_id}/catalog"
Public Const TYPE_UPDATE_YES = "Обновлять"
Public Const TYPE_UPDATE__NO = "Не обновлять"

Sub Кнопка1_Щелчок()
    On Error GoTo er

    Call init
    
    Dim user_form As New Setup
    user_form.Show

    Exit Sub
er:
    MsgBox "Error: " & Err.Description & Err.Number
End Sub
' Инициализация настроек
Sub init()
On Error GoTo er

    Set ws_store = Worksheets("Магазины")
    Set ws_skus = Worksheets("Товар")
    Set ws_price = Worksheets("Цены")

Exit Sub
er:
    MsgBox "Error: init [" & Erl & "] " & Err.Description
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
Function HTML_Post_Content(strURL As String, request As String)
    html = ""
    Set MyBrowser = CreateObject("MSXML2.XMLHTTP")
    
    MyBrowser.Open "POST", strURL, False
    MyBrowser.setRequestHeader "User-Agent", "Mozilla/5.0"
    MyBrowser.setRequestHeader "Content-type", "application/json"
    MyBrowser.setRequestHeader "Accept", "application/json"
    MyBrowser.send request
    
    If MyBrowser.Status = 200 Then
        html = MyBrowser.responseText
       ' Убираем спец символы
        html = Replace(html, Chr(9), "")
        html = Replace(html, Chr(10), "")
        html = Replace(html, Chr(13), "")
    End If
    
    HTML_Post_Content = html
    Set MyBrowser = Nothing
End Function
' Преобразование текста в JSON обьект
Function JSON_From_Text(json_test As String) As Object
    Set JSON_From_Text = JsonConverter.ParseJson(json_test)
End Function
' Сохранение контента в файл
Sub html_to_file(html As String)
    On Error GoTo er
    
    Open ThisWorkbook.Path & "\Output.html" For Output As #1
    Print #1, html
    Close #1
        
Exit Sub
er:
    MsgBox "Error: html_to_file [" & Erl & "] " & Err.Description
End Sub
' Получение списка магазинов
Sub get_stores_list()
    On Error GoTo er
    Dim response As String
    response = HTML_Get_Content(SRV & GET_STORES)
    'Call html_to_file(response)
    
    Dim json As Object
    Set json = JSON_From_Text(response)

    ' Заполнение табличной части
    i = 2
    For Each item In json
        ws_store.cells(i, 1) = id_prepare(CStr(item("id")), False)
        ws_store.cells(i, 2) = item("cityName")
        ws_store.cells(i, 3) = item("address")
        ws_store.cells(i, 4) = item("type")
        ws_store.cells(i, 5) = TYPE_UPDATE__NO
        With ws_store.cells(i, 5).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=TYPE_UPDATE__NO & "," & TYPE_UPDATE_YES
        End With
        i = i + 1
    Next
    
Exit Sub
er:
    MsgBox "Error: get_stores_list [" & Erl & "] " & Err.Source & " " & Err.Description
End Sub
' Получение / обновление списка товаров по ID
Sub get_skus_list()
    On Error GoTo er
    Dim query As String
    query = "{""skuCodes"":  [" & skus_get_all_id() & "]}"
    Dim url As String
    url = Replace(SRV & STORE_SKUS_LIST, "{store_id}", "0073")
    MsgBox url
    Dim response As String
    response = HTML_Post_Content(url, query)
    'Call html_to_file(response)
    
    Dim json As Object
    Set json = JSON_From_Text(response)
    
    ' перезаписываем все значения
    Dim i As Long
    i = 1
    For Each item In json
        i = i + 1
        Call skus_update_cell(i, item)
    Next item
Exit Sub
er:
    MsgBox "Error: get_skus_list [" & Erl & "] " & Err.Source & " " & Err.Description
End Sub
' Заполнение ряда ячеек информацией о товаре
Sub skus_update_cell(i As Long, item)
    On Error GoTo er
    
    ws_skus.cells(i, 1).Value = id_prepare(CStr(item("code")), False)
    ws_skus.cells(i, 2).Value = item("title")
    ws_skus.cells(i, 3).Value = item("regularPrice")
    ws_skus.cells(i, 4).Value = item("discountPrice")
    ws_skus.cells(i, 5).Value = item("skuWeight")
    ws_skus.cells(i, 6).Value = item("categories")("group")("name")
    'ws_skus.cells(i, 7).Value = item("webUrl")
    Call ws_skus.Hyperlinks.Add(ws_skus.cells(i, 7), item("webUrl"), "", "", "Ссылка")
    
    ' Добавление примечания
    Dim cell As range
    Set cell = ws_skus.cells(i, 2) ' выделяем колонку для изображения
    cell.ClearComments ' Очистка примечания
    If cell.comment Is Nothing Then ' Если примечание создано, то не обновляем картикну
        Dim comment As comment
        Set comment = cell.AddComment ' Создание примечания
        comment.Shape.Fill.UserPicture (item("image")("medium"))
        comment.Shape.Height = 220 ' Устанавливаем размеры окна примечания
        comment.Shape.Width = 220
    End If
Exit Sub
er:
    MsgBox "Error: skus_update_cell [" & Erl & "] " & Err.Source & " " & Err.Description
End Sub
' Накопление информации о ценах по магазинам
Sub get_price_list()
    On Error GoTo er
    Dim store_id As String
    Dim url As String
    Dim response As String
    Dim json As Object
    
    Dim list As ListObject
    Set list = ws_store.ListObjects("store")
    
    Dim query As String
    query = "{""skuCodes"":  [" & skus_get_all_id() & "]}"
    Dim i As Long
    i = ws_price.cells(Rows.Count, 1).End(xlUp).Row
    
    For Each Row In list.ListColumns("Действие").DataBodyRange
        If Row.Value = TYPE_UPDATE_YES Then
            store_id = CStr(Intersect(Row.EntireRow, list.ListColumns("ID").DataBodyRange))
            store_id = id_prepare(store_id, True)
            
            url = Replace(SRV & STORE_SKUS_LIST, "{store_id}", store_id)
        
            response = HTML_Post_Content(url, query)
            Set json = JSON_From_Text(response)
            
            For Each item In json
                i = i + 1
                ws_price.cells(i, 1).Value = Now
                ws_price.cells(i, 2).Value = item("title")
                ws_price.cells(i, 3).Value = id_prepare(CStr(item("code")), False)
                ws_price.cells(i, 4).Value = id_prepare(store_id, False)
                ws_price.cells(i, 5).Value = item("regularPrice")
                ws_price.cells(i, 6).Value = item("discountPrice")
            Next item
            
        End If
    Next Row
    
Exit Sub
er:
    MsgBox "Error: get_price_list [" & Erl & "] " & Err.Source & " " & Err.Description
End Sub
Function id_prepare(id As String, is_clean As Boolean) As String
    If is_clean Then
        id_prepare = Replace(id, "_", "")
    Else
        id_prepare = "_" & id & "_"
    End If
End Function
Function skus_get_all_id() As String
    Dim out As String
    Dim list As ListObject
    Set list = ws_skus.ListObjects("skus")
    
    For Each item In list.ListColumns("ID").DataBodyRange
        If out <> "" Then out = out & ", "
        out = out & """" & id_prepare(CStr(item.Value), True) & """"
    Next item
    
    skus_get_all_id = out
End Function
