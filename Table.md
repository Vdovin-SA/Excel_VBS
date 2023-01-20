# Перебор всех ячеек таблицы

## Простая таблица
Желательно указывать начальную позицию таблицы.
Желательно хранить настройки табличной части.
Создавать именованные списки, что бы упростить маппинг колонок
```
Sub Table_Simple()
    On Error GoTo er
    Const List_Name = "List1"
    
    Dim ws As Worksheet 
    Set ws = Worksheets(List_Name)
    
    Dim r_max As Long, c_max As Long
    r_max = ws.Cells(Rows.Count, 1).End(xlUp).Row
    c_max = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Перебор таблицы / всех заполненных ячеек
    For r = 1 To r_max
        For c = 1 To c_max
            MsgBox r & " " & c & " is " & ws.Cells(r, c).Value
        Next
    Next
    
    Exit Sub
er:
    MsgBox "Error: Table_Simple() [" & Err.Number & "] " & Err.Description
End Sub
```

## Умная таблица
Не надо знать стартовую позицию таблицы. 
Возможно размещать несколько таблиц на одном листе и обращаться к ним по имени.
Возможно обращаться к колонкам по имени.
Возможно создание автоматической строки итогов.
```
Sub Table_Smart()
    On Error GoTo er
    Const List_Name = "List2"
    Const Table_Name = "Users"
    Const Collumn_Name = "Name"

    Dim list As ListObject
    Set list = Worksheets(List_Name).ListObjects(Table_Name)
    
    ' Перебор ячеек определенного столбца 
    For Each Item In list.ListColumns(Collumn_Name).DataBodyRange
        MsgBox Collumn_Name & " is " & CStr(Item.Value)
    Next Item
    
    ' Перебор всех заголовков
    For c = 1 To list.ListColumns.Count
        MsgBox c & " is " & list.ListColumns(c).Name
    Next
    
    ' Перебор всех ячеек табличной части
    For r = 1 To list.ListRows.Count
        For c = 1 To list.ListColumns.Count
            MsgBox r & " " & c & " is " & list.DataBodyRange.Cells(r, c).Value
        Next
    Next

Exit Sub
er:
    MsgBox "Error: Table_Smart() [" & Err.Number & "] " & Err.Description
End Sub
```

## Выделенный диапазон
```
Sub Table_Range()
    On Error GoTo er
    
    Dim r_start As Long, r_end As Long
    r_start = Selection.Cells(1).Row
    r_end = Selection.Cells(Selection.Cells.Count).Row
    
    Dim c_start As Long, c_end As Long
    c_start = Selection.Column
    c_end = c_start + Selection.Columns.Count - 1
    MsgBox r_start & ":" & c_start & " -> " & r_end & ":" & c_end
    
    ' Перебор таблицы
    For r = r_start To r_end
        For c = c_start To c_end
            MsgBox r & " " & c & " is " & Cells(r, c).Value
        Next
    Next
    
    Exit Sub
er:
    MsgBox "Error: Table_Range() [" & Err.Number & "] " & Err.Description
End Sub
```
