# Перебор всех ячеек таблицы

## Простая таблица
```
some code
```

## Умная таблица
Не надо знать стартовую пазицию таблицы. Возможно размещать несколько таблиц на одном листе и обращаться к ним по имени.
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
some code
```
