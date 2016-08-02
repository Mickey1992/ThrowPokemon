Attribute VB_Name = "模块1"
Public Const MIN_CP As Integer = 500
Public Const MIN_IVPERF As Integer = 60

Sub getDeleteList()
    Dim analyze_sheet As Worksheet
    Dim delete_sheet As Worksheet
    Dim all_pokemon_sheet As Worksheet
    Set all_pokemon_sheet = Worksheets("Sheet")
    Set analyze_sheet = Worksheets.Add(After:=Worksheets("Sheet"))
    analyze_sheet.Name = "pokemon_count"
    
    Call getPokemonId(all_pokemon_sheet, analyze_sheet)
    Call getUselessPokemon(delete_sheet)
    Call countAllPokemon(analyze_sheet, all_pokemon_sheet, delete_sheet)
    Call savePokemon(delete_sheet, analyze_sheet)
    
    delete_sheet.Activate
End Sub

'获取当前持有的所有pokemon的ID
Function getPokemonId(from_sheet As Variant, to_sheet As Variant)
    'Copy All Pokemon ID
    last_row = from_sheet.Cells(Rows.Count, 1).End(xlUp).Row
    Worksheets("Sheet").Range("A1:A" + CStr(last_row)).Copy to_sheet.Range("A:A")
    
    'Remove Duplicated Pokemon ID
    to_sheet.Range("A:A").RemoveDuplicates Columns:=Array(1), Header:=xlYes
End Function

'计算指定sheet内一种pokemon的持有个数
Function countPokemon(count_sheet As Variant, pokemon_id As Integer)
    result = Application.WorksheetFunction.CountIf(count_sheet.[A:A], pokemon_id)
    countPokemon = result
End Function

'计算各种pokemon的持有总数以及(CP<=500,IVPerf<=60%)的数量
Function countAllPokemon(count_sheet As Variant, _
                         all_pokemon_sheet As Variant, _
                         to_delete_sheet As Variant)
    With count_sheet
        .Activate
        Dim max_row As Integer
        max_row = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set Header
        Cells(1, 2).Value = "pokemon总数"
        Cells(1, 3).Value = "(CP<=500,IVPerf<=60%)pokemon总数"
        
        'Count
        For i = 2 To max_row
            Cells(i, 2).Value = countPokemon(all_pokemon_sheet, Cells(i, 1).Value)
            Cells(i, 3).Value = countPokemon(to_delete_sheet, Cells(i, 1).Value)
        Next
    End With
End Function


'抽出CP<=500,IVPerf<=60%的pokemon
Function getUselessPokemon(ByRef delete_sheet As Variant)
    Set delete_sheet = Worksheets.Add(After:=Worksheets("Sheet"))
    delete_sheet.Name = "pokemons_tobe_thrown"
    Set ws = Worksheets("Sheet")
    With ws
        .Activate
        Dim max_row As Integer
        Dim max_column As Integer
        Dim insert_row As Integer
        insert_row = 2
        max_row = Cells(Rows.Count, 1).End(xlUp).Row
        max_column = Cells(1, Columns.Count).End(xlToLeft).Column
              
        'Copy Header
        Range(Cells(1, 1), Cells(1, max_column)).Copy delete_sheet.Cells(1, 1)
        
        'Copy Pokemon Data
        insert_row = 2
        For i = 2 To max_row
            If ((Cells(i, 6).Value <= MIN_CP) And (Cells(i, 12).Value <= MIN_IVPERF)) Then
                .Range(Cells(i, 1), Cells(i, max_column)).Copy delete_sheet.Cells(insert_row, 1)
                insert_row = insert_row + 1
            End If
        Next
        
        'Sort ID,IvPerf (desc)
        delete_sheet.Range("A1:N" & CStr(insert_row - 1)).Sort _
                        Key1:=delete_sheet.Range("A2"), Order1:=xlAscending, _
                        Key2:=delete_sheet.Range("L2"), Order2:=xlDescending, _
                        Header:=xlYes
    End With
End Function

'如果某一种pokemon都在delete list里，则将完美度最高的从list种移除
Function savePokemon(to_delete_sheet As Variant, count_sheet As Variant)
    Dim d_max_row As Integer
    d_max_row = to_delete_sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    With count_sheet
        .Activate
        Dim max_row As Integer
        max_row = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 1 To max_row
            If (Cells(i, 3).Value = Cells(i, 2).Value) Then
                Dim delete_row As Integer
                delete_row = to_delete_sheet.Range("A1:A" & CStr(d_max_row)).Find(Cells(i, 1).Value).Row
                to_delete_sheet.Rows(delete_row).EntireRow.Delete
                d_max_row = to_delete_sheet.Cells(Rows.Count, 1).End(xlUp).Row
            End If
        Next
    End With
End Function
