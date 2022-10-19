Private Sub Workbook_Open()

    Dim GM_Semana As Worksheet, BD As Worksheet
    Set GM_Semana = Sheets("GM_Semana")
    Set BD = Sheets("BD")
    
    ' ---Runs modules that delete data and formatting since the last time---
    delete_data.delete_data
    apagar_gm_dia.apagar_gm_dia
    
    GM_Semana.Activate
    GM_Semana.Range("B4:BA4").Select
    GM_Semana.Range(Selection, Selection.End(xlDown)).Select
    
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    '---Refreshes the powerquery table and sorts the C column in ascending order---
    BD.Activate
    BD.ListObjects(1).Refresh
    
    Columns("C:C").Select
        Selection.TextToColumns Destination:=Range("C:C"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 4), TrailingMinusNumbers:=True
            
    ActiveWorkbook.Worksheets("BD").ListObjects("Registos").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("BD").ListObjects("Registos").Sort.SortFields.Add _
            Key:=Range("Registos[[#All],[Data inicio visita]]"), SortOn:=xlSortOnValues _
            , Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("BD").ListObjects("Registos").Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
           
    '---Runs the dedicated module for each person---
    Person1.run
    Person2.run
    Person3.run
    Person4.run
    Person5.run
    Person6.run
    Person7.run
End Sub
