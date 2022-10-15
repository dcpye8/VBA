Attribute VB_Name = "AnaVaz"
Sub run()

nome_BD = "Ana Vaz"
linha_gm_semana = 4
nome_folha = "ANA VAZ"

Dim BaseDados As Worksheet, i As Integer, folha_gm As Worksheet, dict_mes, GM_Semana As Worksheet

Set BaseDados = Sheets("BD")
Set folha_gm = Sheets(nome_folha)
Set GM_Semana = Sheets("GM_Semana")
Set dict_mes = CreateObject("Scripting.Dictionary")

dict_mes.Add 1, "JANEIRO"
dict_mes.Add 2, "FEVEREIRO"
dict_mes.Add 3, "MARÇO"
dict_mes.Add 4, "ABRIL"
dict_mes.Add 5, "MAIO"
dict_mes.Add 6, "JUNHO"
dict_mes.Add 7, "JULHO"
dict_mes.Add 8, "AGOSTO"
dict_mes.Add 9, "SETEMBRO"
dict_mes.Add 10, "OUTUBRO"
dict_mes.Add 11, "NOVEMBRO"
dict_mes.Add 12, "DEZEMBRO"

ActiveWorkbook.Worksheets("BD").ListObjects("Registos") _
    .QueryTable.Refresh BackgroundQuery:=False

i = 1
k = 3
mes = 3

folha_gm.Activate
folha_gm.Range("A3:F3").Select
folha_gm.Range(Selection, Selection.End(xlDown)).Select


With Selection
    Selection.ClearContents
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlCenter

End With

With Selection.Font
    .Name = "Tahoma"
    .Size = 9
    .ColorIndex = xlAutomatic
End With

With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With


Do Until BaseDados.Cells(i, 1).Value = ""

If BaseDados.Cells(i, 1).Value = nome_BD Then
    
    mes_visita = Month(BaseDados.Cells(i, 3).Value)
    
    If dict_mes(mes_visita) = folha_gm.Cells(mes, 1).Value Then
        
        folha_gm.Activate
        folha_gm.Cells(k, 1).Value = BaseDados.Cells(i, 6).Value   'Cliente
        folha_gm.Cells(k, 2).Value = BaseDados.Cells(i, 7).Value   'Classificação Cliente
        folha_gm.Cells(k, 3).Value = BaseDados.Cells(i, 8).Value   'Tipo de Visita
        folha_gm.Cells(k, 4).Value = BaseDados.Cells(i, 10).Value  'Coleção
        
        duracao_visita = BaseDados.Cells(i, 4).Value
        dia_visita = Day(BaseDados.Cells(i, 3).Value)
        
        dia_final = dia_visita + duracao_visita
        
        folha_gm.Cells(k, 5).Value = dia_visita
        folha_gm.Cells(k, 6).Value = dia_final
    
        k = k + 1
        
        dia_semana = WorksheetFunction.WeekNum(BaseDados.Cells(i, 3).Value)
        
        If dia_semana = 0 Then
        Else
            GM_Semana.Activate
            GM_Semana.Cells(linha_gm_semana, dia_semana + 1).Select
            
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 10498160
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    

    Else
        mes = k
        folha_gm.Activate
        folha_gm.Cells(k, 1).Value = dict_mes(mes_visita)
        
        
        Range(folha_gm.Cells(k, 1), folha_gm.Cells(k, 6)).Select
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.499984740745262
            .PatternTintAndShade = 0
        End With
        
        k = k + 1
        
        folha_gm.Cells(k, 1).Value = BaseDados.Cells(i, 6).Value   'Cliente
        folha_gm.Cells(k, 2).Value = BaseDados.Cells(i, 7).Value   'Classificação Cliente
        folha_gm.Cells(k, 3).Value = BaseDados.Cells(i, 8).Value   'Tipo de Visita
        folha_gm.Cells(k, 4).Value = BaseDados.Cells(i, 10).Value  'Coleção
        
        duracao_visita = BaseDados.Cells(i, 4).Value
        dia_visita = Day(BaseDados.Cells(i, 3).Value)
        
        dia_final = dia_visita + duracao_visita
                
        folha_gm.Cells(k, 5).Value = dia_visita
        folha_gm.Cells(k, 6).Value = dia_final
    
        k = k + 1
        
        dia_semana = WorksheetFunction.WeekNum(BaseDados.Cells(i, 3).Value)
        
        If dia_semana = 0 Then
        Else
            GM_Semana.Activate
            GM_Semana.Cells(linha_gm_semana, dia_semana + 1).Select
            
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 10498160
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        
    End If

End If

i = i + 1

Loop
End Sub


