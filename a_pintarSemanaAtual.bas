Attribute VB_Name = "a_pintarSemanaAtual"
Option Explicit

Sub semana_atual()
Attribute semana_atual.VB_ProcData.VB_Invoke_Func = " \n14"
'
' semana_atual Macro
'
'------------- JUNK IT--------------------

'Defininerá variaveis
Dim DataHist                  As Date
Dim DataCad                   As Date
Dim UltDia                    As Date
'Dim mard                      As Worksheet


Dim ws_dep                    As Worksheet
Dim state                     As Worksheet

Dim historico_auditoria       As Worksheet
Dim indicadores_m             As Worksheet
Dim SemAtual                  As Integer
Dim SemAno                    As Integer
Dim SemRest                   As Integer
Dim ItensTot                  As Integer
Dim ItensMed                  As Integer
Dim SelItem                   As Integer
Dim i                         As Integer
Dim j                         As Integer
Dim ItensTotDep               As Integer
Dim counter                   As Integer
Dim ItensTotFull As Integer
Dim condicao    As String
Dim mesCompleto() As String
Dim DataCompleta() As String
Dim dia As String
Dim mes As String
Dim ano As String
Dim mes_ind As String
Dim db_per_dep  As Double
Dim ItensPer    As Double
Dim msgResp As VbMsgBoxResult
Dim d
'---------------------------- --------------------


Dim historico                 As Worksheet
Dim Menu                      As Worksheet
Dim aux                       As Worksheet
Dim deps As Variant
Dim size As Integer
Dim dep As String
Dim dep0010 As Integer
Dim dep0020 As Integer
Dim dep0030 As Integer
Dim dep0040 As Integer
Dim dep0041 As Integer
Dim dep0050 As Integer
Dim dep0060 As Integer
Dim dep0080 As Integer




'Selecionará uma variável já declarada e atribuirá a ela um caminho para uma guia
Set aux = Sheets("Auxiliar")
Set historico = Sheets("historico")
Set Menu = Sheets("MENU")




Call desbloquear_guias

    d = Date
    DataCompleta = Split(d, "/")
    dia = DataCompleta(0)
    mes_ind = DataCompleta(1)
    ano = DataCompleta(2)
    d = mes_ind & "/" & dia & "/" & ano  'DATA CONVERTIDA PARA PADRAO USA
historico.Select
    '==EXCLUIR=== APENAS P TESTE
    'd = "11/" & CStr(Range("J13").Value) & "/2022"
    '===========================
    
    Range("A1").Select
    If ActiveCell.ListObject.ShowAutoFilter = False Then
       Range("A1:G1").AutoFilter
    Else
        Range("A1:G1").AutoFilter
        Range("A1:G1").AutoFilter
    End If
    
'Array armazenando strings
    deps = Array("0010", "0020")



size = 0

While (size <= UBound(deps))
        dep = deps(size)
        historico.Select
        ActiveSheet.ListObjects("historico").Range.AutoFilter Field:=3, Criteria1:= _
        dep
            
    '************************** FILTRA DATAS NO MES ATUAL *************************
        historico.Activate

          ActiveSheet.Range("$G$1:$G$" & Range("G1").End(xlDown).Row).AutoFilter Field:=7, Operator:= _
               xlFilterValues, Criteria2:=Array(1, d)
    '*******************************************************************************
        aux.Activate
        Range("H:H").Delete
    
        historico.Activate
        Range("G1").Select
        Range(Selection, Selection.End(xlDown)).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
    
      
        '********************************* DATA + HORA --->  DATA **************
        aux.Activate
        Range("h1").Select
        Selection.PasteSpecial
        Range("h1").Select
        While ActiveCell.Value <> ""
        
        
        mesCompleto = Split(Selection.Value, " ")
        Selection.Value = mesCompleto
        ActiveCell.Offset(1, 0).Select
        Wend
        
        '********************* RETIRAR DUPLICATAS e ESTABELECER O NÚMERO DA SEMANA DO MÊS ******
        
        Range("H1").Select
        If Selection.Value <> "" Then
        Range(Selection, Selection.End(xlDown)).Select
        ActiveSheet.Range("$H$1:$H$" & Range("H1").End(xlDown).Row).RemoveDuplicates Columns:=1, Header:=xlNo
        Range("h1").Select
            If ActiveCell.Offset(1, 0) <> "" Then
            j = Range("H1").End(xlDown).Row
            Else
            j = ActiveCell.Row
            End If
            
        End If
        If ActiveCell.Value = "" Then
            ActiveCell.Value = ""
            j = -1
        End If
        
        
    Select Case dep
    
'        Case "0010"
'         dep0010 = j
'        Case "0020"
'        dep0020 = j
'
'        Case "0030"
'        dep0030 = j
'
'        Case "0040"
'         dep0040 = j
            
        Case "0010"
        dep0010 = j
        
'        Case "0050"
'        dep0050 = j
'
        Case "0020"
        dep0020 = j
        
'        Case "0080"
'        dep0080 = j
        
    End Select





size = size + 1
Wend




    If (dep0010 = -1) And (dep0020 = -1) Then
    
        Menu.Select
        Range("A5").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight2
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With

        Range("B5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With

        Range("C5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With
    
        Range("D5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With
    
    ElseIf (dep0010 = 1) And (dep0020 = 1) Then

        Menu.Select
        Range("B5").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight2
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        
        Range("A5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With

        Range("C5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With
    
        Range("D5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With
        

        
    ElseIf (dep0010 = 2) And (dep0020 = 2) Then

        Menu.Select
        Range("C5").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight2
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        
                Range("A5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With

        Range("B5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With
    
        Range("D5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With
    
    ElseIf (dep0010 = 3) And (dep0020 = 3) Then
          
        Menu.Select
        Range("D5").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight2
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        
                Range("A5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With

        Range("B5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With
    
        Range("C5").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With
        
    End If
    
    Menu.Select


                Call bloquear_guias

End Sub
Sub paintWhite()
Attribute paintWhite.VB_ProcData.VB_Invoke_Func = " \n14"
'
' paintWhite Macro
'

'
    Range("D5").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
End Sub
