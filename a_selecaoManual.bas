Attribute VB_Name = "a_selecaoManual"
Option Explicit

Sub SelecaoAutomatica()
'Defininerá variaveis
Dim DataHist                  As Date
Dim DataCad                   As Date
Dim UltDia                    As Date
'Dim mard                      As Worksheet
Dim historico                 As Worksheet
Dim Menu                      As Worksheet
Dim ws_dep                    As Worksheet
Dim state                     As Worksheet
Dim dep0010                   As Worksheet
Dim dep0020                   As Worksheet
Dim dep0030                   As Worksheet
Dim dep0041                   As Worksheet
Dim dep0060                   As Worksheet
Dim dep0080                   As Worksheet
Dim historico_auditoria       As Worksheet
Dim indicadores_m             As Worksheet
Dim SemAtual                  As Integer
Dim SemAno                    As Integer
Dim SemRest                   As Integer
'Dim ItensTot                  As Integer
Dim ItensTot                  As Long
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
Dim db_per_dep  As Double
Dim ItensPer    As Double
Dim msgResp As VbMsgBoxResult
Dim d


'Selecionará uma variável já declarada e atribuirá a ela um caminho para uma guia
    Set mard = Sheets("MARD")
    Set historico = Sheets("historico")
    Set Menu = Sheets("MENU")



    'Definirá um novo valor para a variável str_dep
        If auto_run = "" Then
            Call desbloquear_guias
            str_dep = Menu.Range("D2").Value
            Menu.Range("H:H").Delete

            mard.Select
            Range("A1").Select
            Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
        End If

        'auto_run = "off"


     'Selecionará uma variável já declarada e atribuirá a ela um caminho para uma guia
        Set ws_dep = Sheets(str_dep)
        Set state = Sheets("state")



        mard.Select
        Range("A1").Select

        If ActiveCell.ListObject.ShowAutoFilter = False Then
           Range("C1:D1").AutoFilter
        Else
            Range("A1:I1").AutoFilter
            Range("A1:I1").AutoFilter
        End If


    'Atribuir a variável DataHist o valor de Data Atual+Hora,min,s Atual
        DataHist = Now

    'Na coluna 3, seleciona-se o filtro procurando pelo valor contido na variável str_dep
        ActiveSheet.ListObjects("MARD").Range.AutoFilter Field:=3, Criteria1:= _
            str_dep
    'Na coluna 3, seleciona-se o filtro procurando pelo valores diferentes de zero ou "<>0"
        ActiveSheet.ListObjects("MARD").Range.AutoFilter Field:=4, Criteria1:="<>0" _
            , Operator:=xlAnd

        historico.Activate

        Range("A1").Select
    'Configuração repete-se exatamente como já ocorrida no código
        If ActiveCell.ListObject.ShowAutoFilter = False Then
           Range("A1:H1").AutoFilter
        Else
            Range("A1:H1").AutoFilter
            Range("A1:H1").AutoFilter
        End If


    'Limpará a ordenação dos campos
        ActiveWorkbook.Worksheets("historico").ListObjects("historico").Sort.SortFields.Clear
    'Deixará a coluna "Data-Hora" do maior para o menor = "Ascending"
        ActiveWorkbook.Worksheets("historico").ListObjects("historico").Sort.SortFields. _
        Add2 Key:=Range("historico[[#All],[DATA-HORA]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    'Irá configurar a coluna "DATA-HORA" do mais antigo para o mais novo
        ActiveWorkbook.Worksheets("historico").ListObjects("historico").Sort.Apply

        ActiveSheet.ListObjects("historico").Range.AutoFilter Field:=3, Criteria1:= _
        str_dep





        If Range("A1048576").End(xlUp).Row = 2 And Range("A1048576").End(xlUp) = "" Then
            Range("A1048576").End(xlUp).Select

            If str_dep = "0050" Then

            Range("A1").Select
            'Configuração repete-se exatamente como já ocorrida no código
            If ActiveCell.ListObject.ShowAutoFilter = False Then
               Range("A1:H1").AutoFilter
            Else
                Range("A1:H1").AutoFilter
                Range("A1:H1").AutoFilter
            End If

            End If
        Else
        '1050 Ao invés de usar CurrentRegion, usar seleção de tabela
            Range("A1").CurrentRegion.Offset(1, 0).SpecialCells(xlCellTypeVisible).Rows(1).Select

            If str_dep = "0050" Then  '                                                                  ADAPTAÇÃO PARA O DEP 50!

                Range("A1").Select
                'Configuração repete-se exatamente como já ocorrida no código
                If ActiveCell.ListObject.ShowAutoFilter = False Then
                   Range("A1:H1").AutoFilter
                Else
                    Range("A1:H1").AutoFilter
                    Range("A1:H1").AutoFilter
                End If

            Range("A1").CurrentRegion.Offset(1, 0).SpecialCells(xlCellTypeVisible).Rows(1).Select


            End If
        End If



                                                    '********************************
                                                    '             0050
                                                    '********************************
        If str_dep = "0050" Then

            mard.Activate
    '--------- bloco responsável por dizer quantos itens visíveis estão sendo apresentados na presente planilha ---------------
            Range("P1").Select
            Application.CutCopyMode = False
            ActiveCell.FormulaR1C1 = "=AGGREGATE(3,7,MARD[[#All],[MATNR]])"
            ItensTot = Range("P1").Value
            ItensPer = ItensTot - 1 'TOTAL MARD
    '-------------------------------------------------------------------------------------------------------------------------:.

            Range("A1", "F" & Range("A1048576").End(xlUp).Row - 1).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy
            historico.Activate

            If Range("A1").Offset(1, 0) = "" Then
                    Range("A2").Select
            Else
                    Range("A1").End(xlDown).Offset(1, 0).Select
            End If

            ActiveSheet.Paste
            Range("G" & ActiveCell.Offset(0, 6).Row, "G" & Range("G1048576").End(xlUp).Row).Select
            Selection = DataHist

    '--------- bloco responsável por dizer quantos itens visíveis estão sendo apresentados na presente planilha ---------------
            ItensTot = Selection.Row
            i = Range("G1048576").End(xlUp).Row
            ItensTot = ((i - ItensTot) + 1)  'TOTAL HIST

    '-------------------------------------------------------------------------------------------------------------------------:.

        End If

    'Se depósito for = "0050", então pula para mostrar os dados e finalizar o código                                                          \!/ 0050 \!/
    If str_dep <> "0050" Then
    
    If (str_dep = "0041") Or (str_dep = "0060") Then



        historico.Activate

    If str_dep = "0041" Then

    'Na coluna 3, seleciona-se o filtro procurando pelo valor contido na variável str_dep
    str_dep = "0030"
        historico.Activate

        Range("A1").Select
    'Configuração repete-se exatamente como já ocorrida no código



        ActiveSheet.ListObjects("historico").Range.AutoFilter Field:=3, Criteria1:= _
        str_dep

    End If

    On Error GoTo here
        ActiveSheet.Range("$G$1:$G$" & Range("G1").End(xlDown).Row).AutoFilter Field:=7, Operator:= _
            xlFilterValues, Criteria2:=Array(1, Month(Now()) - 1 & "/1/2023")
    '*******************************************************************************
    GoTo keepIt
'--------------------- ESCOPO PARA TRATAMENTO DE ERRO ------------------------
here:

MsgBox "Ocorreu um Erro!", vbCritical
Stop
If Month(Now()) = 1 Then

        ActiveSheet.Range("$G$1:$G$" & Range("G1").End(xlDown).Row).AutoFilter Field:=7, Operator:= _
            xlFilterValues, Criteria2:=Array(1, Month(Now()) & "/1/2023")

End If
'-----------------------------------------------------------------------------

'On Error GoTo -1

keepIt:
        Menu.Activate
        Range("H:H").Delete

        historico.Activate
        Range("G1").Select
        Range(Selection, Selection.End(xlDown)).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy


        '********************************* DATA + HORA --->  DATA **************
        Menu.Activate
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
            i = Range("H1").End(xlDown).Row
            Else
            i = 1
            i = i + 1
            counter = i
            End If



        If str_dep = "0030" Then

        str_dep = "0041"

        End If

        End If

On Error GoTo 0



        ' \!/
    If str_dep = "0041" Then

    'Na coluna 3, seleciona-se o filtro procurando pelo valor contido na variável str_dep

        historico.Activate

        Range("A1").Select
    'Configuração repete-se exatamente como já ocorrida no código



        ActiveSheet.ListObjects("historico").Range.AutoFilter Field:=3, Criteria1:= _
        str_dep

    End If

    '************************** FILTRA DATAS NO MES ATUAL *************************
        historico.Activate
        ActiveSheet.Range("$G$1:$G$" & Range("G1").End(xlDown).Row).AutoFilter Field:=7, Operator:= _
            xlFilterValues, Criteria2:=Array(1, Month(Now()) & "/1/2023")
    '*******************************************************************************
        Menu.Activate
        Range("H:H").Delete

        historico.Activate
        Range("G1").Select
        Range(Selection, Selection.End(xlDown)).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy


        '********************************* DATA + HORA --->  DATA **************
        Menu.Activate
        Range("h1").Select
        Selection.PasteSpecial
        Range("h1").Select
        While ActiveCell.Value <> ""


        mesCompleto = Split(Selection.Value, " ")
        Selection.Value = mesCompleto
        ActiveCell.Offset(1, 0).Select
        Wend

        '********************* RETIRAR DUPLICATAS e ESTABELECER O NÚMERO DA SEMANA DO MÊS ******
        'Define o "i" para o cálculo do valor de itens que será selecionado para o inventário
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

    End If
    '==================================================================================================


    '---------------------------------- VERIFICAÇÃO P REPOR DEP 0080 -------------
    If str_dep = "0080" Or str_dep = "0041" Then

    '************************** FILTRA DATAS NO MES ATUAL *************************
        historico.Activate
        ActiveSheet.Range("$G$1:$G$" & Range("G1").End(xlDown).Row).AutoFilter Field:=7, Operator:= _
            xlFilterValues, Criteria2:=Array(1, Month(Now()) & "/1/2023")
    '*******************************************************************************
        Menu.Activate
        Range("H:H").Delete

        historico.Activate
        Range("G1").Select
        Range(Selection, Selection.End(xlDown)).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy


        '********************************* DATA + HORA --->  DATA **************
        Menu.Activate
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




    End If

        If ActiveCell = "" Or (str_dep = "0080" And j = -1) Or (str_dep = "0041" And i = 4 And j = -1) Or (str_dep = "0060" And i = 4 And j = -1) Then     'Or (str_dep = "0041" And j <= 4 And i = 4)                       '\!/ REVISAR LÓGICA COM O GREG!!! \!/
            ws_dep.Activate
            Range("A1:XFD1048576").Delete
            mard.Activate
            'Range("A:F").Copy
            Range("A1", "F" & Range("A1").End(xlDown).Row).SpecialCells(xlCellTypeVisible).Copy
            ws_dep.Activate

            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

            ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = _
            "Tabela" & str_dep

            Range("Tabela" & str_dep & "[#All]").Select
            ActiveSheet.ListObjects("Tabela" & str_dep).TableStyle = "TableStyleMedium7"

        Else
    'Conta a partir da célula ativa 1 para baixo e 6 para direita, o valor será armazenado em DataCad
            historico.Activate
            DataCad = Range("A1").Offset(1, 6)

            mard.Activate

            If ActiveCell.ListObject.ShowAutoFilter = False Then
               Range("C1:D1").AutoFilter
            Else
                Range("C1:D1").AutoFilter
                Range("C1:D1").AutoFilter
            End If

    'Seleciona o depósito
            ActiveSheet.ListObjects("MARD").Range.AutoFilter Field:=3, Criteria1:= _
            str_dep

    'Selecionará na coluna 4, tudo que é diferente de 0
            ActiveSheet.ListObjects("MARD").Range.AutoFilter Field:=4, Criteria1:="<>0" _
            , Operator:=xlAnd




    'Fomata a data no formato determinado
            ActiveSheet.ListObjects("MARD").Range.AutoFilter Field:=9, Criteria1:= _
            ">" & Format(DataCad, "mm/dd/yyyy hh:mm:ss"), Operator:=xlAnd


            Range("A1").CurrentRegion.Offset(1, 0).SpecialCells(xlCellTypeVisible).Rows(1).Select

            If ActiveCell <> "" Then

                'Cola na coluna 10 e 11 os nomes respectivamente "VERIFICACAO_HIST" e "VERIFICACAO_" & str_dep
                Range("K1") = "VERIFICACAO_HIST"
                Range("L1") = "VERIFICACAO_" & str_dep

                Range("K2:K" & Range("A1").End(xlDown).Row) = "=VLOOKUP(""00000000000"" & RIGHT(RC[-10],7),historico!C[-10],1,0)"
                Range("L2:L" & Range("A1").End(xlDown).Row) = "=VLOOKUP(""00000000000"" & RIGHT(RC[-11],7)," & str_dep & "!C[-11],1,0)"


                ActiveSheet.ListObjects("MARD").Range.AutoFilter Field:=11, Criteria1:="#N/D"
                ActiveSheet.ListObjects("MARD").Range.AutoFilter Field:=12, Criteria1:="#N/D"

                Range("A1").CurrentRegion.Offset(1, 0).SpecialCells(xlCellTypeVisible).Rows(1).Select

                If ActiveCell <> "" Then
                'If ActiveCell.Offset(1, 0) <> "" Then
                'Seleciona matriz e copia
                    Range("A2:F" & ActiveCell.Offset(0, 5).End(xlDown).Row).Copy   'Range("A2:F19").copy

                    ws_dep.Activate
                    Range("A1").End(xlDown).Offset(1, 0).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False

                End If

            End If


        End If


        ws_dep.Activate


        Range("A1").Select


                                                    '********************************
                                                    '   0010        -         0020
                                                    '********************************

        If str_dep = "0010" Or str_dep = "0020" Or str_dep = "0030" Then

            SemAtual = WorksheetFunction.WeekNum(Date, 2)

            UltDia = Year(Date) & "/" & 12 & "/" & 31

            SemAno = WorksheetFunction.WeekNum(UltDia, 2)

            SemRest = SemAno - SemAtual

            ItensTot = WorksheetFunction.CountA(Range("A2:A" & Range("A2").End(xlDown).Row))
            ItensTotFull = ItensTot
            ItensMed = WorksheetFunction.RoundUp(ItensTot / SemRest, 0)

            i = 1

            While i >= 1 And i <= ItensMed
                ItensTot = WorksheetFunction.CountA(Range("A2:A" & Range("A2").End(xlDown).Row)) '||Em TESTE
                i = i + 1

                ws_dep.Activate

                SelItem = WorksheetFunction.RandBetween(1, ItensTot)

                Range("A" & SelItem + 1 & ":F" & SelItem + 1).Select

                historico.Activate
    
                Range("A1").Select
                If ActiveCell.ListObject.ShowAutoFilter = False Then
                   Range("C1:D1").AutoFilter
                Else
                    Range("C1:D1").AutoFilter
                    Range("C1:D1").AutoFilter
                End If

                ws_dep.Activate

                Selection.Cut

                historico.Activate

                If Range("A1").Offset(1, 0) = "" Then
                    Range("A2").Select
                Else
                    Range("A1").End(xlDown).Offset(1, 0).Select
                End If



                ActiveSheet.Paste

                ActiveCell.Offset(0, 6) = DataHist


                ws_dep.Activate

                Selection.Delete Shift:=xlUp

            Wend




    End If





    '******************************************************************************
    '***************************APENAS SEMANA 1 ***********************
    '******************************************************************************

    If str_dep = "0041" Or str_dep = "0060" Then


    '___________________________SEMANA 1_________________________
            mard.Activate
            Range("A1").Select

            If ActiveCell.ListObject.ShowAutoFilter = False Then
               Range("C1:D1").AutoFilter
            Else
                Range("C1:D1").AutoFilter
                Range("C1:D1").AutoFilter
            End If

    'Seleciona o depósito
            ActiveSheet.ListObjects("MARD").Range.AutoFilter Field:=3, Criteria1:= _
            str_dep

    'Selecionará na coluna 4, tudo que é diferente de 0
            ActiveSheet.ListObjects("MARD").Range.AutoFilter Field:=4, Criteria1:="<>0" _
            , Operator:=xlAnd


        
                    Range("P1").Select
                    Application.CutCopyMode = False
                    ActiveCell.FormulaR1C1 = "=AGGREGATE(3,7,MARD[[#All],[MATNR]])"
                    ItensTot = Range("P1").Value
                    ItensTot = ItensTot - 1
        'Itens presentes no DEP atual


        ws_dep.Activate
        ItensTotDep = WorksheetFunction.CountA(Range("A2:A" & Range("A1048576").End(xlUp).Row))


            If ItensTot = ItensTotDep Then
            counter = 1
'                MsgBox "Calcularemos 2/4 do total"
                db_per_dep = 4
            'MsgBox "Estamos na semana " & counter
            End If
    '____________________________________________________________________________

    '_____
    ' \!/

    '************************** FILTRA DATAS NO MES ATUAL *************************
        historico.Activate
        ActiveSheet.Range("$G$1:$G$" & Range("G1").End(xlDown).Row).AutoFilter Field:=7, Operator:= _
            xlFilterValues, Criteria2:=Array(1, Month(Now()) & "/1/2023")
    '*******************************************************************************
        Menu.Activate
        Range("H:H").Delete

        historico.Activate


        Range("G1").Select
        Range(Selection, Selection.End(xlDown)).Offset(1, 0).SpecialCells(xlCellTypeVisible).Copy



        '********************************* DATA + HORA --->  DATA **************
        Menu.Activate
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
    If (str_dep <> "0041" And i <> 4) Or (str_dep = "0041" And j <= 4) Or (str_dep = "0060" And j <= 4) Then
        If Selection.Value <> "" Then
        Range(Selection, Selection.End(xlDown)).Select
        ActiveSheet.Range("$H$1:$H$" & Range("H1").End(xlDown).Row).RemoveDuplicates Columns:=1, Header:=xlNo
        Range("h1").Select

            If ActiveCell.Offset(1, 0) <> "" Then
            i = Range("H1").End(xlDown).Row
            i = i + 1
            counter = i
            Else
            i = 1
            i = i + 1
            counter = i
            End If


        End If
    End If



        '_______________________SEMANA 2, 3 E 4 __________________________________________________________

        '********************* ESTABELECER VALOR P CALC DE MÉDIA COM BASE NA SEMANA ***********************

        If ActiveCell.Value <> "" Then
            'If (str_dep <> "0041" And i <> 4) Or (str_dep = "0041" And j <= 4) Or (str_dep = "0060" And i = 4) Then
            If (str_dep <> "0041" And i <> 4) Or (str_dep = "0060" And i = 4) Then
            MsgBox "Esta é a semana número " & i


                If i = 2 Then
                MsgBox "Calcularemos 1/3 do total"
                db_per_dep = 3
                ElseIf i = 3 Then
                MsgBox "Calcularemos 1/2 do total"
                db_per_dep = 2
                ElseIf i = 4 Then
                MsgBox "Calcularemos 1 do total"
                db_per_dep = 1
                Else
                MsgBox "Esta é a 5º semana, não há mais valores para serem inventariados. O código será finalizado!"
                Range("h1").Select
                Range("H:H").Delete
                Exit Sub
                End If
            End If
        End If

        Range("H:H").Delete
        '*****************************************************************
        '//////////////////////////////////////////////////////////FIM///////
    End If

    ' \!/






    '*******************************  DEP 60 ***************************************************

    If str_dep = "0060" Then
            ws_dep.Activate
            ItensTot = WorksheetFunction.CountA(Range("A2:A" & Range("A2").End(xlDown).Row))
            ItensTotFull = ItensTot
            ItensPer = Application.WorksheetFunction.RoundUp(ItensTot / db_per_dep, 0)

            i = 1

            While i >= 1 And i <= ItensPer
               ItensTot = WorksheetFunction.CountA(Range("A2:A" & Range("A2").End(xlDown).Row))
               i = i + 1
                ws_dep.Activate

                SelItem = WorksheetFunction.RandBetween(1, ItensTot)
                Range("A" & SelItem + 1 & ":F" & SelItem + 1).Select
                historico.Activate

                Range("A1").Select

                '***** NECESSÁRIO FILTRO NO LOOPING? **********

                If ActiveCell.ListObject.ShowAutoFilter = False Then
                   Range("C1:D1").AutoFilter
                Else
                    Range("C1:D1").AutoFilter
                    Range("C1:D1").AutoFilter
                End If

                ws_dep.Activate
                Selection.Cut
                historico.Activate

                If Range("A1").Offset(1, 0) = "" Then
                    Range("A2").Select
                Else
                    Range("A1").End(xlDown).Offset(1, 0).Select
                End If

                ActiveSheet.Paste
                ActiveCell.Offset(0, 6) = DataHist
                ws_dep.Activate
                Selection.Delete Shift:=xlUp

            Wend

    End If







    If str_dep = "0080" Or str_dep = "0050" Or str_dep = "0041" Then
        Select Case str_dep

            Case "0050"
                '--------------------------------
                '              0050
                '--------------------------------
                    db_per_dep = 1   '100%


            Case "0080"
                '--------------------------------
                '              0080
                '--------------------------------
                    db_per_dep = 0.5 '50%
                    
                    
            Case "0041"
                '--------------------------------
                '              0041
                '--------------------------------
                    db_per_dep = 0.5 '50%
                    'MsgBox "Calcularemos 50% do total", vbInformation
                    
                    

        End Select
    End If





        '********************************
        '              0050
        '********************************

     If str_dep = "0050" Then


            ItensTot = WorksheetFunction.CountA(Range("A2:A" & Range("A2").End(xlDown).Row))
            ItensPer = Application.WorksheetFunction.RoundUp(ItensTot * db_per_dep, 0)
            i = 1

            While i >= 1 And i <= ItensPer
                ItensTot = WorksheetFunction.CountA(Range("A2:A" & Range("A2").End(xlDown).Row))
                i = i + 1
                ws_dep.Activate
                SelItem = WorksheetFunction.RandBetween(1, ItensTot)
                Range("A" & SelItem + 1 & ":F" & SelItem + 1).Select
                historico.Activate

                Range("A1").Select

                If ActiveCell.ListObject.ShowAutoFilter = False Then
                   Range("C1:D1").AutoFilter
                Else
                    Range("C1:D1").AutoFilter
                    Range("C1:D1").AutoFilter
                End If

                ws_dep.Activate
                Selection.Cut
                historico.Activate

                If Range("A1").Offset(1, 0) = "" Then
                    Range("A2").Select
                Else
                    Range("A1").End(xlDown).Offset(1, 0).Select
                End If

                ActiveSheet.Paste
                'ActiveCell.Offset(0, 6) = DataHist
                ActiveCell.Offset(0, 6) = DataHist
                ws_dep.Activate
                Selection.Delete Shift:=xlUp

            Wend

    End If

        '********************************
        '              0080 e   0041
        '********************************
        state.Activate
        condicao = Range("A1").Value

        ws_dep.Activate
        Range("A1").Select

    If str_dep = "0080" And condicao = "S" Or str_dep = "0041" And condicao = "S" Then


     If j = 1 Then
     db_per_dep = 1
     End If

    If db_per_dep = "0.5" Then
    MsgBox "Metade dos itens do depósito " & str_dep & " será inventariado!", vbInformation
    ElseIf db_per_dep = "1" Then
    MsgBox "O restante dos itens no depósito " & str_dep & " será inventariado!", vbInformation
    End If
    

            ItensTot = WorksheetFunction.CountA(Range("A2:A" & Range("A2").End(xlDown).Row))
            ItensPer = Application.WorksheetFunction.RoundUp(ItensTot * db_per_dep, 0)
            ItensTotFull = ItensTot
            i = 1

            While i >= 1 And i <= ItensPer
                'Atualiza o número de itens total a cada rodada
                ItensTot = WorksheetFunction.CountA(Range("A2:A" & Range("A2").End(xlDown).Row))
                i = i + 1
                ws_dep.Activate
                SelItem = WorksheetFunction.RandBetween(1, ItensTot)
                Range("A" & SelItem + 1 & ":F" & SelItem + 1).Select
                historico.Activate

                Range("A1").Select

                If ActiveCell.ListObject.ShowAutoFilter = False Then
                   Range("C1:D1").AutoFilter
                Else
                    Range("C1:D1").AutoFilter
                    Range("C1:D1").AutoFilter
                End If

                ws_dep.Activate
                Selection.Cut
                historico.Activate

                If Range("A1").Offset(1, 0) = "" Then
                    Range("A2").Select
                Else
                    Range("A1").End(xlDown).Offset(1, 0).Select
                End If

                ActiveSheet.Paste
                ActiveCell.Offset(0, 6) = DataHist
                ws_dep.Activate
                Selection.Delete Shift:=xlUp

            Wend



    End If




 '============================================================================================================================

        '\!/ APÓS INVENTARIAR O DEPÓSITO SELECIONADO, IRÁ ATUALIZAR OS DADOS DE DEPÓSITO(S) PARA AQUELE DIA \!/

        '-------------------------- RESETAR FILTRO HISTÓRICO --------------------------

        historico.Activate
        Range("A1").Select
        If ActiveCell.ListObject.ShowAutoFilter = False Then
           Range("C1:D1").AutoFilter
        Else
            Range("C1:D1").AutoFilter
            Range("C1:D1").AutoFilter
        End If


        ActiveSheet.ListObjects("historico").Range.AutoFilter Field:=3, Criteria1:= _
        str_dep

        '----------------------------Atualizar Estoque Antes de Inventariar
        d = Date
        DataCompleta = Split(d, "/")
        dia = DataCompleta(0)

        mes = DataCompleta(1)

        ano = DataCompleta(2)

        d = mes & "/" & dia & "/" & ano  'DATA CONVERTIDA PARA PADRAO USA

        '************************** FILTRA DATAS NO DIA DE HOJE *************************
        ActiveSheet.Range("$G$1:$G$" & Range("G1").End(xlDown).Row).AutoFilter Field:=7, Operator:= _
            xlFilterValues, Criteria2:=Array(2, d)
            ''Criteria2:=Array(2, d)' a opção dois significa, basicamente, que irá pegar o dia e a opção 1 é = pegar tudo no mês e 3 = pegar tudo no ano

        '*******************************************************************************


        '------------------------------------------------------------------------------
                Range("A1").CurrentRegion.Offset(1, 0).SpecialCells(xlCellTypeVisible).Rows(1).Select
                If ActiveCell <> "" Then
                Dim sDep As String
                sDep = str_dep
                ActiveCell.Offset(0, 3).Select
                'ActiveCell.FormulaR1C1 = "=VLOOKUP([@QTD],MARD!C[-3]:C[8],4,0)"
                    ActiveCell.FormulaR1C1 = _
                 "=XLOOKUP(historico[@CODIGO]&" & """" & str_dep & """" & ",MARD[id],MARD[LABST])"

                Selection.Copy
                j = ActiveCell.Row

                Range("D" & j & ":D" & Range("D1048576").End(xlUp).Row).Select

                Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
                Selection.Copy

                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False

                End If

                Range("A1").Select
                If ActiveCell.ListObject.ShowAutoFilter = False Then
                   Range("C1:D1").AutoFilter
                Else
                    Range("C1:D1").AutoFilter
                    Range("C1:D1").AutoFilter
                End If

        '---------------------------------------------------------------------------------

'==================================================================================================================================













    If auto_run <> "on" Then
                Call UnlockManualLocal

                state.Activate
                condicao = Range("A1").Value

                If Range("A1").Value = "S" Then
                    Range("A1").Value = "N"
                Else
                     Range("A1").Value = "S"
                End If

                ws_dep.Activate
                Range("A1").Select

                Call bloquear_guias

                historico.Activate
                Range("A1").Select
                ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
                , AllowFiltering:=True
    End If


    End If




    '-------------------------------------- RELATÓRIOS DE CONCLUSÃO \!/ MANUAL \!/ ----------------------------------------------

'MENSAGENS DEP 0010 | 0020 | 0030
    If str_dep = "0010" Or str_dep = "0020" Or str_dep = "0030" Then
        If ItensMed <= 1 Then
           If auto_run = "" Then
           MsgBox ItensMed & " material foi inventariado para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!" & Chr(10) & "A base de cálculo foi realizada sobre as " & SemRest & " semanas restantes!"
           End If
           If auto_run = "on" Then
                If str_dep = "0010" Then
                    msgCompleta(1) = CStr(ItensMed & " material foi inventariado para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!" & Chr(10) & "A base de cálculo foi realizada sobre as " & SemRest & " semanas restantes!")
                    ElseIf str_dep = "0020" Then
                    msgCompleta(2) = CStr(ItensMed & " material foi inventariado para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!" & Chr(10) & "A base de cálculo foi realizada sobre as " & SemRest & " semanas restantes!")
                    Else
                    msgCompleta(3) = CStr(ItensMed & " material foi inventariado para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!" & Chr(10) & "A base de cálculo foi realizada sobre as " & SemRest & " semanas restantes!")
                End If
           End If

        Else
            If auto_run = "" Then
            MsgBox ItensMed & " materiais foram inventariados para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!" & Chr(10) & "A base de cálculo foi realizada sobre as " & SemRest & " semanas restantes!"
            End If
            If auto_run = "on" Then
                If str_dep = "0010" Then
                    msgCompleta(1) = CStr(ItensMed & " materiais foram inventariados para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!" & Chr(10) & "A base de cálculo foi realizada sobre as " & SemRest & " semanas restantes!")
                    ElseIf str_dep = "0020" Then
                    msgCompleta(2) = CStr(ItensMed & " materiais foram inventariados para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!" & Chr(10) & "A base de cálculo foi realizada sobre as " & SemRest & " semanas restantes!")
                    Else
                    msgCompleta(3) = CStr(ItensMed & " materiais foram inventariados para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!" & Chr(10) & "A base de cálculo foi realizada sobre as " & SemRest & " semanas restantes!")
                End If
            End If

    End If

'MENSAGENS DEP 0041 | 0060
    ElseIf str_dep = "0041" Or str_dep = "0060" Then
        If ItensPer <= 1 And auto_run = "" Then
            If str_dep = "0041" And condicao = "N" Then
            MsgBox "Essa semana não será inventariada para o depósito " & str_dep & ", tente na próxima semana." & Chr(10) & "O estado atual para a semana referente ao depósito 0041 pode ser checada na guia MENU", vbInformation
            ElseIf str_dep = "0041" And condicao = "S" Then
            MsgBox ItensPer & " material foi inventariado para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!"
            ElseIf auto_run = "" Then
           MsgBox ItensPer & " material foi inventariado para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!" ' & Chr(10) & "Esta é a semana número " & counter & " !"
        End If
        
        

            If auto_run = "on" Then
                If str_dep = "0041" And condicao = "N" Then
                    msgCompleta(4) = CStr(ItensPer & " material foi inventariado para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!") ' & Chr(10) & "Esta é a semana número " & counter & " !")
                ElseIf str_dep = "0041" And condicao = "S" Then
                msgCompleta(4) = CStr("Essa semana não será inventariada para o depósito " & str_dep & ", tente na próxima semana." & Chr(10) & "O estado atual para a semana referente ao depósito 0041 pode ser checada na guia MENU")
                ElseIf str_dep = "0060" Then
                    msgCompleta(6) = CStr(ItensPer & " material foi inventariado para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!") ' & Chr(10) & "Esta é a semana número " & counter & " !")
                End If
            End If

        Else
            If auto_run = "" Then
            MsgBox ItensPer & " materiais foram inventariados para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!" '& Chr(10) & "Esta é a semana número " & counter & " !"
            End If

            If auto_run = "on" Then
            
                If str_dep = "0041" And condicao = "S" Then
                    msgCompleta(4) = CStr(ItensPer & " materiais foram inventariados para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!") ' & Chr(10) & "Esta é a semana número " & counter & " !")
                ElseIf str_dep = "0041" And condicao = "N" Then
                    msgCompleta(4) = CStr("Essa semana não será inventariada para o depósito " & str_dep & ", tente na próxima semana.") ' & Chr(10) & "O estado atual para a semana referente ao depósito 0080 pode ser checada na guia MENU")
                ElseIf str_dep = "0060" Then
                    msgCompleta(6) = CStr(ItensPer & " materiais foram inventariados para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!") ' & Chr(10) & "Esta é a semana número " & counter & " !")
                End If
            End If

        End If

'MENSAGENS DEP 0050
    ElseIf str_dep = "0050" Then
        If ItensTot <= 1 Then
           If auto_run = "" Then
           MsgBox ItensPer & " material foi inventariado para o depósito número " & str_dep & ",  na guia MARD, de um total de " & ItensTot & " item presente no histórico!"
           End If

            If auto_run = "on" Then
                If str_dep = "0050" Then
                    msgCompleta(5) = CStr(ItensPer & " material foi inventariado para o depósito número " & str_dep & ",  na guia MARD, de um total de " & ItensTot & " item presente no histórico!")
                End If
            End If

        Else
            If auto_run = "" Then
            MsgBox ItensPer & " materiais foram inventariados para o depósito número " & str_dep & ",  na guia MARD, de um total de " & ItensTot & " itens presentes no histórico!"
            End If

            If auto_run = "on" Then
                If str_dep = "0050" Then
                    msgCompleta(5) = CStr(ItensPer & " materiais foram inventariados para o depósito número " & str_dep & ",  na guia MARD, de um total de " & ItensTot & " itens presentes no histórico!")
                End If
            End If

        End If

'MENSAGENS DEP 0080
    ElseIf str_dep = "0080" And condicao = "S" Then
        If ItensTotFull <= 1 Then
           If auto_run = "" Then
           MsgBox ItensPer & " material foi inventariado para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!"
           End If

            If auto_run = "on" Then
                If str_dep = "0080" Then
                    msgCompleta(7) = CStr(ItensPer & " material foi inventariado para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " item!")
                End If
            End If

        Else
           If auto_run = "" Then
            MsgBox ItensPer & " materiais foram inventariados para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!"
           End If

            If auto_run = "on" Then
                If str_dep = "0080" Then
                    msgCompleta(7) = CStr(ItensPer & " materiais foram inventariados para o depósito número " & str_dep & ", de um total de " & ItensTotFull & " itens!")
                End If
            End If

        End If
    ElseIf str_dep = "0080" And condicao = "N" Then
           If auto_run = "" Then
           MsgBox "Essa semana não será inventariada para o depósito " & str_dep & ", tente na próxima semana." & Chr(10) & "O estado atual para a semana referente ao depósito 0080 pode ser checada na guia MENU"
           End If

            If auto_run = "on" Then
                If str_dep = "0080" Then
                    msgCompleta(7) = CStr("Essa semana não será inventariada para o depósito " & str_dep & ", tente na próxima semana.") ' & Chr(10) & "O estado atual para a semana referente ao depósito 0080 pode ser checada na guia MENU")
                End If
            End If

End If



End Sub
