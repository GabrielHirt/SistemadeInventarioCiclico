Attribute VB_Name = "a_selecaoAutomatica"
Option Explicit

Public str_dep  As String
Public auto_run As String
Public mard     As Worksheet
Public msgCompleta(1 To 7) As String
Public LOADING  As Worksheet


Sub run()
    'Seleção Automática
    
    '==========================================
    
    Dim rA          As Worksheet
    Dim state       As Worksheet
    Dim historico  As Worksheet
    Dim condicao    As String
    Dim msgResp As VbMsgBoxResult
    Dim newHour As Date
    Dim newMinute As Date
    Dim newSecond As Date
    Dim waitTime As Date
    
    
    Set mard = Sheets("MARD")
    
    msgResp = MsgBox("Deseja executar o comando de seleção automática?", vbYesNo)
    
If msgResp = vbYes Then
    
        Set state = Sheets("state")
        Set rA = Sheets("runALL")
        Set historico = Sheets("historico")
        Set LOADING = Sheets("LOADING")
    '==========================================

    
        '============DESBLOQUEAR GUIAS===============:.
    '    runALL.Unprotect
    '    state.Unprotect
        Call desbloquear_guias
        '============================================:.
        
        mard.Select
        Range("A1").Select
        Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
        
        
        auto_run = "on"
        rA.Select
        Range("B2").Value = "0010"
        str_dep = "0010"
        
        
    
        
        
    '______________________________________________________________________ .
    '                IF SCOUP                                               . .
    '______________________________________________________________________ . . .
        Dim l As String
        While (str_dep <> "")
        auto_run = "on"
            If str_dep = "0010" Then        '// anual - dinâmico
            
            Application.ScreenUpdating = True

            
            LOADING.Activate

           ' Dim StartingTime As Single
            'StartingTime = Timer

            

            
            Range("F4") = "INVENTÁRIO CÍCLICO AUTOMÁTICO"
            Range("C13") = "O processo de inventário para o depósito " & str_dep & " está em andamento"
            Range("C9") = "CARREGANDO"
            Range("R7") = ":D"
            Range("J8") = "Aguarde"
            
            '============= TIMER ================================:.
            newHour = Hour(Now())
            newMinute = Minute(Now())
            newSecond = Second(Now()) + 1
            waitTime = TimeSerial(newHour, newMinute, newSecond)
            Application.Wait waitTime
            '============= END TIMER ================================:.
            'MsgBox Format((Timer - StartingTime) / 86400, "hh:mm:ss")
            Application.ScreenUpdating = False
            
            'MsgBox "Este é o depósito " & str_dep
                Call SelecaoAutomatica
                     
                rA.Select
                Range("B2").Value = "0020"
                str_dep = "0020"
                
                    
            ElseIf str_dep = "0020" Then   '// anual - dinâmico
            
            Application.ScreenUpdating = True
            
            LOADING.Activate
            Range("F4") = "INVENTÁRIO CÍCLICO AUTOMÁTICO"
            Range("C13") = "O processo de inventário para o depósito " & str_dep & " está em andamento"
            
            '============= TIMER ================================:.
            newHour = Hour(Now())
            newMinute = Minute(Now())
            newSecond = Second(Now()) + 1
            waitTime = TimeSerial(newHour, newMinute, newSecond)
            Application.Wait waitTime
            '============= END TIMER ================================:.
            
            Application.ScreenUpdating = False
            
             'MsgBox "Este é o depósito " & str_dep
            
                Call SelecaoAutomatica
                     
                rA.Select
                Range("B2").Value = "0030"
                str_dep = "0030"
            
            ElseIf str_dep = "0030" Then    '// anual - dinâmico
            
            Application.ScreenUpdating = True
            
            LOADING.Activate
            Range("F4") = "INVENTÁRIO CÍCLICO AUTOMÁTICO"
            Range("C13") = "O processo de inventário para o depósito " & str_dep & " está em andamento"
            '============= TIMER ================================:.
            newHour = Hour(Now())
            newMinute = Minute(Now())
            newSecond = Second(Now()) + 1
            waitTime = TimeSerial(newHour, newMinute, newSecond)
            Application.Wait waitTime
            '============= END TIMER ================================:.
            
            Application.ScreenUpdating = False
            

            ' MsgBox "Este é o depósito " & str_dep
                Call SelecaoAutomatica
                     
                rA.Select
                Range("B2").Value = "0041"
                str_dep = "0041"
            
            ElseIf str_dep = "0041" Then   '// 25% semanal - 100% mensal
            
            Application.ScreenUpdating = True
            
            LOADING.Activate
            Range("F4") = "INVENTÁRIO CÍCLICO AUTOMÁTICO"
            Range("C13") = "O processo de inventário para o depósito " & str_dep & " está em andamento"
            '============= TIMER ================================:.
            newHour = Hour(Now())
            newMinute = Minute(Now())
            newSecond = Second(Now()) + 1
            waitTime = TimeSerial(newHour, newMinute, newSecond)
            Application.Wait waitTime
            '============= END TIMER ================================:.
            
            Application.ScreenUpdating = False
            
           ' MsgBox "Este é o depósito " & str_dep
                Call SelecaoAutomatica
                     
                rA.Select
                Range("B2").Value = "0050"
                str_dep = "0050"
            
            ElseIf str_dep = "0050" Then   '// 100% semanal
            
            Application.ScreenUpdating = True
            
            LOADING.Activate
            Range("F4") = "INVENTÁRIO CÍCLICO AUTOMÁTICO"
            Range("C13") = "O processo de inventário para o depósito " & str_dep & " está em andamento"
            '============= TIMER ================================:.
            newHour = Hour(Now())
            newMinute = Minute(Now())
            newSecond = Second(Now()) + 1
            waitTime = TimeSerial(newHour, newMinute, newSecond)
            Application.Wait waitTime
            '============= END TIMER ================================:.
            
            Application.ScreenUpdating = False
            
            ' MsgBox "Este é o depósito " & str_dep
                Call SelecaoAutomatica
                     
                rA.Select
                Range("B2").Value = "0060"
                str_dep = "0060"
            
            ElseIf str_dep = "0060" Then   '// 25% semanal - 100% mensal
            
            Application.ScreenUpdating = True
            
            LOADING.Activate
            Range("F4") = "INVENTÁRIO CÍCLICO AUTOMÁTICO"
            Range("C13") = "O processo de inventário para o depósito " & str_dep & " está em andamento"
            '============= TIMER ================================:.
            newHour = Hour(Now())
            newMinute = Minute(Now())
            newSecond = Second(Now()) + 1
            waitTime = TimeSerial(newHour, newMinute, newSecond)
            Application.Wait waitTime
            '============= END TIMER ================================:.
            
            Application.ScreenUpdating = False
            
            ' MsgBox "Este é o depósito " & str_dep
                Call SelecaoAutomatica
                     
                rA.Select
                Range("B2").Value = "0080"
                str_dep = "0080"
            
            ElseIf str_dep = "0080" Then    '// 50% a cada 2 semanas - 100% mês
            
            Application.ScreenUpdating = True
            
            LOADING.Activate
            Range("F4") = "INVENTÁRIO CÍCLICO AUTOMÁTICO"
            Range("C13") = "O processo de inventário para o depósito " & str_dep & " está em andamento"
            '============= TIMER ================================:.
            newHour = Hour(Now())
            newMinute = Minute(Now())
            newSecond = Second(Now()) + 1
            waitTime = TimeSerial(newHour, newMinute, newSecond)
            Application.Wait waitTime
            '============= END TIMER ================================:.
            
            Application.ScreenUpdating = False
            
            ' MsgBox "Este é o depósito " & str_dep
                Call SelecaoAutomatica
                     
                rA.Select
                Range("B2").Value = "0010"
                str_dep = ""
            
            End If
        Wend
        
    '______________________________________________________________________ .
    '               END OF SCOUP                                            . .
    '______________________________________________________________________ . . .
        
        
    
                    '--------------------------------- Mudar state
                    state.Activate
                    condicao = Range("A1").Value
                
                    If Range("A1").Value = "S" Then
                        Range("A1").Value = "N"
                    Else
                        Range("A1").Value = "S"
                    End If
                
                    rA.Select
                    Range("B2").Select
                    '---------------------------------------------
                    
                    MsgBox "DEPÓSITO 0010" & Chr(10) & msgCompleta(1) & Chr(10) & Chr(10) & "DEPÓSITO 0020" & Chr(10) & msgCompleta(2) & Chr(10) & Chr(10) & "DEPÓSITO 0030" & Chr(10) & msgCompleta(3) & Chr(10) & Chr(10) & "DEPÓSITO 0041" & Chr(10) & msgCompleta(4) & Chr(10) & Chr(10) & "DEPÓSITO 0050" & Chr(10) & msgCompleta(5) & Chr(10) & Chr(10) & "DEPÓSITO 0060" & Chr(10) & msgCompleta(6) & Chr(10) & Chr(10) & "DEPÓSITO 0080" & Chr(10) & msgCompleta(7)
                        
                        '============UNLOCK RANGE===============:.
                        
                            Call UnlockManualLocal
                        
                        '============DESBLOQUEAR GUIAS===============:.
                            Call bloquear_guias
                            
                        '==========UNLOCK FILTERS=====================:.
                        
                        historico.Activate
                        Range("A1").Select
                        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
                        , AllowFiltering:=True
                        '============================================:.
                        
                        '============SALVAR RELATÓRIO===============:.
                            Call saveNotePadData
                        '============================================:.
                        
                    LOADING.Activate
                    Range("F4") = ""
                    Range("C13") = "Nenhum processo de inventário está em execução"
                    Range("C9") = "------------"
                    Range("R7") = ""
                    Range("J8") = "N/A"
                    Application.ScreenUpdating = True
                    
                    MsgBox "INVENTÁRIO CÍCLICO AUTOMÁTICO CONCLUÍDO"
Else
                    MsgBox "A EXECUÇÃO DO INVENTÁRIO CÍCLICO AUTOMÁTICO FOI CANCELADA"
    Exit Sub
    
            
End If

End Sub
