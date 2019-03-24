
    Dim escolhida
    Dim aleatorio
    Dim aberta
    Dim portafinal
    Dim segundaporta
    Dim porta1aberta
    Dim porta2aberta
    Dim porta3aberta
    Dim pasta
    Dim NomeDaFoto
    Dim CaminhoCompleto
    

Private Sub ToggleButton1_Click()

    escolhida = Range("A3")
    aleatorio = Range("A4")
    segundaporta = Range("A5")
    
    
    If (ToggleButton1.Value = True And segundaporta = 0) Then
        
        If (escolhida = 1) Then
            If (aleatorio = 1) Then
                porta2aberta = 1
                Call mudarimagem2
            End If
            
            If (aleatorio = 2) Then
                porta3aberta = 1
                Call mudarimagem3
            End If
        End If
        
        If (escolhida = 2) Then
            porta3aberta = 1
            Call mudarimagem3
        End If
    
        If (escolhida = 3) Then
            porta2aberta = 1
            Call mudarimagem2
        End If
    End If
       
    If (segundaporta = 1) Then
        porta1aberta = 1
        Call mudarimagem1
    End If

    Range("A5") = 1
      
End Sub


Private Sub ToggleButton2_Click()

    escolhida = Range("A3")
    aleatorio = Range("A4")
    segundaporta = Range("A5")
    
    
    If (ToggleButton2.Value = True And segundaporta = 0) Then
        
        If (escolhida = 2) Then
            If (aleatorio = 1) Then
                porta1aberta = 1
                Call mudarimagem1
            End If
            
            If (aleatorio = 2) Then
                porta3aberta = 1
                Call mudarimagem3
            End If
        End If
        
        If (escolhida = 1) Then
            porta3aberta = 1
            Call mudarimagem3
        End If
    
        If (escolhida = 3) Then
            porta1aberta = 1
            Call mudarimagem1
        End If
    End If
       
    If (segundaporta = 1) Then
        porta2aberta = 1
        Call mudarimagem2
    End If

    Range("A5") = 1
    
End Sub


Private Sub ToggleButton3_Click()

    escolhida = Range("A3")
    aleatorio = Range("A4")
    segundaporta = Range("A5")
    
    
    If (ToggleButton3.Value = True And segundaporta = 0) Then
        
        If (escolhida = 3) Then
            If (aleatorio = 1) Then
                porta1aberta = 1
                Call mudarimagem1
            End If
            
            If (aleatorio = 2) Then
                porta2aberta = 1
                Call mudarimagem2
            End If
        End If
        
        If (escolhida = 1) Then
            porta2aberta = 1
            Call mudarimagem2
        End If
    
        If (escolhida = 2) Then
            porta1aberta = 1
            Call mudarimagem1
        End If
    End If
       
    If (segundaporta = 1) Then
        porta3aberta = 1
        Call mudarimagem3
    End If

    Range("A5") = 1
    
End Sub






Sub mudarimagem1()

   
    segundaporta = Range("A5")
       
    'Limpa a Imagem anterior
    ToggleButton1.Picture = LoadPicture()

    'Captura o caminho em que o "xlsm" está e concatena com a pasta "fotos"
    pasta = ThisWorkbook.Path & "\Fotos\"
    
    If (escolhida = 1) Then
        'Nome da Foto sem a extensão
        NomeDaFoto = "porta aberta premiada"
    
    Else
        NomeDaFoto = "porta aberta cabra"
    End If
    
    
    If (porta1aberta = 0) Then
        'Nome da Foto sem a extensão
        NomeDaFoto = "porta fechada"
    End If
        
    'Montamos o cainho completo
    CaminhoCompleto = pasta & NomeDaFoto & ".bmp"

    
    ToggleButton1.Picture = LoadPicture(CaminhoCompleto)
    
End Sub

Sub mudarimagem2()

   
    segundaporta = Range("A5")
       
    'Limpa a Imagem anterior
    ToggleButton2.Picture = LoadPicture()

    'Captura o caminho em que o "xlsm" está e concatena com a pasta "fotos"
    pasta = ThisWorkbook.Path & "\Fotos\"
    
    If (escolhida = 2) Then
        'Nome da Foto sem a extensão
        NomeDaFoto = "porta aberta premiada"
    
    Else
        NomeDaFoto = "porta aberta cabra"
    End If
    
    
    If (porta2aberta = 0) Then
        'Nome da Foto sem a extensão
        NomeDaFoto = "porta fechada"
    End If
        
    'Montamos o cainho completo
    CaminhoCompleto = pasta & NomeDaFoto & ".bmp"

    
    ToggleButton2.Picture = LoadPicture(CaminhoCompleto)
    
End Sub

Sub mudarimagem3()
    
    

   
    segundaporta = Range("A5")
       
    'Limpa a Imagem anterior
    ToggleButton3.Picture = LoadPicture()

    'Captura o caminho em que o "xlsm" está e concatena com a pasta "fotos"
    pasta = ThisWorkbook.Path & "\Fotos\"
    
    If (escolhida = 3) Then
        'Nome da Foto sem a extensão
        NomeDaFoto = "porta aberta premiada"
    
    Else
        NomeDaFoto = "porta aberta cabra"
    End If
    
    
    If (porta3aberta = 0) Then
        'Nome da Foto sem a extensão
        NomeDaFoto = "porta fechada"
    End If
        
    'Montamos o cainho completo
    CaminhoCompleto = pasta & NomeDaFoto & ".bmp"

    
    ToggleButton3.Picture = LoadPicture(CaminhoCompleto)

End Sub



Sub resetarporta()

    Range("A3") = (1 + Int(Rnd * 3))
    Range("A4") = (1 + Int(Rnd * 2))

    Range("A5") = 0
    porta1aberta = 0
    Call mudarimagem1
    ToggleButton1.Value = False
    
    Range("A5") = 0
    porta2aberta = 0
    Call mudarimagem2
    ToggleButton2.Value = False
    
    Range("A5") = 0
    porta3aberta = 0
    Call mudarimagem3
    ToggleButton3.Value = False
    
    
    Range("A5") = 0
    


End Sub



Sub automatizar()
    

    Range("E24:T5024").ClearContents
    Range("A1") = 0
    
    
    resposta = MsgBox("A simulação irá fazer 1000 tentativas" + vbCrLf + "Pressione 'Esc' para cancelar antes (Pode apresentar erros)", 1)
    
    If (resposta = 2) Then GoTo fim
    
    
    
    
    
     For i = Range("A1") To 1000
        Range("A3") = (1 + Int(Rnd * 3))    'Porta Premiada
        Range("A4") = (1 + Int(Rnd * 2))    'Número Aleatório 1
        
        escolhida = Range("A3")
        aleatorio = Range("A4")
        
        
        Range("B3") = (1 + Int(Rnd * 3))       'Porta Escolhida pra ser aberta
        aberta = Range("B3")
    
            If (aberta = 1) Then                'Porta escolhida 1
                If (escolhida = 1) Then         'Porta premiada 1
                    If (aleatorio = 1) Then     'Será exibida a porta 2
                        portafinal = 3
                    End If
                    If (aleatorio = 2) Then     'Será exibida a porta 3
                        portafinal = 2
                    End If
                End If
                If (escolhida = 2) Then         'Porta premiada 2
                    portafinal = 2
                End If
                If (escolhida = 3) Then         'Porta premiada 3
                    portafinal = 3
                End If
            End If
            
            If (aberta = 2) Then                'Porta escolhida 2
                If (escolhida = 2) Then         'Porta premiada 2
                    If (aleatorio = 1) Then     'Será exibida a porta 1
                        portafinal = 3
                    End If
                    If (aleatorio = 2) Then     'Será exibida a porta 3
                        portafinal = 1
                    End If
                End If
                If (escolhida = 1) Then         'Porta premiada 1
                    portafinal = 1
                End If
                If (escolhida = 3) Then         'Porta premiada 3
                    portafinal = 3
                End If
            End If
            
            If (aberta = 3) Then                'Porta escolhida 3
                If (escolhida = 3) Then         'Porta premiada 3
                    If (aleatorio = 1) Then     'Será exibida a porta 1
                        portafinal = 2
                    End If
                    If (aleatorio = 2) Then     'Será exibida a porta 2
                        portafinal = 1
                    End If
                End If
                If (escolhida = 1) Then         'Porta premiada 1
                    portafinal = 1
                End If
                If (escolhida = 2) Then         'Porta premiada 2
                    portafinal = 2
                End If
            End If
            
            
            
        ActiveSheet.Cells(24 + Range("A1"), 5).Select
        Cells(24 + Range("A1"), 8) = escolhida                  'Mostra qual porta foi escolhida
        Cells(24 + Range("A1"), 5) = i + 1                      'Mostra a quantidade de jogadas
        If (escolhida = portafinal) Then
            If (i = 0) Then
                Cells(24, 10) = 1
            Else
                Cells(24 + Range("A1"), 10) = Cells(23 + Range("A1"), 10) + 1
            End If
        Else
            If (i = 0) Then
                Cells(24, 10) = 0
            Else
                Cells(24 + Range("A1"), 10) = Cells(23 + Range("A1"), 10)
            End If
        End If
        
        Cells(24 + Range("A1"), 17) = aberta
        
        Cells(24 + Range("A1"), 14) = Cells(24 + Range("A1"), 10) / Cells(24 + Range("A1"), 5)
        
        Range("A1") = Range("A1") + 1
    
        Next i
        
fim:
            
End Sub


Sub trocaporta()
    

    Range("E24:T5024").ClearContents
    Range("A1") = 0
    
    resposta = MsgBox("A simulação irá fazer 1000 tentativas" + vbCrLf + "Pressione 'Esc' para cancelar antes (Pode apresentar erros)", 1)
    
    If (resposta = 2) Then GoTo fim
    
     For i = Range("A1") To 1000
        Range("A3") = (1 + Int(Rnd * 3))    'Porta Premiada
                
        escolhida = Range("A3")
        
        Range("B3") = (1 + Int(Rnd * 3))       'Porta Escolhida pra ser aberta
        aberta = Range("B3")
    
        portafinal = aberta
                    
        ActiveSheet.Cells(24 + Range("A1"), 5).Select
        Cells(24 + Range("A1"), 8) = escolhida                  'Mostra qual porta foi escolhida
        Cells(24 + Range("A1"), 5) = i + 1                      'Mostra a quantidade de jogadas
        If (escolhida = portafinal) Then
            If (i = 0) Then
                Cells(24, 10) = 1
            Else
                Cells(24 + Range("A1"), 10) = Cells(23 + Range("A1"), 10) + 1
            End If
        Else
            If (i = 0) Then
                Cells(24, 10) = 0
            Else
                Cells(24 + Range("A1"), 10) = Cells(23 + Range("A1"), 10)
            End If
        End If
        
        Cells(24 + Range("A1"), 17) = aberta
        
        Cells(24 + Range("A1"), 14) = Cells(24 + Range("A1"), 10) / Cells(24 + Range("A1"), 5)
        
        Range("A1") = Range("A1") + 1
        
        Next i
        
fim:

End Sub

