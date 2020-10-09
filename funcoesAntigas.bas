Attribute VB_Name = "Módulo1"
'Password: gumemura
'Todas as funções tem que ter a seguinte linha no começo
'   ActiveSheet.Unprotect "gumemura"

'E essa no final

'   ActiveSheet.Protect Contents:=True

'Na primeira linha, estamos destravando a planilha
'Na ultima, a retornando ao estado de proteção

Function ValorBotaoCancelar() As Integer
    'retorna 0, que é o valor retornado por uma input box quando o botao 'cacelar' é apertado
    ValorBotaoCancelar = 0
End Function

Function AcaoFoiCancelada(valor As String) As Boolean
    If valor = "Falso" Then
        AcaoFoiCancelada = True
    Else
        AcaoFoiCancelada = False
    End If
End Function

Sub Entrada()
Attribute Entrada.VB_Description = "Dar entrada em produtos"
Attribute Entrada.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveSheet.Unprotect "gumemura"
    'Dim serve para declarar variaveis
    'No VBA o int não é muito usado, mas sim o Long
    Dim N As Long, i As Long, code As Double, loteQuantidadeDeCelulas As Long, checkLote As String
    
    'Pedindo para selecionar a coluna que tera a quantidade inserida
    Set lote = Application.InputBox("Selecione uma célula do lote a ser incrementado", "ENTRADA", Type:=8)
    
    loteQuantidadeDeCelulas = lote.Cells.Count
    
    'Checando se o intervalo selecionado tem apenas uma celula
    'Se nao, dar mensagem de erro
    If loteQuantidadeDeCelulas > 1 Then
        MsgBox ("Seleciona apenas uma célula!")
    Else
        ' Pega a ultima coluna que tenha algum dado da linha A (linha A)
        NumRows = Cells(1, Columns.Count).End(xlToLeft).Column
    
        ' Checando se a celula esta entre o intervalo das colunas de lote
        If lote.Cells(1, 1).Column <= NumRows - 3 And lote.Cells(1, 1).Column >= 10 Then
            'Abre um caixa de dialogo para coletar o numero do codigo de barras
            'O Type:=1 diz que o input sera um numero. Sem isso, seria uma String
            code = Application.InputBox("", "ENTRADA", Type:=1)
            
            'Pega a ultima linha com dado da coluna 3 (coluna C)
            N = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row
            
            While code <> ValorBotaoCancelar
                For i = 1 To N
                    'Percorrendo cada celula da coluna com os codigos de barra
                    If Cells(i, "C").Value = code Then
                        Cells(i, lote.Cells(1, 1).Column).Value = Cells(i, lote.Cells(1, 1).Column).Value + 1
                        Exit For
                    End If
                Next i
            
                If i = N + 1 Then
                    MsgBox "Produto nao registrado", vbCritical
                End If
                
                code = Application.InputBox("", "ENTRADA", Type:=1)
            Wend
        Else
            MsgBox ("Célula selecionada não é de um lote!")
        End If
    End If
    
    ActiveSheet.Protect Contents:=True
End Sub

Sub Saida()
Attribute Saida.VB_Description = "Dar saida"
Attribute Saida.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveSheet.Unprotect "gumemura"
    Dim N As Long, i As Long, code As Double
    code = Application.InputBox("", "SAÍDA", Type:=1)

    N = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row
    
    While code <> ValorBotaoCancelar
        
        For i = 1 To N
            If Cells(i, "C").Value = code Then
                Cells(i, "E").Value = Cells(i, "E").Value + 1
                Exit For
            End If
        Next i
    
        If i = N + 1 Then
            MsgBox "Produto nao registrado", vbCritical
        End If
        code = Application.InputBox("", "SAÍDA", Type:=1)
    Wend
    ActiveSheet.Protect Contents:=True
End Sub
'Criando uma coluna para um novo lote
Sub Novo_Lote()
    ActiveSheet.Unprotect "gumemura"
    
    Dim nomeNovoLote As String
    
    nomeNovoLote = Application.InputBox("", "Nome do novo lote", Type:=2)
    
    If StrPtr(nomeNovoLote) <> 0 Then
        ' Pega a ultima coluna que tenha algum dado da linha A (linha A)
        NumRows = Cells(1, Columns.Count).End(xlToLeft).Column
        
        ' Pega a coluna e subtrai seu numero por 2 (as duas ultimas colunas sao 'rolos' e 'total de componentes')
        ' Essa nova coluna esta oculta. Ela é o parametro para a criação de novas colunas usadas para novos lotes
        Cells(2, NumRows - 2).Select
        
        'Cria nova coluna
        Selection.EntireColumn.Insert
        Cells(2, NumRows - 2).Value = nomeNovoLote
    End If
    
    ActiveSheet.Protect Contents:=True
End Sub
'Criando uma linha para um novo componente
Sub Novo_Comp()
    Dim AcaoCancelada As Boolean
    Dim compName As String, quantPorRolo As String, codigoDeBarras As String
    
    AcaoCancelada = False
    
    ActiveSheet.Unprotect "gumemura"
    'Worksheet.Protect gumemura, UserInterfaceOnly:=True
    
    'Encontrando a ultima linha com conteudo
    N = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Pega a ultima coluna que tenha algum dado da linha A (linha A)
    ' Isso serve para termos o correto parametro para copiarmos as equações de uma linha já existente para a nova linha
    lastColunm = Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Selecionando a linha abaixo dessa ultima (a linha N, acima encontrada)
    Cells(N + 1, 1).Select
    
    'Criando nova lina
    Selection.EntireRow.Insert
    
    'Nome do componente
    compName = Application.InputBox("", "Nome do componente", Type:=2)
    Cells(N + 1, "A").Value = compName

    If AcaoFoiCancelada(compName) = False Then
        'Quantidade por rolo
        quantPorRolo = Application.InputBox("", "Quantidade de componentes por rolo", Type:=1)
        Cells(N + 1, "B").Value = quantPorRolo
    Else
        quantPorRolo = "Falso"
        AcaoCancelada = True
    End If

    If AcaoFoiCancelada(quantPorRolo) = False Then
        'Codigo de barras
        codigoDeBarras = Application.InputBox("", "Código de barras", Type:=1)
        Cells(N + 1, "C").Value = codigoDeBarras
    Else
        codigoDeBarras = "Falso"
        AcaoCancelada = True
    End If
    
    If AcaoFoiCancelada(codigoDeBarras) = False Then
        'Total de rolos copia
        Cells(N, "F").Select
        Selection.Copy
        Cells(N + 1, "F").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        'Total de componentes copia
        Cells(N, "G").Select
        Selection.Copy
        Cells(N + 1, "G").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        'Total de rolos original
        Cells(N, lastColunm - 1).Select
        Selection.Copy
        Cells(N + 1, lastColunm - 1).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        'Total de componentes original
        Cells(N, lastColunm).Select
        Selection.Copy
        Cells(N + 1, lastColunm).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
    Else
        AcaoCancelada = True
    End If
    
    If AcaoCancelada = True Then
        'Selecionando a linha abaixo dessa ultima (a linha N, acima encontrada)
        Cells(N + 1, 1).Select
        
        'Criando nova lina
        Selection.EntireRow.Delete
    End If

    ActiveSheet.Protect Contents:=True
End Sub

Sub Remover_componente()
    ActiveSheet.Unprotect "gumemura"
    
    Dim confirmRemoveComp As Integer, compName As String
 
    confirmRemoveComp = MsgBox("Tem certeza que quer remover o componente?", vbQuestion + vbYesNo + vbDefaultButton2, "Remover componente")
    
    If confirmRemoveComp = vbYes Then
        Set componente_A_Eliminar = Application.InputBox("Selecione a célula com o nome do componente a ser removido", "Remover componente", Type:=8)
        compName = ActiveSheet.Cells(componente_A_Eliminar.Cells(1, 1).Row, "A").Value
        confirmRemoveComp = MsgBox("Componente a ser removido" + vbNewLine + vbNewLine + vbTab + compName + vbNewLine + vbNewLine + "Tem certeza?", vbQuestion + vbYesNo + vbDefaultButton2, "Remover componente")
        
        If confirmRemoveComp = vbYes Then
            componente_A_Eliminar.EntireRow.Delete
        End If
    End If
    ActiveSheet.Protect Contents:=True
End Sub

Sub Remover_Lote()
    ActiveSheet.Unprotect "gumemura"
        Dim confirmRemoveLote As Integer, loteName As String
        
        confirmRemoveLote = MsgBox("Tem certeza que quer remover o lote inteiro?", vbQuestion + vbYesNo + vbDefaultButton2, "Remover lote")
        If confirmRemoveLote = vbYes Then
            Set lote_A_Eliminar = Application.InputBox("Selecione a célula com o nome do lote a ser removido", "Remover lote", Type:=8)
            loteName = ActiveSheet.Cells(2, lote_A_Eliminar.Cells(1, 1).Column).Value
            confirmRemoveLote = MsgBox("Lote a ser removido" + vbNewLine + vbNewLine + vbTab + loteName + vbNewLine + vbNewLine + "Tem certeza?", vbQuestion + vbYesNo + vbDefaultButton2, "Remover lote")
            
            If confirmRemoveLote = vbYes Then
                Columns(lote_A_Eliminar.Cells(1, 1).Column).EntireColumn.Delete
            End If
        
        End If
        
    ActiveSheet.Protect Contents:=True
End Sub

Sub CriarBotao()
    ActiveSheet.Unprotect "gumemura"
    Dim nomeDaPlaca As String
    
    nomeDaPlaca = Application.InputBox("", "Nome da Placa", vbNullString, Type:=2)
    
    If nomeDaPlaca <> "" And nomeDaPlaca <> "Falso" Then
        Dim botaoPlaca As Button, colunaAtiva As Integer
        
        ' Pega a ultima coluna que tenha algum dado da linha A (linha A)
        lastColunm = Cells(1, Columns.Count).End(xlToLeft).Column
        
        colunaAtiva = lastColunm + 1
        
        Cells(1, colunaAtiva) = nomeDaPlaca
        
        'ActiveSheet.Range(Cells(1, colunaAtiva), Cells(2, colunaAtiva)).Merge
        
        Set celulaAlvo = ActiveSheet.Range(Cells(1, colunaAtiva), Cells(2, colunaAtiva))
        Set botaoPlaca = ActiveSheet.Buttons.Add(celulaAlvo.Left, celulaAlvo.Top, celulaAlvo.Width, celulaAlvo.Height)
        With botaoPlaca
          .OnAction = "TesteBotao"
          .Caption = nomeDaPlaca
          .Name = nomeDaPlaca
        End With
    ElseIf nomeDaPlaca = "" Then
        MsgBox "Nome vazio"
    End If
    
    'Verificando se o botao cancelar foi apertado
    'If nomeDaPlaca = "Falso" Then
        'MsgBox "Cancelou"
    'Else
        'Verificando se a string é vazia
        'If nomeDaPlaca = "" Then
            'MsgBox "String Vazia"
        'Else
            'MsgBox nomeDaPlaca
        'End If
    'End If
    ActiveSheet.Protect Contents:=True
End Sub



Sub TesteBotao()
    Dim colunaAtiva As Integer, i As Long, quantidadeMontavel As Integer, compNecessario As Long, compEmEstoque As Long, tempQuantMont As Integer
    Dim relatorio As String
    
    relatorio = ""
    quantidadeMontavel = 10000
    tempQuantMont = 0
    
    'Finding column of button
    colunaAtiva = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column
    
    'Encontrando a ultima linha com conteudo
    ultimaLinha = ActiveSheet.Cells(Rows.Count, colunaAtiva).End(xlUp).Row
    
    For i = 3 To ultimaLinha
        
        If Cells(i, colunaAtiva).Value <> "" And Cells(i, colunaAtiva).Value > 0 Then
            compNecessario = Cells(i, colunaAtiva).Value
            compEmEstoque = Cells(i, 2).Value
            
            tempQuantMont = compEmEstoque \ compNecessario
            
            'Gerando relatorio
            'Nome do componente + Em estoque + Comps necessarios para uma placa + Divisao
            relatorio = relatorio + CStr(Cells(i, 1).Value) + vbTab + CStr(Cells(i, 2).Value) + vbTab + CStr(Cells(i, colunaAtiva).Value) + vbTab + CStr(tempQuantMont) + vbNewLine
            
            If tempQuantMont <= quantidadeMontavel Then
                quantidadeMontavel = tempQuantMont
            End If
        End If
    Next i
    
    Dim nomePlaca  As String
    nomePlaca = CStr(Cells(1, colunaAtiva).Value)
    
    MsgBox nomePlaca + vbNewLine + vbNewLine + "Podem ser montadas " + CStr(quantidadeMontavel) + " placas", , nomePlaca
    MsgBox relatorio, , nomePlaca
End Sub

