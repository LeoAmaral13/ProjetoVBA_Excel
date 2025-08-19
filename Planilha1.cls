VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Ajuste estes nomes se necessário:
Private Const NOME_PT As String = "Tabela dinâmica4"
Private Const CAMPO_NOME As String = "Nome_Projeto"
Private Const CAMPO_ID As String = "ID_Projeto"
Private Const NOME_CAIXA_COMENT As String = "txtComentarios"

Private Sub txtPesquisa_Change()
    Dim pt As PivotTable
    Dim pfNome As PivotField
    Dim pfID As PivotField
    Dim pi As PivotItem
    Dim idSel As String
    Dim textoFiltro As String
    Dim encontrou As Boolean
    Dim i As Long

    textoFiltro = Me.txtPesquisa.Text

    Set pt = Me.PivotTables("Tabela dinâmica4")
    Set pfNome = pt.PivotFields("Nome_Projeto")
    Set pfID = pt.PivotFields("ID_Projeto")

    Application.ScreenUpdating = False
    pt.ManualUpdate = True

    '=== (1) Limpa filtro atual ===
    On Error Resume Next
    pfNome.ClearAllFilters
    pfNome.EnableMultiplePageItems = True
    For Each pi In pfNome.PivotItems
        pi.Visible = True
    Next pi
    On Error GoTo 0

    encontrou = False

    '=== (2) Aplica filtro somente se houver texto ===
    If Len(textoFiltro) > 0 Then
        ' Oculta os que NÃO contêm o texto
        For i = 1 To pfNome.PivotItems.count
            If InStr(1, pfNome.PivotItems(i).Caption, textoFiltro, vbTextCompare) = 0 Then
                pfNome.PivotItems(i).Visible = False
            Else
                encontrou = True
            End If
        Next i

        ' Se não houver nenhuma correspondência, volta tudo e evita o erro 1004
        If Not encontrou Then
            For Each pi In pfNome.PivotItems
                pi.Visible = True
            Next pi
            MsgBox "Nenhum projeto encontrado com este termo.", vbExclamation
            pt.ManualUpdate = False
            Application.ScreenUpdating = True
            Call MostrarComentariosPorProjetoFiltrado("")   ' limpa a caixa
            Exit Sub
        End If
    End If

    pt.ManualUpdate = False
    pt.RefreshTable

    '=== (3) Pega o ID correspondente ao item visível do campo Nome_Projeto ===
    idSel = ""
    For i = 1 To pfNome.PivotItems.count
        If pfNome.PivotItems(i).Visible Then
            idSel = pfID.PivotItems(i).Caption
            Exit For
        End If
    Next i

    Application.ScreenUpdating = True

    '=== (4) Atualiza os comentários ===
    Call MostrarComentariosPorProjetoFiltrado(idSel)
End Sub


Private Sub AtualizarComentariosEntrada()
    Const NOME_PT As String = "Tabela dinâmica4"
    Const NOME_ABA As String = "Entrada"

    Dim pt As PivotTable
    Dim pfID As PivotField
    Dim pi As PivotItem
    Dim idSelecionado As String

    Set pt = ThisWorkbook.Worksheets(NOME_ABA).PivotTables(NOME_PT)
    Set pfID = pt.PivotFields("ID_Projeto")

    For Each pi In pfID.PivotItems
        If pi.Visible Then
            idSelecionado = CStr(pi.Caption)
            Exit For
        End If
    Next pi

    ' ?? agora passa o ID capturado para a macro de comentários
    MostrarComentariosPorProjetoFiltrado idSelecionado
End Sub

