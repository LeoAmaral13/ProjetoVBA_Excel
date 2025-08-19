Attribute VB_Name = "Módulo1"
Option Explicit

' ==== CONFIGURAÇÕES DO DASHBOARD ====
Private Const AREA_PERMITIDA As String = "A1:P27"
Private Const ZOOM_FIXO As Long = 100
Private Const SENHA As String = "dashboard"
Public DashboardAtivo As Boolean

' -----------------------------------------------------------------------------
'                                 DASHBOARD
' -----------------------------------------------------------------------------
Public Sub AtivarModoDashboard()
    Application.ScreenUpdating = False
    Application.DisplayFullScreen = True
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    Application.DisplayScrollBars = False
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        AplicarScroll ws
        BloquearSelecao ws
    Next ws
    
    AplicarTravamento Application.ActiveWindow, ActiveSheet
    Application.ScreenUpdating = True
    DashboardAtivo = True
End Sub

Public Sub DesativarModoDashboard()
    Dim ws As Worksheet, wn As Window
    Application.ScreenUpdating = False
    
    Application.DisplayFullScreen = False
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    Application.DisplayScrollBars = True
    
    For Each wn In Application.Windows
        wn.DisplayWorkbookTabs = True
        wn.DisplayHeadings = True
        wn.DisplayGridlines = True
    Next wn
    
    For Each ws In ThisWorkbook.Worksheets
        ws.ScrollArea = ""
        ws.Unprotect Password:=SENHA
        ws.EnableSelection = xlNoRestrictions
    Next ws
    
    Application.ScreenUpdating = True
    DashboardAtivo = False
End Sub

Public Sub AplicarTravamento(ByVal wn As Window, ByVal Sh As Object)
    On Error Resume Next
    If Not wn Is Nothing Then
        With wn
            .DisplayWorkbookTabs = False
            .DisplayHeadings = False
            .DisplayGridlines = False
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            .Zoom = ZOOM_FIXO
        End With
    End If
    
    If TypeOf Sh Is Worksheet Then
        AplicarScroll Sh
        BloquearSelecao Sh
    End If
End Sub

Private Sub AplicarScroll(ByVal ws As Worksheet)
    ws.ScrollArea = AREA_PERMITIDA
End Sub

Private Sub BloquearSelecao(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=SENHA
    On Error GoTo 0
    ws.Protect Password:=SENHA, _
               UserInterfaceOnly:=True, _
               DrawingObjects:=False, _
               AllowFiltering:=True, _
               AllowUsingPivotTables:=True
    ws.EnableSelection = xlUnlockedCells
End Sub

' -----------------------------------------------------------------------------
'                 MOSTRAR COMENTÁRIOS FILTRADOS POR ID DO PROJETO
' -----------------------------------------------------------------------------
Public Sub MostrarComentariosPorProjetoFiltrado(ByVal IDProjeto As String)
    Dim wsDash As Worksheet, wsData As Worksheet
    Dim lo As ListObject
    Dim idxID As Long, idxComentario As Long, idxDataConclusao As Long
    Dim arrDados As Variant
    Dim tmpData() As Variant, tmpComent() As Variant
    Dim i As Long, j As Long, n As Long
    Dim keyData As Variant, keyComent As Variant
    Dim texto As String
    Dim partes() As String, k As Long

    ' Aba da caixa de comentários
    Set wsDash = ThisWorkbook.Sheets("Entrada")

    ' Limpa a caixa
    wsDash.OLEObjects("txtComentarios").Object.Text = ""

    ' Se nenhum ID informado
    If Len(Trim(IDProjeto)) = 0 Then
        wsDash.OLEObjects("txtComentarios").Object.Text = "(nenhum projeto selecionado)"
        Exit Sub
    End If

    ' Referência da tabela de comentários
    Set wsData = ThisWorkbook.Sheets("Historico_Comentarios")
    Set lo = wsData.ListObjects("Tabela2")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        wsDash.OLEObjects("txtComentarios").Object.Text = "Nenhum comentário encontrado."
        Exit Sub
    End If

    ' Índices das colunas
    idxID = lo.ListColumns("ID_Projeto").Index
    idxComentario = lo.ListColumns("Comentario").Index
    idxDataConclusao = lo.ListColumns("Data de Conclusão").Index

    arrDados = lo.DataBodyRange.Value
    ReDim tmpData(1 To UBound(arrDados, 1))
    ReDim tmpComent(1 To UBound(arrDados, 1))
    n = 0

    ' Filtra comentários do projeto
    For i = 1 To UBound(arrDados, 1)
        If Trim(CStr(arrDados(i, idxID))) = Trim(IDProjeto) Then
            n = n + 1
            tmpData(n) = arrDados(i, idxDataConclusao)
            tmpComent(n) = arrDados(i, idxComentario)
        End If
    Next i

    If n = 0 Then
        wsDash.OLEObjects("txtComentarios").Object.Text = "(nenhum comentário encontrado)"
        Exit Sub
    End If

    ' Ordenação por data (descendente)
    For i = 2 To n
        keyData = tmpData(i)
        keyComent = tmpComent(i)
        j = i - 1
        Do While j >= 1
            If SafeCompareDates(tmpData(j), keyData) < 0 Then
                tmpData(j + 1) = tmpData(j)
                tmpComent(j + 1) = tmpComent(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        tmpData(j + 1) = keyData
        tmpComent(j + 1) = keyComent
    Next i

    ' Monta texto final
    For i = 1 To n
        If SafeFormatDate(tmpData(i)) <> "" Then
            texto = texto & "Data de Conclusão: " & SafeFormatDate(tmpData(i)) & vbCrLf
        End If
        If Len(Trim(tmpComent(i))) > 0 Then
            texto = texto & "Comentário:" & vbCrLf
            partes = Split(tmpComent(i), vbLf)
            For k = LBound(partes) To UBound(partes)
                If Trim(partes(k)) <> "" Then
                    texto = texto & "  - " & Trim(partes(k)) & vbCrLf
                End If
            Next k
        End If
        texto = texto & String(40, "-") & vbCrLf
    Next i

    wsDash.OLEObjects("txtComentarios").Object.Text = texto
End Sub

' -----------------------------------------------------------------------------
'    Obtém o ID_Projeto correspondente ao Nome_Projeto filtrado na Pivot
' -----------------------------------------------------------------------------
Public Function ObterIDPorNomeProjeto(ByVal nomeProjeto As String) As String
    Dim pt As PivotTable
    Dim pfNome As PivotField, pfID As PivotField
    Dim i As Long

    Set pt = ThisWorkbook.Worksheets("Entrada").PivotTables("Tabela dinâmica4")
    Set pfNome = pt.PivotFields("Nome_Projeto")
    Set pfID = pt.PivotFields("ID_Projeto")

    For i = 1 To pfNome.PivotItems.count
        On Error Resume Next
        If pfNome.PivotItems(i).Visible Then
            If StrComp(pfNome.PivotItems(i).Caption, nomeProjeto, vbTextCompare) = 0 Then
                ObterIDPorNomeProjeto = pfID.PivotItems(i).Caption
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next i

    ObterIDPorNomeProjeto = ""
End Function

' -----------------------------------------------------------------------------
'                   Funções auxiliares para datas
' -----------------------------------------------------------------------------
Private Function SafeCompareDates(ByVal a As Variant, ByVal b As Variant) As Long
    Dim ha As Boolean, hb As Boolean, da As Date, db As Date
    ha = IsDate(a)
    hb = IsDate(b)

    If ha Then da = CDate(a)
    If hb Then db = CDate(b)

    If ha And hb Then
        If da < db Then
            SafeCompareDates = -1
        ElseIf da > db Then
            SafeCompareDates = 1
        Else
            SafeCompareDates = 0
        End If
    ElseIf ha And Not hb Then
        SafeCompareDates = 1
    ElseIf Not ha And hb Then
        SafeCompareDates = -1
    Else
        SafeCompareDates = 0
    End If
End Function

Private Function SafeFormatDate(ByVal v As Variant) As String
    If IsDate(v) Then
        SafeFormatDate = Format$(CDate(v), "dd/mm/yyyy")
    Else
        SafeFormatDate = CStr(v)
    End If
End Function


