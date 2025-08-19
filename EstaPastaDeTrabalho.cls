VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Call AtivarModoDashboard
End Sub



Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If DashboardAtivo = True Then
        Call AplicarTravamento(Application.ActiveWindow, Sh)
    End If
End Sub

'=== Adicionar Comentário ===
Sub AdicionarComentario(ByVal id As String, ByVal Comentario As String)
    Dim ws As Worksheet, ultimaLinha As Long
    Set ws = ThisWorkbook.Sheets("Comentarios")
    
    ultimaLinha = ws.Cells(ws.Rows.count, "A").End(xlUp).Row + 1
    ws.Cells(ultimaLinha, 1).Value = id
    ws.Cells(ultimaLinha, 2).Value = Now
    ws.Cells(ultimaLinha, 3).Value = Comentario
End Sub

'=== Ver Comentário (último) ===
Function VerComentario(ByVal id As String) As String
    Dim ws As Worksheet, ultimaLinha As Long, i As Long
    Set ws = ThisWorkbook.Sheets("Comentarios")
    
    ultimaLinha = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    For i = ultimaLinha To 2 Step -1
        If ws.Cells(i, 1).Value = id Then
            VerComentario = ws.Cells(i, 3).Value
            Exit Function
        End If
    Next i
    
    VerComentario = "(nenhum comentário encontrado)"
End Function

'=== Histórico de Comentários ===
Function HistoricoDeComentarios(ByVal id As String) As String
    Dim ws As Worksheet, ultimaLinha As Long, i As Long
    Dim texto As String
    Set ws = ThisWorkbook.Sheets("Comentarios")
    
    ultimaLinha = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    For i = 2 To ultimaLinha
        If ws.Cells(i, 1).Value = id Then
            texto = texto & Format(ws.Cells(i, 2).Value, "dd/mm/yyyy HH:nn") _
                    & " - " & ws.Cells(i, 3).Value & vbCrLf
        End If
    Next i
    
    HistoricoDeComentarios = texto
End Function

