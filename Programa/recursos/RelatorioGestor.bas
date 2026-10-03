Attribute VB_Name = "RelatorioGestor"
Option Explicit

Public Sub InstalarBotaoPDF()
    Dim g As Worksheet, b As Shape, area As Range
    Set g = ThisWorkbook.Worksheets("Gestor")
    Set area = g.Range("R2:V3")
    On Error Resume Next
    Set b = g.Shapes("BotaoResumoPDF")
    On Error GoTo 0
    If b Is Nothing Then
        Set b = g.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
        b.Name = "BotaoResumoPDF"
    End If
    b.OnAction = "GerarResumoPDF"
    b.Fill.ForeColor.RGB = RGB(115, 222, 201)
    b.Line.Visible = 0
    b.TextFrame.Characters.Text = "GERAR RESUMO PDF"
    b.TextFrame.Characters.Font.Name = "Arial"
    b.TextFrame.Characters.Font.Size = 12
    b.TextFrame.Characters.Font.Bold = True
    b.TextFrame.Characters.Font.Color = RGB(16, 25, 35)
    b.TextFrame.HorizontalAlignment = -4108
    b.TextFrame.VerticalAlignment = -4108
    b.DrawingObject.PrintObject = False
    area.ClearContents
    area.Interior.Color = RGB(16, 25, 35)
End Sub

Public Sub GerarResumoPDF()
    Dim g As Worksheet, pasta As String, base As String
    Dim arquivo As String, nome As String, n As Long
    On Error GoTo Falha
    Set g = ThisWorkbook.Worksheets("Gestor")
    base = ThisWorkbook.Path
    If Len(base) = 0 Or InStr(1, base, "://", vbTextCompare) > 0 Then
        base = Environ$("USERPROFILE") & "\Documents"
    End If
    pasta = base & "\Relatorios_PDF"
    If Len(Dir$(pasta, vbDirectory)) = 0 Then MkDir pasta
    nome = "Resumo_Gestor_" & Format$(Now, "yyyy-mm-dd_hh-nn-ss")
    arquivo = pasta & "\" & nome & ".pdf"
    Do While Len(Dir$(arquivo)) > 0
        n = n + 1
        arquivo = pasta & "\" & nome & "_" & CStr(n) & ".pdf"
    Loop
    g.Calculate
    g.PageSetup.RightFooter = "Emitido em " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    g.ExportAsFixedFormat Type:=xlTypePDF, Filename:=arquivo, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    Exit Sub
Falha:
    MsgBox "Nao foi possivel gerar o PDF." & vbCrLf & Err.Description & vbCrLf & _
        "Confira se a pasta permite gravacao e tente novamente.", vbExclamation, "Resumo do gestor"
End Sub
