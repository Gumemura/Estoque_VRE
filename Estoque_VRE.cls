VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function Contador_Placas(ByVal col As Long, nome As String)
    Dim quant As Double, temp As Double, i As Long, cont As Long
    N = ActiveSheet.Cells(Rows.Count, col).End(xlUp).Row
    
    cont = 0
    For i = 3 To N
        If Cells(i, col).Value > 0 Then
            temp = Cells(i, "D").Value \ Cells(i, col).Value
            
            If cont = 0 Then
                cont = 1
                quant = temp
            ElseIf temp < quant Then
                quant = temp
            End If
        End If
    Next i
    
    MsgBox "Podem ser montadas " & quant & " placas", 0, nome
End Function

Private Sub CommandButton1_Click()
    Dim rng As Range
    Set rng = ActiveSheet.Range("G1")
    With ActiveSheet.OLEObjects("CommandButton1")
        .Top = rng.Top
        .Left = rng.Left
        .Width = rng.Width
        .Height = rng.RowHeight
    End With
    
    Call Contador_Placas(rng.Column, ActiveSheet.OLEObjects("CommandButton1").Object.Caption)
End Sub

Private Sub CommandButton2_Click()
    Dim rng As Range
    Set rng = ActiveSheet.Range("H1")
    With ActiveSheet.OLEObjects("CommandButton2")
        .Top = rng.Top
        .Left = rng.Left
        .Width = rng.Width
        .Height = rng.RowHeight
    End With
    
    Call Contador_Placas(rng.Column, ActiveSheet.OLEObjects("CommandButton2").Object.Caption)
End Sub
