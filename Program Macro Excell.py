Private Sub Button37_Click()
   Range("B23").Value = 0 '0
    While Range("B23").Value < 10
    Range("B23").Value = Range("23").Value + 0.05
    DoEvents
   Wend

End Sub
