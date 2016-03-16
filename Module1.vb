Private Sub CommandButton1_Click()
'Заполнение массива, представляющего всех желающих играть в баскетбол'
For i = 1 To 30
Cells(1, i) = Int((61 * Rnd) + 140)
Next i
End Sub

Private Sub CommandButton2_Click()
'Расчет роста самого низкого человека в ачестве потенциального игрока в баскетбол'
Min = 200
For i = 1 To 30
If (Cells(1, i) > 180) And (Cells(1, i) < Min) Then
Min = Cells(1, i)
End If
Next i
MsgBox (Min)
End Sub

Private Sub CommandButton3_Click()
'Закрытие формы'
UserForm1.Hide
End Sub