Private Sub Form_Load()
red.Max = 255
red.Min = 0

green.Max = 255
green.Min = 0

blue.Max = 255
blue.Min = 0
End Sub

Private Sub renk()
renk_ekrani.BackColor = RGB(red.Value, green.Value, blue.Value)
End Sub

Private Sub red_Scroll()
Call renk
End Sub

Private Sub green_Scroll()
Call renk
End Sub

Private Sub blue_Scroll()
Call renk
End Sub
