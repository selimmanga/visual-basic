Private Sub Form_Load()

VScroll1.Max = 255
VScroll1.Min = 0

VScroll2.Max = 255
VScroll2.Min = 0

VScroll3.Max = 255
VScroll3.Min = 0

End Sub

Private Sub renk()
Picture1.BackColor = RGB(VScroll1.Value, VScroll2.Value, VScroll3.Value)
End Sub


Private Sub VScroll1_Scroll()
Call renk
End Sub

Private Sub VScroll2_Scroll()
Call renk
End Sub

Private Sub VScroll3_Scroll()
Call renk
End Sub
