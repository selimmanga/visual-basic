Dim dakika As Integer
Dim saniye As Integer

Private Sub basla_buton_Click()

If Text2.Text = "" And Text1.Text = "" And gerisay.Value = True Then
a = MsgBox("Değer girilmeden geriye doğru saydırmazsınız.", vbCritical, "Dikkat")
Timer1.Enabled = False
End If
    
If Text1.Text = "" Then Text1.Text = 0
If Text2.Text = "" Then Text2.Text = 0

Timer1.Interval = 100
Timer1.Enabled = True

dakika = CInt(Text1.Text)
saniye = CInt(Text2.Text)
End Sub

Private Sub durdur_buton_Click()
Timer1.Enabled = False
End Sub

Private Sub sifirla_buton_Click()
Text1.Text = 0
Text2.Text = 0
Timer1.Enabled = False
End Sub

Private Sub Text2_Change()
If Text2.Text > 60 Then Text2.Text = 60
End Sub

Private Sub Timer1_Timer()

If gerisay.Value = True Then

    saniye = saniye - 1
    
    Text1.Text = dakika
    Text2.Text = saniye
    
    If saniye = 0 Then
    saniye = 59
    dakika = dakika - 1
    End If
    
    If Text1.Text = 0 And Text2.Text = 0 Then
    Timer1.Enabled = False
    End If
    
End If

If ilerisay.Value = True Then

    saniye = saniye + 1
    
    Text1.Text = dakika
    Text2.Text = saniye
    
    If saniye = 60 Then
    saniye = 0
    dakika = dakika + 1
    End If
    
End If
End Sub
