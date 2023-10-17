Dim ay As Integer
Dim ayy As String

Private Sub buton_hesapla_Click()

Dim taksit As Integer
Dim yl As Integer
    
tablo.ColAlignment(2) = flexAlignLeftCenter
    
With tablo
.TextMatrix(0, 0) = "Taksit"
.TextMatrix(0, 1) = "Miktar"
.TextMatrix(0, 2) = "Tarih"
.TextMatrix(0, 3) = "Durum"
    
.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 1500
.ColWidth(3) = 1500
.Width = 5400
End With

ay = Month(Date)
yl = Year(Date)

tablo.Rows = taksit_sayisi.Text + 1
taksit = CInt(fiyat.Text) / CInt(taksit_sayisi.Text)

For t = 1 To taksit_sayisi.Text

ay = ay + 1

If ay > 12 Then
yl = yl + 1
ay = 1
End If
    
Call tarih

With tablo
.TextMatrix(t, 0) = "Taksit" & t
.TextMatrix(t, 1) = taksit
.TextMatrix(t, 2) = yl & " " & ayy
.TextMatrix(t, 3) = "Ödenmedi"
    
.Row = t
.Col = 3
.CellBackColor = vbRed
.CellForeColor = vbWhite
End With
    
Next

fark = taksit * (taksit_sayisi.Text - 1)
fark2 = fiyat.Text - fark
tablo.TextMatrix(taksit_sayisi.Text, 1) = fark2

End Sub


Private Sub tarih()
If ay = 1 Then ayy = "Ocak"
If ay = 2 Then ayy = "Şubat"
If ay = 3 Then ayy = "Mart"
If ay = 4 Then ayy = "Nisan"
If ay = 5 Then ayy = "Mayıs"
If ay = 6 Then ayy = "Haziran"
If ay = 7 Then ayy = "Temmuz"
If ay = 8 Then ayy = "Ağustos"
If ay = 9 Then ayy = "Eylül"
If ay = 10 Then ayy = "Ekim"
If ay = 11 Then ayy = "Kasım"
If ay = 12 Then ayy = "Aralık"
End Sub

Private Sub buton_odeme_Click()

With tablo
.TextMatrix(Label5.Caption, 2) = 0
.TextMatrix(Label5.Caption, 3) = "Ödendi"
.Col = 3
.Row = Label5.Caption
.CellBackColor = vbBlue
End With

End Sub


Private Sub tablo_Click()
miktar.Text = tablo.TextMatrix(tablo.Row, 1)
taksit_yazi.Text = tablo.TextMatrix(tablo.Row, 0)
Label5.Caption = tablo.Row

If tablo.Row < 1 Then
aa = MsgBox("Bu satırı kullanamazsınız!", vbCritical, "Dikkat")
tablo.Row = tablo.Rows - 1
End If

If tablo.TextMatrix(tablo.Row, 3) = "Ödendi" Then
a = MsgBox("Bu taksit ödenmiş!", vbCritical, "Dikkat")
tablo.Row = tablo.Rows - 1
End If
End Sub
