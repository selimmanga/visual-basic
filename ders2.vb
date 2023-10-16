Dim ay As Integer
Dim ayy As String

Private Sub Command1_Click()

MSFlexGrid1.ColAlignment(2) = flexAlignLeftCenter
Dim taksit As Integer
Dim yl As Integer

MSFlexGrid1.TextMatrix(0, 0) = "Taksit"
MSFlexGrid1.TextMatrix(0, 1) = "Miktar"
MSFlexGrid1.TextMatrix(0, 2) = "Tarih"
MSFlexGrid1.TextMatrix(0, 3) = "Durum"

MSFlexGrid1.ColWidth(0) = 1000
MSFlexGrid1.ColWidth(1) = 1000
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 1500
MSFlexGrid1.Width = 5400


ay = Month(Date)
yl = Year(Date)

MSFlexGrid1.Rows = Text2.Text + 1
taksit = CInt(Text1.Text) / CInt(Text2.Text)

For t = 1 To Text2.Text

ay = ay + 1

If ay > 12 Then
yl = yl + 1
ay = 1
End If

Call tarih
MSFlexGrid1.TextMatrix(t, 0) = "Taksit" & t
MSFlexGrid1.TextMatrix(t, 1) = taksit
MSFlexGrid1.TextMatrix(t, 2) = yl & " " & ayy
MSFlexGrid1.TextMatrix(t, 3) = "Ödenmedi"

MSFlexGrid1.Row = t
MSFlexGrid1.Col = 3
MSFlexGrid1.CellBackColor = vbRed
MSFlexGrid1.CellForeColor = vbWhite
Next

fark = taksit * (Text2.Text - 1)
fark2 = Text1.Text - fark
MSFlexGrid1.TextMatrix(Text2.Text, 1) = fark2
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

Private Sub Command2_Click()

With MSFlexGrid1

.TextMatrix(Label5.Caption, 2) = 0
.TextMatrix(Label5.Caption, 3) = Ödendi
.Col = 3
.Row = Label5.Caption
.CellBackColor = vbBlue

End With
End Sub


Private Sub MSFlexGrid1_Click()

Text3.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
Text4.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
Label5.Caption = MSFlexGrid1.Row

If MSFlexGrid1.Row < 1 Then
aa = MsgBox("Bu satırı kullanamazsınız!", vbCritical, "Dikkat")
MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
End If

If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = "Ödendi" Then
a = MsgBox("Bu taksit ödenmiş!", vbCritical, "Dikkat")
MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
End If

End Sub
