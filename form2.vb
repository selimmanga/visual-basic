Private Sub firta_Click()
Form2.Show
End Sub

Private Sub formkapat_Click()
Form2.Unload
End Sub

Private Sub listele_buton_Click()
Dim liste As New ADODB.Recordset
Dim say As Integer
Dim metin As String

say = 0

liste.Open "select * from tablo1", "dsn=masa", adOpenStatic

With MSFlexGrid1
.Rows = liste.RecordCount + 1
.Cols = 3

For t = 1 To liste.RecordCount
.TextMatrix(t, 0) = liste("ad")
.TextMatrix(t, 1) = liste("soyad")
.TextMatrix(t, 2) = liste("yas")
say = say + CInt(liste("yas"))
'Yeni kayıta geçmesi için bunu yazmamız gerekir yoksa hepsine aynı veriyi yazdırır
liste.MoveNext
Next
End With

metin = "Listedeki isimlerin yaş ortalaması " & CInt(say / liste.RecordCount) & "'dır."
Label1.Caption = metin
End Sub
