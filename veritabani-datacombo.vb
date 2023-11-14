Private Sub hesapla_Click()
Dim ortalama As New ADODB.Recordset
Dim liste As New ADODB.Recordset
Dim listele As New ADODB.Recordset

ortalama.Open "select avg(nott) as ort from tablo2", "dsn=masa", adOpenStatic
liste.Open "select * from tablo2", "dsn=masa", adOpenStatic
listele.Open "select distinct(bolum) from tablo2", "dsn=masa", adOpenStatic

Label2.Caption = CInt(ortalama("ort")) 'Yukardaki sorguya "as ort" eklersek expr1000 yerine ort yazabiliriz

With grid

.Rows = liste.RecordCount + 1
.Cols = 4

For t = 1 To liste.RecordCount
.TextMatrix(t, 0) = liste("ad")
.TextMatrix(t, 1) = liste("bolum")
.TextMatrix(t, 2) = liste("nott")

If CInt(liste("nott")) < CInt(ortalama("ort")) Then
.Row = t
.Col = 2
.CellBackColor = vbRed
.CellForeColor = vbWhite
Else
.Row = t
.Col = 2
.CellBackColor = vbBlue
.CellForeColor = vbWhite
End If

.TextMatrix(t, 3) = liste("cins")

liste.MoveNext
Next

.TextMatrix(0, 0) = "Ad"
.TextMatrix(0, 1) = "Bölüm"
.TextMatrix(0, 2) = "Not"
.TextMatrix(0, 3) = "Cinsiyet"

End With

Set DataCombo1.RowSource = listele
DataCombo1.ListField = "bolum"
End Sub

'Sadece listede seçilen bölümü grid'de yazdırır
Private Sub DataCombo1_Click(Area As Integer)
Dim a As New ADODB.Recordset
a.Open "select * from tablo2 where bolum = '" & DataCombo1 & "'", "dsn=masa", adOpenStatic

With grid
.Rows = a.RecordCount + 1
.Cols = 4
For t = 1 To a.RecordCount
.TextMatrix(t, 0) = a("ad")
.TextMatrix(t, 1) = a("bolum")
.TextMatrix(t, 2) = a("nott")
.TextMatrix(t, 3) = a("cins")
a.MoveNext
Next
End With
End Sub
