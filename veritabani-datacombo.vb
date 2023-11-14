Private Sub hesapla_Click()
Dim ortalama As New ADODB.Recordset
Dim imlec_veriler As New ADODB.Recordset
Dim imlec_liste As New ADODB.Recordset

'Sorgular
ortalama.Open "select avg(nott) as ort from tablo2", "dsn=masa", adOpenStatic
imlec_veriler.Open "select * from tablo2", "dsn=masa", adOpenStatic
imlec_liste.Open "select distinct(bolum) from tablo2", "dsn=masa", adOpenStatic

ortalama_yazi.Caption = CInt(ortalama("ort")) 'Yukardaki sorguya "as ort" eklersek expr1000 yerine ort yazabiliriz

With grid
.Rows = imlec_veriler.RecordCount + 1
.Cols = 4

For t = 1 To imlec_veriler.RecordCount
.TextMatrix(t, 0) = imlec_veriler("ad")
.TextMatrix(t, 1) = imlec_veriler("bolum")
.TextMatrix(t, 2) = imlec_veriler("nott")

If CInt(imlec_veriler("nott")) < CInt(ortalama("ort")) Then
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

.TextMatrix(t, 3) = imlec_veriler("cins")
imlec_veriler.MoveNext
Next

.TextMatrix(0, 0) = "Ad"
.TextMatrix(0, 1) = "Bölüm"
.TextMatrix(0, 2) = "Not"
.TextMatrix(0, 3) = "Cinsiyet"

End With

Set DataCombo1.RowSource = imlec_liste
DataCombo1.ListField = "bolum"
End Sub

'Sadece listede seçilen bölümü grid'de yazdırır
Private Sub DataCombo1_Click(Area As Integer)

Dim imlec_bolum As New ADODB.Recordset
imlec_bolum.Open "select * from tablo2 where bolum = '" & DataCombo1 & "'", "dsn=masa", adOpenStatic

With grid
.Rows = imlec_bolum.RecordCount + 1
.Cols = 4

For t = 1 To imlec_bolum.RecordCount
.TextMatrix(t, 0) = imlec_bolum("ad")
.TextMatrix(t, 1) = imlec_bolum("bolum")
.TextMatrix(t, 2) = imlec_bolum("nott")
.TextMatrix(t, 3) = imlec_bolum("cins")
imlec_bolum.MoveNext
Next
End With
End Sub
