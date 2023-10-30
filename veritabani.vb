Private Sub kaydet_Click()
Dim kaydet As New ADODB.Recordset

'Çoklu kullanıcılar için adOpenStatic kullanıyoruz
kaydet.Open "insert into tablo1(ad, soyad, yas) values ('" & ad.Text & "', '" & soyad.Text & "','" & yas.Text & "')", "dsn=masa", adOpenStatic
aa = MsgBox("Kayıt başarılı", vbInformation, "Bilgi")
End Sub
