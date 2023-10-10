Private Sub buton_Click()

'Butona basıldığında seçilen resmi formdaki resim ile değiştirdik
'LoadPicture(dosya yolu & "\" & dosya adı) şeklinde seçilen dosyaların yolunu yazdık
resim.Picture = LoadPicture(dizin.Path & "\" & dosyalar.FileName)
End Sub

Private Sub Dizin_Click()
'Dizin'e tıklatıldığında dosyaların yolunu dizinin yoluna eşitledik
dosyalar.Path = dizin.Path
End Sub

Private Sub surucu_Change()
'Sürücü değiştirildiğinde dizinin yolunu sürücüye eşitle dedik
dizin.Path = surucu
End Sub

Private Sub kapla()
'Bir prosedür oluşturduk ve "kapla" adını verdik. Bu prosedürü istediğimiz yerde kullanabileceğiz
'Alt kısımda bu prosedür uygulandığında uygulanan prosedürde resmin hangi özellikleri alacağını belirledik
resim.Top = 0
resim.Left = 0
resim.Width = Form1.Width
resim.Height = Form1.Height
End Sub

Private Sub Form_Resize()
'Prosedürü burda kullandık ve form ekranını küçültüp büyültürken eklenen resmin tekrar boyutlanmasını yazdığımız prosedüre göre ayarladık
Call kapla
End Sub

Private Sub Form_Active()
'Form aktif olduğunda oluşturduğumuz kapla prosedürünü kullandık
'Yani prosedürdeki resim özellikleri, ekran açıldığında mevcut resim buna göre ayarlanacak
Call kapla
End Sub

