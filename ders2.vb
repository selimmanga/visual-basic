'Değişken tanımlarken "dim değişken adı as değişken türü" şeklinde tanımlarız, ÖR: Dim sayi As Integer
'Dim sayi As Integer için:
'Değişkan adı: sayi 
'Değişken Türü: Integer

Dim ay As Integer
Dim ayy As String

Private Sub buton_hesapla_Click()

Dim taksit As Integer
Dim yl As Integer
    
'Oluşturduğumuz tablodaki hücreye yazılacak olan metinleri sol tarafa yasladık ve merkeze aldık, sol tarafa yaslamanın idsi = 2
tablo.ColAlignment(2) = flexAlignLeftCenter

'with metodu ile tagi aynı olan kodları bir arada yazabiliriz. 
'Aşağıdaki kodların hepsinde tablo değişkeni kullanıldığı için "With tablo - End With" arasında yazabildik. 
'Normal şartlarda hepsinin formatı "tablo.TextMatrix" olmalıydı, bu metod ile daha rahat yazabiliriz
'TextMatrix fonksiyonu hücrelerde işlem yapabilmemizi sağlayan bir fonksiyondur. TextMatrix(satır, sütun) şeklinde parametreler içerir
'Hücrelerin sayımı 0'dan başlar dolayısıyla 1. hücre için değerimiz 0 olur. TextMatrix(0, 0) kodunda 1. satır 1. sütun anlamına gelir.
With tablo
.TextMatrix(0, 0) = "Taksit"
.TextMatrix(0, 1) = "Miktar"
.TextMatrix(0, 2) = "Tarih"
.TextMatrix(0, 3) = "Durum"

'ColWidth kodu türkçeye çevirisiyle de anlaşılacağı gibi "sütun genişliği" anlamına gelir
'ColWidth(sütun id) şeklinde kullanılan bir fonksiyondur ve aynı şekil 0'dan sayıma başlar
.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 1500
.ColWidth(3) = 1500
.Width = 5400
End With

'Oluşturduğumuz ay ve yl değişkenlerini tarih formatına dönüştürdük, 
'Month(Date) fonksiyonu ay değişkenini 12 ay için kullanabileceğimiz bir değişken haline getirmiş oldu
ay = Month(Date)
'Year(Date) fonksiyonu yl değişkenini yıl için kullanabileceğimiz bir değişken haline getirmiş oldu    
yl = Year(Date)

'rows = satırlar 
'cols = sütunlar
'Hesaplarken girdiğimiz taksit sayısına göre tabloda hücre oluşturacağımız için, her bir taksit sayısına bir satır eklettirdik.
tablo.Rows = taksit_sayisi.Text + 1
    
'Taksit dediğimiz şeyin formülü fiyat/taksit sayisi olduğundan taksit değişkenimize bu formülü uyguladık
'CInt fonksiyonu ile bu değerleri integer'a çevirdik
'CInt = Convert Integer (integer'a dönüştür)
taksit = CInt(fiyat.Text) / CInt(taksit_sayisi.Text)

'Bir döngü oluşturalım ve t adında değişkenimiz olsun
' t'yi 1'den girilen taksit sayısı değerine kadar saydıralım
'ÖR: Girilen taksit sayısı 10 ise t değişkeni 1'den 10'a kadar saymaya başlayacaktır ve bu sayım sırasında aşağıdaki kodlar çalışacak
For t = 1 To taksit_sayisi.Text

' "ay" değişkenini her seferinde 1 arttıran kodumuz
'Döngü boyunca gerçekleşecektir, yani t 1'den 10'a kadar sayıyorsa ay = 10'a kadar gelecektir
ay = ay + 1

'Döngü boyunca gerçekleşecektir, 
'Eğer "ay" değişkeninin değeri 12'den büyük olursa yıl için oluşturduğumuz "yl" değeri 1 artacak ve "ay" değişkeni tekrar 1'e eşitlenecektir
'Bunu yapmamızın sebebi toplam 12 ay olması ve her 12 ay sonrasında 1 yıl geçmesi gerektiği için
If ay > 12 Then
yl = yl + 1
ay = 1
End If
    
'Oluşturduğumuz tarih prosedürünü burada devreye sokuyoruz, 
'İçerisinde ay değişkenine ait 12 adet verdiğimiz ay isimleri var (Ocak, Şubat, Mart...) ve bu döngüde kullanılacak
Call tarih

'With metodunu yukarda anlattığımız şekilde döngünün içinde de kullanabiliriz
'Buradaki mantık TextMatrix ile hücrelere döngü boyunca işlem yapmaktır
't değişkenin döngü boyunca her aldığı değer için hücrenin sütunlarına girilecek olan değerleri aşağıda belirliyoruz
    
'Örnek olarak; TextMatrix(t, 0) = "Taksit" & t 
'Bu kodun amacı:
'Diyelim ki t = 2, t = 2 olduğunda TextMatrix(2, 0) = "Taksit" & t kodunu çalıştır, bu kodun anlamı ise yukarda da bahsettiğim gibi,
'2. satırın 0. sütununa "Taksit" yaz ve kaçıncı taksit olduğunu "Taksit" yazısına ekle. Yani t = 2 için Taksit2 yazacaktır.
With tablo
.TextMatrix(t, 0) = "Taksit" & t
.TextMatrix(t, 1) = taksit
.TextMatrix(t, 2) = yl & " " & ayy
.TextMatrix(t, 3) = "Ödenmedi"
    
.Row = t
.Col = 3

'CellBackColor fonksiyonu ile renk değişimi yapabiliyoruz. 
'Renklerin ingilizce adını kullanmakla birlikte en başına "visual basic" kısaltması olan vb ekini koyuyoruz. Aksi takdirde çalışmayacaktır
.CellBackColor = vbRed
.CellForeColor = vbWhite
End With
'With metodunu sonlandırıyoruz
    
Next
'Döngümüzü sonlandırıyoruz

'fark adında bir değişken oluşturuyoruz ve taksit ile taksit sayısının 1 eksiğini çarpıp bu değişkene eşitliyoruz
fark = taksit * (taksit_sayisi.Text - 1)

'fark2 adında yeni bir değişken oluşturuyoruz ve girilen fiyattan yukarda hesapladığımız ilk fark değerini çıkarıp bu değişkene eşitliyoruz
fark2 = fiyat.Text - fark

'Tablomuzda, taksit sayısı hangisiyse o satırın 1. sütununa bulduğumuz "fark2" değerini yazdırıyoruz
tablo.TextMatrix(taksit_sayisi.Text, 1) = fark2
End Sub

'Yukardaki döngüde kullandığımız prosedürü burada oluşturuyoruz.
'Prosedürlerin genel amacı aynı kodları tekrarlı ve uzun bir şekilde yazmamak için tek isim altında toplamaktır
'tarih() adındaki prosedürümüzü her yazdığımızda altındaki kodlar devreye girecektir ve bu işimizi oldukça kolaylaştırır
Private Sub tarih()

'Burda yazılan kodlarda ise eğer "ay" değişkeni 1 ise "ayy" değişkenimizin adı "Ocak" olacaktır, bu şekilde 12 ay için kodlarımızı yazdık
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

'Bir hafızaya ihtiyacımız olduğu için görünmeyen label oluşturup frame'in içine koyduk, bu kullanıcı tarafından görülmeyecek şekilde ayarlanmıştır
' "Ödeme" butonuna basıldığında olacak olanları aşağıda yazdık
'Ödeme yapıldıktan sonra "Ödenmedi" yazan kısımlara mavi renkte "Ödendi" yazdırdık ve ödenmesi gereken borcu 0'a eşitledik
With tablo
.TextMatrix(Label5.Caption, 2) = 0
.TextMatrix(Label5.Caption, 3) = "Ödendi"
.Col = 3
.Row = Label5.Caption
.CellBackColor = vbBlue
End With

End Sub

'Tablo'da tıklanılan hücrenin Taksit sayısını ve ödenmesi gereken miktarını textbox'lara yazdırdık                                            
Private Sub tablo_Click()
miktar.Text = tablo.TextMatrix(tablo.Row, 1)
taksit_yazi.Text = tablo.TextMatrix(tablo.Row, 0)
Label5.Caption = tablo.Row

'Eğer tablo satırı 1'den küçükse 
'Yani en baştaki başlık için kullandığımız hücrelere tıklandıysa buralarda işlem yapılmayacağını kullanıcıya uyarı mesajı ile bildirdik                                       
If tablo.Row < 1 Then
aa = MsgBox("Bu satırı kullanamazsınız!", vbCritical, "Dikkat")
tablo.Row = tablo.Rows - 1
End If

'Tabloda tıklanılan hücredeki zaten ödendiyse tekrar tıkladığımızda kullanıcıya bu borcun ödendiğini bildiren bir mesaj gönderdik                                                
If tablo.TextMatrix(tablo.Row, 3) = "Ödendi" Then
a = MsgBox("Bu taksit ödenmiş!", vbCritical, "Dikkat")
tablo.Row = tablo.Rows - 1
End If
End Sub
