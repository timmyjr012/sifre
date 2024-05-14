'Şifre Çözücü Sistem'

'Burada Ekrana yazı yazdırdık
mesaj = MsgBox("Bu Kisma Adinizi Yazabilirsiniz...")

' sifre adlı bir obje tanımladık.
set sifre = WScript.CreateObject("WScript.Shell")

'Burada Kullanıcıdan Çözümlenecek Metni İstedik.
sifretur = inputbox("Cozulecek metni girin")

'Burada Kullanıcının girdiği metni ters çevirdik.
sifretur = StrReverse(sifretur)

'Burada Cozumleme rakamını istedik.
kacli = inputbox("Cozumlenme Rakami < 1 - 5 arasi rakam girin >")

'Burada Kontrol yaptik
if kacli = "" Then
MsgBox("Bos Birakilmaz!!")
kacli = inputbox("Cozumlenme Rakami < 1 - 5 arasi rakam girin >")
MsgBox("2 kere hata yaptiginizdan dolayi program kapatildi.")
wscript.quit()
Else 
If kacli <= 5 Then
			else
			MsgBox("Lutfen 1 ve 5 arasi bir rakam girdiginizden emin olun")
		kacli = inputbox("Kacli Sifrelensin < 1 - 5 arasi >")
		If kacli <= 5 Then
		
		Else
		
		MsgBox("2 kere hata yaptiginizdan dolayi program kapatildi.")
		wscript.quit()
		End If
		
        End If
End If

'Burada Not defterini açmasını sağladık.
sifre.Run "%windir%\\notepad"

'Burada programımızı kapatıp arka planda çalışmasını sağladık.
wscript.sleep 500

'Burada yukarıda yazdığımız işlemleri yerine getirmesini sağladık
sifre.sendkeys encode(sifretur)



'Yukardaki azalış sorgusunu yerine getirdi.
function encode(azalis)
For yazdir = 1 To Len(azalis)
not_defteri = Mid(azalis, yazdir, 1)
not_defteri = Chr(Asc(not_defteri)-kacli)
coded = coded & not_defteri
Next
encode = coded 

'Uygulama Kapandı
End Function