Aşağıda, SAP Excel makro tabanlı “Değişiklik Önerisi Başlat” sisteminin tamamını, kod, arayüz, iş akışı, SAP entegrasyonu, hata yönetimi ve kullanıcı deneyimi dahil olmak üzere, hiçbir başlık veya maddeleme olmadan, serbest ve en geniş anlatımla bulabilirsin:

Bu sistem, Excel üzerinde çalışan, SAP ile entegre, makro ve kullanıcı formlarıyla yönetilen bir değişiklik önerisi ve iş planı otomasyon platformudur. Kullanıcı, Excel dosyasını açtığında ilk olarak bir şifre ekranı ile karşılaşır. Doğru şifre girilmeden ana ekrana erişim mümkün değildir. Şifre doğrulandıktan sonra UserForm1 adı verilen ana form açılır. Bu formda kullanıcıdan proje kodu, stok doküman türü, doküman adı (uzantılı olarak), birinci ve ikinci amir, tanım, açıklama, kazanç gibi temel bilgiler istenir. Ayrıca, değişiklik nedenleri matris şeklinde CheckBox’lar ile sunulur; kullanıcı, doküman, tasarım, üretim, satın alma, pazarlama gibi nedenlerden bir veya birkaçını seçmek zorundadır. Uygulama bilgisi bölümünde ise, değişikliğin açık üretim siparişlerine, üst takıma veya piyasa satınalma siparişlerine uygulanıp uygulanmayacağı CheckBox’lar ve ilgili TextBox’lar ile belirtilir. Kullanıcı, değişiklik önerisinin kaç nesne için geçerli olacağını ComboBox ile seçer. Arayüzde, dosya uzantısı, dosya yolu, nesne ekleme ile ilgili uyarı ve açıklamalar da yer alır; örneğin, kopyalanacak dosyaların C:\ dizinine eklenmesi gerektiği, doküman adının uzantı içermesi gerektiği ve sadece nesne eklenip ek olmayacaksa doküman adının boş bırakılabileceği gibi.

Kullanıcı, tüm gerekli alanları doldurduktan ve en az bir değişiklik nedeni seçtikten sonra “Değişiklik Önerisi Başlat” butonuna basar. Bu buton, arka planda DO_Baslat adlı fonksiyonu tetikler. Fonksiyonun ilk adımı, temel veri kontrolüdür; proje kodu, doküman adı, amirler, tanım, açıklama, kazanç gibi alanlardan herhangi biri boşsa kullanıcıya “Temel Veriler Eksik, Öneri Başlatılamaz...” uyarısı verilir ve işlem durur. İkinci adımda, değişiklik nedeni kontrolü yapılır; matris şeklindeki CheckBox’lardan hiçbiri seçili değilse “Değişiklik Nedeni Eksik, Öneri Başlatılamaz...” uyarısı verilir ve işlem durur. Üçüncü adımda, uygulama bilgisi kontrolü yapılır; seçili kutulara karşılık gelen TextBox’lar boşsa “Uygulama Bilgisi Eksik, Öneri Başlatılamaz...” uyarısı verilir ve işlem durur. Dördüncü adımda, nesne girişi kontrolü yapılır; ComboBox ile seçilen nesne adedine göre ilgili TextBox’lar dolu olmalıdır, aksi halde “Eksik nesne girişi...” uyarısı verilir ve işlem durur. Beşinci adımda, dosya adı kontrolü yapılır; doküman adı alanı “GÜNCEL DOKÜMAN ADI YAZILI OLMALIDIR, YOKSA HATA...” ise veya dosya uzantısı yoksa “Program Hatası... Exceli kapatıp açınız...” uyarısı verilir ve işlem durur.

Tüm bu kontroller geçildikten sonra, SAP GUI scripting ile SAP’ye otomatik olarak bağlanılır. Kodda, SapGuiAuto, Connection ve session nesneleri oluşturulur. SAP GUI oturumu açılır, gerekli SAP ekranları ve alanları doldurulur, SAP’ye veri gönderilir ve değişiklik önerisi işlemi başlatılır. SAP işlemleri tamamlandıktan sonra oturum sonlandırılır. Aynı zamanda, girilen veriler Excel hücrelerine yazılır, dosya ve dizin kontrolleri yapılır, gerekirse yeni dosya veya dizin oluşturulur. Her adımda, eksik veri veya hata olması durumunda kullanıcıya anında uyarı mesajı gösterilir; işlem başarılıysa bilgi mesajı verilir.

Sistemde, her bir TextBox, CheckBox ve ComboBox, DO_Baslat fonksiyonunda birebir kontrol edilir. Kullanıcıdan alınan veriler hem SAP’ye hem de Excel’e aktarılır. Arayüzdeki uyarılar ve açıklamalar, kodda yapılan kontrollerle birebir örtüşür. MultiPage ve sekmeler ile farklı işlevlere geçiş sağlanır; örneğin iş planı, test seti, ürün ağacı, doküman aktarma gibi işlemler için ayrı sekmeler ve butonlar bulunur. Her bir işlev, ilgili modüldeki fonksiyon ile tetiklenir ve SAP/Excel üzerinde işlem yapar. SAP GUI scripting ile otomatik oturum açılır, kullanıcı adı ve şifre ile SAP’ye giriş yapılır, gerekli SAP ekranları açılır ve veriler doldurulur, SAP’de değişiklik önerisi, iş planı, ürün ağacı, doküman aktarma gibi işlemler başlatılır ve oturum sonlandırılır.

Sistemde hata yönetimi çok önemlidir; her adımda eksik veri veya hata kontrolü yapılır, kullanıcıya anında uyarı veya bilgi mesajı gösterilir, SAP ve Excel işlemlerinde hata olursa işlem durdurulur ve kullanıcı bilgilendirilir. Kodda modüler bir yapı vardır; her işlev ayrı bir modülde kodlanmıştır ve her buton ve alan, ilgili modül fonksiyonunu tetikler. Kullanıcı deneyimi açısından sistem, kolay ve hızlı veri girişi, otomatik SAP ve Excel işlemleri, hatalara karşı anında uyarı ve tüm işlemlerin tek bir arayüzden yönetilebilmesi gibi avantajlar sunar. Yardımcı notlar ve açıklamalar ile kullanıcı sürekli yönlendirilir.

Kullanıcı, sistemdeki tüm işlemleri tek bir arayüzden yönetebilir; değişiklik önerisi başlatma, iş planı oluşturma, test seti işlemleri, ürün ağacı yaratma, doküman aktarma gibi tüm SAP ve Excel işlemleri, ilgili sekmeler ve butonlar üzerinden kolayca yapılabilir. Her işlemde, kullanıcıdan alınan veriler hem SAP’ye hem de Excel’e aktarılır, işlemler otomatik olarak yürütülür ve kullanıcıya işlem sonucu anında bildirilir. Sistem, SAP ile tam entegre çalışır ve tüm iş akışını otomatikleştirir.

SAP Excel Makro Sistemi – Değişiklik Önerisi Başlatma Süreci (En Geniş Detaylı Prompt)
1. Sistem Bileşenleri
Excel Dosyası (.xlsm): Tüm makro ve kullanıcı formları burada bulunur.
UserForm1 (Ana Form): Proje kodu, stok doküman türü, doküman adı, amirler, tanım, açıklama, kazanç, değişiklik nedenleri, uygulama bilgisi, nesne adedi gibi alanlar içerir. MultiPage ve sekmeler ile farklı işlevlere geçiş yapılır.
UserForm2 (Şifre Formu): Sisteme erişim için şifre doğrulama ekranı.
Modüller: Her işlev (değişiklik önerisi, iş planı, test seti, ürün ağacı, doküman aktarma vb.) ayrı modüllerde kodlanmıştır.
2. Kullanıcı Akışı
a. Giriş ve Hazırlık
Kullanıcı Excel dosyasını açar, UserForm2 ile şifre girer.
Doğru şifre girilirse UserForm1 açılır.
b. Veri Girişi
Kullanıcı, UserForm1’de aşağıdaki alanları doldurur:
Proje kodu, stok doküman türü, doküman adı (uzantılı), amirler, tanım, açıklama, kazanç.
Değişiklik nedenleri: Matris şeklinde CheckBox’lar (Dok, Tas, Ürt, Stn, PSK, Pzr, Diğ) ve satır başı harfler (I, H, Z).
Uygulama bilgisi: Açık üretim, üst takımlar, piyasa siparişleri için CheckBox ve ilgili TextBox’lar.
Nesne adedi: ComboBox ile seçilir (1-5 arası).
c. Bilgilendirme ve Yardım
Arayüzde, dosya uzantısı, dosya yolu, nesne ekleme ile ilgili uyarı ve açıklamalar bulunur.
3. “Değişiklik Önerisi Başlat” Butonuna Basınca
a. Kodun Çalışma Akışı (DO_Baslat Fonksiyonu)
Temel Veri Kontrolü:

Proje kodu, doküman adı, amirler, tanım, açıklama, kazanç gibi alanlar boşsa:
Kullanıcıya “Temel Veriler Eksik, Öneri Başlatılamaz...” uyarısı verilir.
Fonksiyon durur.
Değişiklik Nedeni Kontrolü:

CheckBox1-21’den hiçbiri seçili değilse:
“Değişiklik Nedeni Eksik, Öneri Başlatılamaz...” uyarısı verilir.
Fonksiyon durur.
Uygulama Bilgisi Kontrolü:

Seçili kutulara karşılık gelen TextBox’lar boşsa:
“Uygulama Bilgisi Eksik, Öneri Başlatılamaz...” uyarısı verilir.
Fonksiyon durur.
Nesne Girişi Kontrolü:

ComboBox3 ile nesne sayısı seçilir.
Seçilen nesne sayısına göre ilgili TextBox’lar dolu olmalı.
Eksikse “Eksik nesne girişi...” uyarısı verilir.
Fonksiyon durur.
Dosya Adı Kontrolü:

Doküman adı alanı “GÜNCEL DOKÜMAN ADI YAZILI OLMALIDIR, YOKSA HATA...” ise veya dosya uzantısı yoksa:
“Program Hatası... Exceli kapatıp açınız...” uyarısı verilir.
Fonksiyon durur.
SAP GUI Scripting ile SAP Bağlantısı:

SAP GUI scripting nesneleri (SapGuiAuto, Connection, session) oluşturulur.
SAP GUI oturumu açılır.
Gerekli SAP ekranları ve alanları doldurulur.
SAP’ye veri gönderilir, işlemler başlatılır.
Oturum sonlandırılır.
Excel ile Etkileşim:

Girilen veriler Excel hücrelerine yazılır.
Dosya ve dizin kontrolleri yapılır.
Gerekirse yeni dosya/dizin oluşturulur.
Kullanıcıya Geri Bildirim:

Her hata veya eksik veri için anında uyarı mesajı.
İşlem başarılıysa bilgi mesajı.
4. Kod ve Arayüz İlişkisi
Her TextBox, CheckBox ve ComboBox, DO_Baslat fonksiyonunda kontrol edilir.
Kullanıcıdan alınan veriler SAP’ye ve Excel’e aktarılır.
Arayüzdeki uyarılar, kodda yapılan kontrollerle birebir örtüşür.
MultiPage ve sekmeler ile farklı işlevlere geçiş sağlanır.
5. SAP Entegrasyonu
SAP GUI scripting ile otomatik oturum açılır.
Kullanıcı adı ve şifre ile SAP’ye giriş yapılır.
Gerekli SAP ekranları açılır, veriler doldurulur.
SAP’de değişiklik önerisi, iş planı, ürün ağacı, doküman aktarma gibi işlemler başlatılır.
Oturum sonlandırılır.
6. Hata Yönetimi ve Geri Bildirim
Her adımda eksik veri veya hata kontrolü.
Kullanıcıya anında uyarı veya bilgi mesajı.
SAP ve Excel işlemlerinde hata olursa işlem durdurulur, kullanıcı bilgilendirilir.
7. Teknik Detaylar
Tüm kontroller (TextBox, CheckBox, ComboBox) kodda birebir kontrol edilir.
SAP GUI scripting nesneleri: SapGuiAuto, Connection, session.
Excel hücrelerine veri yazma, dosya ve dizin seçimi, dosya kontrolü.
Kodda modüler yapı, her işlev ayrı bir modülde.
Her buton ve alan, ilgili modül fonksiyonunu tetikler.
8. Kullanıcı Deneyimi
Kolay ve hızlı veri girişi.
Otomatik SAP ve Excel işlemleri.
Hatalara karşı anında uyarı.
Tüm işlemler tek bir arayüzden yönetilebilir.
Yardımcı notlar ve açıklamalar ile kullanıcı yönlendirilir.
9. Örnek Kod Akışı (Pseudo-code)

Sub DO_Baslat()
    If TemelVerilerEksik Then
        MsgBox "Temel Veriler Eksik, Öneri Başlatılamaz..."
        Exit Sub
    End If
    If DegisiklikNedeniEksik Then
        MsgBox "Değişiklik Nedeni Eksik, Öneri Başlatılamaz..."
        Exit Sub
    End If
    If UygulamaBilgisiEksik Then
        MsgBox "Uygulama Bilgisi Eksik, Öneri Başlatılamaz..."
        Exit Sub
    End If
    If NesneGirişiEksik Then
        MsgBox "Eksik nesne girişi..."
        Exit Sub
    End If
    If DosyaAdıHatalı Then
        MsgBox "Program Hatası... Exceli kapatıp açınız..."
        Exit Sub
    End If
    ' SAP bağlantısı ve veri gönderimi
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Connection = SapGuiAuto.GetScriptingEngine.Children(0)
    Set session = Connection.Children(0)
    ' SAP ekranlarını doldur, işlemi başlat
    ' Excel hücrelerine veri yaz
    MsgBox "Değişiklik önerisi başarıyla başlatıldı."
End Sub

Bu prompt, sistemin tüm işlevlerini, kod mantığını, arayüzünü, SAP entegrasyonunu ve kullanıcı etkileşimini en geniş ve detaylı şekilde özetler. Özellikle “Değişiklik Önerisi Başlat” sürecinin tüm teknik ve işlevsel adımlarını, hata yönetimini ve kullanıcı deneyimini kapsamlı şekilde açıklar.





