from gib import *
from  termcolor import *

while True:
    cprint("Lütfen Yapmak İstediğiniz İşlemi Seçiniz: ", "blue")
    cprint("1-) Fatura Kesme \n2-) Fatura İndirme \n3-) Zipten Açma \n4-) Fatura Okuma - Excel Yazma \n5-) 2-3-4-5 Numaralı işlemler \n6-) Çıkış ","green")
    scanner = int(input("\n"))
    user_name = input("User Name: ")
    password = input("Password: ")
    a = Gib(user_name, password)
    if scanner == 1:
        cprint("Fatura Kesme İşlemi Başlamıştır...","red")
        cprint("Şuanlık Bu Fonkisyonumuz Çalışmamaktadır.... ")
        cprint("Fatura Kesimi Tamamlandı...","red")
    elif scanner == 2:
        cprint("Fatura İndirme İşlemi Başlamıştır...","red")
        a.openGib()
        a.goModulTaslak()
        a.downloadInvoinces()
        cprint("Fatura İndirme Tamamlandı...", "red")
    elif scanner == 3:
        cprint("Dosyalar Zipten Çıkarılıyor","red")
        a.unZip()
        cprint("Fatura Zipten Çıkarma Tamamlandı...", "red")
    elif scanner == 4:
        cprint("Faturalarınız Xml'den okunuyor...","red")
        a.files()
        a.writeXlsx()
        cprint("Faturaların :Ok Tamamlandı...", "red")
    elif scanner == 5:
        cprint("Toplu işleminiz başlamıştır....""blue")
        a.openGib()
        a.goModulTaslak()
        a.downloadInvoinces()
        a.unZip()
        a.files()
        a.writeXlsx()
        cprint("Fatura İşlemleri Tamamlandı...", "red")
    elif scanner == 6:
        cprint("Program Sonlandırılıyor...")
        time.sleep(1)
        break
    else:
        cprint("Hatalı Tuşlama Yaptınız Tekrar Deneyiniz.","red")
