CATIA
1.adım: ürünü catia da açtıktan sonra start>machining>surface machining yolunu izleyerek burda
ürün seçilerek şekildeki gibi kütüğü çizilir ve onaylanır. Bu işlem kullanıcının sipariş edicek bütün 
parçaların kütüğünün çıkarılmasıyla diğer adıma geçilir.
2.adım: şekildeki gibi ürünün kütüğü oluştuktan sonra window sekmesinden ürüne dönülür.
3.adım 
Bu adımda measure de gerekli yerlerin aktif edilmesi gerekir yoksa BBLX gibi değerler hata vercektir 
bu şekilde aktif edilir
Eğer meansure interia bulunmazsa bu kodu makroda çalıştırılarakta açılması sağlabilir
Sub dff()
CATIA.StartCommand ("Measure Inertia")
End Sub 
Adım 4
Makro dizini açılır ve select tuşuna basılıp programımız hangi dizideyse seçilir
4.adım
Bu adımda toolbar eklenir 
Burda toolbar seçilir ve new den yeni toolbar eklenir ve adı verilir 
Daha sonra commands sekmesine gidilir ve soldaki alandan makros seçilir ve daha önce kütüphaneye 
eklediğimiz makromuzu burda icon da ekleyerek makroyu tolbara sürükleyip bırakabilirz ve toolbarda 
makromuzu icon halinde görülür.
Son adım1:
Bu adımda önceden oluşturduğumuz kütükler isterse tekli istenirsede ctrl tuşuna basılı şekilde seçilir 
ve icona basarak programın çalışması sağlanır
Dim oSel As Selection
Set oSel = CATIA.ActiveDocument.Selection
Bu adımda kullanıcının seçtiği kütükleri selection ifadesiyle osel objesine atıyoruz.
Dim oSe2 As Selection
Set oSe2 = CATIA.ActiveDocument.Selection
Dim Name As String
Name değişkeni osel değişkenine atanan ürünün adıdır.
iCount = oSel.Count
kaç adet select ettiğimizi iCount değişkenine atadık.
ReDim copies(iCount)
YourFile = "C:\Users\sefa_\OneDrive\Masaüstü\Yeni Microsoft Excel Çalışma Sayfası.xlsm" 
Your file bizim excel dosyamızın bilgisayarda kayıtlı olduğu dizini temsil eder.
Dim EXCEL As Object
Set EXCEL = CreateObject("Excel.Application")
EXCEL.workbooks.Open YourFile
EXCEL.Visible = True
'Set myworkbook = EXCEL.worksheets.Add
Bu dizin catiada excelin tanıtılması ve program çalışırken görünür hale gelip verileri almasını 
sağlıyor.
'MsgBox "Number of shapes found: " & icount
Burda kaç adet ürün seçildi bilgisini ekrana yazdırmak için 
For A = 1 To iCount
Set copies(A) = oSel.Item(A).Value
Next
Bu döngüde copies arrayinin içine seçtiğimiz ürünlerin isimlerini atıyoruz ve icount değeriylede bu 
döngüyü döndürüyoruz
For i = 1 To iCount
Name = copies(i).Name
i.arraydeki ürün adını name e atıyoruz ve her döngüde array indexi değişeceğinden bir sonraki 
seçili ürünün adını name değişlkenine atıyor
oSe2.Search ("Name=" & Name & "*,all")
bu adımda search metoduyla arrayimizin içine attığımız isimi bularak o ürünü select ediyoruz.
Dim ProdDoc
Set ProdDoc = CATIA.ActiveDocument
Dim Prod
Set Prod = ProdDoc.Product
CATIA.StartCommand ("Measure Inertia")
Kütüklerin bilgilerini almak için catiada meansure inertia programını kullanmak için run ediyoruz.
Dim Xdim
Dim Ydim
Dim Zdim
Xdim = Prod.Parameters.Item("BBLx").Value
Ydim = Prod.Parameters.Item("BBLy").Value
Zdim = Prod.Parameters.Item("BBLz").Value
BBLx,BBLy,BBLz meansure programındaki,oluşan kütüğün uzunlarını atamamıza yarıyan komut 
dizisi.
Xdim = CStr(Round(Xdim, 0)+10)
Ydim = CStr(Round(Ydim, 0)+10)
Zdim = CStr(Round(Zdim, 0)+10)
Burda ürünün x, y , z değerlerini bulduktandan sonra round formülüyle sayıları en yakın tam sayıya 
yuvarlıyoruz ve işletmenin istediği üzere her kenara işleme payı +10 mm olarak ekliyoruz.
EXCEL.Cells(3, 3).Value = Now
EXCEL.Cells(i + 2, 4).Value = i
EXCEL.Cells(i + 2, 5).Value = Xdim
EXCEL.Cells(i + 2, 6).Value = Ydim
EXCEL.Cells(i + 2, 7).Value = Zdim
Bu dizinde exceldeki ilgili hücrelere zaman ve x,y,z değerlerini yazdırıyoruz
'MsgBox "X = " & Xdim _
'& vbCr & "Y = " & Ydim _
'& vbCr & "Z = " & Zdim
İsternirse bu adımda ürün boyutları ekrana yazdırılabilir.
Next
'MsgBox "seçilen parça sayısı: " & iCount
Parça sayısını ekrana yazdırmak için.
Excel
Burda parça sayısı ve o parçanın boyutlarını program çalışınca otomatik ilgili satırlara yazıyor ve mail 
atmak için gönderilmeden üzere işlem tarihini yazıyor ve şekilde x yazan yere kullanıcı mail 
göndereceği adresi yazıp gönder butonuna basıp maili gönderebiliyor.
Excel makro kodu:
Private Sub CommandButton1_Click()
Dim Sayfa As Worksheet
 Dim Alan As Range
 Dim daralan As Range
 If Cells(2, 2) = "" Then GoTo HATA
 On Error GoTo HATA
 With Application
 .ScreenUpdating = False
 .EnableEvents = False
 End With
 saydir = WorksheetFunction.CountIf(Range("D:D"), "<>") + 1
 DinamikAlan = "c2:" & "G" & saydir
 Set Alan = Worksheets("Sayfa1").Range(DinamikAlan)
 
 Set Sayfa = ActiveSheet
 With Alan
 .Parent.Select
 Set daralan = ActiveCell
 .Select
 ActiveWorkbook.EnvelopeVisible = True
 With .Parent.MailEnvelope
 .Introduction = "ÜRÜN SİPARİŞ BİLGİLERİ"
 With .Item
 .To = Cells(2, 2)
 .CC = Cells(3, 2)
 .Subject = Cells(1, 2)
 
 .Send
 End With
 End With
 daralan.Select
 End With
 
 Sayfa.Select
HATA:
 With Application
 .ScreenUpdating = True
 .EnableEvents = True
 End With
 
End Sub
Makro saf kodu
Sub CATMain()
Dim oSel As Selection
Set oSel = CATIA.ActiveDocument.Selection
Dim oSe2 As Selection
Set oSe2 = CATIA.ActiveDocument.Selection
Dim Name As String
iCount = oSel.Count
ReDim copies(iCount)
YourFile = "C:\Users\USER\Desktop\Yeni Microsoft Excel Çalışma Sayfası.xlsm"
Dim EXCEL As Object
Set EXCEL = CreateObject("Excel.Application")
EXCEL.workbooks.Open YourFile
EXCEL.Visible = True
'Set myworkbook = EXCEL.worksheets.Add
'MsgBox "Number of shapes found: " & icount
For A = 1 To iCount
Set copies(A) = oSel.Item(A).Value
Next
For i = 1 To iCount
'oSel.Item(i).Value.Name
'Set copies(i) = oSel.Item(i).Value
'oSel.Add copies(i)
Name = copies(i).Name
'MsgBox oSel.Item(i).Value.Name
oSe2.Search ("Name=" & Name & "*,all")
Dim ProdDoc
Set ProdDoc = CATIA.ActiveDocument
Dim Prod
Set Prod = ProdDoc.Product
CATIA.StartCommand ("Measure Inertia")
Dim Xdim
Dim Ydim
Dim Zdim
Xdim = Prod.Parameters.Item("BBLx").Value
Ydim = Prod.Parameters.Item("BBLy").Value
Zdim = Prod.Parameters.Item("BBLz").Value
Xdim = CStr(Round(Xdim, 0)+10)
Ydim = CStr(Round(Ydim, 0)+10)
Zdim = CStr(Round(Zdim, 0)+10)
EXCEL.Cells(3, 3).Value = Now
EXCEL.Cells(i + 2, 4).Value = i
EXCEL.Cells(i + 2, 5).Value = Xdim
EXCEL.Cells(i + 2, 6).Value = Ydim
EXCEL.Cells(i + 2, 7).Value = Zdim
'MsgBox "X = " & Xdim _
'& vbCr & "Y = " & Ydim _
'& vbCr & "Z = " & Zdim
Next
'MsgBox "seçilen parça sayısı: " & iCount
'EXCEL.Application.Quit
End Sub
