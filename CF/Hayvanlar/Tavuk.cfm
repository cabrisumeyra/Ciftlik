<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>

<style>
    .container{
            background-color: #f2e5e5a2;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            border-radius: 10px;
        }
            
</style>


<cfparam name="form.CiftlikID" default="">
<cfparam name="form.HayvanId" default="">
<cfparam name="form.HayvanCesidi" default="">
<cfparam name="form.HayvanTuru" default="">
<cfparam name="form.Cinsiyet" default="">
<cfparam name="form.Mahsul" default="">
<cfparam name="form.GebelikDurumu" default="">
<cfparam name="form.YavrulukDurumu" default="">
<cfparam name="form.RecordDate" default="">
<cfparam name="form.History" default="">
<cfobject component="CFC\addGenelHayvan" name="addYeniHayvan">
<cfobject component="CFC\addGenelHayvan" name="GebelikGuncelleme">
<cfobject component="CFC\addGenelHayvan" name="YavrulukGuncelleme">


<!--Sayfa başındaki resim ve başlıklar -->
<div class="container-fluid ">
    <div class="row" >
        <div name="anasayfa" onclick=" document.location='..\index.cfm' " class="col-md-2">
            <cfimage name ="isresim" 
            action="writeToBrowser" 
            source="D:\cabri\workcube\Odev_Ciftlik\img\farm.png" 
            alt="Resim" width="100">
        </div>
        <div class="col-md-7 text-center my-3">
            <h3>TAVUKLARIM</h3>
        </div>
        <div align="right" class="col-md-3">
            <cfimage name ="tavukresim" 
                action="writeToBrowser" 
                source="D:\cabri\workcube\Odev_Ciftlik\img\civcivli.png" 
                alt="Resim" width="125" height="125">
        </div>
    </div>
</div>

<div class="container-fluid">
    <div class="row ">
<!--Genel Tavuk Sayı- -->
<!--- HayvanCesidi]
      ,[HayvanTuru]
      ,[HayvanSayisi]
      ,[YavruHayvanSayisi]
      ,[OluHayvanSayisi]--->

      <cfquery name="HayvanSayi" datasource="Ciftlik">
        SELECT
        HayvanCesidi, HayvanTuru, COUNT(HayvanTuru) as ToplamHayvan
        FROM 
        HayvanDetay
        WHERE HayvanCesidi = 'Tavuk'
        GROUP BY HayvanCesidi, HayvanTuru
    </cfquery>
    <!--Tavuk Sayılarını gösteren tablo -->
    <div class="container p-3 my-3 col-md-6">
        <span>
            <table widht="99%">
                <td align="left" height="40"><h4>Tavuk Sayım Bilgileri</h4></td>
        </table>
        <table class="table table-bordered table-sm small" border="1" bordercolor="#99999" width="99%" align="center">
            <tr>
                <td><b>No</b></td>
                <td><b>Hayvan Çeşidi</b></td>
                <td><b>Hayvan Türü</b></td>
                <td><b>Toplam Hayvan Sayısı</b></td>
            </tr>
            <cfoutput query="HayvanSayi">
                <tr>
                    <td>#currentrow#</td>
                    <td>#HayvanCesidi#</td>
                    <td>#HayvanTuru#</td>  
                    <td>#ToplamHayvan#</td> 
                </tr>
            </cfoutput>
        </table>
        </span>
    </div>

    <!--- Cinsiyet tablosu --->
    <cfquery name="CinsiyetSayi" datasource="Ciftlik">
        SELECT
        HayvanTuru, COUNT(CASE WHEN Cinsiyet = 'Dişi' THEN 1 ELSE NULL END) AS CinsiyetDisi, COUNT(CASE WHEN Cinsiyet = 'Erkek' THEN 1 ELSE NULL END) AS CinsiyetErkek
        FROM 
        HayvanDetay
        WHERE HayvanCesidi = 'Tavuk'
        GROUP BY HayvanCesidi, HayvanTuru
    </cfquery>
    <!--Tavuk Sayılarını gösteren tablo -->
    <div class="container p-3 my-3 col-md-6">
        <span>
            <table widht="99%">
                <td align="left" height="40"><h4>Tavuk Cinsiyet Bilgileri</h4></td>
        </table>
        <table class="table table-bordered table-sm small" border="1" bordercolor="#99999" width="99%" align="center">
            <tr>
                <td><b>No</b></td>
                <td><b>Hayvan Türü</b></td>
                <td><b>Dişi</b></td>
                <td><b>Erkek</b></td>
            </tr>
            <cfoutput query="CinsiyetSayi">
                <tr>
                    <td>#currentrow#</td>
                    <td>#HayvanTuru#</td>  
                    <td>#CinsiyetDisi#</td> 
                    <td>#CinsiyetErkek#</td> 
                </tr>
            </cfoutput>
        </table>
        </span>
    </div>
    </div>
</div>






<!--Form ve hayvan gözüken/ hayvan eklenme tablosu -->
<div class="container-fluid" >
    <div class="row">
        
        <!--Hayvan ekleme formu -->
        <div class="container p-3 my-3 col-md-3 ">
            <table align="left" widht="99%">
                <tr>
                    <td height="40" align="left"><h4>Yeni Tavuk Ekleme</h4></td>
                </tr>
            </table>
            <cfform name="YeniHayvanEkleme" action="..\CFC\addGenelHayvan.cfc?method=addYeniHayvan" method="post" >
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                <tr>
                <td>Ciftlik Id</td>
                <td>: <cfinput type="text" name="CiftlikID" required="true"></td>
                </tr>
                <tr>
                    <td>Hayvan Cesidi</td>
                    <td>: <select name="HayvanCesidi">
                        <option value="Tavuk">Tavuk</option>
                </tr>
                <tr>
                    <td>Hayvan Turu</td>
                    <td>: <cfinput type="text" name="HayvanTuru" maxlength="100" required="true"></td>
                </tr>
                <tr>
                    <td>HayvanNo</td>
                    <td>: <cfinput type="text" name="HayvanNo" required="true"></td>
                </tr>
                <tr>
                    <td>Cinsiyet</td>
                    <td>: <select name="Cinsiyet">
                        <option value="Erkek">Erkek</option>
                        <option value="Dişi">Disi</option>
                    </td>
                </tr>
                <tr>
                    <td>Mahsul</td>
                    <td>: <cfinput type="text" name="Mahsul" maxlength="50" required="true"></td>
                </tr>
                <tr>
                    <td>Yavruluk Durumu</td>
                    <td>: <select name="YavrulukDurumu">
                        <option value="Yavru Değil">Yavru Değil</option>
                        <option value="Yavru">Yavru</option>
                    </td>
                </tr>
                <tr>
                    <td>Kayit Tarihi</td>
                    <td>: <cfinput type="date" name="RecordDate" required="true"></td>
                </tr>
                <tr>
                    <td>Güncellenme Tarihi</td>
                    <td>: <cfinput type="date" name="History" required="true"></td>
                </tr>
                <tr>
                    <td>Kayit Ediniz.</td>
                    <td><cfinput type="submit" name="save" value="Kaydet"></td>
                </tr>
                </table>
            </cfform>
        </div>

        <!--Formdaki boş yerleri doldur uyarısı -->
            <cfoutput>
                <script>
                function validateForm() {
                    var elements = document.getElementsByName("saveform");
                
                    for (var i = 0; i < elements.length; i++) {
                    var inputValue = elements[i].value.trim();
                  
                    if (inputValue === "") {
                        alert("Lütfen giriş alanını doldurun!");
                        return false; // Formu gönderme işlemi durdurulur
                    }
                    }
                    return true; // Formu gönderme işlemine devam edilir
                }
                </script>
            </cfoutput>
        
        <!--SQL sorgu alanı yalnızca tavukları çekiyor -->
            <cfquery name="HayvanDetay" datasource="Ciftlik">
                SELECT
                CiftlikID, 
                HayvanId,
                HayvanCesidi,
                HayvanTuru,
                HayvanNo,
                Cinsiyet,
                Mahsul,
                RecordDate,
                History
                FROM 
                HayvanDetay
                WHERE HayvanCesidi = 'Tavuk'
            </cfquery>
            <br><br><br>

        <!--Tavukları gösteren tablo -->
        <div class="container p-3 my-3 col-md-8">
            <span>
                <table widht="99%">
                    <td align="left" height="40"><h4>Genel Tavuk Bilgileri</h4></td>
            </table>
            <table class="table table-bordered table-sm small" border="1" bordercolor="#99999" width="99%" align="center">
                <tr>
                    <td><b>No</b></td>
                    <td><b>Ciftlik ID</b></td>
                    <td><b>Hayvan ID</b></td>
                    <td><b>Hayvan Çeşidi</b></td>
                    <td><b>Hayvan Türü</b></td>
                    <td><b>Hayvan NO</b></td>
                    <td><b>Cinsiyet</b></td>
                    <td><b>Mahsül</b></td>
                    <td><b>Kayıt Tarihi</b></td>
                    <td><b>Güncellenme Tarihi</b></td>
                </tr>
                <cfoutput query="HayvanDetay">
                    <tr>
                        <td>#currentrow#</td>
                        <td>#CiftlikID#</td> 
                        <td>#HayvanId#</td> 
                        <td>#HayvanCesidi#</td>
                        <td>#HayvanTuru#</td>  
                        <td>#HayvanNo#</td> 
                        <td>#Cinsiyet#</td> 
                        <td>#Mahsul#</td> 
                        <td>#dateFormat(RecordDate, 'dd/mm/yyyy')#</td>
                        <td>#dateFormat(History, 'dd/mm/yyyy')#</td>
                    </tr>
                </cfoutput>
            </table>
            </span>
        </div>
    </div>
</div>

<!--GEBELİK DURUMU -->
<div class="container-fluid" >
    <div class="row">
        
        <!--Hayvan gebelik güncelleme formu -->
        <div class="container p-3 my-3 col-md-3 ">
            <table align="left" widht="99%">
                <tr>
                    <td height="40" align="left"><h4>Kuluçka Dönemindeki Tavuklar</h4></td>
                </tr>
            </table>
            <cfform name="gebeHayvanGuncelleme" action="..\CFC\addGenelHayvan.cfc?method=GebelikGuncelleme" method="post" >
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                <tr>
                    <td>HayvanNo</td>
                    <td>: <cfinput type="text" name="HayvanNo" required="true"></td>
                </tr>
                <tr>
                    <td>Gebelik Durumu</td>
                    <td>: <select name="GebelikDurumu">
                        <option value="Gebe">Gebe</option>
                        <option value="Gebe Değil">Gebe Değil</option>
                    </td>
                </tr>
                <tr>
                    <td>Güncelleme Tarihi</td>
                    <td>: <cfinput type="date" name="History" required="true"></td>
                </tr>
                <tr>
                    <td>Güncelleyiniz.</td>
                    <td><cfinput type="submit" name="save" value="Kaydet"></td>
                </tr>
                </table>
            </cfform>
        </div>

        <!--SQL sorgu alanı yalnızca gebe tavukları çekiyor -->
            <cfquery name="HayvanDetay" datasource="Ciftlik">
                SELECT
                CiftlikID, 
                HayvanTuru,
                HayvanNo,
                GebelikDurumu,
                RecordDate,
                History
                FROM 
                HayvanDetay
                WHERE HayvanCesidi = 'Tavuk' AND GebelikDurumu = 'Gebe'
            </cfquery>
            <br><br><br>

        <!--GEBE Tavukları gösteren tablo -->
        <div class="container p-3 my-3 col-md-5">
            <span>
                <table widht="99%">
                    <td align="left" height="40"><h4>Kuluçka Dönemindeki Tavuk Bilgileri</h4></td>
                </table>
                <table name="gebelik" class="table table-bordered table-sm small" border="1" bordercolor="#99999" width="99%" align="center">
                <tr>
                    <td><b>No</b></td>
                    <td><b>Ciftlik ID</b></td>
                    <td><b>Hayvan Türü</b></td>
                    <td><b>Hayvan NO</b></td>
                    <td><b>GebelikDurumu</b></td>
                    <td><b>Kayıt Tarihi</b></td>
                    <td><b>Güncellenme Tarihi</b></td>
                </tr>
                <cfoutput query="HayvanDetay">
                    <tr>
                        <td>#currentrow#</td>
                        <td>#CiftlikID#</td> 
                        <td>#HayvanTuru#</td>  
                        <td>#HayvanNo#</td> 
                        <td>#GebelikDurumu#</td> 
                        <td>#dateFormat(RecordDate, 'dd/mm/yyyy')#</td>
                        <td>#dateFormat(History, 'dd/mm/yyyy')#</td>
                    </tr>
                </cfoutput>
                </table>
            </span>
        </div>
    </div>
</div>



<!--Civciv DURUMU -->
<div class="container-fluid" >
    <div class="row">
        
        <!--Hayvan yetişkinlik güncelleme formu -->
        <div class="container p-3 my-3 col-md-3 ">
            <table align="left" widht="99%">
                <tr>
                    <td height="40" align="left"><h4>Yetişkin Civcivler Güncelleme</h4></td>
                </tr>
            </table>
            <cfform name="civcivHayvanGuncelleme" action="..\CFC\addGenelHayvan.cfc?method=YavrulukGuncelleme" method="post" >
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                <tr>
                    <td>HayvanNo</td>
                    <td>: <cfinput type="text" name="HayvanNo" required="true"></td>
                </tr>
                <tr>
                    <td>Yetişkinlik Durumu</td>
                    <td>: <select name="YavrulukDurumu">
                        <option value="Yavru Değil">Yavru Değil</option>
                        <option value="Yavru">Yavru</option>
                    </td>
                </tr>
                <tr>
                    <td>Güncelleme Tarihi</td>
                    <td>: <cfinput type="date" name="History" required="true"></td>
                </tr>
                <tr>
                    <td>Güncelleyiniz.</td>
                    <td><cfinput type="submit" name="save" value="Kaydet"></td>
                </tr>
                </table>
            </cfform>
        </div>

    <!--SQL sorgu alanı yalnızca civcivleri çekiyor -->
    <cfquery name="HayvanDetay" datasource="Ciftlik">
        SELECT
        CiftlikID, 
        HayvanTuru,
        HayvanNo,
        YavrulukDurumu,
        RecordDate,
        History
        FROM 
        HayvanDetay
        WHERE HayvanCesidi = 'Tavuk' AND YavrulukDurumu = 'Yavru'
    </cfquery>
    <br><br><br>

    <!--Yavru Tavukları gösteren tablo -->
    <div class="container p-3 my-3 col-md-5">
        <span>
            <table widht="99%">
                <td align="left" height="40"><h4>Civciv Bilgileri</h4></td>
            </table>
            <table name="gebelik" class="table table-bordered table-sm small" border="1" bordercolor="#99999" width="99%" align="center">
                <tr>
                    <td><b>No</b></td>
                    <td><b>Ciftlik ID</b></td>
                    <td><b>Hayvan Türü</b></td>
                    <td><b>Hayvan NO</b></td>
                    <td><b>YavrulukDurumu</b></td>
                    <td><b>Kayıt Tarihi</b></td>
                    <td><b>Güncellenme Tarihi</b></td>
                </tr>
                <cfoutput query="HayvanDetay">
                    <tr>
                        <td>#currentrow#</td>
                        <td>#CiftlikID#</td> 
                        <td>#HayvanTuru#</td>  
                        <td>#HayvanNo#</td> 
                        <td>#YavrulukDurumu#</td> 
                        <td>#dateFormat(RecordDate, 'dd/mm/yyyy')#</td>
                        <td>#dateFormat(History, 'dd/mm/yyyy')#</td>
                    </tr>
                </cfoutput>
            </table>
        </span>
    </div>
</div>


















