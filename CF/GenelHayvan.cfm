    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>

    <style>
        .container{
                background-color: #f2e5e5a2;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                border-radius: 10px;
            }
                
    </style>
<div name="anasayfa" onclick=" document.location='index.cfm' ">
    <cfimage name ="isresim" 
    action="writeToBrowser" 
    source="..\img\farm.png" 
    alt="Resim" width="100">
</div>


<cfparam name="form.CiftlikID" default="">
<cfparam name="form.HayvanCesidi" default="">
<cfparam name="form.HayvanTuru" default="">
<cfparam name="form.BeslenmeDurumu" default="">
<cfparam name="form.RecordDate" default="">
<cfobject component="CFC\addGenelHayvan" name="BeslenmeGuncelleme">


<div class="row ">
    <div class="text-center"><h3>Hayvanlarım</h3></div>
    <div class="p-3 my-3 col-md-8 h-100">
        <cfoutput>
            <div name="genel" class="d-flex justify-content-center">
                                
                <div name="Tavuk " class=" p-1 m-5  col-sm-2  text-black" onclick=" document.location='./Hayvanlar/Tavuk.cfm' ">
                    <cfimage name ="tavukresim" 
                        action="writeToBrowser" 
                        source="..\img\tavuk_ikon.png" 
                        alt="Resim" width="100">
                    <h3>Tavuklarım</h3>
                </div>
        
                <div name="İnek " class=" p-1 m-5 col-sm-2 text-black" onclick="document.location='./Hayvanlar/Inek.cfm'">
                    <cfimage name ="inekresim" 
                        action="writeToBrowser" 
                        source="..\img\inek_ayna.png" 
                        alt="Resim" width="100">
                    <h3>İneklerim</h3>
                </div>
        
                <div name="Köpek " class=" p-1 m-5 col-sm-2 text-black" onclick=" document.location='./Hayvanlar/Kopek.cfm'">
                    <cfimage name ="kopekresim" 
                        action="writeToBrowser" 
                        source="..\img\kopek_ikon.png" 
                        alt="Resim" width="100">
                    <h3>Köpeklerim</h3>
                </div>
            </div>
        </cfoutput>
    </div>
    <!--Genel  Sayı- -->
        
    <cfquery name="HayvanSayi" datasource="Ciftlik">
        SELECT
        HayvanCesidi, COUNT(HayvanCesidi) as ToplamHayvan, SUM(CASE WHEN YavrulukDurumu = 'Yavru' THEN 1 ELSE 0 END) as YavruHayvanSayisi
        FROM 
        HayvanDetay
        GROUP By HayvanCesidi
    </cfquery>
    <!--Hayvan Sayılarını gösteren tablo -->
    <div class="container p-3 my-3 col-md-3 h-100">
        <span>
            <table widht="99%">
                <td align="left" height="40"><h4>Hayvan Sayım Bilgileri</h4></td>
        </table>
        <table class="table table-bordered table-sm small" border="1" bordercolor="#99999" width="99%" align="center">
            <tr>
                <td><b>No</b></td>
                <td><b>Hayvan Çeşidi</b></td>
                <td><b>Toplam Hayvan Sayısı</b></td>
                <td><b>Yavru Hayvan Sayısı</b></td>
            </tr>
            <cfoutput query="HayvanSayi">
                <tr>
                    <td>#currentrow#</td>
                    <td>#HayvanCesidi#</td>
                    <td>#ToplamHayvan#</td>
                    <td>#YavruHayvanSayisi#</td> 
                </tr>
            </cfoutput>
        </table>
        </span>
    </div>
</div>

        
<div class="container-fluid ">
    <div class="row">
        <!--Hayvan besleme formu -->
        <div class="container p-3 my-3 col-md-3">
            <table align="left" widht="99%">
                <tr>
                    <td height="40" align="left"><h4>Beslenme Durumu Güncelleme</h4></td>
                </tr>
            </table>
            <cfform name="BeslenmeGuncelleme" action="CFC\addGenelHayvan.cfc?method=BeslenmeGuncelleme" method="post" >
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                    <tr>
                        <td>Ciftlik Id</td>
                        <td>: <cfinput type="text" name="CiftlikID" required="true"></td>
                    </tr>
                    <tr>
                        <td>Hayvan Cesidi</td>
                        <td>: <select name="HayvanCesidi">
                            <option value="Tavuk">Tavuk</option>
                            <option value="İnek">İnek</option>
                            <option value="Kopek">Köpek</option>
                    </tr>
                    <tr>
                        <td>Hayvan Turu</td>
                        <td>: <cfinput type="text" name="HayvanTuru" maxlength="100" required="true"></td>
                    </tr>
                    <tr>
                        <td>Beslenme Durumu</td>
                        <td>: <select name="BeslenmeDurumu">
                            <option value="Beslenmedi">Beslenmedi</option>
                            <option value="Beslendi">Beslendi</option>
                    </tr>
                    <tr>
                        <td>Güncellenme Tarihi</td>
                        <td>: <cfinput type="date" name="RecordDate" required="true"></td>
                    </tr>
                    <tr>
                        <td>Kayit Ediniz.</td>
                        <td><cfinput type="submit" name="save" value="Kaydet"></td>
                    </tr>
                </table>
            </cfform>
        </div>

        <!--Beslenme Durumu- -->
        <cfquery name="HayvanBeslenme" datasource="Ciftlik">
            SELECT
            CiftlikID, HayvanCesidi, HayvanTuru, BeslenmeDurumu, RecordDate
            FROM 
            GenelHayvan
        </cfquery>

        <!--Hayvan Sayılarını gösteren tablo -->
        <div class="container p-3 my-3 col-md-5">
            <span>
                <table widht="99%">
                    <td align="left" height="40"><h4>Hayvan Beslenme Durumları</h4></td>
                </table>
                <table class="table table-bordered table-sm small" border="1" bordercolor="#99999" width="99%" align="center">
                    <tr>
                        <td><b>No</b></td>
                        <td><b>CiftlikID</b></td>
                        <td><b>Hayvan Çeşidi</b></td>
                        <td><b>Hayvan Turu</b></td>
                        <td><b>Beslenme Durumu</b></td>
                        <td><b>Kayıt Tarihi</b></td>
                    </tr>
                    <cfoutput query="HayvanBeslenme">
                    <tr>
                        <td>#currentrow#</td>
                        <td>#CiftlikID#</td>
                        <td>#HayvanCesidi#</td>
                        <td>#HayvanTuru#</td>
                        <td>#BeslenmeDurumu#</td>
                        <td>#RecordDate#</td> 
                    </tr>
                    </cfoutput>
                </table>
            </span>
        </div>
    </div>
</div>
