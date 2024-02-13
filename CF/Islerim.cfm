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

<div class="container-fluid ">
    <div class="row">
        <!--Hayvan besleme formu -->
        <div class="container p-3 my-3 col-md-3">
            <table align="left" widht="99%">
                <tr>
                    <td height="40" align="left"><h4>Beslenme Durumu Güncelleme</h4></td>
                </tr>
            </table>
            <cfform name="BeslenmeGuncelleme" action="add\updateGenelHayvan.cfm" method="post" >
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
<cfset today = CreateDate(Year(Now()), Month(Now()), Day(Now()))>
<div name="HayvanBilgileri" class="container-fluid">
    <div class="row">
            <cfquery name="VeterinerRandevusu" datasource="Ciftlik">
                SELECT CiftlikID, FinansID, HayvanCesidi, VeterinerRandevu, Aciklama, RecordDate
                FROM Finans
                WHERE VeterinerRandevu = <cfqueryparam value="#today#" cfsqltype="CF_SQL_DATE">
            </cfquery>
    
            <div name="HarcananEkleme" class="container p-3 m-3 col-md-7" >
				<table><h6>Veteriner Takip Bilgileri</h6></table>
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                    <tr>
                        <td>Çiftlik ID</td>
                        <td>Finans ID</td>
                        <td>Hayvan Çesidi</td>
                        <td>Veteriner Randevusu</td>
                        <td>Açıklama</td>
                        <td>Kayıt Tarihi</td>
                    </tr>
                    <cfoutput query="VeterinerRandevusu">
                        <tr>
                            <td>#CiftlikID#</td>
                            <td>#FinansID#</td>
                            <td>#HayvanCesidi#</td>
                            <td>#VeterinerRandevu#</td>
                            <td>#Aciklama#</td>
                            <td>#RecordDate#</td>
                        </tr>
                    </cfoutput>
                </table>
            </div>
    </div>
</div>
