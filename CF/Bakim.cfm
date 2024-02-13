<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>

    <style>
        .container{
                background-color: #e8f25fa2;
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
    <cfparam name="form.IlacAlinan" default="">
    <cfparam name="form.VeterinerMasraf" default="">
    <cfparam name="form.Aciklama" default="">
    <cfparam name="form.VeterinerRandevu" default="">
    <cfparam name="form.HayvanID" default="">
    <cfparam name="form.SaglikDurumu" default="">
    <cfparam name="form.HastalikAdi" default="0">
    <cfparam name="form.AsiTarihi" default="">
    <cfparam name="form.TedaviTarihi" default="">
    <cfparam name="form.History" default="">
    
    <cfif structKeyExists(form, "CiftlikID")>
        <cfobject component="CFC\addBakim" name="Bakimekleme">
        <cfobject component="CFC\addBakim" name="BakimGuncelle">
    </cfif>


<div>
    <h3>Bakım ve Tedavi</h3>
</div> 


<div name="HayvanBilgileri" class="container-fluid">
    <div class="row">
        <div class="container p-3 m-3 col-md-3">
            <cfform name="HarcamaGuncelleme" action="CFC\addBakim.cfc?method=BakimGuncelle" method="post">
                <table>
                    <tr><td><h6>Hasta Hayvan Bilgilerini Ekleyiniz</h6></td></tr>
                </table>
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                    <tr>
                        <td>Çiftlik ID</td>
                        <td>:<select name="CiftlikID">
                            <option value="1">1</option>
                            <option value="2">2</option>
                        </td>
                    </tr>
					<tr>
                        <td>Hayvan ID</td>
                        <td>:<cfinput name="HayvanID" type="text" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>HayvanCesidi</td>
                        <td>:<select name="HayvanCesidi">
                            <option value="Tavuk">Tavuk</option>
                            <option value="Inek">Inek</option>
                            <option value="Kopek">Köpek</option>
                        </td>
                    </tr>
					<tr>
                        <td>Sağlık Durumu</td>
                        <td>:<select name="SaglikDurumu">
                            <option value="Sağlıklı">Sağlıklı</option>
                            <option value="Sağlıksız">Sağlıksız</option>
                        </td>
                    </tr>
					<tr>
                        <td>Hastalık Adı</td>
                        <td>:<cfinput name="HastalikAdi" type="text" maxlength="10" value="0"></td>
                    </tr>
					<tr>
                        <td>Aşı Tarihi</td>
                        <td>:<cfinput type="date" name="AsiTarihi"></td>
                    </tr>
					<tr>
                        <td>Tedavi Tarihi</td>
                        <td>:<cfinput type="date" name="TedaviTarihi"></td>
                    </tr>
					<tr>
                        <td>Açıklama</td>
                        <td>:<cfinput name="Aciklama" type="text" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Güncelleme Tarihi</td>
                        <td>:<cfinput type="date" name="History"></td>
                    </tr>
                    <tr>
                        <td>Kayıt ediniz.</td>
                        <td>:<cfinput name="save" type="submit" value="kaydet"></td>
                    </tr>
                </table>
            </cfform>
        </div>

            <cfquery name="HayvanEkleme" datasource="Ciftlik">
                SELECT CiftlikID, HayvanID, HayvanCesidi, SaglikDurumu, HastalikAdi, AsiTarihi, TedaviTarihi, Aciklama, History
                FROM HayvanDetay
            </cfquery>

            <div name="HarcananEkleme" class="container p-3 m-3 col-md-7">
                <details><summary><table><h6>Hasta Hayvan Bilgileri</h6></table></summary>
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                    <tr>
                        <td>Çiftlik ID</td>
                        <td>Hayvan ID</td>
                        <td>Hayvan Çesidi</td>
                        <td>Sağlık Durumu</td>
                        <td>Hastalık Adı</td>
                        <td>Aşı Tarihi</td>
						<td>Tedavi Tarihi</td>
                        <td>Açıklama</td>
                        <td>Güncelleme Tarihi</td>
                    </tr>
                    <cfoutput query="HayvanEkleme">
                        <tr>
                            <td>#CiftlikID#</td>
                            <td>#HayvanID#</td>
                            <td>#HayvanCesidi#</td>
                            <td>#SaglikDurumu#</td>
                            <td>#HastalikAdi#</td>
                            <td>#AsiTarihi#</td>
							<td>#TedaviTarihi#</td>
                            <td>#Aciklama#</td>
                            <td>#History#</td>
                        </tr>
                    </cfoutput>
                </table></details>
            </div>
    </div>
</div>





<div name="HayvanBilgileri" class="container-fluid">
    <div class="row">
        <div class="container p-3 m-3 col-md-3">
            <cfform name="HarcamaGuncelleme" action="CFC\addBakim.cfc?method=Bakimekleme" method="post">
                <table>
                    <tr><td><h6>Veteriner Takip Bilgileri</h6></td></tr>
                </table>
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                    <tr>
                        <td>Çiftlik ID</td>
                        <td>:<select name="CiftlikID">
                            <option value="1">1</option>
                            <option value="2">2</option>
                        </td>
                    </tr>
                    <tr>
                        <td>HayvanCesidi</td>
                        <td>:<select name="HayvanCesidi">
                            <option value="Tavuk">Tavuk</option>
                            <option value="Inek">Inek</option>
                            <option value="Kopek">Köpek</option>
                        </td>
                    </tr>
					<tr>
                        <td>Ilaç Alınan</td>
                        <td>:<cfinput type="text" name="IlacAlinan" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Veteriner Masrafı</td>
                        <td>:<cfinput type="text" name="VeterinerMasraf" maxlength="10" value="0"></td>
                    </tr>
					<tr>
                        <td>Veteriner Randevu</td>
                        <td>:<cfinput type="date" name="VeterinerRandevu" maxlength="10" value="0"></td>
                    </tr>
					<tr>
                        <td>Açıklama</td>
                        <td>:<cfinput name="Aciklama" type="text" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Kayıt Tarihi</td>
                        <td>:<cfinput type="date" name="RecordDate"></td>
                    </tr>
                    <tr>
                        <td>Kayıt ediniz.</td>
                        <td>:<cfinput name="save" type="submit" value="kaydet"></td>
                    </tr>
                </table>
            </cfform>
        </div>

            <cfquery name="HarcananEkleme" datasource="Ciftlik">
                SELECT CiftlikID, FinansID, HayvanCesidi, IlacAlinan, VeterinerMasraf, VeterinerRandevu, Aciklama, RecordDate
                FROM Finans
            </cfquery>
    
            <div name="HarcananEkleme" class="container p-3 m-3 col-md-7">
				<details><summary><table><h6>Veteriner Takip Bilgileri</h6></table></summary>
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                    <tr>
                        <td>Çiftlik ID</td>
                        <td>Finans ID</td>
                        <td>Hayvan Çesidi</td>
                        <td>Ilaç Alınan</td>
                        <td>Veteriner Masrafı</td>
                        <td>Veteriner Randevu</td>
                        <td>Açıklama</td>
                        <td>Kayıt Tarihi</td>
                    </tr>
                    <cfoutput query="HarcananEkleme">
                        <tr>
                            <td>#CiftlikID#</td>
                            <td>#FinansID#</td>
                            <td>#HayvanCesidi#</td>
                            <td>#IlacAlinan#</td>
                            <td>#VeterinerMasraf#</td>
                            <td>#VeterinerRandevu#</td>
                            <td>#Aciklama#</td>
                            <td>#RecordDate#</td>
                        </tr>
                    </cfoutput>
                </table></details>
            </div>
    </div>
</div>