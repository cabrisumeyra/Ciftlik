<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>

    <style>
        .container{
                background-color: #81e96fa2;
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
<cfparam name="form.CiftlikAd" default="">
<cfparam name="form.KapaliAlan" default="">
<cfparam name="form.AcikAlan" default="">
<cfparam name="form.SahiplikDurumu" default="">
<cfparam name="form.HasatTarih" default="">
<cfparam name="form.SiloKapasite" default="">
<cfparam name="form.CalisanSayisi" default="">
<cfparam name="form.BinaSigortaDurumu" default="">
<cfparam name="form.RecordDate" default="">

<cfif structKeyExists(form, "CiftlikID")>
    <cfobject component="CFC\addCiftlik" name="CiftlikEkleme">
</cfif>


<div class="p-1 m-1">
    <h3>Çiftlik Bilgileri</h3>
</div>
 

<div name="ciftlikBilgileri" class="container-fluid">
    <div class="row">
        <div class="container p-3 m-3 col-md-3">
            <cfform name="ciftlikBilgileri" action="CFC\addCiftlik.cfc?method=CiftlikEkleme" method="post">
                <table>
                    <tr><td><h6> Yeni Çiftlik Bilgilerinizi Güncelleyiniz.</h6></td></tr>
                </table>
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                    <tr>
                        <td>Çiftlik ID</td>
                        <td>:<cfinput name="CiftlikID" type="text" maxlength="10">
                        </td>
                    </tr>
                    <tr>
                        <td>Çiftlik Ad</td>
                        <td>:<cfinput name="CiftlikAd" type="text" maxlength="10"></td>
                    </tr>
                    <tr>
                        <td>Kapalı Alan</td>
                        <td>:<cfinput type="text" name="KapaliAlan" maxlength="10" default="0"></td>
                    </tr>
                    <tr>
                        <td>Açık Alan</td>
                        <td>:<cfinput type="text" name="AcikAlan" maxlength="10" default="0"></td>
                    </tr>
                    <tr>
                        <td>Sahiplik Durumu</td>
                        <td>:<cfinput type="text" name="SahiplikDurumu" maxlength="10" default="0"></td>
                    </tr>
                    <tr>
                        <td>Hasat Tarih</td>
                        <td>:<cfinput type="date" name="HasatTarih" maxlength="10" default="0"></td>
                    </tr>
                    <tr>
                        <td>Silo Kapasite</td>
                        <td>:<cfinput type="text" name="SiloKapasite" maxlength="10" default="0"></td>
                    </tr>
                    <tr>
                        <td>Çalışan Sayısı</td>
                        <td>:<cfinput type="text" name="CalisanSayisi" maxlength="10" default="0"></td>
                    </tr>
                    <tr>
                        <td>Bina Sigorta Durumu</td>
                        <td>:<cfinput type="text" name="BinaSigortaDurumu" maxlength="10" default="0"></td>
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
        
        <cfquery name="ciftliklerim" datasource="Ciftlik">
            SELECT CiftlikID, CiftlikAd, KapaliAlan, AcikAlan, SahiplikDurumu, HasatTarih, SiloKapasite, CalisanSayisi, BinaSigortaDurumu, RecordDate
            FROM CiftlikBilgileri
        </cfquery>

        <div name="uretimtablosu" class="container p-3 m-3 col-md-8">
            <table><h6>Çiftliklerim</h6></table>
            <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                <tr>
                    <td>Çiftlik ID</td>
                    <td>Çiftlik Ad</td>
                    <td>Kapalı Alan</td>
                    <td>Açık Alan</td>
                    <td>Sahiplik Durumu</td>
                    <td>Hasat Tarih</td>
                    <td>Silo Kapasite</td>
                    <td>Çalışan Sayısı</td>
                    <td>Bina Sigorta Durumu</td>
                    <td>Kayıt Tarihi</td>
                </tr>
                <cfoutput query="ciftliklerim">
                    <tr>
                        <td>#CiftlikID#</td>
                        <td>#CiftlikAd#</td>
                        <td>#KapaliAlan#</td>
                        <td>#AcikAlan#</td>
                        <td>#SahiplikDurumu#</td>
                        <td>#HasatTarih#</td>
                        <td>#SiloKapasite#</td>
                        <td>#CalisanSayisi#</td>
                        <td>#BinaSigortaDurumu#</td>
                        <td>#RecordDate#</td>
                    </tr>
                </cfoutput>
            </table>
        </div>
    </div>
</div>