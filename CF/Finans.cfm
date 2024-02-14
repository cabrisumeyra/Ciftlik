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
<cfparam name="form.YemAlinan" default="">
<cfparam name="form.YemHarcanan" default="">
<cfparam name="form.IlacAlinan" default="">
<cfparam name="form.VeterinerMasraf" default="">
<cfparam name="form.Aciklama" default="">
<cfparam name="form.MahsulCesidi" default="">
<cfparam name="form.SatisAdetUcreti" default="">
<cfparam name="form.GunlukSatisUcreti" default="">
<cfparam name="form.GunlukSatisMiktari" default="">
<cfparam name="form.VeterinerRandevu" default="">

<cfobject component="CFC\addFinans" name="FinansEkleme">
<cfobject component="CFC\addFinans" name="FinansGuncelle">



<div>
    <h3>Finansal Bilgiler</h3>
</div>
<div name="finansBilgileri" class="container-fluid">
    <div class="row">
        <div class="container p-3 m-3 col-md-3">
            <cfform name="HarcamaGuncelleme" action="CFC\addFinans.cfc?method=FinansEkleme" method="post">
                <table>
                    <tr><td><h6>Finansal Bilgilerinizi Güncelleyiniz.</h6></td></tr>
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
                        <td>Yem Alınan</td>
                        <td>:<cfinput name="YemAlinan" type="text" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Yem Harcanan</td>
                        <td>:<cfinput name="YemHarcanan" type="text" maxlength="10" value="0"></td>
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
                        <td>Mahsul Çeşidi</td>
                        <td>:<select name="MahsulCesidi">
                            <option value="NULL">Boş</option>
                            <option value="Piliç">Piliç</option>
                            <option value="Yumurta">Yumurta</option>
                            <option value="Süt">Süt</option>
                        </td>
                    </tr>
                    <tr>
                        <td>Satış adet ücreti</td>
                        <td>:<cfinput type="text" name="SatisAdetUcreti" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Günlük Satış Miktarı</td>
                        <td>:<cfinput type="text" name="GunlukSatisMiktari" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Günlük Satış Ücreti</td>
                        <td>:<cfinput type="text" name="GunlukSatisUcreti" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Açıklama</td>
                        <td>:<cfinput type="text" name="Aciklama" maxlength="1000" value="0"></td>
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
                SELECT CiftlikID, FinansID, HayvanCesidi, YemAlinan, IlacAlinan, VeterinerMasraf, Aciklama, RecordDate
                FROM Finans
            </cfquery>
    
            <div name="HarcananEkleme" class="container p-3 m-3 col-md-7">
                <table><h6>Harcama Bilgileri</h6></table>
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                    <tr>
                        <td>Çiftlik ID</td>
                        <td>Finans ID</td>
                        <td>Hayvan Çesidi</td>
                        <td>YemAlinan</td>
                        <td>Ilaç Alınan</td>
                        <td>Veteriner Masrafı</td>
                        <td>Açıklama</td>
                        <td>Kayıt Tarihi</td>
                    </tr>
                    <cfoutput query="HarcananEkleme">
                        <tr>
                            <td>#CiftlikID#</td>
                            <td>#FinansID#</td>
                            <td>#HayvanCesidi#</td>
                            <td>#YemAlinan#</td>
                            <td>#IlacAlinan#</td>
                            <td>#VeterinerMasraf#</td>
                            <td>#Aciklama#</td>
                            <td>#RecordDate#</td>
                        </tr>
                    </cfoutput>
                </table>
            </div>
    </div>
</div>

            <cfquery name="YemGuncelleme" datasource="Ciftlik">
                SELECT CiftlikID, FinansID, HayvanCesidi, YemAlinan, YemHarcanan, (YemAlinan - YemHarcanan) AS KalanYem, Aciklama, RecordDate
                FROM Finans
            
            </cfquery>

<div class="container-fluid">
    <div class="row">
        <div name="YemEkleme" class="container p-3 m-3 col-md-5">
                <table><h6>Yem Bilgileri</h6></table>
                <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                    <tr>
                        <td>Çiftlik ID</td>
                        <td>Finans ID</td>
                        <td>Hayvan Çesidi</td>
                        <td>Yem Alınan</td>
                        <td>Yem Harcanan</td>
                        <td>Kalan Yem Miktarı</td>
                        <td>Açıklama</td>
                        <td>Kayıt Tarihi</td>
                    </tr>
                    <cfoutput query="YemGuncelleme">
                        <tr>
                            <td>#CiftlikID#</td>
                            <td>#FinansID#</td>
                            <td>#HayvanCesidi#</td>
                            <td>#YemAlinan#</td>
                            <td>#YemHarcanan#</td>
                            <td>#KalanYem#</td>
                            <td>#Aciklama#</td>
                            <td>#RecordDate#</td>
                        </tr>
                    </cfoutput>
                </table>
        </div>

        <cfquery name="MahsulGuncelleme" datasource="Ciftlik">
            SELECT CiftlikID, FinansID, HayvanCesidi, MahsulCesidi, SatisAdetUcreti, GunlukSatisMiktari, GunlukSatisUcreti, Aciklama, RecordDate
            FROM Finans
            WHERE MahsulCesidi IS NOT NULL
        </cfquery>


        <div name="MahsulEkleme" class="container p-3 m-3 col-md-6">
            <table><h6>Mahsul Satış Bilgileri</h6></table>
            <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                <tr>
                    <td>Çiftlik ID</td>
                    <td>Finans ID</td>
                    <td>Hayvan Çeşidi</td>
                    <td>Mahsul Çeşidi</td>
                    <td>Satış adet ücreti</td>
                    <td>Günlük Satış Miktarı</td>
                    <td>Günlük Satış Ücreti</td>
                    <td>Açıklama</td>
                    <td>Kayıt Tarihi</td>
                </tr>
                <cfoutput query="MahsulGuncelleme">
                    <tr>
                        <td>#CiftlikID#</td>
                        <td>#FinansID#</td>
                        <td>#HayvanCesidi#</td>
                        <td>#MahsulCesidi#</td>
                        <td>#SatisAdetUcreti#</td>
                        <td>#GunlukSatisMiktari#</td>
                        <td>#GunlukSatisUcreti#</td>
                        <td>#Aciklama#</td>
                        <td>#RecordDate#</td>
                    </tr>
                </cfoutput>
            </table>
    </div>
    </div>
</div>
















