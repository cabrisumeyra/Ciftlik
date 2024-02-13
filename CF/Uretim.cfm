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
<cfparam name="form.MahsulCesidi" default="">
<cfparam name="form.ToplanmaDurumu" default="">
<cfparam name="form.GunlukUretimAdeti" default="">
<cfparam name="form.HasarliMahsulAdedi" default="">

<cfif structKeyExists(form, "CiftlikID")>
    <cfobject component="CFC\addUretim" name="UretimGuncelleme">
</cfif>

<div class="p-1 m-1"> 
    <h3>Üretim Bilgileri</h3>
</div>
<div name="uretimBilgileri" class="container-fluid">
    <div class="row">
        <div class="container p-3 m-3 col-md-3">
            <cfform name="UretimGuncelleme" action="CFC\addUretim.cfc?method=UretimGuncelleme" method="post">
                <table>
                    <tr><td><h6>Üretim Bilgilerinizi Güncelleyiniz.</h6></td></tr>
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
                        <td>Mahsul Çeşidi</td>
                        <td>:<select name="MahsulCesidi">
                            <option value="Süt">Süt</option>
                            <option value="Yumurta">Yumurta</option>
                            <option value="Piliç">Piliç</option>
                        </td>
                    </tr>
                    <tr>
                        <td>Toplanma Durumu</td>
                        <td>:<select name="ToplanmaDurumu">
                            <option value="Toplandı">Toplandı</option>
                            <option value="Toplanmadı">Toplanmadı</option>
                        </td>
                    </tr>
                    <tr>
                        <td>Günlük Üretim Adeti</td>
                        <td>:<cfinput name="GunlukUretimAdeti" type="text" maxlength="10"></td>
                    </tr>
                    <tr>
                        <td>Hasarlı Mahsul Adedi</td>
                        <td>:<cfinput type="text" name="HasarliMahsulAdedi" maxlength="10" default="0"></td>
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
        
        <cfquery name="uretimGuncel" datasource="Ciftlik">
            SELECT CiftlikID, UretimID, MahsulCesidi, ToplanmaDurumu, GunlukUretimAdeti, HasarliMahsulAdedi, RecordDate
            FROM URETİM
        </cfquery>

        <div name="uretimtablosu" class="container p-3 m-3 col-md-7">
            <table><h6>Üretim Bilgileri</h6></table>
            <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                <tr>
                    <td>Çiftlik ID</td>
                    <td>Üretim ID</td>
                    <td>Mahsul Çesidi</td>
                    <td>Toplanma Durumu</td>
                    <td>Günlük Üretim Adeti</td>
                    <td>Hasarlı Mahsul Adedi</td>
                    <td>Kayıt Tarihi</td>
                </tr>
                <cfoutput query="uretimGuncel">
                    <tr>
                        <td>#CiftlikID#</td>
                        <td>#UretimID#</td>
                        <td>#MahsulCesidi#</td>
                        <td>#ToplanmaDurumu#</td>
                        <td>#GunlukUretimAdeti#</td>
                        <td>#HasarliMahsulAdedi#</td>
                        <td>#RecordDate#</td>
                    </tr>
                </cfoutput>
            </table>
        </div>
    </div>
</div>
