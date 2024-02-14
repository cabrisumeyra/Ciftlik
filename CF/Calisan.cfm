<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
<style>
        .container{
                background-color: #7e8ff1a2;
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
<cfparam name="form.CalisanIsim" default="">
<cfparam name="form.CalisanSoyad" default="">
<cfparam name="form.CalisanYas" default="">
<cfparam name="form.Maas" default="">
<cfparam name="form.GirisTarihi" default="">
<cfparam name="form.CikisTarihi" default="">

<cfif structKeyExists(form, "CalisanIsim")>
    <cfobject component="CFC\addCalisan" name="CalisanEkle">
</cfif>


<div class="p-1 m-1"><h4>Çalışanlarım</h4></div>
<div class="container-fluid">
    <div class="row">
        <div name="CalisanEkle" class="container p-2 m-2 col-md-2">
            <cfform name="CalisanEkle" action="CFC\addCalisan.cfc?method=CalisanEkle" method="post">
                <table>
                    <tr><td><h6>Çalışan Ekleme</h6></td></tr>
                </table>
                <table><table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                    <tr>
                        <td>Çiftlik ID</td>
                        <td>:<select name="CiftlikID">
                            <option value="1">1</option>
                            <option value="2">2</option>
                        </td>
                    </tr>
                    <tr>
                        <td>Çalışan Isim</td>
                        <td>:<cfinput type="text" name="CalisanIsim" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Çalışan Soyisim</td>
                        <td>:<cfinput type="text" name="CalisanSoyad" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Çalışan Yaş</td>
                        <td>:<cfinput type="text" name="CalisanYas" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Maas</td>
                        <td>:<cfinput type="text" name="Maas" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Giriş Tarihi</td>
                        <td>:<cfinput type="date" name="GirisTarihi" ></td>
                    </tr>
                    <tr>
                        <td>Kayıt ediniz.</td>
                        <td>:<cfinput name="save" type="submit" value="kaydet"></td>
                    </tr>
                </table>
            </cfform>
        </div>

        <div name="CikisEkleme" class="container p-2 m-2 col-md-2">
            <cfform name="CalisanEkle" action="CFC\addCalisan.cfc?method=CikisEkle" method="post">
                <table>
                    <tr><td><h6>Çalışan Çıkışı Ekleme</h6></td></tr>
                </table>
                <table><table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                    <tr>
                        <td>Çiftlik ID</td>
                        <td>:<select name="CiftlikID">
                            <option value="1">1</option>
                            <option value="2">2</option>
                        </td>
                    </tr>
                    <tr>
                        <td>Çalışan Isim</td>
                        <td>:<cfinput type="text" name="CalisanIsim" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Çalışan Soyisim</td>
                        <td>:<cfinput type="text" name="CalisanSoyad" maxlength="10" value="0"></td>
                    </tr>
                    <tr>
                        <td>Çıkış Tarihi</td>
                        <td>:<cfinput type="date" name="CikisTarihi" ></td>
                    </tr>
                    <tr>
                        <td>Kayıt ediniz.</td>
                        <td>:<cfinput name="save" type="submit" value="kaydet"></td>
                    </tr>
                </table>
            </cfform>
        </div>

        <cfquery name="TumCalisanlar" datasource="Ciftlik">
            SELECT CiftlikID, CalisanID, CalisanIsim, CalisanSoyad, CalisanYas, Maas, GirisTarihi, CikisTarihi
            FROM Calisan
        </cfquery>

        <div name="calisangoster" class="container p-3 m-3 col-md-7">
            <table><h6>Çalışan Bilgileri</h6></table>
            <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
                <tr>
                    <td>Çiftlik ID</td>
                    <td>Çalışan ID</td>
                    <td>Çalışan İsim</td>
                    <td>Çalışan Soyisim</td>
                    <td>Çalışanın Yaşı</td>
                    <td>Maaş</td>
                    <td>Giriş Tarihi</td>
                    <td>Çıkış Tarihi</td>
                </tr>
                <cfoutput query="TumCalisanlar">
                    <tr>
                        <td>#CiftlikID#</td>
                        <td>#CalisanID#</td>
                        <td>#CalisanIsim#</td>
                        <td>#CalisanSoyad#</td>
                        <td>#CalisanYas#</td>
                        <td>#Maas#</td>
                        <td>#GirisTarihi#</td>
                        <td>#CikisTarihi#</td>
                    </tr>
                </cfoutput>
            </table>
        </div>

    </div>
</div>


