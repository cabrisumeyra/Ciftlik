<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>

    <style>
        .container{
                background-color: #c8e39da2;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                border-radius: 10px;
            }
        .genis-form input[type="text"] {
                width: 90%;
            }
                
    </style>
    <div name="anasayfa" onclick=" document.location='index.cfm' ">
        <cfimage name ="isresim" 
        action="writeToBrowser" 
        source="..\img\farm.png" 
        alt="Resim" width="100">
    </div>
    
<div name="arama" class="container p-4 m-4 ">
    <cfform method="get" class="genis-form">
        <input type="text" name="arama_kelimesi" placeholder="Arama kelimesi girin">
        <input type="submit" value="Ara">
    </cfform>
</div>

<cfparam name="url.arama_kelimesi" default="">

<cfif NOT StructKeyExists(URL, "arama_kelimesi") OR Trim(URL.arama_kelimesi) eq "">
    <cfquery name="Hayvanlarim" datasource="Ciftlik">
        SELECT CiftlikID, HayvanID, HayvanCesidi, HayvanTuru, Cinsiyet, Mahsul, GebelikDurumu, YavrulukDurumu, SaglikDurumu, HastalikAdi, AsiTarihi, TedaviTarihi, Aciklama, History
        FROM HayvanDetay
    </cfquery>
<cfelse>
    <cfquery name="Hayvanlarim" datasource="Ciftlik">
        SELECT CiftlikID, HayvanID, HayvanCesidi, HayvanTuru, Cinsiyet, Mahsul, GebelikDurumu, YavrulukDurumu, SaglikDurumu, HastalikAdi, AsiTarihi, TedaviTarihi, Aciklama, History
        FROM HayvanDetay
        WHERE CiftlikID LIKE <cfqueryparam value="%#url.arama_kelimesi#%" cfsqltype="CF_SQL_VARCHAR">
        OR HayvanCesidi LIKE <cfqueryparam value="%#url.arama_kelimesi#%" cfsqltype="CF_SQL_VARCHAR">
        OR HayvanTuru LIKE <cfqueryparam value="%#url.arama_kelimesi#%" cfsqltype="CF_SQL_VARCHAR">
        OR Cinsiyet LIKE <cfqueryparam value="%#url.arama_kelimesi#%" cfsqltype="CF_SQL_VARCHAR">
        OR Mahsul LIKE <cfqueryparam value="%#url.arama_kelimesi#%" cfsqltype="CF_SQL_VARCHAR">
        OR GebelikDurumu LIKE <cfqueryparam value="%#url.arama_kelimesi#%" cfsqltype="CF_SQL_VARCHAR">
        OR YavrulukDurumu LIKE <cfqueryparam value="%#url.arama_kelimesi#%" cfsqltype="CF_SQL_VARCHAR">
        OR SaglikDurumu LIKE <cfqueryparam value="%#url.arama_kelimesi#%" cfsqltype="CF_SQL_VARCHAR">
        OR HastalikAdi LIKE <cfqueryparam value="%#url.arama_kelimesi#%" cfsqltype="CF_SQL_VARCHAR">
        OR Aciklama LIKE <cfqueryparam value="%#url.arama_kelimesi#%" cfsqltype="CF_SQL_VARCHAR">
        OR History LIKE <cfqueryparam value="%#url.arama_kelimesi#%" cfsqltype="CF_SQL_VARCHAR">
    </cfquery>
</cfif>

<div name="HarcananEkleme" class="container p-3 m-3 ">
    <table><h6>Hasta Hayvan Bilgileri</h6></table>
    <table class="table table-bordered table-sm small" align="center" widht="99%" border="1" bordercolor="#99999">
        <tr>
            <td>Çiftlik ID</td>
            <td>Hayvan ID</td>
            <td>Hayvan Çesidi</td>
            <td>HayvanTuru</td>
            <td>Cinsiyet</td>
            <td>Mahsul</td>
            <td>GebelikDurumu</td>
            <td>YavrulukDurumu</td>
            <td>Sağlık Durumu</td>
            <td>Hastalık Adı</td>
            <td>Aşı Tarihi</td>
            <td>Tedavi Tarihi</td>
            <td>Açıklama</td>
            <td>Güncelleme Tarihi</td>
        </tr>
        <cfoutput query="Hayvanlarim">
            <tr>
                <td>#CiftlikID#</td>
                <td>#HayvanID#</td>
                <td>#HayvanCesidi#</td>
                <td>#HayvanTuru#</td>
                <td>#Cinsiyet#</td>
                <td>#Mahsul#</td>
                <td>#GebelikDurumu#</td>
                <td>#YavrulukDurumu#</td>
                <td>#SaglikDurumu#</td>
                <td>#HastalikAdi#</td>
                <td>#AsiTarihi#</td>
                <td>#TedaviTarihi#</td>
                <td>#Aciklama#</td>
                <td>#History#</td>
            </tr>
        </cfoutput>
    </table>
</div>