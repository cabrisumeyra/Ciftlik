<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
    <style>
        .duyuru{
                background-color: #e7f17ea2;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                border-radius: 10px;
            }
</style>
</head>
<body>
</body>
</html>

<div class="row">
    <div name="anasayfa" onclick=" document.location='index.cfm' " class="col-md-9">
        <cfimage name ="isresim" 
        action="writeToBrowser" 
        source="..\img\farm.png" 
        alt="Resim" width="100">
    </div>
    
</div>
    
    <div name="genel" class="container d-flex justify-content-center">
        <div class="row text-center">
            <div name="GunIsleri " class="  p-3 m-3  col-sm-2 " onclick=" document.location='Islerim.cfm' ">
                <h3>Bugünkü İşlerim</h3>
                <cfimage name ="isresim" 
                        action="writeToBrowser" 
                        source="..\img\index\isler.png" 
                        alt="Resim" width="100">
            </div>
    
            <div name="GunUretim " class=" p-3 m-3 col-sm-2 " onclick=" document.location='Uretim.cfm' ">
                <h3>Günlük Üretim Bilgileri</h3>
                <cfimage name ="isresim" 
                        action="writeToBrowser" 
                        source="..\img\index\uretim.png" 
                        alt="Resim" width="100">
            </div>
    
            <div name="GunSatis " class=" p-3 m-3 col-sm-2 " onclick=" document.location='Finans.cfm' ">
                <h3>Günlük Satış Bilgileri</h3>
                <cfimage name ="isresim" 
                        action="writeToBrowser" 
                        source="..\img\index\finans.png" 
                        alt="Resim" width="100">
            </div>
    
            <div name="Bakim " class=" p-3 m-3 col-sm-2 " onclick=" document.location='Bakim.cfm' ">
                <h3>Bakım Takibi</h3>
                <cfimage name ="isresim" 
                        action="writeToBrowser" 
                        source="..\img\index\bakim.png" 
                        alt="Resim" width="100">
            </div>
    
            <div name="Hayvanlarim " class=" p-3 m-3 col-sm-2 " onclick="document.location='GenelHayvan.cfm'">
                <h3>Hayvanlarım</h3>
                <cfimage name ="isresim" 
                        action="writeToBrowser" 
                        source="..\img\index\hayvanlar.png" 
                        alt="Resim" width="100">
            </div>
    
            <div name="Ciftliklerim " class=" p-3 m-3 col-sm-2 " onclick="document.location='CiftlikBilgi.cfm'">
                <h3>Çiftliklerim</h3>
                <cfimage name ="isresim" 
                        action="writeToBrowser" 
                        source="..\img\index\ciftlik.png" 
                        alt="Resim" width="100">
            </div>
    
            <div name="Raporlar" class=" p-3 m-3 col-sm-2 " onclick="document.location='Raporlama.cfm'">
                <h3>Raporlar</h3>
                <cfimage name ="isresim" 
                        action="writeToBrowser" 
                        source="..\img\index\rapor.png" 
                        alt="Resim" width="100">
            </div>
    
            <div name="CalisanEkrani" class=" p-3 m-3 col-sm-2 " onclick="document.location='Calisan.cfm'">
                <h3>Çalışanlarım</h3>
                <cfimage name ="isresim" 
                        action="writeToBrowser" 
                        source="..\img\index\calisan.png" 
                        alt="Resim" width="100">
            </div>
            <div name="dashboard" class=" p-3 m-3 col-sm-2 " onclick="document.location='Dashboard.cfm'">
                <h3>Dashboard</h3>
                <cfimage name ="isresim" 
                        action="writeToBrowser" 
                        source="..\img\index\dashboard.png" 
                        alt="Resim" width="100">
            </div>
            <div name="excel-veri-aktarimi" class="p-2 m-2 col-md-2">
                <cfimage name ="isresim" 
                action="writeToBrowser" 
                source="..\img\excel.png" 
                alt="Resim" width="30">

                <cfform method="post"  enctype="multipart/form-data" action="excel.cfm">
                    <input type="file" name="file_11" id="file_11">
                    <input type="hidden"  name="FileName" id="FileName">
                    <input type="hidden"  name="is_submitted" id="is_submitted">
                    <input type="submit">
                </cfform>
        
            </div>
        </div>
    </div>

