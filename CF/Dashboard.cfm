<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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

<cfquery name="HayvanSayi" datasource="Ciftlik">
    SELECT
        HayvanCesidi, 
        COUNT(HayvanCesidi) as ToplamHayvan
    FROM 
        HayvanDetay
    GROUP BY HayvanCesidi
</cfquery>
<cfquery name="HayvanTuruSayisi" datasource="Ciftlik">
    SELECT
     HayvanTuru, COUNT(HayvanTuru) as HayvanCesitleri
    FROM 
    HayvanDetay
    GROUP By HayvanTuru
</cfquery>
<cfquery name="Uretim" datasource="Ciftlik">
    SELECT MahsulCesidi, GunlukUretimAdeti
    FROM Uretim
    WHERE MahsulCesidi IS NOT NULL
</cfquery>

<div class="container-fluid">
    <div class="row">
        <div class="p-3 m-3 col-md-2">
            
            <canvas id="HayvanSayisi" ></canvas>
            <script>
                // Veritabanından alınan verileri JavaScript değişkenlerine aktar
                var labels = [];
                var data = [];
                <cfoutput query="HayvanSayi">
                    labels.push("#HayvanSayi.HayvanCesidi#");
                    data.push(#HayvanSayi.ToplamHayvan#);
                </cfoutput>

                // Veri seti
                var dataset = {
                    labels: labels,
                    datasets: [{
                        label: 'Donut Grafik Örneği',
                        data: data,
                        backgroundColor: [
                            'rgba(255, 99, 132)',
                            'rgba(54, 162, 235)',
                            'rgba(255, 206, 86)'
                        ],
                        borderColor: [
                            'rgba(255, 99, 132, 1)',
                            'rgba(54, 162, 235, 1)',
                            'rgba(255, 206, 86, 1)'
                        ],
                        borderWidth: 1
                    }]
                };

                // Seçenekler
                var options = {
                    responsive: true,
                    maintainAspectRatio: false,
                    title: {
                        display: true,
                        text: 'Hayvan Sayısı'
                    }
                };

                // Donut Grafik oluştur
                var ctx = document.getElementById('HayvanSayisi').getContext('2d');
                var HayvanSayisi = new Chart(ctx, {
                    type: 'doughnut',
                    data: dataset,
                    options: options
                });
            </script>
        </div>
        <div class="p-3 m-3 col-md-8">
            
            <canvas id="HayvanTuruSayisi" ></canvas>
            <script>
                // Veritabanından alınan verileri JavaScript değişkenlerine aktar
                var labels = [];
                var data = [];
                <cfoutput query="HayvanTuruSayisi">
                    labels.push("#HayvanTuruSayisi.HayvanTuru#");
                    data.push(#HayvanTuruSayisi.HayvanCesitleri#);
                </cfoutput>

                // Veri seti
                var dataset = {
                    labels: labels,
                    datasets: [{
                        label: 'Hayvan Türlerine Göre Hayvan Sayıları',
                        data: data,
                        backgroundColor: [
                            'rgba(255, 99, 132)',
                            'rgba(54, 162, 235)',
                            'rgba(255, 206, 86)'
                        ],
                        borderColor: [
                            'rgba(255, 99, 132, 1)',
                            'rgba(54, 162, 235, 1)',
                            'rgba(255, 206, 86, 1)'
                        ],
                        borderWidth: 1
                    }]
                };

                // Seçenekler
                var options = {
                    responsive: true,
                    maintainAspectRatio: false,
                    title: {
                        display: true,
                        text: 'Hayvan Sayısı'
                    }
                };

                // Donut Grafik oluştur
                var ctx = document.getElementById('HayvanTuruSayisi').getContext('2d');
                var HayvanTuruSayisi = new Chart(ctx, {
                    type: 'line',
                    data: dataset,
                    options: options
                });
            </script>
        </div>
        <div class="p-3 m-3 col-md-5">
            
            <canvas id="Uretim" ></canvas>
            <script>
                // Veritabanından alınan verileri JavaScript değişkenlerine aktar
                var labels = [];
                var data = [];
                <cfoutput query="Uretim">
                    labels.push("#Uretim.MahsulCesidi#");
                    data.push(#Uretim.GunlukUretimAdeti#);
                </cfoutput>

                // Veri seti
                var dataset = {
                    labels: labels,
                    datasets: [{
                        label: 'Mahsul Satış Ücretleri',
                        data: data,
                        backgroundColor: [
                            'rgba(255, 99, 132)',
                            'rgba(54, 162, 235)',
                            'rgba(255, 206, 86)'
                        ],
                        borderColor: [
                            'rgba(255, 99, 132, 1)',
                            'rgba(54, 162, 235, 1)',
                            'rgba(255, 206, 86, 1)'
                        ],
                        borderWidth: 1
                    }]
                };

                // Seçenekler
                var options = {
                    responsive: true,
                    maintainAspectRatio: false,
                    title: {
                        display: true,
                        text: 'Hayvan Sayısı'
                    }
                };

                // Donut Grafik oluştur
                var ctx = document.getElementById('Uretim').getContext('2d');
                var HayvanTuruSayisi = new Chart(ctx, {
                    type: 'bar',
                    data: dataset,
                    options: options
                });
            </script>
        </div>
    </div>
</div>


