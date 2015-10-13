google.load('visualization', '1', {'packages': ['geochart']});
google.setOnLoadCallback(drawRegionsMap);

function drawRegionsMap() {
    // data ตัวหน้าเป็น รหัส หรือ ชื่อก็ได้ ดูจาก https://en.wikipedia.org/wiki/ISO_3166-2:TH    
	var data = google.visualization.arrayToDataTable([
        ['province', 'valueTitle'],  // ตัวแรกเหมือนจะเป็น Caption
        ['TH-57' , 883], 
        ['Chiang Mai' , 1784], // ชื่อ เป็น ไทย หรือ Eng ก็ได้ (แต่ต้องสะกดถูกต้อง และเป็นฃื่อทางการ 
        ['กรุงเทพมหานคร' , 784],
        ['ลำปาง' , 619],
        ['ลำพูน' , 619],
        ['พระนครศรีอยุธยา' , 619],
        ['สุพรรณบุรี' , 619],
        ['Surat Thani',888],
    ]);
	var options = {'region':'TH',
                   'resolution':'provinces' ,
                   //'colors': ['#FF0000', '#00FF00']
	};
    var chart = new google.visualization.GeoChart(
                            document.getElementById('chart_div')
    );
    chart.draw(data, options);
};
    
