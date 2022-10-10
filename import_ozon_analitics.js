function OZON_import() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
              var ss = SpreadsheetApp.getActiveSpreadsheet();
              var sheet = ss.getSheetByName("OZON_API_import");
              var today = new Date();
              var dd = String(today.getDate()).padStart(2, '0');
              var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
              var yyyy = today.getFullYear();
              today = yyyy + '-' + mm + '-' + dd;
              if (mm < 2) {
                mm = mm + 11;
                yyyy = yyyy - 1;
                mm = String(mm).padStart(2,'0');
              } else {
                mm = mm-1;
                mm = String(mm).padStart(2,'0');                
              }
              var date_from = yyyy + '-' + mm + '-' + dd;
              var j = Number(sheet.getRange(1,15).getValue());
              var r1 = sheet.getLastRow();
              var c1 = sheet.getMaxColumns();
              if (r1<2){r1 = 2}
              if (j==0) 
              {sheet.getRange(2,1,r1,c1).clearContent();}
              else
              {sheet.getRange((j-1)*1000+2,1,r1,c1).clearContent();}
          for (var iter = 1000; iter == 1000; j = j+1){
          var offset = j*1000;
          sheet.getRange(1,15).setValue(j);
                var data = {
                  "date_from": date_from,
                  "date_to": today,
                  "metrics": [
                  "ordered_units",
                  "revenue",
                  "adv_sum_all",
                  "returns"
                  ],
                  "dimension": [
                  "sku",
                  "day"
                  ],
                  "filters": [],
                  "sort": [
                    /////{
                    /////"key": "sku",
                    /////"order": "DESC"
                    /////}
                  ],
                  "limit": 1000,
                  "offset": offset
                    }
                var options = {
                  'method': 'post',
                  'payload': JSON.stringify(data),
                  'headers': {
                  'Client-Id': '59173',
                  'Api-Key': '72ba412a-efd7-46c6-9e16-9f0fe9ad5551'
                  }
                }
                var response = UrlFetchApp.fetch('https://api-seller.ozon.ru/v1/analytics/data', options);
                var response = JSON.parse(response.getContentText());
                var response_iter = new Array(response.result.data.length);
                for (var k=0; k<=response.result.data.length-1 ;k=k+1)
                {
                  response_iter[k] = new Array(6);
                }
                var lastRow = sheet.getLastRow();
              for (i=0; i<=response.result.data.length-1; i=i+1) {
                response_iter[i][0] = response.result.data[i].dimensions[0].id;
                response_iter[i][1] = response.result.data[i].dimensions[1].id;
                response_iter[i][2] = response.result.data[i].metrics[0];
                response_iter[i][3] = response.result.data[i].metrics[1];
                response_iter[i][4] = response.result.data[i].metrics[2];
                response_iter[i][5] = response.result.data[i].metrics[3];
              }
              sheet.getRange(lastRow+1,1,response_iter.length,6).setValues(response_iter);
              iter = response.result.data.length;
              response_iter.splice(0,response_iter.length);
              
}
copy_clear(lastRow+iter,6,"OZON_API_Import","import2");
sheet.getRange(1,15).setValue("0");

}
