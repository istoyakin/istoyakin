function import_order() {
              var ss = SpreadsheetApp.getActiveSpreadsheet();
              var sheet = ss.getSheetByName("OZON_API_import_orders");
              if (sheet.getMaxRows() < 70000){
              sheet.insertRowsAfter(sheet.getMaxRows(), 70000);
              }
              var today = new Date();
              var dd = String(today.getDate()).padStart(2, '0');
              var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
              var yyyy = today.getFullYear();
              today = yyyy + '-' + mm + '-' + dd;
              if (mm < 3) {
                mm = mm + 10;
                yyyy = yyyy - 1;
                mm = String(mm).padStart(2,'0');
              } else {
                mm = mm-2;
                mm = String(mm).padStart(2,'0');                
              }
              var j = Number(sheet.getRange(1,12).getValue());
              if (j=="") {j=0;}
              var r1 = sheet.getMaxRows();
              var c1 = sheet.getMaxColumns();
              if (r1<2){r1 = 2}
              if (j==0) 
                {
                sheet.getRange(1,1,r1,c1).clearContent();
                sheet.getRange(1,1).setValue("Артикул");
                sheet.getRange(1,2).setValue("Кол-во");
                sheet.getRange(1,3).setValue("Цена");
                sheet.getRange(1,4).setValue("Дата");
                sheet.getRange(1,5).setValue("Регион");
                sheet.getRange(1,6).setValue("Склад");
                sheet.getRange(1,7).setValue("За вычетом комиссии");
                sheet.getRange(1,8).setValue("...");
                sheet.getRange(1,9).setValue("...");
                sheet.getRange(1,10).setValue("Доставка");
                sheet.getRange(1,11).setValue("FRESH");
                }
              else
              {sheet.getRange((j-1)*1000+2,1,r1,c1).deleteCells;}
              var date_from = yyyy + '-' + mm + '-' + dd;
      for (var iter = 1000;iter==1000; j=j+1){
                 var offset = j*1000;
                 sheet.getRange(1,12).setValue(j);
                var data = {
                    'dir': 'asc',
                    'filter': {
                        "since": date_from+"T00:00:01Z",
                        "to": today+"T23:59:59Z"
                    },
                    "limit": 1000,
                    "offset": offset,
                    "translit": true,
                    "with": {
                        "analytics_data": true,
                        "financial_data": true
                    }
                }
                var options = {
                  'method': 'post',
                  'payload': JSON.stringify(data),
                  'headers': {
                  'Client-Id': '',
                  'Api-Key': ''
                  }
                };
                var response = UrlFetchApp.fetch('https://api-seller.ozon.ru/v2/posting/fbo/list', options);
                
                // Logger.log(response.getContentText());
                response = JSON.parse(response);
                var lastRow = sheet.getLastRow();
                var response_iter = new Array(response.result.length);
                for (var k=0; k<=response.result.length-1 ;k=k+1)
                {
                  response_iter[k] = new Array(10);
                }
              for (var i=0; i<=response.result.length-1; i=i+1) 
              {
                for (var l=0; l<=response.result[i].products.length - 1; l = l+1)
                {
                response_iter[i+l][0] = response.result[i].products[l].offer_id;
                response_iter[i+l][1] = response.result[i].products[l].quantity;
                response_iter[i+l][2] = response.result[i].products[l].price;
                response_iter[i+l][3] = response.result[i].created_at;
                response_iter[i+l][4] = response.result[i].analytics_data.region;
                response_iter[i+l][5] = response.result[i].analytics_data.warehouse_name;
                response_iter[i+l][6] = response.result[i].financial_data.products[l].payout;
                response_iter[i+l][7] = response.result[i].financial_data.products[l].item_services.marketplace_service_item_fulfillment;
                response_iter[i+l][8] = response.result[i].financial_data.products[l].item_services.marketplace_service_item_direct_flow_trans;
                response_iter[i+l][9] = response.result[i].financial_data.products[l].item_services.marketplace_service_item_deliv_to_customer;
                if (response_iter[i+l][5].toString().includes('FRESH') == true){response_iter[i+l][10]='Да';} else {response_iter[i+l][10]= 'Нет';}
                }
                l=0
              }
              iter = response.result.length;
              sheet.getRange(lastRow+1,1,response_iter.length,11).setValues(response_iter);
              response_iter.splice(0,response_iter.length);
      }
    sheet.getRange(1,12).setValue("0");
    new_rename(lastRow+iter,11,"OZON_API_import_orders","import_заказы");


}
