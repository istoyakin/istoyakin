function OZON_import() {
  /// Объявляем перменные книги, листа
              var ss = SpreadsheetApp.getActiveSpreadsheet();
              var sheet = ss.getSheetByName("OZON_API_import");
  /// Выбираем диапазон дат для выгрузки
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
  /// Подготовливаем лист
              var j = Number(sheet.getRange(1,15).getValue()); /// Переменная для запоминания точки итерации, если скрипт работал дольше 6 минут
              var r1 = sheet.getLastRow();
              var c1 = sheet.getMaxColumns();
              if (r1<2){r1 = 2} /// Не трогаем 1ю строку
              if (j==0) 
              {sheet.getRange(2,1,r1,c1).clearContent();}
              else
              {sheet.getRange((j-1)*1000+2,1,r1,c1).clearContent();} /// Если скрпит прервался, удаляем данные до последней успешной итерации
  /// Начинаем итерацию
          for (var iter = 1000; iter == 1000; j = j+1){ /// Выполняем пока количество строк в ответе не станет отличным от 1000 (максимальное количество строк в итерации)
          var offset = j*1000; /// Отступ в ответе от API
          sheet.getRange(1,15).setValue(j); /// записываем номер итерации
  /// Параметры API запроса
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
                  'Client-Id': '',
                  'Api-Key': ''
                  }
                }
   ///Запрос API
                var response = UrlFetchApp.fetch('https://api-seller.ozon.ru/v1/analytics/data', options);
                var response = JSON.parse(response.getContentText());
   ///Создаем массив для записи в него данных
                var response_iter = new Array(response.result.data.length);
                for (var k=0; k<=response.result.data.length-1 ;k=k+1)
                {
                  response_iter[k] = new Array(6);
                }
                var lastRow = sheet.getLastRow();
   /// Записываем нужные данные
              for (i=0; i<=response.result.data.length-1; i=i+1) {
                response_iter[i][0] = response.result.data[i].dimensions[0].id;
                response_iter[i][1] = response.result.data[i].dimensions[1].id;
                response_iter[i][2] = response.result.data[i].metrics[0];
                response_iter[i][3] = response.result.data[i].metrics[1];
                response_iter[i][4] = response.result.data[i].metrics[2];
                response_iter[i][5] = response.result.data[i].metrics[3];
              }
  /// Записываем данные в таблицу
              sheet.getRange(lastRow+1,1,response_iter.length,6).setValues(response_iter);
  /// Записываем колиство строк для последующей проверки в цикле
              iter = response.result.data.length;
  /// Чистим массив
              response_iter.splice(0,response_iter.length);
              
}
  /// Если скрипт завершен успешно - переносим данные в другую таблицу (чтобы таблица не была в постоянном обновлении)
copy_clear(lastRow+iter,6,"OZON_API_Import","import2");
  /// Обновляем показатель итерации
sheet.getRange(1,15).setValue("0");

}
