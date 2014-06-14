var fs = require('fs'),
  xls = require('xls');
fs.readdir('xls', function(err, list) {
  if (err) throw err;
  //console.log(list[0]);
  for (var i = 0; i < list.length; i++) {
    if (!(/(^|.\/)\.+[^\/\.]/g).test(list[i])) {
      (function(filename) {
        console.log(filename)
        require('xls').parse('xls/' + filename, function(err, data) {
          var output = [];
          var rows = data.excel_workbook.sheets.sheet.rows.row;
          for (var i = 2; i < rows.length; i++) {
            var cells = rows[i].cell;
            //console.log(cells);
            var obj = {};
            for (var j = 0; j < cells.length; j++) {
              var o = cells[j];
              if (o['$t'])
                switch (cells[j].col) {
                  case 1:
                    obj['屬性'] = o['$t'];
                    break;
                  case 2:
                    obj['機構名稱'] = o['$t'];
                    break;
                  case 3:
                    obj['負責人'] = o['$t'];
                    break;
                  case 4:
                    obj['地址'] = o['$t'];
                    break;
                  case 5:
                    obj['電話'] = o['$t'];
                    break;
                  case 6:
                    obj['收容對象'] = o['$t'];
                    break;
                  case 7:
                    obj['核定收容人數'] = o['$t'];
                    break;
                  case 8:
                    obj['立案日期'] = o['$t'];
                    break;
                }
            }
            if (Object.keys(obj).length !== 0)
              output.push(obj);
          }
          fs.writeFile('json/' + filename.replace('xls', 'json'), JSON.stringify(output), function(err) {
            if (err) throw err;
            else 
              console.log(filename + ' has been converted to json');
          });
          //console.log(output);
        });
      })(list[i]);
    }
  }
});
