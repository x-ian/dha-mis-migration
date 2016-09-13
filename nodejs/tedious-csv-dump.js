var Connection = require('tedious').Connection;
var fs = require('fs');
var moment = require('moment');
var async = require('async');
var roundTo = require('round-to');
var numeral = require('numeral');

var config = {
  userName: 'root',
  password: 'root',
  server: 'localhost',
  //server: 'IE11WIN7\\SQLEXPRESS',
  // If you are on Microsoft Azure, you need this:
  options: {encrypt: false, database: 'HIVData2', requestTimeout: 0 /*120000*/}
};
var connection = new Connection(config);
connection.on('connect', function(err) {
  if (err) {
    console.log('Connecterror');
    console.log(err);
  } else {
    // If no error, then good to proceed.
    console.log("Connected");

    var table = process.argv[2];
    dumpTable(table);
  }
});

var Request = require('tedious').Request;
var TYPES = require('tedious').TYPES;

function dumpTable(table) {
  var wstream = fs.createWriteStream('c:\\Users\\IEUser\\Desktop\\csv-export-sql\\' + table + '.csv');
 
	if (table === 'code_hfacility_GPS' 
	|| table === 'map_supply_item_cms_code'
	|| table === 'map_user'
	|| table === 'NMCP_DL20_export'
	|| table === 'NMCP_DL21_export'
	|| table === 'NMCP_DL26_export'
	|| table === 'NMCP_DL30_export'
	|| table === 'population'
	|| table === 'pop_sex_district_hiv') {
 query = "SELECT * FROM " + table;
 		
	} else {
 query = "SELECT * FROM " + table + " ORDER BY ID";
 		
	}
  // query = ""SELECT c.CustomerID, c.CompanyName,COUNT(soh.SalesOrderID) AS OrderCount FROM SalesLT.Customer AS c LEFT OUTER JOIN SalesLT.SalesOrderHeader AS soh ON c.CustomerID = soh.CustomerID GROUP BY c.CustomerID, c.CompanyName ORDER BY OrderCount DESC;"
  var start = moment();
  request = new Request(query, function(err, rowCount) {
    if (err) {
      console.log('Requesterror');
       console.log(err);
     } else {
       wstream.end();
       var end = moment();
       console.log(" - Done (" + end.diff(start, 'seconds') + " secs, " + rowCount + " rows)");
       console.log('');
       connection.close();
     }
    });

    process.stdout.write("Dumping " + table);
    // column
    request.on('columnMetadata', function (columns) {
      var header = "";
      columns.forEach(function(column) {
        if (column.colName === 'SSMA_TimeStamp'
			|| column.colName === 'my_timestamp') {
          // do nothing
        } else {
          //header += "\"" + column.colName + " (" + column.type.name + ")\",";
          header += "\"" + column.colName + "\",";
        }
      });
      wstream.write(header.substring(0, header.length - 1) + '\r\n');
    });

    // recordset
    request.on('row', function(columns) {
      var result = "";
      //console.log(columns);
      columns.forEach(function(column) {
        if (column.metadata.colName === 'SSMA_TimeStamp'
			|| column.metadata.colName === 'my_timestamp') {
          // do nothing
        } else if (column.value === null) {
          result += ",";
        } else {
          type = column.metadata.type.name;
          if (type === 'NVarChar') {
            result+= "\"" + column.value.replace(/"/g, '""') + "\",";
          } else if (type === 'DateTime2N') {
            // result += moment(column.value).format('M/D/YYYY H:mm:SS') + ",";
            result += moment.utc(column.value).format('M/D/YYYY H:mm:ss') + ",";
            //result += column.value + ',';
          } else if (type === 'BitN') {
			  if (column.value == true) {
				  result += "1,";
			  } else if (column.value == false) {
				  result += "0,";
			  } else {
				  result += ",";
			  }
		  } else if (type === 'Money') {
			  result += "$" + numeral(column.value).format('0.00') + ",";
		  } else if (type === 'Real') {
			  //result += formatDecimal(column.value) + ',';
			  /*
			  if (column.value < 0) {
				result += numeral(roundTo.up(column.value, 2)).format('0.00') + ",";
			} else {
				result += numeral(roundTo.down(column.value, 2)).format('0.00') + ",";
			  }
			  */
			  
			  var temp = "";
			  temp = numeral(column.value).format('0.00000');
			  result += temp.substring(0, temp.length - 3) + ",";
			  
			  /*
			  if (("" + column.value).indexOf("." == -1)) {
				  result += column.value + '.00,';
			  } else {
				  result += column.value + ',';
			  }*/
			  //result+= column.value + ",";
		  } else if (type === 'FloatN') {
			  //result += formatDecimal(column.value) + ',';
			  //result+= column.value + ",";
			  /*
			  if (column.value < 0) {
				result += numeral(roundTo.up(column.value, 2)).format('0.00') + ",";
			} else {
				result += numeral(roundTo.down(column.value, 2)).format('0.00') + ",";
			  }
			  */
			  
			  var temp = "";
			  temp = numeral(column.value).format('0.00000');
			  result += temp.substring(0, temp.length - 3) + ",";
			  
			  /*
			  if (("" + column.value).indexOf("." == -1)) {
				  result += column.value + '.00,';
			  } else {
				  result += column.value + ',';
			  }
			  */
          } else {
            result+= column.value + ",";
          }
        }
      });
      wstream.write(result.substring(0, result.length - 1) + '\r\n');
    });

    // request.on('doneProc', function(rowCount, more, rows) {
    //   console.log(rowCount + ' rows returned');
    //   wstream.end();
    // });
    connection.execSql(request);

  }

    function formatDecimal(value) {
	  		  var indexOf = ("" + value).indexOf('.');
			  if (indexOf === -1) {
				  // no decimal at all
				  return "" + value + ".00";
			  } else if ((indexOf - 1) === ("" + value).length) {
				  // only one digit after point
				  return "" + value + "0";
			  } else {
				  // 
				  return ("" + value).substring(0, indexOf + 3);
			  }

  }


  function endsWith(str, suffix) {
    return str.indexOf(suffix, str.length - suffix.length) !== -1;
}