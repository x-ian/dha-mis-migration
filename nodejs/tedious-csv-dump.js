var Connection = require('tedious').Connection;
var fs = require('fs');
var moment = require('moment');
var async = require('async');

var config = {
  userName: 'root',
  password: 'root',
  server: 'localhost',
  //server: 'IE11WIN7\\SQLEXPRESS',
  // If you are on Microsoft Azure, you need this:
  options: {encrypt: false, database: 'HIVData', requestTimeout: 0 /*120000*/}
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
  var wstream = fs.createWriteStream('c:\\Users\\IEUser\\Desktop\\' + table + '.csv');
  query = "SELECT * FROM " + table + " ORDER BY ID";
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
        if (column.colName === 'SSMA_TimeStamp') {
          // do nothing
        } else {
          header += "\"" + column.colName + " (" + column.type.name + ")\",";
          //header += "\"" + column.colName + "\",";
        }
      });
      wstream.write(header.substring(0, header.length - 1) + '\r\n');
    });

    // recordset
    request.on('row', function(columns) {
      var result = "";
      //console.log(columns);
      columns.forEach(function(column) {
        if (column.metadata.colName === 'SSMA_TimeStamp') {
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
		  } else if (type === 'Real') {
			  if (("" + column.value).indexOf("." == -1)) {
				  result += column.value + '.00,';
			  } else {
				  result += column.value + ',';
			  }
		  } else if (type === 'FloatN') {
			  if (("" + column.value).indexOf("." == -1)) {
				  result += column.value + '.00,';
			  } else {
				  result += column.value + ',';
			  }
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

  function endsWith(str, suffix) {
    return str.indexOf(suffix, str.length - suffix.length) !== -1;
}