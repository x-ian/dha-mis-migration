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
  var wstream = fs.createWriteStream('c:\\Users\\IEUser\\Desktop\\tabledef.csv', {flags: 'a'});
  query = "SELECT * FROM " + table;
  // query = ""SELECT c.CustomerID, c.CompanyName,COUNT(soh.SalesOrderID) AS OrderCount FROM SalesLT.Customer AS c LEFT OUTER JOIN SalesLT.SalesOrderHeader AS soh ON c.CustomerID = soh.CustomerID GROUP BY c.CustomerID, c.CompanyName ORDER BY OrderCount DESC;"
  request = new Request(query, function(err, rowCount) {
    if (err) {
      console.log('Requesterror');
       console.log(err);
     } else {
       wstream.end();
       connection.close();
     }
    });

	process.stdout.write("Dumping " + table);
	// column
	request.on('columnMetadata', function (columns) {
	  var header = "";
	  columns.forEach(function(column) {
		  header += column.colName + "," + column.type.name + ",";
	  });
	  wstream.write(table + "," + header.substring(0, header.length - 1) + '\r\n');
	});

	connection.execSql(request);

}
