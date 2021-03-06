var sql = require('mssql');

sql.connect("mssql://root:root@localhost/hivdata2").then(function() {
    // Query

    new sql.Request().query('select * from mytable').then(function(recordset) {
        console.dir(recordset);
    }).catch(function(err) {
      console.dir(err);
        // ... query error checks
    });

    // Stored Procedure

    new sql.Request()
    .input('input_parameter', sql.Int, value)
    .output('output_parameter', sql.VarChar(50))
    .execute('procedure_name').then(function(recordsets) {
        console.dir(recordsets);
    }).catch(function(err) {
      console.dir(err);
        // ... execute error checks
    });

    // ES6 Tagged template literals (experimental)

    sql.query`select * from mytable where id = ${value}`.then(function(recordset) {
        console.dir(recordset);
    }).catch(function(err) {
      console.dir(err);
        // ... query error checks
    });
}).catch(function(err) {
  console.dir(err);
    // ... connect error checks
});
