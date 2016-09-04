Attribute VB_Name = "CsvExport"
Option Compare Database
Public Sub ExportAll()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentData
    For Each obj In dbs.AllTables
        If Left(obj.Name, 4) <> "MSys" Then
            DoCmd.TransferText acExportDelim, , obj.Name, obj.Name & ".csv", True, , 65001
            ' DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, obj.Name, obj.Name & ".xls", True
        End If
    Next obj
End Sub
