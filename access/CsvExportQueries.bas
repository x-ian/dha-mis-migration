Attribute VB_Name = "CsvExportQueries"

Option Compare Database

Public Sub CsvExportQueries()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentData
 
    Dim intFile As Integer
    Dim strFile As String
    strFile = "C:\Users\IEUser\Desktop\File.txt"
    intFile = FreeFile
    Open strFile For Output As #intFile

	Dim t as single
	
    On Error Resume Next

    For Each obj In dbs.AllQueries
    
        Dim qdf As DAO.QueryDef
        Set qdf = CurrentDb.QueryDefs(obj.Name)
        'Print #intFile, qdf.Name & " " & qdf.Type
        
        ' select 0, union 128, crosstab 16
        If qdf.Type = 0 Or qdf.Type = 128 Or qdf.Type = 16 Then

            Dim params As Boolean
            params = False
        
            For Each p In qdf.Parameters
                If Err.Number <> 0 Then
                    Print #intFile, "xx Error occured for query " & obj.Name
                    Err.Clear
                    GoTo END_FOR
                End If
                'Debug.Print p.Name & " " & p.Value
                params = True
            Next p
            
            If Not params Then
				t = Timer
                DoCmd.TransferText acExportDelim, , obj.Name, obj.Name & ".csv", True, , 65001
                Print #intFile, "aa query without params " & obj.Name & " export took " & ((Timer - t) / 1000) & " secs"
            Else
                Print #intFile, "bb ignoring query with params " & obj.Name
            End If

        Else
            Print #intFile, "cc ignoring non-scriptable query " & obj.Name
        End If
END_FOR:
    Next obj
        
    Close #intFile
End Sub




