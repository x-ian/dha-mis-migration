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

    Dim t As Single
    Dim t2 As Single
    Dim t3 As Single
    
    t2 = Timer
    'On Error Resume Next

    For Each obj In dbs.AllQueries
    
        Debug.Print obj.Name
        
        Dim qdf As DAO.QueryDef
        Set qdf = CurrentDb.QueryDefs(obj.Name)
        'Print #intFile, "__ " & qdf.name & " " & queryDefType(qdf.Type)
        
        Dim params As Boolean
        params = False
        
        For Each p In qdf.Parameters
            If Err.Number <> 0 Then
                Print #intFile, "xx Error occured for query " & obj.Name & " " & queryDefType(qdf.Type)
                Err.Clear
                GoTo END_FOR
            End If
            'Debug.Print p.Name & " " & p.Value
            params = True
        Next p
            
        If Not params Then
           If qdf.Type <> dbQDelete And qdf.Type <> dbQAction And qdf.Type <> dbQAppend And qdf.Type <> dbQUpdate Then ' Or qdf.Type = 128 Or qdf.Type = 16 Then
                If obj.Name = "calc_test" Or obj.Name = "calc_test_result" Then
                    Debug.Print "Skip calc_test and calc_test_result"
                    Print #intFile, "xx Skip calc_test and calc_test_result"
                Else
                    t = Timer
                    DoCmd.TransferText acExportDelim, , obj.Name, obj.Name & ".csv", True, , 65001
                    Print #intFile, "aa (param-less query)" & vbTab & obj.Name & vbTab & queryDefType(qdf.Type) & vbTab & Round((Timer - t), 2) & vbTab & "secs" & vbTab & Replace(qdf.SQL, vbCrLf, " ")
                End If

            Else
                DoCmd.SetWarnings False
                t3 = Timer
                DoCmd.OpenQuery (obj.Name)
                DoCmd.SetWarnings True
                Print #intFile, "dd (param-less query)" & vbTab & obj.Name & vbTab & queryDefType(qdf.Type) & vbTab & Round((Timer - t3), 2) & vbTab & "secs" & vbTab & Replace(qdf.SQL, vbCrLf, " ")
                'Print #intFile, "cc ignoring non-scriptable query " & obj.Name & " of type " & queryDefType(qdf.Type)
            End If
        Else
            Print #intFile, "bb ignoring query with params " & obj.Name & " of type " & queryDefType(qdf.Type)
        End If

END_FOR:
    Next obj
        
    Debug.Print "total time: " & Round((Timer - t2), 2)
    Close #intFile
End Sub

Function queryDefType(typ As Integer) As String

	' https://msdn.microsoft.com/en-us/library/office/ff192931.aspx

	s = ""
	Select Case typ
		Case 240
			s = "dbQAction"
		Case 64
			s = "dbQAppend"
		Case 160
			s = "dbQCompound"
		Case 16
			s = "dbQCrosstab"
		Case 96
			s = "dbQDDL"
		Case 32
			s = "dbQDelete"
		Case 80
			s = "dbQMakeTable"
		Case 224
			s = "dbQProcedure"
		Case 0
			s = "dbQSelect"
		Case 128
			s = "dbQSetOperation"
		Case 144
			s = "dbQSPTBulk"
		Case 112
			s = "dbQSQLPassThrough"
		Case 48
			s = "dbQUdate"
		Case Else
			s = "Unknown type " & qdf.Type
	End Select

	queryDefType = s

End Function

