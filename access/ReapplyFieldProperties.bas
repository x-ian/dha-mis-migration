Option Compare Database

Function isOdbcLinkedTable(table As String) As Boolean
    Dim td As TableDef
    Dim db As Database
        
    Set db = CurrentDb
    Set td = db.TableDefs(table)
    
'    isOdbcLinkedTable = startsWith(td.Connect, "ODBC;")
    isOdbcLinkedTable = startsWith(td.Connect, ";DATABASE")
End Function

Public Function startsWith(str As String, prefix As String) As Boolean
    startsWith = Left(str, Len(prefix)) = prefix
End Function

Sub main_print_frontend_field_properties()

    Dim td As TableDef
    Dim fd As field
    Dim p As property
    
    'Dim tbl As DAO.TableDef

    For Each td In CurrentDb.TableDefs
        If isOdbcLinkedTable(td.name) Then
        For Each fd In td.Fields
            For Each p In fd.Properties
                If p.name = "Caption" Or p.name = "DecimalPlace" Or p.name = "Format" _
                    Or p.name = "IMESentenceMode" Or p.name = "ShowDatePicker" Or p.name = "TextAlign" _
                    Or p.name = "TextFormat" _
                    Or p.name = "DisplayControl" Or p.name = "RowSourceType" Or p.name = "RowSource" _
                    Or p.name = "BoundColumn" Or p.name = "ColumnCount" Or p.name = "ColumnHeads" _
                    Or p.name = "ColumnWidths" Or p.name = "ListRows" Or p.name = "ListWidth" _
                    Or p.name = "LimitToList" Or p.name = "AllowMultipleValues" Or p.name = "AllowValueListEdits" _
                    Or p.name = "ListItemEditForm" Or p.name = "ShowOnlyRowSource" Then
                    
                    Dim name As String
                    Dim typ As String
                    Dim value As String
                    name = ""
                    typ = ""
                    value = ""
                    name = p.name
                    typ = p.Type
                    value = p.value
                    
                    Debug.Print td.name & "    " & fd.name & "    " & "Name:" & "    " & name & "    " & " Type:" & "    " & typ & "    " & "Value:" & "    " & value
                End If
            Next p
        Next fd
        'GoTo done
        End If
    Next td
done:
End Sub

Sub main_load_frontend_field_properties()

    Dim rs As Recordset
    Dim table As String
    Dim field As String
    Dim property As String
    Dim value As String
    Dim typ As String
    
    Set rs = CurrentDb.OpenRecordset("field_properties")
    'populate the table
    rs.MoveLast
    rs.MoveFirst

    Do While Not rs.EOF
        Debug.Print rs!ID
        table = rs!table
        field = rs!field
        typ = getTypeForProperty("Caption")
            If Not IsNull(rs!Caption) Then
                Call applyProperty(table, field, "Caption", rs!Caption, getTypeForProperty("Caption"))
            End If
            If Not IsNull(rs!DecimalPlace) Then
            End If
            If Not IsNull(rs!Format) Then
            End If
            If Not IsNull(rs!IMESentenceMode) Then
            End If
            If Not IsNull(rs!ShowDatePicker) Then
            End If
            If Not IsNull(rs!TextAlign) Then
            End If
            If Not IsNull(rs!TextFormat) Then
            End If
            If Not IsNull(rs!DisplayControl) Then
            End If
            If Not IsNull(rs!RowSourceType) Then
            End If
            If Not IsNull(rs!RowSource) Then
            End If
            If Not IsNull(rs!BoundColumn) Then
            End If
            If Not IsNull(rs!ColumnCount) Then
            End If
            If Not IsNull(rs!ColumnHeads) Then
            End If
            If Not IsNull(rs!ColumnWidths) Then
            End If
            If Not IsNull(rs!ListRows) Then
            End If
            If Not IsNull(rs!ListWidth) Then
            End If
            If Not IsNull(rs!LimitToList) Then
            End If
            If Not IsNull(rs!AllowMultipleValues) Then
            End If
            If Not IsNull(rs!AllowValueListEdits) Then
            End If
            If Not IsNull(rs!ListItemEditForm) Then
            End If
            If Not IsNull(rs!ShowOnlyRowSource) Then
            End If
        typ = getTypeForProperty(rs!Type)
        Call applyProperty(table, field, property, value, typ)
        rs.MoveNext
    Loop
End Sub

Sub applyProperty(table As String, field As String, py As String, value As String, typ As String)

    Dim db As Database
    Dim f As Dao.field
    Dim p As Dao.property
    
    Set db = CurrentDb
    Set f = db.TableDefs(table).Fields(field)
    
    'On Error Resume Next
    Set p = f.Properties(py)
    f.Properties.Delete py
    On Error GoTo 0
    
    Set p = f.CreateProperty(py, typ, value)
    f.Properties.Append p
    
End Sub

Function getTypeForProperty(property As String)
    Dim typ As String
        Select Case property
            Case "Caption"
                typ = "12"
            Case "DecimalPlace"
                typ = "2"
            Case "Format"
                typ = "10"
            Case "IMESentenceMode"
                typ = "2"
            Case "ShowDatePicker"
                typ = "3"
            Case "TextAlign"
                typ = "2"
            Case "TextFormat"
                typ = "2"
            Case "DisplayControl"
                typ = "3"
            Case "RowSourceType"
                typ = "10"
            Case "RowSource"
                typ = "12"
            Case "BoundColumn"
                typ = "3"
            Case "ColumnCount"
                typ = "3"
            Case "ColumnHeads"
                typ = "1"
            Case "ColumnWidths"
                typ = "3"
            Case "ListRows"
                typ = "3"
            Case "ListWidth"
                typ = "10"
            Case "LimitToList"
                typ = "1"
            Case "AllowMultipleValues"
                typ = "1"
            Case "AllowValueListEdits"
                typ = "1"
            Case "ListItemEditForm"
                typ = "12"
            Case "ShowOnlyRowSource"
                typ = "1"
            Case Else
                typ = "UNKNOWN"
                Debug.Print "Unknown property " & property
        End Select
    getTypeForProperty = typ
End Function

Sub main_save_frontend_field_properties()

    Dim td As TableDef
    Dim fd As field
    Dim p As property
    
  '  With CurrentDb
  '      .Execute "DELETE field_properties.* FROM field_properties;"
  '  End With
    
    For Each td In CurrentDb.TableDefs
        If isOdbcLinkedTable(td.name) Then
        For Each fd In td.Fields
            For Each p In fd.Properties
                If p.name = "Caption" Or p.name = "DecimalPlace" Or p.name = "Format" _
                    Or p.name = "IMESentenceMode" Or p.name = "ShowDatePicker" Or p.name = "TextAlign" _
                    Or p.name = "TextFormat" _
                    Or p.name = "DisplayControl" Or p.name = "RowSourceType" Or p.name = "RowSource" _
                    Or p.name = "BoundColumn" Or p.name = "ColumnCount" Or p.name = "ColumnHeads" _
                    Or p.name = "ColumnWidths" Or p.name = "ListRows" Or p.name = "ListWidth" _
                    Or p.name = "LimitToList" Or p.name = "AllowMultipleValues" Or p.name = "AllowValueListEdits" _
                    Or p.name = "ListItemEditForm" Or p.name = "ShowOnlyRowSource" Then
                    
                    Call createOrUpdateProperty(td.name, fd.name, p.name, p.value)
                End If
            Next p
        Next fd
        End If
    Next td
End Sub

Private Sub createOrUpdateProperty(table As String, field As String, property As String, value As String)

    Dim sql As String
    Dim rs As Recordset
    Dim v2 As String
    
    v2 = Replace(value, """", """""")
    'On Error Resume Next
    ' find existing record and update
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM field_properties where [table] = """ & table & """ AND field = """ & field & """", dbOpenDynaset, dbSeeChanges)
    If rs.EOF Then
        ' create a new one
        sql = "INSERT INTO field_properties ([Table], Field, " & property & ") VALUES (""" & table & """, """ & field & """, """ & v2 & """);"
        
        Debug.Print sql
        CurrentDb.Execute sql
    Else
        ' update existing one
        sql = "UPDATE field_properties SET " & property & " = """ & v2 & """ WHERE [Table] = """ & table & """ AND Field = """ & field & """;"
        Debug.Print sql
        CurrentDb.Execute sql
    End If

End Sub





