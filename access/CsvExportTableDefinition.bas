Attribute VB_Name = "Module2"
Option Compare Database

' stolen/adapted from
'   http://stackoverflow.com/questions/172895/table-creation-ddl-from-microsoft-access
'   http://allenbrowne.com/func-06.html

Public Function TableCreateDefinition(TableDef As TableDef) As String
    Dim fldDef As Field
    Dim FieldIndex As Integer
    Dim fldName As String, fldDataInfo As String
    Dim DDL As String
    Dim TableName As String

    TableName = TableDef.Name
    DDL = TableName
    With TableDef
        For FieldIndex = 0 To .Fields.Count - 1
            Set fldDef = .Fields(FieldIndex)
            With fldDef
                fldName = .Name
                fldDataInfo = "," & FieldTypeName(fldDef)
            End With
            For Each p In fldDef.Properties
                If p.Name = "RowSource" Then
                    fldDataInfo = fldDataInfo & ",""Lookup" & " {" & p & "}"""
                    Exit For
                End If
            Next
        
            DDL = DDL & "," & fldName & fldDataInfo
        Next FieldIndex
    End With
    TableCreateDefinition = DDL
End Function

Sub ExportAllTableDefintions()
    Dim lTbl As Long
    Dim dBase As Database
    Dim Handle As Integer

    Set dBase = CurrentDb

    Handle = FreeFile

    Open "c:\Users\IEUser\Desktop\TableCreateDefinition.csv" For Output Access Write As #Handle

    For lTbl = 0 To dBase.TableDefs.Count - 1
         'If the table name is a temporary or system table then ignore it
        If Left(dBase.TableDefs(lTbl).Name, 1) = "~" Or _
        Left(dBase.TableDefs(lTbl).Name, 4) = "MSYS" Then
            '~ indicates a temporary table
            'MSYS indicates a system level table
        Else
            Print #Handle, TableCreateDefinition(dBase.TableDefs(lTbl))
        End If
    Next lTbl
    Close Handle
    Set dBase = Nothing
End Sub




Function FieldTypeName(fld As DAO.Field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn
End Function


