Attribute VB_Name = "Module1"
Option Compare Database

' stolen from http://stackoverflow.com/questions/172895/table-creation-ddl-from-microsoft-access

Public Function TableCreateDDL(TableDef As TableDef) As String

         Dim fldDef As Field
         Dim FieldIndex As Integer
         Dim fldName As String, fldDataInfo As String
         Dim DDL As String
         Dim TableName As String

         TableName = TableDef.Name
         TableName = Replace(TableName, " ", "_")
         DDL = "create table " & TableName & "(" & vbCrLf
         With TableDef
            For FieldIndex = 0 To .Fields.Count - 1
               Set fldDef = .Fields(FieldIndex)
               With fldDef
                  fldName = .Name
                  fldName = Replace(fldName, " ", "_")
                  Select Case .Type
                     Case dbBoolean
                        fldDataInfo = "nvarchar2"
                     Case dbByte
                        fldDataInfo = "number"
                     Case dbInteger
                        fldDataInfo = "number"
                     Case dbLong
                        fldDataInfo = "number"
                     Case dbCurrency
                        fldDataInfo = "number"
                     Case dbSingle
                        fldDataInfo = "number"
                     Case dbDouble
                        fldDataInfo = "number"
                     Case dbDate
                        fldDataInfo = "date"
                     Case dbText
                        fldDataInfo = "nvarchar2(" & Format$(.Size) & ")"
                     Case dbLongBinary
                        fldDataInfo = "****"
                     Case dbMemo
                        fldDataInfo = "****"
                     Case dbGUID
                        fldDataInfo = "nvarchar2(16)"
                  End Select
               End With
               If FieldIndex > 0 Then
               DDL = DDL & ", " & vbCrLf
               End If
               DDL = DDL & "  " & fldName & " " & fldDataInfo
               Next FieldIndex
         End With
         DDL = DDL & ");"
         TableCreateDDL = DDL
End Function


Sub ExportAllTableCreateDDL()

    Dim lTbl As Long
    Dim dBase As Database
    Dim Handle As Integer

    Set dBase = CurrentDb

    Handle = FreeFile

    Open "c:\Users\IEUser\Desktop\TableCreateDDL.txt" For Output Access Write As #Handle

    For lTbl = 0 To dBase.TableDefs.Count - 1
         'If the table name is a temporary or system table then ignore it
        If Left(dBase.TableDefs(lTbl).Name, 1) = "~" Or _
        Left(dBase.TableDefs(lTbl).Name, 4) = "MSYS" Then
             '~ indicates a temporary table
             'MSYS indicates a system level table
        Else
          Print #Handle, TableCreateDDL(dBase.TableDefs(lTbl))
        End If
    Next lTbl
    Close Handle
    Set dBase = Nothing
End Sub

