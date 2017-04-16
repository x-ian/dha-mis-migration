Option Compare Database
Option Explicit

Public DhaUser As String
Public DbUser As String
Public DbServer As String
Public DbInstance As String
Public DbTrustedConnection As String

Public Function runningOnAccessRuntime() As Boolean
    'runningOnAccessRuntime = SysCmd(acSysCmdRuntime)
    runningOnAccessRuntime = True
End Function

Public Function CurrentUser()
    CurrentUser = DhaUser
End Function

Public Function CurrentDbUser()
    CurrentDbUser = DbUser
End Function

Public Function CurrentServer()
    CurrentServer = DbServer
End Function

Public Function CurrentDbInstance()
    CurrentDbInstance = DbInstance
End Function

Public Function CurrentDbTrustedConnection()
    CurrentDbTrustedConnection = DbTrustedConnection
End Function

Sub main_RelinkAllTables(Server As String, Db As String, Optional stUsername As String, Optional stPassword As String)
    Dim Col As Collection
    Set Col = New Collection
    Col.Add "art_accom"
    Col.Add "art_clinic_obs"
    Col.Add "art_coh_reg_target"
    Col.Add "art_drug_stocks"
    Col.Add "art_person"
    Col.Add "art_sched_day"
    Col.Add "art_sched_person"
    Col.Add "art_sched_set"
    Col.Add "art_sched_site"
    Col.Add "art_sched_team"
    Col.Add "art_staff"
    Col.Add "art_staff_obs"
    Col.Add "art_supervisor"
    Col.Add "code_hdepartment"
    Col.Add "code_hfacility"
    Col.Add "code_year_quarter"
    Col.Add "concept"
    Col.Add "concept_set"
    Col.Add "field_properties"
    Col.Add "htc_obs"
    Col.Add "htc_obs_dimensions"
    Col.Add "htc_person"
    Col.Add "htc_person_obs"
    Col.Add "htc_prov"
    Col.Add "htc_site_obs"
    Col.Add "htc_site_person_obs"
    Col.Add "htc_supervisor"
    Col.Add "map_regimen_supply"
    Col.Add "map_regimen_supply_rule"
    Col.Add "map_scm_site"
    Col.Add "map_user"
    Col.Add "obs"
    Col.Add "obs_dimensions"
    Col.Add "population"
    Col.Add "pop_district"
    Col.Add "pop_map"
    Col.Add "pop_sex_district_hiv"
    Col.Add "psm_dist_batch"
    Col.Add "psm_dist_item"
    Col.Add "psm_dist_round"
    Col.Add "psm_DL_item"
    Col.Add "psm_DL_sheet"
    Col.Add "psm_DL_site"
    Col.Add "psm_relocate"
    Col.Add "psm_relocate_old"
    Col.Add "psm_ro_item"
    Col.Add "psm_ro_sheet"
    Col.Add "psm_stock_report"
    Col.Add "supply_item"
    Col.Add "supply_item_set"
    Col.Add "tblOrgUnit"

    Dim name As Variant
    Dim available As Variant
    Dim localName As String
    Dim Db2 As Database
    Dim td As TableDef
    Set Db2 = CurrentDb
    On Error Resume Next

    For Each name In Col
        localName = CStr(name)
        If name = "concept" Then
            Set td = Db2.TableDefs(localName)
            If IsObject(td) Then
                ' need to break it into 2 conditions as access stupidily evaluates always both expressions, even if the first one is already false
                If td.Connect = "" Then
                    ' this is already a local table, not a linked one
                    ' dataentry frontend: make sure that local concept table remains
                    ' and MSSQL concept table is linked as concept_live
                    localName = "concept_live"
                End If
            End If
        End If
        If name = "concept_set" Then
            Set td = Db2.TableDefs(localName)
            If IsObject(td) Then
                ' need to break it into 2 conditions as access stupidily evaluates always both expressions, even if the first one is already false
                If td.Connect = "" Then
                    ' this is already a local table, not a linked one
                    ' dataentry frontend: make sure that local concept_set table remains
                    ' and MSSQL concept_set table is linked as concept_set_live
                    localName = "concept_set_live"
                End If
            End If
        End If
        
        available = CurrentDb.TableDefs(localName).name
        If Err <> 3265 Then
            Call DeleteTable(localName)
        End If
        If Not AttachDSNLessTable(localName, "dbo." & CStr(name), Server, Db, stUsername, stPassword) Then
'        Call AttachDSNLessTable(localName, "dbo." & CStr(name), "NDX-HAD1\DHA_MIS", "HIVData9", "sa", "dhamis@2016")
            MsgBox ("Error during (re-) linking. Wrong login details? Frontend not usuable; aborting")
            GoTo ERROR
        End If
    Next
    
    Call changeOdbcSourceForSqlPassthroughQueries(Server, Db, stUsername, stPassword)
    
    Call updateConnectionDetails
    Call main_load_frontend_field_properties
    
    Exit Sub
ERROR:
    
End Sub

Public Function databaseAccess() As Boolean
    On Error GoTo ERROR
    Dim qdf As QueryDef
    Dim Rs As Recordset
    Set qdf = CurrentDb.QueryDefs("current_user")
    Set Rs = qdf.OpenRecordset(dbOpenDynaset, dbSeeChanges)
    databaseAccess = True
    Exit Function
ERROR:
    databaseAccess = False
End Function

Public Sub updateConnectionDetails()
    Dim qdf As QueryDef
    Dim Db As Database
    Dim Rs As Recordset
    
    On Error GoTo ERROR
    
    Set Db = CurrentDb
    Set qdf = CurrentDb.QueryDefs("current_user")
    Set Rs = qdf.OpenRecordset(dbOpenDynaset, dbSeeChanges)
    Rs.MoveFirst
    DbUser = Rs!Expr1000

    Set Rs = CurrentDb.OpenRecordset("SELECT [User] FROM map_user where ActiveDirectoryUser = """ & DbUser & """", dbOpenDynaset, dbSeeChanges)
    Rs.MoveFirst
    DhaUser = Rs!User
    
    Dim start, ende
    start = InStr(qdf.Connect, "SERVER=")
    ende = InStr(start + 7, qdf.Connect, ";")
    DbServer = Mid(qdf.Connect, start + 7, ende - (start + 7))
    
    start = InStr(qdf.Connect, "DATABASE=")
    ende = InStr(start + 9, qdf.Connect, ";")
    DbInstance = Mid(qdf.Connect, start + 9, ende - (start + 9))
    
    start = InStr(qdf.Connect, "Trusted_Connection=Yes")
    If start > 0 Then
        DbTrustedConnection = "AD Login"
    Else
        DbTrustedConnection = ""
    End If
    Exit Sub
ERROR:
    MsgBox "Error establishing database access." & vbCrLf & "Check SQL Server login, local map_user table, and/or server availability" & vbCrLf & DbUser & " " & DhaUser & " " & DbServer & " " & DbInstance & " " & DbTrustedConnection
End Sub

Sub main_RelinkAllTablesHardCoded()
    'main_RelinkAllTables "NDX-HAD1\DHA_MIS", "HIVData9", "sa", "dhamis@2016")
    main_RelinkAllTables "IE11WIN7\SQLEXPRESS", "HIVData3"
End Sub


'//Name     :   AttachDSNLessTable
'//Purpose  :   Create a linked table to SQL Server without using a DSN
'//Parameters
'//     stLocalTableName: Name of the table that you are creating in the current database
'//     stRemoteTableName: Name of the table that you are linking to on the SQL Server database
'//     stServer: Name of the SQL Server that you are linking to
'//     stDatabase: Name of the SQL Server database that you are linking to
'//     stUsername: Name of the SQL Server user who can connect to SQL Server, leave blank to use a Trusted Connection
'//     stPassword: SQL Server user password
Function AttachDSNLessTable(stLocalTableName As String, stRemoteTableName As String, stServer As String, stDatabase As String, Optional stUsername As String, Optional stPassword As String)
    On Error GoTo AttachDSNLessTable_Err
    Dim td As TableDef
    Dim stConnect As String
    
    For Each td In CurrentDb.TableDefs
        If td.name = stLocalTableName Then
            CurrentDb.TableDefs.Delete stLocalTableName
        End If
    Next
      
      
    stConnect = createOdbcConnectString(stServer, stDatabase, stUsername, stPassword)
    Set td = CurrentDb.CreateTableDef(stLocalTableName, dbAttachSavePWD, stRemoteTableName, stConnect)
    CurrentDb.TableDefs.Append td
    Debug.Print "added " & stRemoteTableName
    AttachDSNLessTable = True
    Exit Function

AttachDSNLessTable_Err:
    
    AttachDSNLessTable = False
    MsgBox "AttachDSNLessTable encountered an unexpected error: " & Err.Description

End Function

Private Sub changeOdbcSourceForSqlPassthroughQueries(Server As String, Db As String, Optional stUsername As String, Optional stPassword As String)

    Dim obj As AccessObject
    Dim qdf As DAO.QueryDef
    Dim odbc As String
    
    odbc = createOdbcConnectString(Server, Db, stUsername, stPassword)
    
    For Each obj In Application.CurrentData.AllQueries
        Set qdf = CurrentDb.QueryDefs(obj.name)
        If queryDefType(qdf.Type) = "dbQSQLPassThrough" Then
            qdf.Connect = odbc
        End If
    Next obj
    
End Sub

Sub DeleteTable(name As String)
    Dim rel As Relation
    
         '**********
         ' Since the delete action will fail if the
         ' table is participating in any relation, first
         ' find and delete existing relations for table.
         '**********
         For Each rel In CurrentDb.Relations
            If rel.table = name Or rel.ForeignTable = name Then
               Debug.Print name & " | " & rel.name
               CurrentDb.Relations.Delete rel.name
            End If
         Next rel
         '**********
         ' Now, we're ready to delete the table.
         '**********
         'docmd.SetWarnings False
         DoCmd.DeleteObject acTable, name
         Debug.Print name & " deleted"
         'docmd.SetWarnings True
End Sub

Private Function createOdbcConnectString(stServer As String, stDatabase As String, Optional stUsername As String, Optional stPassword As String)
'ODBC;DRIVER=ODBC Driver 11 for SQL Server;SERVER=IE11WIN7\SQLEXPRESS;Trusted_Connection=Yes;APP=Microsoft® Windows® Operating System;DATABASE=HIVData2;;TABLE=dbo.art_accom

    If Len(stUsername) = 0 Then
        '//Use trusted authentication if stUsername is not supplied.
        createOdbcConnectString = "ODBC;DRIVER=ODBC Driver 11 for SQL Server;SERVER=" & stServer & ";Trusted_Connection=Yes;APP=Microsoft® Windows® Operating System;DATABASE=" & stDatabase & ";;"
    Else
        '//WARNING: This will save the username and the password with the linked table information.
        createOdbcConnectString = "ODBC;DRIVER=ODBC Driver 11 for SQL Server;SERVER=" & stServer & ";DATABASE=" & stDatabase & ";UID=" & stUsername & ";PWD=" & stPassword
    End If
End Function

Private Function queryDefType(typ As Integer) As String

' https://msdn.microsoft.com/en-us/library/office/ff192931.aspx
Dim s As String

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
        s = "Unknown type " & typ
End Select

queryDefType = s

End Function


