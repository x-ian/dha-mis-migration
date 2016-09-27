Attribute VB_Name = "RelinkTables"
Option Compare Database

Sub RelinkAllTables()
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
    Col.Add "code_hfacility_GPS"
    Col.Add "code_year_quarter"
    Col.Add "concept"
    Col.Add "concept_set"
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
    Col.Add "map_supply_item_cms_code"
    Col.Add "map_user"
    Col.Add "NMCP_DL20_export"
    Col.Add "NMCP_DL21_export"
    Col.Add "NMCP_DL26_export"
    Col.Add "NMCP_DL30_export"
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
    Col.Add "psm_ro_site"
    Col.Add "psm_stock_report"
    Col.Add "supply_item"
    Col.Add "supply_item_set"
    Col.Add "tblOrgUnit"

    Dim name As Variant

    On Error Resume Next

    For Each name In Col
        Available = CurrentDb.TableDefs(name).name
        If Err <> 3265 Then
            Call DeleteTable(CStr(name))
        End If
        Call AttachDSNLessTable(CStr(name), "dbo." & CStr(name), "IE11WIN7\SQLEXPRESS", "HIVData")
    Next
    
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
      
'ODBC;DRIVER=ODBC Driver 11 for SQL Server;SERVER=IE11WIN7\SQLEXPRESS;Trusted_Connection=Yes;APP=Microsoft® Windows® Operating System;DATABASE=HIVData2;;TABLE=dbo.art_accom

    If Len(stUsername) = 0 Then
        '//Use trusted authentication if stUsername is not supplied.
        stConnect = "ODBC;DRIVER=ODBC Driver 11 for SQL Server;SERVER=" & stServer & ";Trusted_Connection=Yes;APP=Microsoft® Windows® Operating System;DATABASE=" & stDatabase & ";;"
    Else
        '//WARNING: This will save the username and the password with the linked table information.
        stConnect = "ODBC;DRIVER=ODBC Driver 11 for SQL Server;SERVER=" & stServer & ";DATABASE=" & stDatabase & ";UID=" & stUsername & ";PWD=" & stPassword
    End If
    Set td = CurrentDb.CreateTableDef(stLocalTableName, dbAttachSavePWD, stRemoteTableName, stConnect)
    CurrentDb.TableDefs.Append td
    Debug.Print "added " & stRemoteTableName
    AttachDSNLessTable = True
    Exit Function

AttachDSNLessTable_Err:
    
    AttachDSNLessTable = False
    MsgBox "AttachDSNLessTable encountered an unexpected error: " & Err.Description

End Function

Sub DeleteTable(name As String)
         '**********
         ' Since the delete action will fail if the
         ' table is participating in any relation, first
         ' find and delete existing relations for table.
         '**********
         For Each rel In CurrentDb.Relations
            If rel.Table = name Or rel.ForeignTable = name Then
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

