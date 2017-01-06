Option Compare Database

Public Sub main_1a_hardcodeParamsIntoQueries()

    Dim obj As AccessObject
    Dim qdf As DAO.QueryDef

    Dim t As String
    
    Debug.Print "Start"
    
    For Each obj In Application.CurrentData.AllQueries
        Set qdf = CurrentDb.QueryDefs(obj.name)
        
        If qdf.name = "psm_regim_consum_growth" Then GoTo END_FOR_2
        If qdf.name = "art_clinic_old" Then GoTo END_FOR_2
        
        'Debug.Print qdf.name
        t = qdf.name
        
        If qdf.Type = dbQSelect Or qdf.Type = dbQCrosstab Or qdf.Type = dbQSetOperation Then   'And qdf.name = "data_entry_stats_page2" Then
            If hasParams(qdf) And InStr(1, qdf.name, "_original") = 0 Then
        
                If originalQueryExists(qdf.name) Then
                    ' shouldnt happen, seems to be called multiple times
                    t = t & " 11 stopping as original query already exists "
                    Debug.Print t
                    Exit Sub
                Else
                    Dim sql As String
                
                    sql = qdf.sql
                    
                    ' if contains PARAMETERS def, get rid of it
                    If InStr(1, sql, "PARAMETERS") = 1 Then
                        sql = Mid(sql, InStr(sql, ";") + 1, Len(sql))
                    End If
    
                    ' replace all known parameters
                    sql = Replace(sql, "[Enter year]", "2016")
                    sql = Replace(sql, "[Enter quarter]", "1")
                    sql = Replace(sql, "[Enter distribution round]", "29")
                    sql = Replace(sql, "[Enter distribution round number]", "29")
                    sql = Replace(sql, "[Enter first day of report interval]", "#01/01/2016#")
                    sql = Replace(sql, "[Enter last day of report interval]", "#31/03/2016#")
                    sql = Replace(sql, "[forms]![art_clinic_obs_v8]![year_quarter_id]", "65")
                    sql = Replace(sql, "art_clinic_obs_supervisor.year_quarter_id", "65")
                    sql = Replace(sql, "[Forms]![psm_dist_round]![ref_year_quarter_id]", "65")
                    sql = Replace(sql, "[Allocations between: start date]", "#01/01/2016#")
                    sql = Replace(sql, "[Allocations between: end date]", "#31/03/2016#")
                    sql = Replace(sql, "Forms!psm_dist_item_check_frm!psm_dist_round_id", "29")
                    sql = Replace(sql, "[Enter interval start date]", "#01/01/2016#")
                    sql = Replace(sql, "[Enter interval end date]", "#31/03/2016#")
                    sql = Replace(sql, "[Enter first team number]", "1")
                    sql = Replace(sql, "[Enter last team number]", "10")
                    sql = Replace(sql, "[Enter opening date]", "#01/01/2016#")
                    sql = Replace(sql, "[Enter reference date]", "#01/01/2016#")
                    sql = Replace(sql, "[Enter supply interval start date]", "#01/01/2016#")
                    sql = Replace(sql, "[Forms]![psm_dist_round]![consum_date]", "#01/01/2016#")
                    sql = Replace(sql, "[Forms]![psm_dist_round]![dist_round]", "29")
                    sql = Replace(sql, "[Forms]![psm_dist_round]![ref_year_quarter_id]", "65")
                    sql = Replace(sql, "[forms]![art_clinic_obs_v8]![year_quarter_id]", "65")
                    sql = Replace(sql, "art_clinic_obs_supervisor.year_quarter_id", "65")
                    sql = Replace(sql, "forms!psm_dist_round!dist_round", "29")
                    sql = Replace(sql, "Forms!psm_site_stock_consumpt_getval!supply_item_id_getval", "1")
                    sql = Replace(sql, "Forms!psm_relocate!supply_item_id", "1")
                    sql = Replace(sql, "[Enter start date for this round of supervision]", "#01/01/2016#")
                    sql = Replace(sql, "Forms!art_clinic_obs_v8!ID", "1")
                    sql = Replace(sql, "Forms!art_clinic_obs_v8!year_quarter_id", "26")
                    sql = Replace(sql, "[Forms]![psm_site_stock_consumpt_getval]![supply_item_id_getval]", "1")
                    sql = Replace(sql, "Forms!psm_site_stock_consumpt_getval!year_quarter_id_getval", "26")
                    sql = Replace(sql, "Forms!psm_dist_round!ref_year_quarter_id", "65")
                    sql = Replace(sql, "[Forms].[psm_dist_round].[ref_year_quarter_id]", "65")
                    sql = Replace(sql, "Forms!psm_relocate_sheet!psm_relocate_ID_select", "1004")
                    sql = Replace(sql, "[Forms]![psm_ro_sheet]![RO_item_set]", "1")
                    sql = Replace(sql, "Forms!supply_item_set!supply_group_select", "34")
                    sql = Replace(sql, "[Forms]![supply_item_set].[supply_group_select]", "34")
                    sql = Replace(sql, "Forms!supply_item_set.version_set_select", "194")
                
                    sql = Replace(sql, "forms!art_clinic_obs_v8.ID", "18100")
                    sql = Replace(sql, "Forms!art_clinic_obs_v8_startup!art_clinic_obs_id_select", "18100")
                    sql = Replace(sql, "Forms!art_clinic_obs_v8_startup!hdepartment_ID_select", "804")
                    sql = Replace(sql, "Forms!art_clinic_obs_v8!obs_dim_set!obs_set!obs_dimensions_ID_parent", "1")
                    sql = Replace(sql, "Forms!art_clinic_obs_v8!obs_dim_set!obs_dimensions_id_parent", "1")
                    'sql = Replace(sql, "", "")
                    
                    ' whitelisting known ok queries
        '            If qdf.name = "art_obs_clinic_odo" _
         '               Or qdf.name = "" Then
          '          Else
                    
           '         If InStr(InStr(1, sql, "FROM") + 1, sql, "[") > 0 Then
                        ' still contains unknown params, ignore it
                 '       t = t & " 22 Unknown parameter for query, skipping " ' & sql
            '            Debug.Print t
             '           GoTo END_FOR_2
              '      End If
            
               '     End If
                    
                    ' make copy of query
                    t = t & " 33 making copy of query with known params"
                    On Error Resume Next
                    CurrentDb.CreateQueryDef obj.name & "_original", qdf.sql
                    If Err.Number <> 0 Then
                        Err.Clear
                        t = t & " XX error copying query"
                        Debug.Print t
                    End If

                    qdf.sql = sql
                End If
                'Debug.Print t
                'Exit Sub
            Else
                t = t & " 44 doing nothing with select query without params "
            End If
        Else
            t = t & " 55 doing nothing with non-select query"
        End If
        
END_FOR_2:
    Debug.Print t
    Next obj
    
End Sub

Public Sub main_1b_checkAndRestoreOriginalParamQueryNowWithoutParams()

    ' now scan again all converted queries for nested queries and parameters
    For Each obj In Application.CurrentData.AllQueries
        Set qdf = CurrentDb.QueryDefs(obj.name)

        ' check for every _original query if it still requires a parameter
        If InStr(1, qdf.name, "_original") > 0 Then 'And originalQueryExists(qdf.name) Then
        If qdf.name = "art_death_now_prev_neg_original" Then
            Debug.Print qdf.name
            Debug.Print ""
        End If
            If hasParams(qdf) Then
                ' all good, still requires parameter, keep it like that
                Debug.Print qdf.name & " still requires params"
            Else
                ' looks like now the orginal query doesnt need a parameter anymore
                ' assuming that only a nested query had params and restoring original query as it is
                Debug.Print qdf.name & " restore original query"
                DoCmd.DeleteObject acQuery, Left(qdf.name, InStr(qdf.name, "_original") - 1)
                DoCmd.Rename Left(qdf.name, InStr(qdf.name, "_original") - 1), acQuery, obj.name
            End If
        End If
    Next obj

End Sub

Public Sub main_2_restoreQueriesWithParams()

    Dim obj As AccessObject
    Dim qdf As DAO.QueryDef

    Dim t As String
    
    For Each obj In Application.CurrentData.AllQueries
        If originalQueryExists(obj.name) Then
            Set qdf = CurrentDb.QueryDefs(obj.name)
            Debug.Print "restore original query and delete param-less query " & obj.name
            DoCmd.DeleteObject acQuery, obj.name
            DoCmd.Rename obj.name, acQuery, obj.name & "_original"
        End If
    Next obj
    
End Sub

Public Sub main_3_CsvExportQueries()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentData
 
    Dim intFile As Integer
    Dim dir As String
    Dim strFile As String
   ' strFile = "C:\Users\IEUser\Desktop\export\File.txt"
    dir = CurrentProject.Path & "\export"
    If Not DirExists(dir) Then
        MkDir (dir)
    End If
    strFile = dir & "\_file.txt"
    intFile = FreeFile
    Open strFile For Output As #intFile

    Dim t As Single
    Dim t2 As Single
    Dim t3 As Single
    
    t2 = Timer
    On Error GoTo ERROR_HANDLER

    For Each obj In dbs.AllQueries
    
        Debug.Print obj.name
        
        Dim qdf As DAO.QueryDef
        Set qdf = CurrentDb.QueryDefs(obj.name)
        'Print #intFile, "__ " & qdf.name & " " & queryDefType(qdf.Type)
        
        Dim params As Boolean
        params = False
        
        For Each P In qdf.Parameters
            If Err.Number <> 0 Then
                Print #intFile, ";yy;Error occured;" & obj.name & ";" & queryDefType(qdf.Type)
                Err.Clear
                GoTo END_FOR
            End If
            'Debug.Print p.Name & " " & p.Value
            params = True
        Next P
            
        If Not params Then
 '          If qdf.Type <> dbQDelete And qdf.Type <> dbQAction And qdf.Type <> dbQAppend And qdf.Type <> dbQUpdate Then ' Or qdf.Type = 128 Or qdf.Type = 16 Then
           If qdf.Type = dbQSelect Or qdf.Type = dbQSetOperation Or qdf.Type = dbQCrosstab Then
                If obj.name = "art_sched_set_rank" Or obj.name = "calc_test" _
                    Or obj.name = "calc_test_result" Or obj.name = "chk_outc" _
                    Or obj.name = "chk_outc_new" Or obj.name = "chk_outc_new_outlier" _
                    Or obj.name = "chk_surv_reggroup" Or obj.name = "pmtct_data" _
                    Or obj.name = "htc_set_site_district_greater_allq" Or obj.name = "chk_surv_reg" _
                    Or obj.name = "chk_surv_reg_diff" Or obj.name = "chk_surv_de_reg" _
                    Or obj.name = "art_clinic_obs_odo_rank" Or obj.name = "chk_surv_de_reg" _
                    Or obj.name = "report_oi_quart_cum_expos_result_zone" Or obj.name = "chk_art_cohort_outcome_prev_now" _
                    Then
                    Debug.Print "Skip calc_test and calc_test_result"
                    Print #intFile, "xx;Skip;" & obj.name & ";" & queryDefType(qdf.Type)
                Else
                    t = Timer
                    DoCmd.TransferText acExportDelim, , obj.name, dir & "\" & obj.name & ".csv", True, , 65001
                    If originalQueryExists(obj.name) Then
                        Print #intFile, "ab;(param-less query derived from param-query);" & obj.name & ";" & queryDefType(qdf.Type) & ";" & Round((Timer - t), 2) & ";" & "secs;" & Replace(qdf.sql, vbCrLf, " ")
                    Else
                        Print #intFile, "aa;(param-less query);" & obj.name & ";" & queryDefType(qdf.Type) & ";" & Round((Timer - t), 2) & ";" & "secs;" & Replace(qdf.sql, vbCrLf, " ")
                    End If
                End If

            Else
                'DoCmd.SetWarnings False
                't3 = Timer
                'DoCmd.OpenQuery (obj.Name)
                'DoCmd.SetWarnings True
                'Print #intFile, "dd (param-less query)" & vbTab & obj.Name & vbTab & queryDefType(qdf.Type) & vbTab & Round((Timer - t3), 2) & vbTab & "secs" & vbTab & Replace(qdf.SQL, vbCrLf, " ")
                Print #intFile, "cc;ignoring non-scriptable query;" & obj.name & ";" & queryDefType(qdf.Type)
            End If
        Else
            Print #intFile, "bb;ignoring query with params;" & obj.name & ";" & queryDefType(qdf.Type)
        End If

END_FOR:
    GoTo END_FOR_REAL
    
ERROR_HANDLER:
        Print #intFile, "zz;error during call;" & obj.name & ";" & queryDefType(qdf.Type)
    Resume END_FOR_REAL
    
END_FOR_REAL:
    Next obj
        
    Debug.Print "total time: " & Round((Timer - t2), 2)
    Close #intFile
End Sub

Public Sub main_4_printAllParams()

    Dim obj As AccessObject
    Dim qdf As DAO.QueryDef
    Dim params As Boolean
    
    Dim t As String
    
    params = False
    
    Debug.Print "start"
    
    For Each obj In Application.CurrentData.AllQueries
        Set qdf = CurrentDb.QueryDefs(obj.name)
        
        'If qdf.name = "psm_regim_consum_growth" Then GoTo END_FOR_2
        On Error Resume Next
        
        t = qdf.name & ","
        If originalQueryExists(qdf.name) Then 'qdf.Type = dbQSelect Then
            For Each P In qdf.Parameters
                t = t + P.name & "," 'Debug.Print p.name
                params = True
                If P.name <> "[Enter quarter]" And P.name <> "[Enter year]" And P.name <> "[Enter distribution round]" Then
                'Debug.Print p.name
                End If
            Next P
         '   If hasParams(qdf) And InStr(1, qdf.sql, "PARAMETERS") = 1 Then
          '       Debug.Print qdf.name & "," & Left(qdf.sql, InStr(qdf.sql, ";"))
           ' End If
           If params Then
            Debug.Print t
            End If
            params = False
        End If
    Next obj
End Sub

Private Function hasParams(qdf)
    On Error Resume Next
    For Each P In qdf.Parameters
        If Err.Number <> 0 Then
            Debug.Print "error looking into query : parameter " & qdf.name & " : " & P.name & " Errnumber : " & Err.Number
            Err.Clear
        End If
        hasParams = True
        Exit Function
    Next P
    hasParams = False
End Function

Private Function originalQueryExists(name)
    On Error Resume Next
    Set qdf = CurrentDb.QueryDefs(name & "_original")
    If Err.Number <> 0 Then
        Err.Clear
        originalQueryExists = False
        Exit Function
    End If
    originalQueryExists = True
End Function
 
Private Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
End Function

Private Function queryDefType(typ As Integer) As String

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


