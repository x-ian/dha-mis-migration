Option Compare Database

Sub main_MigrateCodeForAllModules()
    
    For Each vbcomp In Application.VBE.ActiveVBProject.VBComponents
        Debug.Print vbcomp.Name
        If vbcomp.Name = "MigrateCodeToMssql" Then
            Debug.Print "Do not replace code in MigrateCodeToMssql"
        Else
            Debug.Print "Replace code in " & vbcomp.Name
            migrateCode (vbcomp.CodeModule)
        End If
    Next vbcomp
End Sub

Sub main_MigrateCodeForActiveCodePane()
    migrateCode Application.VBE.ActiveCodePane.CodeModule
End Sub

' requires reference to Microsoft Visual Basic for Application Extensibility
Private Sub migrateCode(CodeMod As VBIDE.CodeModule)
    
    ' DATA ENTRY DB
    
    'Form_art_clinic_obs_v8
    replaceLine CodeMod, _
        "    Set Rst = dbs.OpenRecordset(""code_year_quarter"", dbOpenDynaset, dbReadOnly)", _
        "    Set Rst = dbs.OpenRecordset(""code_year_quarter"", dbOpenDynaset, dbReadOnly + dbSeeChanges)"
    replaceLine CodeMod, _
        "Set rs = db.OpenRecordset(""code_hdepartment"", dbOpenDynaset)", _
        "Set rs = db.OpenRecordset(""code_hdepartment"", dbOpenDynaset, dbSeeChanges)"
    replaceLine CodeMod, _
        "Set Rs1 = db.OpenRecordset(""code_hdepartment"", dbOpenDynaset)", _
        "Set Rs1 = db.OpenRecordset(""code_hdepartment"", dbOpenDynaset, dbSeeChanges)"
    replaceLine CodeMod, _
        "Set Rs2 = db.OpenRecordset(""code_hfacility"", dbOpenDynaset)", _
        "Set Rs2 = db.OpenRecordset(""code_hfacility"", dbOpenDynaset, dbSeeChanges)"
    replaceLine CodeMod, _
        "Set Rst = dbs.OpenRecordset(""SELECT quarter_stopdate FROM code_year_quarter WHERE ID = "" & CodeYearQuarterID & "" ;"")", _
        "Set Rst = dbs.OpenRecordset(""SELECT quarter_stopdate FROM code_year_quarter WHERE ID = "" & CodeYearQuarterID & "" ;"")"
    'Form_htc_providerObs
    'to be manually replaced
    
    'form_obs_set
    ' to be manually replaced
    
    ' form_psm_ro_item_set
    '
    replaceLine CodeMod, _
        "Set Rst = dbs.OpenRecordset(""psm_ro_item"", dbOpenDynaset)", _
        "Set Rst = dbs.OpenRecordset(""psm_ro_item"", dbOpenDynaset, dbseechanges)"
        
    ' PubFunctions
    replaceLine CodeMod, _
        "Set RstCalc = dbs.OpenRecordset(""concept_set_calc_de"")", _
        "Set RstCalc = dbs.OpenRecordset(""concept_set_calc_de"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "Set RstObs = dbs.OpenRecordset(""report_select_obs_dimensions_ID_obs"")", _
        "Set RstObs = dbs.OpenRecordset(""report_select_obs_dimensions_ID_obs"", , dbseechanges)"
    replaceLine CodeMod, _
        "Set RstObsCalc = dbs.OpenRecordset(""obs_calc"")", _
        "Set RstObsCalc = dbs.OpenRecordset(""obs_calc"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "'Set RstObsDim = Dbs.OpenRecordset(""report_select_obs_dimensions_ID"")", _
        "'Set RstObsDim = Dbs.OpenRecordset(""report_select_obs_dimensions_ID"", , dbSeeChanges)"
    ' already present
    'replaceLine CodeMod, _
    '    "Set RstCalc = dbs.OpenRecordset(""concept_set_calc_de"", , dbSeeChanges)", _
    '    "Set RstCalc = dbs.OpenRecordset(""concept_set_calc_de"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "Set RstObsCalc = dbs.OpenRecordset(""obs_calc"", dbOpenTable)", _
        "Set RstObsCalc = dbs.OpenRecordset(""obs_calc"", dbOpenTable, dbSeeChanges)"
    replaceLine CodeMod, _
        "Set RstObsCalcSubgp = dbs.OpenRecordset(""SELECT sub_group FROM obs_calc GROUP BY sub_group"")", _
        "Set RstObsCalcSubgp = dbs.OpenRecordset(""SELECT sub_group FROM obs_calc GROUP BY sub_group"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "Set Rst1 = dbs.OpenRecordset(""report_select_obs_wide"")", _
        "Set Rst1 = dbs.OpenRecordset(""report_select_obs_wide"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "Set Rst2 = dbs.OpenRecordset(""calc_rule"")", _
        "Set Rst2 = dbs.OpenRecordset(""calc_rule"", , dbSeeChanges)"
        
        
        
    ' analyse DB
    ' form_psm_dist_round
    replaceLine CodeMod, _
        "Set Rst = Dbs.OpenRecordset(""SELECT MAX(dist_round) AS MaxRound FROM psm_dist_round ;"")", _
        "Set Rst = Dbs.OpenRecordset(""SELECT MAX(dist_round) AS MaxRound FROM psm_dist_round ;"", , dbSeeChanges)"
    ' form_psm_relo_item
    replaceLine CodeMod, _
        "    Set rs = db.OpenRecordset(""SELECT MAX(ID) AS MaxID FROM psm_relocate"")", _
        "    Set rs = db.OpenRecordset(""SELECT MAX(ID) AS MaxID FROM psm_relocate"", , dbSeeChanges)"
    ' form_psm_dist_round
    replaceLine CodeMod, _
        "Set Rst = Dbs.OpenRecordset(""SELECT MAX(dist_round) AS MaxRound FROM psm_dist_round ;"")", _
        "Set Rst = Dbs.OpenRecordset(""SELECT MAX(dist_round) AS MaxRound FROM psm_dist_round ;"", , dbSeeChanges)"
    ' form_psm_relo_transact
    replaceLine CodeMod, _
        "    Set rs = db.OpenRecordset(""SELECT MAX(ID) AS MaxID FROM psm_relocate"")", _
        "    Set rs = db.OpenRecordset(""SELECT MAX(ID) AS MaxID FROM psm_relocate"", , dbSeeChanges)"
    ' form_psm_relo_transact_item_review, already present
    'replaceLine CodeMod, _
    '    "    Set rs = db.OpenRecordset("SELECT MAX(ID) AS MaxID FROM psm_relocate")", _
    '    ""
    ' form_psm_relocate
    replaceLine CodeMod, _
        "    Set rs = db.OpenRecordset(""SELECT MAX(ID) AS MaxID FROM psm_relocate"")", _
        "    Set rs = db.OpenRecordset(""SELECT MAX(ID) AS MaxID FROM psm_relocate"", , dbSeeChanges)"
    ' form_psm_relocate_relo, unknown what to do
    'replaceLine CodeMod, _
    '    "    Set rs = Qd.OpenRecordset()", _
    '    "    Set rs = Qd.OpenRecordset()"
    ' form_psm_ro_item_set
    replaceLine CodeMod, _
        "Set Rst = Dbs.OpenRecordset(""psm_ro_item"", dbOpenDynaset)", _
        "Set Rst = Dbs.OpenRecordset(""psm_ro_item"", dbOpenDynaset, dbSeeChanges)"
    ' Module1
    replaceLine CodeMod, _
        "    Set rstNewRefs = Dbs.OpenRecordset(""ANC_samp"", dbOpenDynaset)", _
        "    Set rstNewRefs = Dbs.OpenRecordset(""ANC_samp"", dbOpenDynaset, dbSeeChanges)"
    ' PubFunctions
    replaceLine CodeMod, _
        "Set RstCalc = Dbs.OpenRecordset(""concept_set_calc_de"")", _
        "Set RstCalc = Dbs.OpenRecordset(""concept_set_calc_de"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "Set RstObs = Dbs.OpenRecordset(""report_select_obs_dimensions_ID_obs"")", _
        "Set RstObs = Dbs.OpenRecordset(""report_select_obs_dimensions_ID_obs"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "Set RstObsCalc = Dbs.OpenRecordset(""obs_calc"")", _
        "Set RstObsCalc = Dbs.OpenRecordset(""obs_calc"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "'Set RstObsDim = Dbs.OpenRecordset(""report_select_obs_dimensions_ID"")", _
        "'Set RstObsDim = Dbs.OpenRecordset(""report_select_obs_dimensions_ID"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "Set RstCalc = Dbs.OpenRecordset(""concept_set_calc_de"")", _
        "Set RstCalc = Dbs.OpenRecordset(""concept_set_calc_de"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "Set RstObsCalc = Dbs.OpenRecordset(""obs_calc"", dbOpenTable)", _
        "Set RstObsCalc = Dbs.OpenRecordset(""obs_calc"", dbOpenTable, dbSeeChanges)"
    replaceLine CodeMod, _
        "Set RstObsCalcSubgp = Dbs.OpenRecordset(""SELECT sub_group FROM obs_calc GROUP BY sub_group"")", _
        "Set RstObsCalcSubgp = Dbs.OpenRecordset(""SELECT sub_group FROM obs_calc GROUP BY sub_group"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "Set Rst1 = Dbs.OpenRecordset(""report_select_obs_wide"")", _
        "Set Rst1 = Dbs.OpenRecordset(""report_select_obs_wide"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "Set Rst2 = Dbs.OpenRecordset(""calc_rule"")", _
        "Set Rst2 = Dbs.OpenRecordset(""calc_rule"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "Set Rst = Dbs.OpenRecordset(""report_levels_concept_name_obs_dimensions"")", _
        "Set Rst = Dbs.OpenRecordset(""report_levels_concept_name_obs_dimensions"", , dbSeeChanges)"
    replaceLine CodeMod, _
        "", _
        ""
    replaceLine CodeMod, _
        "", _
        ""

    
End Sub

Private Sub replaceLine(CodeMod As VBIDE.CodeModule, oldString As String, newString As String)
    Dim c01 As String
    c01 = Replace(CodeMod.Lines(1, CodeMod.CountOfLines), oldString, newString)
    CodeMod.DeleteLines 1, CodeMod.CountOfLines
    CodeMod.AddFromString c01
End Sub




