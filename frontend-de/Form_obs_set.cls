VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_obs_set"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Sub CommandSave_Click()
Dim dbs As Database
Dim DbsLink1, DbsLink2 As DAO.Database
Dim t As TableDef
Dim DbPath1, DbPath2, SourceTable1, SourceTable2, TableName
Dim RstObsDim, RstObs As DAO.Recordset
Dim RstObsDimFrm, RstObsFrm, RstChkSum As Recordset
Dim ArtClinicObsId, PeriodReport, SubGroup, ObsDimensionsID As Long
Dim Cancel As Integer

'Create a recordset from the current obs_set form data
    
    Set dbs = CurrentDb()
    Set RstObsFrm = Form_obs_set.Recordset
     
'Cycle through records to check for completeness of entry

    With RstObsFrm
        .MoveFirst
        
        Do While Not .EOF
            If IsNull(!data_value_tmp) Then
                MsgBox "Enter value for: " & !data_element_name, , "Data Entry Required"
                Form_obs_set.Bookmark = .Bookmark
                
                If Me!ctl_type = 361 Then
                    Me.data_value_per.SetFocus
                ElseIf Me!ctl_type = 360 Then
                    Me.data_value_ynb.SetFocus
                Else
                    Me.data_value_txt.SetFocus
                End If
                
                Exit Sub
            Else
                .MoveNext
            End If
        Loop
    End With
    
'Cycle through records of checksum form to check for inconsistencies

    Set RstChkSum = Form_obs_dim_set.obs_set_tmptbl_chksum.Form.RecordsetClone
    
    With RstChkSum
    If .RecordCount > 0 Then
        .MoveFirst
        
        Do While Not .EOF
            If !element_right_sum <> !element_left_sum Then
                MsgBox "Inconsistent check sum: " & !chksum_name, , "Data Validation"
                
                RstObsFrm.FindFirst "[chksum_name] = " & """" & !chksum_name & """"
                If Not RstObsFrm.NoMatch Then
                    Form_obs_set.Bookmark = RstObsFrm.Bookmark
                End If
                
                If Me!ctl_type = 361 Then
                    Me.data_value_per.SetFocus
                ElseIf Me!ctl_type = 360 Then
                    Me.data_value_ynb.SetFocus
                Else
                    Me.data_value_txt.SetFocus
                End If
                
                Exit Sub
            Else
                .MoveNext
            End If
        Loop
    End If
    End With
    
'Check if any data_value_temp field in this section was updated
    
    With RstObsFrm
        .FindFirst " [data_value_changed] = True "
        If .NoMatch Then
            MsgBox "Click OK to move to next section.", vbOKOnly, "No New Data in this Section!"
            GoTo ScrollNext
        End If
    End With
    

'Populate current parameter from Form_art_clinic_obs_v8

    ArtClinicObsId = Form_art_clinic_obs_v8.ID
    
'Populate current parameters from Form_obs_dim_set

    PeriodReport = Form_obs_dim_set.period_report
    SubGroup = Form_obs_dim_set.sub_group

'Open recordset from obs_dimensions and obs tables (work-around to use index for seek on linked tables)

'    Set DbsLink1 = DBEngine(0)(0)
'    TableName = "obs"
'    DbPath1 = Mid(DbsLink1(TableName).Connect, InStr(1, DbsLink1(TableName).Connect, "=") + 1)
    
'    If DbPath1 = "" Then DbPath1 = CurrentDb()
    
'    SourceTable1 = DbsLink1(TableName).SourceTableName
    
'    Set DbsLink1 = DBEngine(0).OpenDatabase(DbPath1)
'    Set RstObs = DbsLink1.OpenRecordset(SourceTable1, DB_OPEN_TABLE)
'    RstObs.Index = "uniq"
    
'    Set DbsLink2 = DBEngine(0)(0)
'    TableName = "obs_dimensions"
'    DbPath2 = Mid(DbsLink2(TableName).Connect, InStr(1, DbsLink2(TableName).Connect, "=") + 1)
'    SourceTable2 = DbsLink2(TableName).SourceTableName
    
'    Set DbsLink2 = DBEngine(0).OpenDatabase(DbPath2)
'    Set RstObsDim = DbsLink2.OpenRecordset(SourceTable2, DB_OPEN_TABLE)
'    RstObsDim.Index = "uniq"

    
'Find obs_dimensions_ID matching current parameters, otherwise insert rec in obs_dimensions table

    Set RstObsDim = CurrentDb.OpenRecordset("select * from obs_dimensions where art_clinic_obs_id = " & ArtClinicObsId & " and period_report = " & PeriodReport & " and sub_group = " & SubGroup & ";", dbOpenDynaset, dbSeeChanges)
    With RstObsDim
'        .Seek "=", ArtClinicObsId, PeriodReport, SubGroup
        If .EOF Then
            .AddNew
            !art_clinic_obs_id = ArtClinicObsId
            !period_report = PeriodReport
            !sub_group = SubGroup
            !User = CurrentUser()
            !TimeStamp = Now()
            .Update
            .Bookmark = RstObsDim.LastModified
            ObsDimensionsID = !ID
        Else
            ObsDimensionsID = ![ID]
        End If
    End With
    
'Add or update all changed records in obs table
    
    With RstObsFrm
        Do While Not .EOF
            .FindFirst "[data_value_changed] = True"
            If .NoMatch Then
                GoTo ScrollNext
            Else
 '               RstObs.Seek "=", ObsDimensionsID, RstObsFrm!data_element
 '               If RstObs.NoMatch Then
                 Set RstObs = CurrentDb.OpenRecordset("select * from obs where obs_dimensions_id = " & ObsDimensionsID & ";", dbOpenDynaset, dbSeeChanges)
                 If RstObs.EOF Then
                    RstObs.AddNew
                    RstObs!obs_dimensions_ID = ObsDimensionsID
                    RstObs!data_element = !data_element
                    RstObs!data_value = !data_value_tmp
                    RstObs.Update
                Else
                    RstObs.Edit
                    RstObs!data_value = !data_value_tmp
                    RstObs.Update
                End If
                Me!data_value_changed = False
            End If
        Loop
    End With

ScrollNext:
    FormSectionSelectScroll

End Sub

Private Sub data_value_per_AfterUpdate()

If Not IsNull(Me.data_value_per) Then
Me.data_value_per.Value = Me.data_value_per.Value / 100
End If

'Flag record as updatedated
Me!data_value_changed = True

End Sub

Private Sub data_value_per_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Me.CurrentRecord = Me.RecordsetClone.RecordCount Then
        Form_obs_dim_set.form_section_select.SetFocus
End If
End Sub

Private Sub data_value_txt_AfterUpdate()

'Flag record as updatedated
    Me!data_value_changed = True

End Sub

Private Sub data_value_txt_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If (KeyCode = 9 Or KeyCode = 13) And Me.CurrentRecord = Me.RecordsetClone.RecordCount Then
        Form_obs_set.CommandSave.SetFocus
    End If

End Sub

Private Sub data_value_ynb_AfterUpdate()

'Flag record as updatedated
    Me!data_value_changed = True

End Sub

Private Sub data_value_ynb_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 66 Or KeyCode = 78 Or KeyCode = 89 Then
        If Me.CurrentRecord = Me.RecordsetClone.RecordCount Then
            Form_obs_set.CommandSave.SetFocus
        Else
            SendKeys "{enter}"
        End If
    End If

End Sub

Public Sub FormSectionSelectScroll()
Dim RequeryCancel As Boolean
Dim Rst As Recordset
Dim Cancel As Integer

RequeryCancel = True
    
    With Form_obs_dim_set.form_section_select
        .SetFocus ' the combobox has to have the focus to set the ListIndex
        If .ListIndex < .ListCount - 1 Then
            .OnClick = Cancel
            .ListIndex = .ListIndex + 1 ' select the next item
            Form_obs_dim_set.Requery
            .OnClick = "[Event Procedure]"
        End If
    End With

Set Rst = Form_obs_set.RecordsetClone
Rst.MoveFirst
Me.Bookmark = Rst.Bookmark
If Me!ctl_type = 361 Then
    Me.data_value_per.SetFocus
ElseIf Me!ctl_type = 360 Then
    Me.data_value_ynb.SetFocus
Else
    Me.data_value_txt.SetFocus
End If

End Sub

Private Sub Parking_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 9 Or KeyCode = 13 Then
    FormSectionSelectScroll
End If
End Sub

Private Sub Form_AfterUpdate()
Dim SubFrmChkSum As SubForm

Set SubFrmChkSum = Form_obs_dim_set.obs_set_tmptbl_chksum
SubFrmChkSum.Requery

End Sub


