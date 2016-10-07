VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_htc_provider_obs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Combo16_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 66 Or KeyCode = 78 Or KeyCode = 89 Then
        SendKeys "{enter}"
    End If
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me.User = CurrentUser()
    Me.TimeStamp = Now()
    Me.UpdateUser = CurrentUser()
    Me.UpdateTimeStamp = Now()
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Me.UpdateUser = CurrentUser()
    Me.UpdateTimeStamp = Now()
End Sub

Private Sub htc_prov_id_AfterUpdate()
Dim dbs As Database
Dim DbsLink As DAO.Database
Dim t As TableDef
Dim DbPath, SourceTable1, SourceTable2, TableName1, TableName2
Dim RstHtcPerson, RstHtcProv As DAO.Recordset
Dim HtcProvId As String



'Open recordset from htc_person (work-around to use index for seek on linked tables)

'    Set DbsLink = DBEngine(0)(0)
'    TableName1 = "htc_person"
'    TableName2 = "htc_prov"
'    DbPath = Mid(DbsLink(TableName1).Connect, InStr(1, DbsLink(TableName1).Connect, "=") + 1)
    
'    If DbPath = "" Then DbPath = CurrentDb()
    
'    SourceTable1 = DbsLink(TableName1).SourceTableName
'    SourceTable2 = DbsLink(TableName2).SourceTableName
    
'    Set DbsLink = DBEngine(0).OpenDatabase(DbPath)
'    Set RstHtcPerson = DbsLink.OpenRecordset(SourceTable1, DB_OPEN_TABLE)
'    Set RstHtcProv = DbsLink.OpenRecordset(SourceTable2, DB_OPEN_TABLE)
'    RstHtcPerson.Index = "htc_prov_id"
'    RstHtcProv.Index = "ID"
    
'Find the htc_person.ID matching the value entered in the htc_prov_ID and insert the htc_person.ID into the htc_person_ID

    HtcProvId = Me!htc_prov_id

    Set RstHtcProv = CurrentDb.OpenRecordset("select id from htc_prov where ID = '" & HtcProvId & "';", dbOpenDynaset, dbSeeChanges)
    
If HtcProvId <> "" Then
    With RstHtcProv
'        .Seek "=", HtcProvId
        If .EOF Then
            MsgBox ("Invalid Counsellor ID!" & Chr(10) & "Letters I, O, Q, S, U, Z" & Chr(10) & "are not permitted in IDs")
            Me.Undo
            Me!name_last_temp.SetFocus
            Me!htc_prov_id.SetFocus
            Exit Sub
        End If
    End With
    
    
    Set RstHtcPerson = CurrentDb.OpenRecordset("select id from htc_person where htc_prov_id = '" & HtcProvId & "';", dbOpenDynaset, dbSeeChanges)
    
    With RstHtcPerson
 '       .Seek "=", HtcProvId
        If Not .EOF Then
            Me!htc_person_id = !ID
            If HtcProvId = "4444" Then
                Me!name_last_temp.SetFocus
            Else
                Me!log_seen.SetFocus
            End If
        Else
            Me!name_last_temp.SetFocus
        End If
    End With
End If

End Sub

Private Sub htc_prov_id_Exit(Cancel As Integer)
Dim dbsCurrent As Database
Dim Qdf_chk_htc_prov_id As QueryDef
Dim rstcount As Integer
Dim HtcProvId As String
Dim chkrst As Variant

If Not IsNull(Me!htc_prov_id) Then
    
    Set dbsCurrent = CurrentDb
    HtcProvId = Me!htc_prov_id
        
    Set Qdf_chk_htc_prov_id = dbsCurrent.CreateQueryDef("")
    With Qdf_chk_htc_prov_id
        .SQL = "SELECT * FROM htc_prov " & _
         "WHERE ID =  '" & HtcProvId & "' AND Printed = TRUE ;"
        Set chkrst = .OpenRecordset()
        rstcount = chkrst.RecordCount()
    End With
        
        If rstcount = 0 Then
        MsgBox ("Invalid Counsellor ID!" & Chr(10) & "Letters I, O, Q, S, U, Z" & Chr(10) & "are not permitted in IDs")
        Me!htc_prov_id.SetFocus
        End
        End If
End If


End Sub

Private Sub log_seen_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 66 Or KeyCode = 78 Or KeyCode = 89 Then
        SendKeys "{enter}"
    End If

End Sub

Private Sub name_first_temp_AfterUpdate()
    
    If Not IsNull(Me!name_first_temp) Then
        Me!name_first_temp = StrConv(Trim(Me!name_first_temp), 3)
    End If

End Sub

Private Sub name_last_temp_AfterUpdate()

    If Not IsNull(Me!name_last_temp) Then
        Me!name_last_temp = StrConv(Trim(Me!name_last_temp), 3)
    End If

End Sub

Private Sub pt_done_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 66 Or KeyCode = 78 Or KeyCode = 89 Then
        SendKeys "{enter}"
    End If

End Sub

Private Sub test_done_known_AfterUpdate()
If Me!test_done_known = 0 Then
    Me!pt_done.SetFocus
End If
End Sub

Private Sub test_done_known_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 66 Or KeyCode = 78 Or KeyCode = 89 Then
        SendKeys "{enter}"
    End If

End Sub


