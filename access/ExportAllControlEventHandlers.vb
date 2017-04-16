Option Compare Database

Public Function ListAllControlsAndTheirEvents(form As Access.form) As String
    Dim prop    As Object
    Dim control As Access.control
    Dim result As String
    
    For Each control In form.Controls
        For Each prop In control.Properties
            ' Category 4 type 8 identifies an event property
            If prop.Category = 4 And prop.Type = 8 Then
                result = result & form.name & "," & control.name & "," & prop.name & "," & prop.value & vbCrLf
            End If
        Next
    Next
    ListAllControlsAndTheirEvents = result
End Function

Sub ListAllControlEvents()
    Dim file As String
    Dim intFile As Integer
    Dim events As String
    
    file = CurrentProject.Path & "\controls.csv"
    intFile = FreeFile
    Open file For Output As #intFile

    For Each obj In Application.CurrentProject.AllForms
        On Error Resume Next
        DoCmd.OpenForm obj.name, acDesign
        On Error GoTo 0
        events = ListAllControlsAndTheirEvents(Forms(0))
        Print #intFile, events
        DoCmd.Close acForm, obj.name
    Next obj
    Close intFile
End Sub

Sub closeAllForms()
    For i = 0 To Forms.Count - 1
        DoCmd.Close acForm, Forms(0).name
    Next i
End Sub
