Attribute VB_Name = "ExportAllCode"
Option Compare Database

Sub ExportAllCode()

For Each c In Application.VBE.VBProjects(1).VBComponents
Select Case c.Type
    Case vbext_ct_ClassModule, vbext_ct_Document
        Sfx = ".cls"
    Case vbext_ct_MSForm
        Sfx = ".frm"
    Case vbext_ct_StdModule
        Sfx = ".bas"
    Case Else
        Sfx = ""
End Select
If Sfx <> "" Then
    c.Export _
        FileName:=CurrentProject.Path & "\" & _
        c.name & Sfx
End If
Next c

End Sub
