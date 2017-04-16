Option Compare Database
'Option Explicit

' Microsoft Visual Basic for Applications Extensibility 5.3: https://msdn.microsoft.com/en-us/library/aa443970(v=vs.60).aspx

Public BEGINNING_LINE_1 As String
Public BEGINNING_LINE_2 As String
Public ENDING_LINE_1a As String
Public ENDING_LINE_1b As String
Public ENDING_LINE_2 As String
Public ENDING_LINE_3 As String
Public ENDING_LINE_4 As String

Sub setStrings()
    BEGINNING_LINE_1 = "    ' GENERATED ERROR HANDLER"
    BEGINNING_LINE_2 = "    If runningOnAccessRuntime Then On Error Goto ERROR_HANDLER_RUNTIME"
    ENDING_LINE_1a = "    Exit Sub"
    ENDING_LINE_1b = "    Exit Function"
    ENDING_LINE_2 = "    ' GENERATED ERROR HANDLER"
    ENDING_LINE_3 = "ERROR_HANDLER_RUNTIME:"
    ENDING_LINE_4 = "    MsgBox ""Error happened"""
End Sub

Sub GenerateErrorHandlerCodeForActiveCodePane()
    setStrings
    RemoveErrorHandler Application.VBE.ActiveCodePane.CodeModule
    AddErrorHandlerBeginning Application.VBE.ActiveCodePane.CodeModule
    AddErrorHandlerEnding Application.VBE.ActiveCodePane.CodeModule
End Sub
 
Sub RemoveErrorHandlerCodeForActiveCodePane()
    setStrings
    RemoveErrorHandler Application.VBE.ActiveCodePane.CodeModule
End Sub
 
Sub GenerateErrorHandlerCodeForAllModules()
    'Dim objComponent As ActiveVBProject.VBComponents
    setStrings
    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
        Debug.Print objComponent.name
        RemoveErrorHandler objComponent.CodeModule
        AddErrorHandlerBeginning objComponent.CodeModule
        AddErrorHandlerEnding objComponent.CodeModule
    Next objComponent
End Sub
 

Private Sub RemoveErrorHandler(CodeMod As VBIDE.CodeModule)
    setStrings
    removeLine CodeMod, BEGINNING_LINE_1
    removeLine CodeMod, BEGINNING_LINE_2
    removeLine CodeMod, ENDING_LINE_1a
    removeLine CodeMod, ENDING_LINE_1b
    removeLine CodeMod, ENDING_LINE_2
    removeLine CodeMod, ENDING_LINE_3
    removeLine CodeMod, ENDING_LINE_4
End Sub

Private Sub AddErrorHandlerEnding(CodeMod As VBIDE.CodeModule)
    Dim c01 As String
    c01 = Replace(CodeMod.Lines(1, CodeMod.CountOfLines), _
        "End Sub", _
        ENDING_LINE_1a & vbCrLf & ENDING_LINE_2 & vbCrLf & ENDING_LINE_3 & vbCrLf & ENDING_LINE_4 & vbCrLf & "End Sub")
    CodeMod.DeleteLines 1, CodeMod.CountOfLines
    CodeMod.AddFromString c01
    
    c01 = Replace(CodeMod.Lines(1, CodeMod.CountOfLines), _
        "End Function", _
        ENDING_LINE_1b & vbCrLf & ENDING_LINE_2 & vbCrLf & ENDING_LINE_3 & vbCrLf & ENDING_LINE_4 & vbCrLf & "End Function")
    CodeMod.DeleteLines 1, CodeMod.CountOfLines
    CodeMod.AddFromString c01
    
End Sub

Private Sub removeLine(CodeMod As VBIDE.CodeModule, line As String)
    Dim c01 As String
    c01 = Replace(CodeMod.Lines(1, CodeMod.CountOfLines), line & vbCrLf, "")
    CodeMod.DeleteLines 1, CodeMod.CountOfLines
    CodeMod.AddFromString c01
End Sub

Sub AddErrorHandlerBeginning(CodeMod As VBIDE.CodeModule)

    Dim ProcName As String
    Dim ModuleName As String
    Dim ProcLine As String
    Dim ProcType As VBIDE.vbext_ProcKind
    Dim StartLine As Long
    Dim Done As Boolean
    Dim ProcBodyLine As Long
    Dim SaveProcName As String
    Dim ConstAtLine As Long
    Dim EndOfDeclaration As Long
    Dim EndOf As Long
    Dim ProcStartLine As String
    Dim i As Integer
    Dim isFunction As Integer
    Dim line As String
    
    ' Skip past any Option statement and any module-level
    ' variable declations. Start at the first procuedure
    ' in the module.
    StartLine = CodeMod.CountOfDeclarationLines + 1

    ' Get the procedure name that is at StartLine.
    ProcName = CodeMod.ProcOfLine(StartLine, ProcType)
    SaveProcName = ProcName
    ModuleName = CodeMod.name

    StartLine = CodeMod.ProcCountLines(ProcName, ProcType)
    
    ' Loop through all procedures in the module.
    Do Until Done
        ' Insert lines at the beginning of proc
        ProcBodyLine = CodeMod.ProcBodyLine(ProcName, ProcType)
        EndOfDeclaration = EndOfDeclarationLines(CodeMod, ProcName, ProcType)
        ProcLine = EndOfCommentOfProc(CodeMod, EndOfDeclaration + 1)
        CodeMod.InsertLines ProcLine + 1, BEGINNING_LINE_1
        CodeMod.InsertLines ProcLine + 2, BEGINNING_LINE_2
        
        ' Skip StartLine to the next proc
        StartLine = ProcBodyLine + CodeMod.ProcCountLines(ProcName, ProcType) + 1
        ProcName = CodeMod.ProcOfLine(StartLine, ProcType)
    
        ' Special handling for the last
        ' procedure in the module in case
        ' it has blank lines following the
        ' end of the procedure body.
        If ProcName = SaveProcName Then
            Done = True
        Else
            SaveProcName = ProcName
        End If
    Loop

End Sub

Function EndOfCommentOfProc(CodeMod As VBIDE.CodeModule, ProcBodyLine As Long) As Long
    Dim Done As Boolean
    Dim LineNum As String
    Dim LineText As String
    
    LineNum = ProcBodyLine

    Do Until Done
        LineNum = LineNum + 1
        LineText = CodeMod.Lines(LineNum, 1)
        If Left(Trim(LineText), 1) = "'" Then
            Done = False
        Else
            Done = True
        End If
    Loop
    EndOfCommentOfProc = LineNum - 1
End Function

Function EndOfDeclarationLines(CodeMod As VBIDE.CodeModule, ProcName As String, ProcType As VBIDE.vbext_ProcKind) As Long
    Dim LineNum As Long
    Dim LineText As String

    LineNum = CodeMod.ProcBodyLine(ProcName, ProcType)
    Do Until Right(CodeMod.Lines(LineNum, 1), 1) <> "_"
        LineNum = LineNum + 1
    Loop

    EndOfDeclarationLines = LineNum - 1 ' added - 1 compare to cpearson's code
End Function

