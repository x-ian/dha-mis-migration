Option Compare Database

' Meta programming
' Change code with code
' Instructing / instrumenting code
' code generation
' for logging, profiling
' Microsoft Visual Basic for Applications Extensibility 5.3: https://msdn.microsoft.com/en-us/library/aa443970(v=vs.60).aspx
' heavily inspired by http://www.cpearson.com/excel/InsertProcedureNames.aspx and http://stackoverflow.com/questions/35459541/adding-line-numbers-to-vba-code-microsoft-access-2016
 
Sub GenerateProfilerCodeForActiveCodePane()
    PrependBlockIntoProcedures Application.VBE.ActiveCodePane.CodeModule
End Sub
 
Sub GenerateProfilerCodeForAllModules()
    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
        Debug.Print objComponent.name
        PrependBlockIntoProcedures objComponent.CodeModule
    Next objComponent
End Sub
 
' Get rid of Profiler code, use search and replace (for Module or Whole Project)
' replace>	Dim mProfiler As New ProfiledMethod< with empty line (Find Whole Words Only, Match Case)
' replace>	mProfiler.ProfiledMethod = BLOCK_MARKER_PROC_NAME< with empty line (Find Whole Words Only, Match Case)
' replace>	Const BLOCK_MARKER_PROC_NAME = "*"< with empty line (Find Whole Words Only, Match Case, With pattern matching)


' InsertProcedureNameIntoProcedure will insert a CONST statement at the top of each
' procedure in the Application.VBE.ActiveCodePane.CodeModule, like
'    Const C_PROC_NAME_illegal = "callit"
' It supports procedures whose Declaration spans more than one line, e.g.,
'       Public Function Test( X As Integer, _
'                             Y As Integer, _
'                             Z As Integer)
' If comment lines appear DIRECTLY below the procedured declaration (no blank lines
' between the declaration and the start of the comments), the CONST statement is
' placed directly below the comment block. If a constant already exists with name
' user specified name, that constant declaration is deleted and replace with the
' new CONST line.
Sub PrependBlockIntoProcedures(CodeMod As VBIDE.CodeModule)

    Dim ProcName As String
    Dim ModuleName As String
    Dim ProcLine As String
    Dim ProcType As VBIDE.vbext_ProcKind
    Dim StartLine As Long
    Dim Done As Boolean
    Dim ProcBodyLine As Long
    Dim SaveProcName As String
    Const ConstName = "BLOCK_MARKER_PROC_NAME"
    Dim ConstAtLine As Long
    Dim EndOfDeclaration As Long

    ' Skip past any Option statement and any module-level
    ' variable declations. Start at the first procuedure
    ' in the module.
    StartLine = CodeMod.CountOfDeclarationLines + 1

    ' Get the procedure name that is at StartLine.
    ProcName = CodeMod.ProcOfLine(StartLine, ProcType)
    SaveProcName = ProcName
    ModuleName = CodeMod.name

    ' Loop through all procedures in the module.
    Do Until Done
    
        ' Get the body proc line (the actual declaration line, ignoring comments)
        ProcBodyLine = CodeMod.ProcBodyLine(ProcName, ProcType)

        ' See if the constant declaration already exists.
        ConstAtLine = ConstNameInProcedure(ConstName, CodeMod, ProcName, ProcType)
        If ConstAtLine > 0 Then
            ' Const line already exist. Assume next x lines belong to it as well.
            ' Delete it and replace it.
            CodeMod.DeleteLines ConstAtLine, 3
            CodeMod.InsertLines ConstAtLine, "    Const " & ConstName & " = " & Chr(34) & ModuleName & "." & ProcName & Chr(34)
            CodeMod.InsertLines ConstAtLine + 1, "    Dim mProfiler As New ProfiledMethod"
            CodeMod.InsertLines ConstAtLine + 2, "    mProfiler.ProfiledMethod = " & ConstName
        Else
            ' Skip past the declaration lines and the comment lines that
            ' immediately follow the declarations (no blank lines between
            ' the declarations and the comments).
            ' Insert the CONST declaration.
            EndOfDeclaration = EndOfDeclarationLines(CodeMod, ProcName, ProcType)
            ProcLine = EndOfCommentOfProc(CodeMod, EndOfDeclaration + 1)
'            CodeMod.InsertLines ProcLine + 1, "    Const " & ConstName & " = " & Chr(34) & ProcName & Chr(34)
            CodeMod.InsertLines ProcLine + 1, "    Const " & ConstName & " = " & Chr(34) & ModuleName & "." & ProcName & Chr(34)
            CodeMod.InsertLines ProcLine + 2, "    Dim mProfiler As New ProfiledMethod"
            CodeMod.InsertLines ProcLine + 3, "    mProfiler.ProfiledMethod = " & ConstName
        End If
  
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

' This returns the line number of the last comment line in a comment block that IMMEDIATELY
' follow the procedure declaration. For example, with the following code
'       Function MyTest(X As Integer, _
'                       Y As Integer, _
'                       Z As Ineger)
'       '''''''''''''''''''''''''''''''''''' START COMMENT BLOCK - NO BLANK LINES ABOVE
'       ' Some Comments
'       '''''''''''''''''''''''''''''''''''' END COMMENT BLOCK
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

' This returns the line number containing the existing constant declaration,
' or 0 if the procedure does not contain the constant declaration.
Function ConstNameInProcedure(ConstName As String, CodeMod As VBIDE.CodeModule, _
    ProcName As String, _
    ProcType As VBIDE.vbext_ProcKind) As Long
    
    Dim LineNum As Long
    Dim LineText As String
    Dim ProcBodyLine As Long

    ProcBodyLine = CodeMod.ProcBodyLine(ProcName, ProcType)
    For LineNum = ProcBodyLine To ProcBodyLine + CodeMod.ProcCountLines(ProcName, ProcType)
        LineText = CodeMod.Lines(LineNum, 1)
        If InStr(LineText, " " & ConstName & " ") > 0 Then
            ConstNameInProcedure = LineNum
            Exit Function
        End If
    Next LineNum
End Function

' This return the line number of the last declation lines. This is used to find
' the end of declarations when the declaration span more than one line of text
' in the code. For example, with the code
'       Function MyTest(X As Integer, _
'                       Y As Integer, _
'                       Z As Integer)
' it will return the line number of "Z As Integer)", the line number of the
' end of the declarations block.
Function EndOfDeclarationLines(CodeMod As VBIDE.CodeModule, ProcName As String, ProcType As VBIDE.vbext_ProcKind) As Long
    Dim LineNum As Long
    Dim LineText As String

    LineNum = CodeMod.ProcBodyLine(ProcName, ProcType)
    Do Until Right(CodeMod.Lines(LineNum, 1), 1) <> "_"
        LineNum = LineNum + 1
    Loop

    EndOfDeclarationLines = LineNum - 1 ' added - 1 compare to cpearson's code
End Function

