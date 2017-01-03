Option Compare Database

' http://codereview.stackexchange.com/questions/70247/vba-code-profiling
'
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public mProfileManager As ProfileManager

Public Function GetProfileManager() As ProfileManager

    If mProfileManager Is Nothing Then
        Set mProfileManager = New ProfileManager
    End If

    Set GetProfileManager = mProfileManager

End Function

Public Sub resetProfileManager()
    Set mProfileManager = Nothing
End Sub

