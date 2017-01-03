

Sub mainProfilerTest()
    'reinit profile manager
    Profiler.resetProfileManager

    'run some time/tests
    test1

    'print results
    Profiler.GetProfileManager.PrintTimes
End Sub

Sub test1()
    Dim mProfiler As New ProfiledMethod
    mProfiler.ProfiledMethod = "test1"


    Dim i As Long
    For i = 0 To 100
        test2
    Next i


End Sub

Sub test2()
    Dim mProfiler As New ProfiledMethod
    mProfiler.ProfiledMethod = "test2"

    test3
    test4

End Sub

Sub test3()
    Dim mProfiler As New ProfiledMethod
    mProfiler.ProfiledMethod = "test3"

    Dim i As Long, j As Long

    For i = 0 To 1000000
        j = 1 + 5
    Next i

End Sub


Sub test4()
    Dim mProfiler As New ProfiledMethod
    mProfiler.ProfiledMethod = "test4"

    Dim i As Long, j As Long

    For i = 0 To 500000
        j = 1 + 5
    Next i

End Sub
