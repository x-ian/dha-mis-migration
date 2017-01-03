Option Compare Database
Option Explicit

Private initTime As Double
Private mProfiledMethod As String

Public Property Let ProfiledMethod(pValue As String)
    mProfiledMethod = pValue
End Property

Private Sub Class_Initialize()
    initTime = GetTickCount
End Sub

Private Sub Class_Terminate()
    GetProfileManager.addMethodCall mProfiledMethod, GetTickCount() - initTime
End Sub

