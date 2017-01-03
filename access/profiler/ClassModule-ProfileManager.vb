Option Compare Database
Option Explicit

' Reference to Microsoft Scripting Runtime
Private m_MethodTotalTimes As Scripting.Dictionary
Private m_MethodTotalCalls As Scripting.Dictionary

Public Sub addMethodCall(p_method As String, p_time As Double)

    If m_MethodTotalTimes.exists(p_method) Then
        m_MethodTotalTimes(p_method) = m_MethodTotalTimes(p_method) + p_time
        m_MethodTotalCalls(p_method) = m_MethodTotalCalls(p_method) + 1
    Else
        m_MethodTotalTimes.Add p_method, p_time
        m_MethodTotalCalls.Add p_method, 1
    End If

End Sub

Public Sub PrintTimes()
    Dim mKey
    For Each mKey In m_MethodTotalTimes.Keys
        Debug.Print mKey & " was called " & m_MethodTotalCalls(mKey) & " times for a total time of " & m_MethodTotalTimes(mKey)
    Next mKey
End Sub

Private Sub Class_Initialize()
    Set m_MethodTotalTimes = New Scripting.Dictionary
    Set m_MethodTotalCalls = New Scripting.Dictionary
End Sub
