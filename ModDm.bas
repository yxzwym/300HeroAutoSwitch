Attribute VB_Name = "ModDm"
' ��Įע��ģ��

Option Explicit

' ��ע����ô�Į api
Private Declare Function SetDllPathA Lib "DmReg.dll" (ByVal path As String, ByVal mode As Long) As Long

' ��ȫ��ʹ�õĴ�Į����
Public dm As Object

' ��ʼ����Į dll
' 0��ʧ�ܣ�1���ɹ�
Public Function InitDm() As Integer
try: On Error GoTo catch
    Call SetDllPathA("Dm.dll", 0)
    Set dm = CreateObject("dm.dmsoft")
finally:
    InitDm = 1
    Exit Function
catch:
    InitDm = 0
    Exit Function
End Function

' ��ȡ��Ϸ�� Hwnd
Public Function GetHwnd() As Long
    Dim hwnds As Variant
    hwnds = Split(dm.EnumWindow(0, "", "", 1 + 4 + 8 + 16), ",")
    
    Dim hwnd As Long
    Dim hwndStr As Variant
    Dim process_path As String
    For Each hwndStr In hwnds
        hwnd = CLng(Val(hwndStr))
        If Len(dm.GetWindowTitle(hwnd)) > 0 Then
            process_path = dm.GetWindowProcessPath(hwnd)
            If InStr(process_path, "300.exe") Then
                GetHwnd = hwnd
                Exit Function
            End If
        End If
    Next
    
    GetHwnd = 0
End Function
