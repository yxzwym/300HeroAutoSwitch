Attribute VB_Name = "ModConfig"
' ��д ini

' д�� ini
' WriteIni "pronum", "num", 123, path

' ����ini
' num = ReadIni("pronum", "num", "0", path)

Option Explicit

' �������ļ�(����ϵͳ�⺯��)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
 ByVal AppName As String, _
 ByVal KeyName As String, _
 ByVal lpDefault As String, _
 ByVal lpReturnString As String, _
 ByVal nSize As Long, _
 ByVal FileName As String) _
 
' д�����ļ�(����ϵͳ�⺯��)
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
ByVal AppName$, _
ByVal KeyName$, _
ByVal keyDefault$, _
ByVal FileName$)

' �������ļ����򻯺������ã�
Public Function ReadIni(ByVal AppName As String, ByVal KeyName As String, ByVal DefaultValue As String, Optional IniPath As String) As String
    Dim buf As String
    Dim ret As Integer
    Dim tmp As String
    buf = String(1024, 0) 'buf=1024��0
    If IniPath = "" Then
        ret = GetPrivateProfileString(AppName, KeyName, DefaultValue, buf, 1024, App.path + "\config.ini")
    Else
        ret = GetPrivateProfileString(AppName, KeyName, DefaultValue, buf, 1024, IniPath)
    End If
    tmp = Mid(buf, 1, ret)
    If InStr(1, tmp, Chr(0)) > 0 Then tmp = Left(tmp, InStr(1, tmp, Chr(0)) - 1)
    ReadIni = tmp
End Function

' д�����ļ����򻯺������ã�
Public Function WriteIni(ByVal AppName As String, ByVal KeyName As String, ByVal KeyValue As String, Optional IniPath As String = "") As Boolean
    If IniPath = "" Then
        WritePrivateProfileString& AppName, KeyName, KeyValue, App.path + "\config.ini"
    Else
        WritePrivateProfileString& AppName, KeyName, KeyValue, IniPath
    End If
End Function

' ������װ
' num����x��
' equip��1-1/1-2
Public Function SetEquip(num As Integer, equip As String)
    Call WriteIni("Equip", Str(num), equip)
End Function

' ��ȡ��x����װ
' num����x��
Public Function GetEquip(num As Integer) As String
    Dim tmp As String
    tmp = ReadIni("Equip", Str(num), "")
    GetEquip = tmp
End Function

' ����һ����װ�ٶ�
' 0�����٣�1������
Public Function SetSlowMode(speed As Integer)
    Call WriteIni("Setting", "SlowMode", speed)
End Function

' ��ȡһ����װ�ٶ�
' 0�����٣�1������
Public Function GetSlowMode() As Integer
    GetSlowMode = ReadIni("Setting", "SlowMode", 0)
End Function

