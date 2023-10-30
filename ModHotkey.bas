Attribute VB_Name = "ModHotkey"
' 热键检测模块

Option Explicit

' 需要返回的热键
Dim retHotkey As Integer

' 检测热键
Public Sub CheckHotKey(hwnd As Long)

    ' 判断Home键
    If dm.getKeyState(vbKeyHome) <> 0 Then
        retHotkey = 255
    End If
    
    ' 判断是第几套配装
    If dm.getKeyState(vbKey1) <> 0 And dm.getKeyState(vbKeyControl) <> 0 Then
        retHotkey = 1
    End If
    If dm.getKeyState(vbKey2) <> 0 And dm.getKeyState(vbKeyControl) <> 0 Then
        retHotkey = 2
    End If
    If dm.getKeyState(vbKey3) <> 0 And dm.getKeyState(vbKeyControl) <> 0 Then
        retHotkey = 3
    End If
    If dm.getKeyState(vbKey4) <> 0 And dm.getKeyState(vbKeyControl) <> 0 Then
        retHotkey = 4
    End If
    
    ' 等到所有按键松开后再进行操作
    If retHotkey <> 0 And dm.getKeyState(vbKeyControl) = 0 Then
    
        Select Case retHotkey
            Case 255
                ' 装备截图
                Call ScreenEquip(hwnd)
            Case 1
                ' 快捷换第一套装备
                Call QuickEquip(hwnd, 1)
            Case 2
                ' 快捷换第二套装备
                Call QuickEquip(hwnd, 2)
            Case 3
                ' 快捷换第三套装备
                Call QuickEquip(hwnd, 3)
            Case 4
                ' 快捷换第四套装备
                Call QuickEquip(hwnd, 4)
        End Select
    
        retHotkey = 0
    End If
End Sub
