Attribute VB_Name = "ModHotkey"
' �ȼ����ģ��

Option Explicit

' ��Ҫ���ص��ȼ�
Dim retHotkey As Integer

' ����ȼ�
Public Sub CheckHotKey(hwnd As Long)

    ' �ж�Home��
    If dm.getKeyState(vbKeyHome) <> 0 Then
        retHotkey = 255
    End If
    
    ' �ж��ǵڼ�����װ
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
    
    ' �ȵ����а����ɿ����ٽ��в���
    If retHotkey <> 0 And dm.getKeyState(vbKeyControl) = 0 Then
    
        Select Case retHotkey
            Case 255
                ' װ����ͼ
                Call ScreenEquip(hwnd)
            Case 1
                ' ��ݻ���һ��װ��
                Call QuickEquip(hwnd, 1)
            Case 2
                ' ��ݻ��ڶ���װ��
                Call QuickEquip(hwnd, 2)
            Case 3
                ' ��ݻ�������װ��
                Call QuickEquip(hwnd, 3)
            Case 4
                ' ��ݻ�������װ��
                Call QuickEquip(hwnd, 4)
        End Select
    
        retHotkey = 0
    End If
End Sub
