Attribute VB_Name = "ModFunction"
' ִ�в���ģ��

Option Explicit

' ��������
Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long

' ȫ�־��
Dim hwnd300 As Long
' ����ģʽ
Dim isSlow As Boolean
' �ֱ���
Dim screenW, screenH As Integer

' ��ȡװ��
Public Function ScreenEquip(hwnd As Long)
    ' ���洰�ھ��
    hwnd300 = hwnd
    ' �ֱ���
    screenW = dm.GetScreenWidth
    screenH = dm.GetScreenHeight
    
    ' �󶨺�̨
    Call dm.BindWindow(hwnd300, "dx2", "windows", "windows", 0)
    
    ' �ֱ����ж�
    If screenW = 1920 Then
        Call dm.Capture(1119, 332, 1532, 821, "equip.bmp")
    Else
        Call dm.Capture(1439, 512, 1853, 1001, "equip.bmp")
    End If
    ' ����
    Call dm.Beep(1000, 100)
    ' ����װ��
    Call ScreenEquipDecode
    
    ' ����󶨺�̨
    Call dm.UnBindWindow
End Function

' ����װ��
Public Function ScreenEquipDecode()
    ' ͼƬ�����ھͲ���Ҫ����
    If dm.IsFileExist("equip.bmp") = 0 Then
        Exit Function
    End If

    ' �ֱ���
    screenW = dm.GetScreenWidth
    screenH = dm.GetScreenHeight
    
    Dim i As Integer, j As Integer
    Dim X As Integer, Y As Integer
    Dim step As Integer
    ' �ֱ����ж�
    If screenW = 1920 Then
        X = 1119
        Y = 332
    Else
        X = 1439
        Y = 512
    End If
    
    ' ��������
    For i = 0 To 6
        For j = 0 To 5
            FormMain.PictureEquip.PaintPicture LoadPicture("equip.bmp"), 0, 0, 65, 65, j * 65 + j * 4, i * 65 + i * 5, 65, 65
            FormMain.ImageBag(i * 6 + j).Picture = FormMain.PictureEquip.Image
        Next j
    Next i
    
    ' ����װ��
    For i = 1 To 4
        Dim equip As String
        Select Case i
            Case 1
                equip = FormMain.TextEquip1.Text
            Case 2
                equip = FormMain.TextEquip2.Text
            Case 3
                equip = FormMain.TextEquip3.Text
            Case 4
                equip = FormMain.TextEquip4.Text
        End Select
        
        ' �ж�����
        If equip <> "" Then
            ' ��Ϊ�գ��ָ���װ
            Dim arr() As String
            arr = Split(equip, ";")
            
            Dim e() As String
            Dim row As Integer, col As Integer
            
            ' ����װ��
            For j = LBound(arr) To UBound(arr)
                Dim arr2() As String
                Dim k As Integer
                arr2() = Split(equip, "/")
                
                For k = 0 To 5
                    If k <= UBound(arr2) Then
                        e() = Split(arr2(k), "-")
                        row = Val(e(0))
                        col = Val(e(1))
                        FormMain.PictureEquip.PaintPicture LoadPicture("equip.bmp"), 0, 0, 65, 65, (col - 1) * 65 + (col - 1) * 4, (row - 1) * 65 + (row - 1) * 5, 65, 65
                        Select Case i
                            Case 1
                                FormMain.ImageEquip1(k).Picture = FormMain.PictureEquip.Image
                            Case 2
                                FormMain.ImageEquip2(k).Picture = FormMain.PictureEquip.Image
                            Case 3
                                FormMain.ImageEquip3(k).Picture = FormMain.PictureEquip.Image
                            Case 4
                                FormMain.ImageEquip4(k).Picture = FormMain.PictureEquip.Image
                        End Select
                    Else
                        Select Case i
                            Case 1
                                FormMain.ImageEquip1(k).Picture = LoadPicture("")
                            Case 2
                                FormMain.ImageEquip2(k).Picture = LoadPicture("")
                            Case 3
                                FormMain.ImageEquip3(k).Picture = LoadPicture("")
                            Case 4
                                FormMain.ImageEquip4(k).Picture = LoadPicture("")
                        End Select
                    End If
                    
                Next k
            Next j
        Else
            ' Ϊ�գ�ɾ����װ
            For k = 0 To 5
                Select Case i
                    Case 1
                        FormMain.ImageEquip1(k).Picture = LoadPicture("")
                    Case 2
                        FormMain.ImageEquip2(k).Picture = LoadPicture("")
                    Case 3
                        FormMain.ImageEquip3(k).Picture = LoadPicture("")
                    Case 4
                        FormMain.ImageEquip4(k).Picture = LoadPicture("")
                End Select
            Next k
        End If
    Next i
    
End Function

' �����л�ս��װ��
' num���ڼ���
Public Function QuickEquip(hwnd As Long, num As Integer)
    ' ��ȡװ������
    Dim equip As String
    Select Case num
        Case 1
            equip = FormMain.TextEquip1.Text
        Case 2
            equip = FormMain.TextEquip2.Text
        Case 3
            equip = FormMain.TextEquip3.Text
        Case 4
            equip = FormMain.TextEquip4.Text
    End Select
    
    ' ���洰�ھ��
    hwnd300 = hwnd
    ' �ֱ���
    screenW = dm.GetScreenWidth
    screenH = dm.GetScreenHeight
    
    ' ����ģʽ
    isSlow = FormMain.CbSlow.Value = 1
    ' �����
    Call dm.SetWindowState(hwnd300, 1)
    
    ' �򿪱���
    dm.KeyPress vbKeyO
    
    ' �ж�����
    If equip = "" Then
        ' Ϊ�գ�ֻ��װ��
        Call DownEquip
    Else
        ' ��Ϊ�գ��ָ���װ
        Dim arr() As String
        Dim i As Integer
        arr = Split(equip, ";")
        
        ' ������װ
        For i = LBound(arr) To UBound(arr)
            ' ����װ��
            Call DownEquip
            ' Ȼ����װ��
            Call UpEquip(arr(i))
        Next i
    End If
    
    ' �رձ���
    dm.KeyPress vbKeyO
    ' �ٴμ����
    ' Call dm.SetWindowState(hwnd300, 1)
End Function

' ������װ��
Public Function DownEquip()
    ' �󶨺�̨
    Call dm.BindWindow(hwnd300, "normal", "windows", "windows", 0)
    ' ����ǰ̨��������
    Call dm.SetWindowState(hwnd300, 10)
    
    Dim X As Integer, Y As Integer
    Dim step As Integer, offset As Integer
    Dim minDelay As Long, maxDelay As Long
    offset = 10
    
    
    If isSlow Then
        minDelay = 300
        maxDelay = 330
    Else
        minDelay = 100
        maxDelay = 120
    End If
    ' �ֱ����ж�
    If screenW = 1920 Then
        X = 960
        Y = 380
        step = 75
    Else
        X = 1294
        Y = 568
        step = 75
    End If
    
    ' �µ�����װ��
    Delay minDelay, maxDelay
    If isSlow Then
        minDelay = 150
        maxDelay = 200
    Else
        minDelay = 15
        maxDelay = 25
    End If
    
    dm.moveToEx X, Y, offset, offset
    dm.RightClick
    Delay minDelay, maxDelay
    dm.moveToEx X + step, Y, offset, offset
    dm.RightClick
    Delay minDelay, maxDelay
    dm.moveToEx X, Y + step, offset, offset
    dm.RightClick
    Delay minDelay, maxDelay
    dm.moveToEx X + step, Y + step, offset, offset
    dm.RightClick
    Delay minDelay, maxDelay
    dm.moveToEx X, Y + step * 2, offset, offset
    dm.RightClick
    Delay minDelay, maxDelay
    dm.moveToEx X + step, Y + step * 2, offset, offset
    dm.RightClick
    
    ' ����װ����һ��
    Delay minDelay, maxDelay
    
    ' �������ǰ̨��������
    Call dm.SetWindowState(hwnd300, 11)
    ' ����󶨺�̨
    Call dm.UnBindWindow
End Function

' ������װ��
Public Function UpEquip(equip As String)
    ' �������λ��
    Dim mouseX, mouseY
    Call dm.GetCursorPos(mouseX, mouseY)
    
    ' �󶨺�̨
    Call dm.BindWindow(hwnd300, "dx", "dx", "windows", 0)
    ' ���μ������룬��ֹ����
    Call BlockInput(True)
    
    Dim X As Integer, Y As Integer
    Dim step As Integer, offset As Integer
    Dim minDelay As Long, maxDelay As Long
    offset = 10
    
    If isSlow Then
        minDelay = 260
        maxDelay = 300
    Else
        minDelay = 120
        maxDelay = 140
    End If
    ' �ֱ����ж�
    If screenW = 1920 Then
        X = 1150
        Y = 366
        step = 68
    Else
        X = 1473
        Y = 545
        step = 68
    End If
    
    ' �и�װ��
    Dim arr() As String
    arr() = Split(equip, "/")
    Dim i As Integer
    Dim e() As String
    Dim row As Integer, col As Integer
    
    ' ����װ��
    For i = LBound(arr) To UBound(arr)
        e() = Split(arr(i), "-")
        row = Val(e(0))
        col = Val(e(1))
        
        dm.moveToEx X + (col - 1) * step, Y + (row - 1) * step, offset, offset
        Delay minDelay, maxDelay
        dm.RightClick
    Next i
    
    ' ����װ����һ��
    Delay minDelay, maxDelay
    
    ' �ָ���������
    Call BlockInput(False)
    ' ����󶨺�̨
    Call dm.UnBindWindow
    
    ' ��ԭ���λ��
    Call dm.moveTo(mouseX, mouseY)
End Function

' �ַ�������
Public Function CountChar(s As String, char As String) As Integer
    Dim arr
    arr = Split(s, char)
    CountChar = UBound(arr) - LBound(arr)
End Function

' ɾ������������
Public Function RemoveArrayItem(StrArray() As String, Index As Integer) As String()
    Dim X As Integer
    For X = Index To UBound(StrArray) - 1
        StrArray(X) = StrArray(X + 1)
    Next
    If UBound(StrArray) > 0 Then
        ReDim Preserve StrArray(UBound(StrArray) - 1)
    Else
        Erase StrArray()
    End If
    RemoveArrayItem = StrArray()
End Function

' �ж������Ƿ�Ϊ��
Public Function IsNotEmpty(ByVal sArray As Variant) As Boolean
    Dim i     As Long
    IsNotEmpty = True
    On Error GoTo lerr:
    i = UBound(sArray)
    Exit Function
lerr:
    IsNotEmpty = False
End Function
