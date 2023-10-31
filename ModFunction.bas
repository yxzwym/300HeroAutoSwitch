Attribute VB_Name = "ModFunction"
' 执行操作模块

Option Explicit

' 屏蔽输入
Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long

' 全局句柄
Dim hwnd300 As Long
' 慢速模式
Dim isSlow As Boolean
' 分辨率
Dim screenW, screenH As Integer

' 获取装备
Public Function ScreenEquip(hwnd As Long)
    ' 保存窗口句柄
    hwnd300 = hwnd
    ' 分辨率
    screenW = dm.GetScreenWidth
    screenH = dm.GetScreenHeight
    
    ' 绑定后台
    Call dm.BindWindow(hwnd300, "dx2", "windows", "windows", 0)
    
    ' 分辨率判断
    If screenW = 1920 Then
        Call dm.Capture(1119, 332, 1532, 821, "equip.bmp")
    Else
        Call dm.Capture(1439, 512, 1853, 1001, "equip.bmp")
    End If
    ' 蜂鸣
    Call dm.Beep(1000, 100)
    ' 解析装备
    Call ScreenEquipDecode
    
    ' 解除绑定后台
    Call dm.UnBindWindow
End Function

' 解析装备
Public Function ScreenEquipDecode()
    ' 图片不存在就不需要解析
    If dm.IsFileExist("equip.bmp") = 0 Then
        Exit Function
    End If

    ' 分辨率
    screenW = dm.GetScreenWidth
    screenH = dm.GetScreenHeight
    
    Dim i As Integer, j As Integer
    Dim X As Integer, Y As Integer
    Dim step As Integer
    ' 分辨率判断
    If screenW = 1920 Then
        X = 1119
        Y = 332
    Else
        X = 1439
        Y = 512
    End If
    
    ' 解析背包
    For i = 0 To 6
        For j = 0 To 5
            FormMain.PictureEquip.PaintPicture LoadPicture("equip.bmp"), 0, 0, 65, 65, j * 65 + j * 4, i * 65 + i * 5, 65, 65
            FormMain.ImageBag(i * 6 + j).Picture = FormMain.PictureEquip.Image
        Next j
    Next i
    
    ' 解析装备
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
        
        ' 判断配置
        If equip <> "" Then
            ' 不为空，分隔套装
            Dim arr() As String
            arr = Split(equip, ";")
            
            Dim e() As String
            Dim row As Integer, col As Integer
            
            ' 遍历装备
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
            ' 为空，删除配装
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

' 快速切换战场装备
' num：第几套
Public Function QuickEquip(hwnd As Long, num As Integer)
    ' 获取装备配置
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
    
    ' 保存窗口句柄
    hwnd300 = hwnd
    ' 分辨率
    screenW = dm.GetScreenWidth
    screenH = dm.GetScreenHeight
    
    ' 慢速模式
    isSlow = FormMain.CbSlow.Value = 1
    ' 激活窗口
    Call dm.SetWindowState(hwnd300, 1)
    
    ' 打开背包
    dm.KeyPress vbKeyO
    
    ' 判断配置
    If equip = "" Then
        ' 为空，只下装备
        Call DownEquip
    Else
        ' 不为空，分隔套装
        Dim arr() As String
        Dim i As Integer
        arr = Split(equip, ";")
        
        ' 遍历套装
        For i = LBound(arr) To UBound(arr)
            ' 先下装备
            Call DownEquip
            ' 然后上装备
            Call UpEquip(arr(i))
        Next i
    End If
    
    ' 关闭背包
    dm.KeyPress vbKeyO
    ' 再次激活窗口
    ' Call dm.SetWindowState(hwnd300, 1)
End Function

' 快速下装备
Public Function DownEquip()
    ' 绑定后台
    Call dm.BindWindow(hwnd300, "normal", "windows", "windows", 0)
    ' 屏蔽前台按键干扰
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
    ' 分辨率判断
    If screenW = 1920 Then
        X = 960
        Y = 380
        step = 75
    Else
        X = 1294
        Y = 568
        step = 75
    End If
    
    ' 下掉所有装备
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
    
    ' 下完装备等一下
    Delay minDelay, maxDelay
    
    ' 解除屏蔽前台按键干扰
    Call dm.SetWindowState(hwnd300, 11)
    ' 解除绑定后台
    Call dm.UnBindWindow
End Function

' 快速上装备
Public Function UpEquip(equip As String)
    ' 保存鼠标位置
    Dim mouseX, mouseY
    Call dm.GetCursorPos(mouseX, mouseY)
    
    ' 绑定后台
    Call dm.BindWindow(hwnd300, "dx", "dx", "windows", 0)
    ' 屏蔽键鼠输入，防止干扰
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
    ' 分辨率判断
    If screenW = 1920 Then
        X = 1150
        Y = 366
        step = 68
    Else
        X = 1473
        Y = 545
        step = 68
    End If
    
    ' 切割装备
    Dim arr() As String
    arr() = Split(equip, "/")
    Dim i As Integer
    Dim e() As String
    Dim row As Integer, col As Integer
    
    ' 遍历装备
    For i = LBound(arr) To UBound(arr)
        e() = Split(arr(i), "-")
        row = Val(e(0))
        col = Val(e(1))
        
        dm.moveToEx X + (col - 1) * step, Y + (row - 1) * step, offset, offset
        Delay minDelay, maxDelay
        dm.RightClick
    Next i
    
    ' 换完装备等一下
    Delay minDelay, maxDelay
    
    ' 恢复键鼠输入
    Call BlockInput(False)
    ' 解除绑定后台
    Call dm.UnBindWindow
    
    ' 还原鼠标位置
    Call dm.moveTo(mouseX, mouseY)
End Function

' 字符串计数
Public Function CountChar(s As String, char As String) As Integer
    Dim arr
    arr = Split(s, char)
    CountChar = UBound(arr) - LBound(arr)
End Function

' 删除数组中索引
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

' 判断数组是否为空
Public Function IsNotEmpty(ByVal sArray As Variant) As Boolean
    Dim i     As Long
    IsNotEmpty = True
    On Error GoTo lerr:
    i = UBound(sArray)
    Exit Function
lerr:
    IsNotEmpty = False
End Function
