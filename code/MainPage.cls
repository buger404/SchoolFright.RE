VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(255, 0, 0, 0) '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Page.DrawImage "LOGO.png", GW / 2, GH / 2, alpha:=Abs((GetTickCount Mod 3000) / 1500 - 1), pos:=posOnCenter
    
    Page.Writes "> 关于 <", GW - 140, 20, 24, argb(255, 255, 255, 255), 120, 40, StringAlignmentCenter
    If CheckMouse2 = mMouseUp Then
        ECore.SimpleMsg "Error 404,喜欢把BUG写出艺术的智障BUG专家。" & vbCrLf & "2019级垃圾苦逼高一学生，现已转C# Unity。" & vbCrLf & vbCrLf & "该游戏属于2017年的早期作品“校园惊魂”的复刻" & vbCrLf & "使用了最新的开源项目Emerald进行重置" & vbCrLf & vbCrLf & "Github: https://github.com/buger404", "关于", StrArray("好")
        Mouse.State = 0: Mouse.button = 0
        Exit Sub
    End If
    
    Page.Writes "该游戏由Error404 2017年的“校园惊魂”游戏使用Emerald重置", GW / 2, GH - 50, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    If CheckMouse(0, 0, GW, GH) = mMouseUp Then
        ECore.NewTransform transFallDark, , "GamePage"
        CGamePage.LoadEvent 0
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.path & "\assets\ui", 0.8
    '创建页面
    ECore.Add Page, "MainPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
