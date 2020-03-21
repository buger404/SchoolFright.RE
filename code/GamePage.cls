VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GamePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Private Type Img
        X As Long
        y As Long
        Name As String
        FadeTime As Long
        Visible As Boolean
    End Type
    Private Type buttons
        Content As String
        Tag As String
    End Type
    Dim Imgs() As Img, Btn() As buttons
    Dim TimeLine As Long, ChoiceTime As Long, TimeCount As Long, GameTime As Long
    Dim Passage As String, BKColor As Long, DisableUI As Boolean, Died As Boolean
    Dim LastChoice As String, LastChoiceTime As Long
'======================================
Public Sub Choice(index As Integer)
    Dim Params() As String
    ReDim Imgs(0)
    Died = False
    ECore.NewTransform transFadeIn, 1000: LastChoice = Btn(index).Content: LastChoiceTime = GetTickCount
    If InStr(Btn(index).Tag, "dead*") Then
        DisableUI = True: Died = True: ChoiceTime = 0
        Params = Split(Btn(index).Tag, "dead*")
        Passage = Params(1)
        ReDim Btn(0)
        RunStoryMovie App.path & "\assets\animator\Event" & TimeLine & "-" & index + 1 & ".txt"
        FakeWait 1000
        ReDim Imgs(0)
        MusicList.Play "GameOver.wav"
        Btn(0).Tag = TimeLine: Btn(0).Content = "从上一个位置复活"
        TimeLine = 0:  GameTime = GetTickCount
        DisableUI = False
    ElseIf Btn(index).Tag = "sendwin" Then
        ShellExecuteA GHwnd, "open", Replace("http://connect.qq.com/widget/shareqq/index.html?url=https://rsinter.000webhostapp.com/school&suin=827466598&_wv=2098179&f=3&desc=&title=校园惊魂---我通关啦！&summary=面对魔鬼勇敢不屈！花了#Time#终于完美通关~撒花~撒花~&pics=https://rsinter.000webhostapp.com/school/LOGO2.png&flash=&site=QQ%E7%BA%A2%E5%8C%85&style=201&width=32&height=32", "#Time#", GameTime & "秒"), "", "", 1
    Else
        TimeCount = TimeCount + 1: TimeLine = Val(Btn(index).Tag)
        ChoiceTime = IIf(TimeLine = 50, 0, GetTickCount)
        LoadEvent TimeLine
    End If
End Sub
Public Sub LoadEvent(ByVal index As String)
    If Died Then Exit Sub

    Dim Chara As String
    Chara = App.path & "\assets\character.ini"
    
    If index = 0 Then
        GameTime = GetTickCount
        BGM.Create App.path & "\Music\Backround.mp3"
        BGM.Play
    End If
    
    Dim Count As Long, Code As String
    Passage = ReadINI("Event" & index, "msg", Chara)
    Count = Val(ReadINI("Event" & index, "max", Chara))
    ReDim Btn(0)
    RunStoryMovie App.path & "\assets\animator\Event" & index & ".txt"

    ReDim Btn(Count - 1)
    For I = 0 To Count - 1
        Code = ReadINI("Event" & index, "button" & I + 1, Chara)
        Btn(I).Content = Split(Code, ";")(0): Btn(I).Tag = Split(Code, ";")(1)
    Next
    
    ChoiceTime = IIf(TimeLine = 50, 0, GetTickCount)
End Sub
Public Sub FakeWait(ByVal Time As Long)
    For I = 1 To Time / 24
        ECore.Display: DoEvents
        Sleep 16
    Next
End Sub
Public Sub ExecCode(ByVal Code As String)
    Dim Cmd() As String, Params() As String
    Cmd = Split(Code, " ")
    Debug.Print Now, "Exec:" & Code
    Select Case Cmd(0)
        Case "Exit"
            Params = Split(Cmd(1), ";")
            Select Case Params(1)
                Case "正常": Imgs(Val(Params(0))).Visible = False
                Case "淡出": Imgs(Val(Params(0))).Visible = False: Imgs(Val(Params(0))).FadeTime = GetTickCount: FakeWait 1000
            End Select
        Case "Play"
            MusicList.Play (Cmd(1))
        Case "Wait"
            FakeWait Val(Cmd(1))
        Case "PlayB": BGM.Create (App.path & "\Music\" & Cmd(1)): BGM.Play
        Case "Load"
            Params = Split(Cmd(1), ";")
            Dim d() As String, e() As String, S As Integer
            d = Split(Params(0), "|"): e = Split(d(2), ",")
            For S = 1 To Val(d(1))
                ReDim Preserve Imgs(UBound(Imgs) + 1)
                With Imgs(UBound(Imgs))
                    .Name = d(0) & ".png"
                    If e(0) = "Screen" Then
                        .X = Val(e(2)): .y = Val(e(3))
                    Else
                        Dim u As Integer
                        For u = 0 To 3
                            e(u) = Replace(Replace(e(u), "#Width", 10515), "#Height", 6195)
                        Next
                        Dim p As Object, fi As String
                        Set p = CreateObject("MSScriptControl.ScriptControl")
                        p.language = "VBScript"
                        .X = p.Eval(e(2)): .y = p.Eval(e(3)): fi = GetTickCount
                        .X = .X / 10515 * GW: .y = .y / 6195 * GH
                        Page.Res.newImage App.path & "\assets\story\" & .Name, .X, .y, fi & ".png"
                        .Name = fi & ".png"
                        .X = p.Eval(e(0)): .y = p.Eval(e(1))
                        .X = .X / 10515 * GW: .y = .y / 6195 * GH
                        Set p = Nothing
                    End If
                    Select Case Params(1)
                        Case "正常": .Visible = True
                        Case "淡入": .Visible = True: .FadeTime = GetTickCount: FakeWait 1000
                    End Select
                End With
            Next
        Case "SetBackColor"
            Params = Split(Cmd(1), ",")
            BKColor = argb(255, Val(Params(0)), Val(Params(1)), Val(Params(2)))
    End Select
End Sub
Public Sub RunStoryMovie(ByVal File As String)
    'On Error GoTo ErrHandle
    DisableUI = True
    
    Dim Inputs As String
    
    ReDim Imgs(0)
    Open File For Input As #1
    Do While Not EOF(1)
        Line Input #1, Inputs
        Call ExecCode(Inputs)
        ECore.Display: DoEvents
    Loop
    Close #1

    DisableUI = False: ECore.NewTransform transFadeIn, 1000

    Exit Sub
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox "There was something wrong with the story movie ." & vbCrLf & "FileName: " & File & vbCrLf & "Line: " & UBound(Imgs) & vbCrLf & "MovieCode: " & Code & vbCrLf & "ErrorNumber:" & Err.Number, 48
        End
    End If
End Sub

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    Dim I As Integer
    
    If Not BGM.IsPlaying Then BGM.Play
    
    Page.Clear BKColor '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Dim Alp As Single, pro As Single
    For I = 1 To UBound(Imgs)
        Alp = IIf(Imgs(I).Visible, 1, 0)
        If Imgs(I).FadeTime <> 0 Then
            pro = (GetTickCount - Imgs(I).FadeTime) / 1000
            If pro > 1 Then pro = 1: Imgs(I).FadeTime = 0
            Alp = IIf(Imgs(I).Visible, 0, 1) + Cubic(pro, 0, 1, 1, 1) * IIf(Imgs(I).Visible, 1, -1)
        End If
        Page.DrawImage Imgs(I).Name, Imgs(I).X, Imgs(I).y, alpha:=Alp
    Next
    
    If Not DisableUI Then Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(100, 20, 0, 0)
    
    Page.DrawImage "back.png", 0, 0, alpha:=IIf(DisableUI, 0.5, 0.3)
    
    Dim cpro As Single, CS As Long
    If LastChoiceTime <> 0 Then
        cpro = (GetTickCount - LastChoiceTime) / 1000
        If cpro > 1 Then cpro = 1: LastChoiceTime = 0
        cpro = Cubic(cpro, 0, 0, 0, 1)
        Page.Writes LastChoice, GW / 2, GH / 2 - 64 / 2, 64, argb((1 - cpro) * 255, 255, 0, 0), align:=StringAlignmentCenter
    End If
    
    Page.Writes ToTime((GetTickCount - GameTime) / 1000), 20, 20, 24, argb(255, 255, 255, 255), style:=FontStyleBold
    
    If DisableUI Then Exit Sub
    
    Dim BY As Long, m As Integer
    BY = GH - 100
    
    For I = 0 To UBound(Btn)
        m = CheckMouse(0, BY - 3, GW, 40)
        If m Then Page.Paint 0, 0, BY - 3, GW, 40, argb(100, 255, 0, 0)
        Page.Writes IIf(m, "> ", "") & Btn(I).Content & IIf(m, " <", ""), GW / 2, BY, 20, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        BY = BY - 45
        If m = 3 Then
            Choice I
            Exit Sub
        End If
    Next
    
    Page.Writes Passage, 150, BY - 70, 20, argb(255, 255, 255, 255), GW - 300, 0, StringAlignmentCenter, FontStyleRegular

    If ChoiceTime = 0 Then Exit Sub

    Dim SavePro As Single
    SavePro = (GetTickCount - ChoiceTime) / 10000
    If SavePro > 1 Then SavePro = 1: Choice (UBound(Btn))
    SavePro = Cubic(SavePro, 0, 1, 1, 1)
    
    Page.Paint 0, GW / 2, GH - 19, GW * (1 - SavePro), 21, argb(255 * SavePro, 255, 0, 0), pos:=posOnCenter
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
    Page.Res.NoLoadCrash = True
    Page.Res.NewImages App.path & "\assets\bg", 1166, 777
    Page.Res.NewImages App.path & "\assets\story", 1166, 777
    '创建页面
    ECore.Add Page, "GamePage"
    BKColor = argb(255, 0, 0, 0)
    ReDim Imgs(0): ReDim Btn(0)
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
