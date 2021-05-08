VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "学工系统批量考勤2021.3.20"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   23340
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer9 
      Left            =   5400
      Top             =   120
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   11033
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer Timer8 
      Left            =   4800
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "签退"
      Height          =   735
      Left            =   7200
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Timer Timer7 
      Left            =   4200
      Top             =   120
   End
   Begin VB.Timer Timer6 
      Left            =   3600
      Top             =   120
   End
   Begin VB.Timer Timer5 
      Left            =   3000
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Left            =   2400
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Left            =   1800
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "签到"
      Height          =   735
      Left            =   7200
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6135
      Left            =   9000
      TabIndex        =   0
      Top             =   480
      Width           =   14175
      ExtentX         =   25003
      ExtentY         =   10821
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "作者： 150149"
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type WSADATA
        wversion As Integer
        wHighVersion As Integer
        szDescription(0 To 256) As Byte
        szSystemStatus(0 To 128) As Byte
        iMaxSockets As Integer
        iMaxUdpDg As Integer
        lpszVendorInfo As Long
    End Type
    Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
    Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
    Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHostname As String) As Long
    Private Const WS_VERSION_REQD = &H101
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
Public dengdai As Integer
Public dijige As Integer
Public isqiandao As Integer

    Public Function IsConnectedState() As Boolean
        Dim udtWSAD As WSADATA
        Call WSAStartup(WS_VERSION_REQD, udtWSAD)
        IsConnectedState = CBool(gethostbyname("www.baidu.com"))
        Call WSACleanup
    End Function

Private Sub Command1_Click()
    If IsConnectedState Then
        WebBrowser1.Navigate "http://stu.dgut.edu.cn/student/partwork/attendance.jsp"
    
        dijige = 0
    
        Timer3.Enabled = True
        Timer3.Interval = 1000
        
        Command1.Enabled = False
        Command2.Enabled = False
        isqiandao = 1
    Else
        MsgBox "网络未连接"
    End If
End Sub

Private Sub Command2_Click()
    If IsConnectedState Then
        WebBrowser1.Navigate "http://stu.dgut.edu.cn/student/partwork/attendance.jsp"
        
        dijige = 0
    
        Timer3.Enabled = True
        Timer3.Interval = 1000
        
        Command1.Enabled = False
        Command2.Enabled = False
        isqiandao = 0
    Else
        MsgBox "网络未连接"
    End If
End Sub

Private Sub Form_Load()
    WebBrowser1.Silent = True
    dengdai = 0
    
    ListView1.ListItems.Clear               '清空列表
    ListView1.ColumnHeaders.Clear           '清空列表头
    ListView1.View = lvwReport              '设置列表显示方式
    ListView1.GridLines = True              '显示网络线

    ListView1.FullRowSelect = False        '选择整行
    
    ListView1.ColumnHeaders.Add , , "学号", 2000
    ListView1.ColumnHeaders.Add , , "密码", 2000
    ListView1.ColumnHeaders.Add , , "状态", 2000
    
    Dim tempStr As String '定义变量tempStr为字符串
    Dim strs
    Dim i As Integer
    i = 0
    Open "账号数据.txt" For Input As #1 '打开文件
    While Not EOF(1)  '读取到结束
        Line Input #1, tempStr '读取一行到变量tempStr
        strs = Split(tempStr, " ")
        i = i + 1
        ListView1.ListItems.Add , , strs(0)
        ListView1.ListItems(i).SubItems(1) = strs(1)
        
    Wend '未结束继续
    Close #1 '关闭
End Sub

Private Sub Timer1_Timer()
    dengdai = dengdai + 1
    If WebBrowser1.Busy Or dengdai < 3 Then
        Debug.Print ("等待网页加载完成点击登录")
        Exit Sub
    Else
        Timer1.Enabled = False

        WebBrowser1.Document.querySelector("div.login-btn-text").Click
        dengdai = 0
        Timer2.Enabled = True
        Timer2.Interval = 500
    End If
End Sub

Private Sub Timer2_Timer()
    dengdai = dengdai + 1
    If WebBrowser1.Busy Or dengdai < 3 Then
        Debug.Print ("等待网页加载完成填写账号密码")
        Exit Sub
    Else
        Timer2.Enabled = False
        dengdai = 0
        
        dijige = dijige + 1
        If dijige = 0 Then dijige = 1
        If dijige > ListView1.ListItems.Count Then
            Timer2.Enabled = False
            Command1.Enabled = True
            Command2.Enabled = True
            Exit Sub
        End If
        WebBrowser1.Document.querySelector("#username").Value = ListView1.ListItems(dijige).Text
        WebBrowser1.Document.querySelector("#casPassword").Value = ListView1.ListItems(dijige).SubItems(1)

        Dim vDoc, X, VTag
        Dim timea As Integer
        timea = 0
        Set vDoc = WebBrowser1.Document
        Debug.Print ("所有的标签数量: " & CStr(vDoc.All.Length))
        For X = 0 To vDoc.All.Length - 1 '检测所有标签
            timea = timea + 1
            If UCase(vDoc.All(X).TagName) = "BUTTON" Then  '找到input标签
                Set VTag = vDoc.All(X)
                If VTag.Type = "submit" Then VTag.Click '点击提交了，一切都OK了
                timea = timea + 1
            End If
        Next X
                
        WebBrowser1.Navigate "http://stu.dgut.edu.cn/student/partwork/attendance.jsp"
        dengdai = 0
        Timer3.Enabled = True
        Timer3.Interval = 500

    End If
End Sub

Private Sub Timer3_Timer()
    dengdai = dengdai + 1
    If WebBrowser1.Busy Or dengdai < 3 Then
        Debug.Print ("等待网页加载完成进入签到签退页面")
        Exit Sub
    Else
        Timer3.Enabled = False
        dengdai = 0
        WebBrowser1.Navigate "http://stu.dgut.edu.cn/student/partwork/attendance.jsp"
        
        Timer4.Enabled = True
        Timer4.Interval = 100
    End If
End Sub

Private Sub Timer4_Timer()
    dengdai = dengdai + 1
    If WebBrowser1.Busy Or dengdai < 3 Then
        Debug.Print ("等待网页加载完成选择岗位")
        Exit Sub
    Else
        Timer4.Enabled = False
        dengdai = 0
        
        Dim str As String
        str = WebBrowser1.Document.documentelement.outerhtml

        Debug.Print ("登录过期检测: " & CStr(InStr(str, "欢迎使用东莞理工学院中央认证登录系统")))
        If InStr(str, "欢迎使用东莞理工学院中央认证登录系统") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Debug.Print ("检测到登陆过期")
            WebBrowser1.Navigate "http://stu.dgut.edu.cn/"
            Timer2.Enabled = True
            Timer2.Interval = 100
            dengdai = 0
            Exit Sub
        End If

        Dim vDoc, X, VTag
        Dim timea As Integer
        timea = 0
        Set vDoc = WebBrowser1.Document
        
        Debug.Print ("所有的标签数量: " & CStr(vDoc.All.Length))
        For X = 0 To vDoc.All.Length - 1 '检测所有标签
            timea = timea + 1
            
            Debug.Print UCase(vDoc.All(X).TagName)
            
            If UCase(vDoc.All(X).TagName) = "SELECT" Then '找到input标签
                    Set VTag = vDoc.All(X)
                    VTag.selectedIndex = 1
                    VTag.onchange
            End If
            
        Next X
        
        dengdai = 0
        If isqiandao = 1 Then
            Timer5.Enabled = True
            Timer5.Interval = 500
        ElseIf isqiandao = 0 Then
            Timer6.Enabled = True
            Timer6.Interval = 500
        End If
    End If
End Sub

Private Sub Timer5_Timer()
    dengdai = dengdai + 1
    If WebBrowser1.Busy Or dengdai < 3 Then
        Debug.Print ("等待网页加载完成点击签到按钮")
        Exit Sub
    Else
        Timer5.Enabled = False
        dengdai = 0
        
        Dim str As String
        str = WebBrowser1.Document.documentelement.outerhtml

        Debug.Print ("登录过期检测: " & CStr(InStr(str, "欢迎使用东莞理工学院中央认证登录系统")))
        If InStr(str, "欢迎使用东莞理工学院中央认证登录系统") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Timer5.Enabled = False
            Debug.Print ("检测到登陆过期")
            WebBrowser1.Navigate "http://stu.dgut.edu.cn/"
            Timer2.Enabled = True
            Timer2.Interval = 500
            dengdai = 0
            Exit Sub
        End If

        Dim vDoc, X, VTag
        Dim timea As Integer
        timea = 0
        Set vDoc = WebBrowser1.Document
        
        Debug.Print ("所有的标签数量: " & CStr(vDoc.All.Length))
        For X = 0 To vDoc.All.Length - 1 '检测所有标签
            timea = timea + 1
            
            If UCase(vDoc.All(X).TagName) = "INPUT" Then '找到input标签
                Set VTag = vDoc.All(X)
                If VTag.Type = "button" And VTag.Value = "签到" Then '找到确定按钮。
                    VTag.onclick
                    Debug.Print "点击按钮完成"
                    timea = timea + 1
                End If
            End If

        Next X
        Timer7.Enabled = True
        Timer7.Interval = 500

    End If
End Sub

Private Sub Timer6_Timer()
    dengdai = dengdai + 1
    If WebBrowser1.Busy Or dengdai < 3 Then
        Debug.Print ("等待网页加载完成点击签退按钮")
        Exit Sub
    Else
        Timer6.Enabled = False
        dengdai = 0
        
        Dim str As String
        str = WebBrowser1.Document.documentelement.outerhtml

        Debug.Print ("登录过期检测: " & CStr(InStr(str, "欢迎使用东莞理工学院中央认证登录系统")))
        If InStr(str, "欢迎使用东莞理工学院中央认证登录系统") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Timer5.Enabled = False
            Timer6.Enabled = False
            Debug.Print ("检测到登陆过期")
            WebBrowser1.Navigate "http://stu.dgut.edu.cn/"
            Timer2.Enabled = True
            Timer2.Interval = 500
            dengdai = 0
            Exit Sub
        End If

        Dim vDoc, X, VTag
        Dim timea As Integer
        timea = 0
        Set vDoc = WebBrowser1.Document
        
        Debug.Print ("所有的标签数量: " & CStr(vDoc.All.Length))
        For X = 0 To vDoc.All.Length - 1 '检测所有标签
            timea = timea + 1
            
            If UCase(vDoc.All(X).TagName) = "INPUT" Then '找到input标签
                Set VTag = vDoc.All(X)
                If VTag.Type = "button" And VTag.Value = "签退" Then '找到确定按钮。
                    VTag.onclick
                    Debug.Print "点击按钮完成"
                    timea = timea + 1
                End If
            End If

        Next X

        Timer8.Enabled = True
        Timer8.Interval = 500
    End If
End Sub

Private Sub Timer7_Timer()
    dengdai = dengdai + 1
    If WebBrowser1.Busy Or dengdai < 3 Then
        Debug.Print ("等待网页加载完成检测是否签到完成")
        Exit Sub
    Else
        Timer7.Enabled = False
        dengdai = 0
        
        Dim str As String
        str = WebBrowser1.Document.documentelement.outerhtml

        If InStr(str, "上岗签到完成") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Timer5.Enabled = False
            Timer6.Enabled = False
            Debug.Print ("检测到签到完成")
            If dijige = 0 Then dijige = 1
            ListView1.ListItems(dijige).SubItems(2) = "上岗签到完成"
            WebBrowser1.Navigate "https://cas.dgut.edu.cn/user/logout?service=http://stu.dgut.edu.cn:80"
            'Timer2.Enabled = True
            'Timer2.Interval = 1000
            dengdai = 0
            
            Timer9.Enabled = True
            Timer9.Interval = 500
            Exit Sub
        ElseIf InStr(str, "无需重复签到") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Timer5.Enabled = False
            Timer6.Enabled = False
            Debug.Print ("检测到签到重复")
            If dijige = 0 Then dijige = 1
            ListView1.ListItems(dijige).SubItems(2) = "签到完成确认"
            WebBrowser1.Navigate "http://stu.dgut.edu.cn/logout.jsp"
            'Timer2.Enabled = True
            'Timer2.Interval = 1000
            dengdai = 0
            Command1.Enabled = True
            
            Timer9.Enabled = True
            Timer9.Interval = 500
            Exit Sub
        End If

    End If
End Sub

Private Sub Timer8_Timer()
    dengdai = dengdai + 1
    If WebBrowser1.Busy Or dengdai < 3 Then
        Debug.Print ("等待网页加载完成检测是否签退完成")
        Exit Sub
    Else
        Timer8.Enabled = False
        dengdai = 0
        
        Dim str As String
        str = WebBrowser1.Document.documentelement.outerhtml

        If InStr(str, "下岗签退完成") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Timer5.Enabled = False
            Timer6.Enabled = False
            Debug.Print ("检测到签退完成")
            If dijige = 0 Then dijige = 1
            ListView1.ListItems(dijige).SubItems(2) = "下岗签退完成"
            WebBrowser1.Navigate "http://stu.dgut.edu.cn/logout.jsp"
            'Timer2.Enabled = True
            'Timer2.Interval = 1000
            dengdai = 0
            
            Timer9.Enabled = True
            Timer9.Interval = 500
            
            Exit Sub
        End If

    End If
End Sub

Private Sub Timer9_Timer()
    dengdai = dengdai + 1
    If WebBrowser1.Busy Or dengdai < 3 Then
        Debug.Print ("等待网页加载完成进入下一个签到")
        Exit Sub
    Else
        Timer9.Enabled = False
        dengdai = 0
        
        WebBrowser1.Navigate "http://stu.dgut.edu.cn/student/partwork/attendance.jsp"
        
        Timer3.Enabled = True
        Timer3.Interval = 500

    End If
End Sub

