VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "ѧ��ϵͳ��������2021.3.20"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   23340
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "ǩ��"
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
      Caption         =   "ǩ��"
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
      Caption         =   "���ߣ� 150149"
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
        MsgBox "����δ����"
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
        MsgBox "����δ����"
    End If
End Sub

Private Sub Form_Load()
    WebBrowser1.Silent = True
    dengdai = 0
    
    ListView1.ListItems.Clear               '����б�
    ListView1.ColumnHeaders.Clear           '����б�ͷ
    ListView1.View = lvwReport              '�����б���ʾ��ʽ
    ListView1.GridLines = True              '��ʾ������

    ListView1.FullRowSelect = False        'ѡ������
    
    ListView1.ColumnHeaders.Add , , "ѧ��", 2000
    ListView1.ColumnHeaders.Add , , "����", 2000
    ListView1.ColumnHeaders.Add , , "״̬", 2000
    
    Dim tempStr As String '�������tempStrΪ�ַ���
    Dim strs
    Dim i As Integer
    i = 0
    Open "�˺�����.txt" For Input As #1 '���ļ�
    While Not EOF(1)  '��ȡ������
        Line Input #1, tempStr '��ȡһ�е�����tempStr
        strs = Split(tempStr, " ")
        i = i + 1
        ListView1.ListItems.Add , , strs(0)
        ListView1.ListItems(i).SubItems(1) = strs(1)
        
    Wend 'δ��������
    Close #1 '�ر�
End Sub

Private Sub Timer1_Timer()
    dengdai = dengdai + 1
    If WebBrowser1.Busy Or dengdai < 3 Then
        Debug.Print ("�ȴ���ҳ������ɵ����¼")
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
        Debug.Print ("�ȴ���ҳ���������д�˺�����")
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
        Debug.Print ("���еı�ǩ����: " & CStr(vDoc.All.Length))
        For X = 0 To vDoc.All.Length - 1 '������б�ǩ
            timea = timea + 1
            If UCase(vDoc.All(X).TagName) = "BUTTON" Then  '�ҵ�input��ǩ
                Set VTag = vDoc.All(X)
                If VTag.Type = "submit" Then VTag.Click '����ύ�ˣ�һ�ж�OK��
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
        Debug.Print ("�ȴ���ҳ������ɽ���ǩ��ǩ��ҳ��")
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
        Debug.Print ("�ȴ���ҳ�������ѡ���λ")
        Exit Sub
    Else
        Timer4.Enabled = False
        dengdai = 0
        
        Dim str As String
        str = WebBrowser1.Document.documentelement.outerhtml

        Debug.Print ("��¼���ڼ��: " & CStr(InStr(str, "��ӭʹ�ö�ݸ��ѧԺ������֤��¼ϵͳ")))
        If InStr(str, "��ӭʹ�ö�ݸ��ѧԺ������֤��¼ϵͳ") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Debug.Print ("��⵽��½����")
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
        
        Debug.Print ("���еı�ǩ����: " & CStr(vDoc.All.Length))
        For X = 0 To vDoc.All.Length - 1 '������б�ǩ
            timea = timea + 1
            
            Debug.Print UCase(vDoc.All(X).TagName)
            
            If UCase(vDoc.All(X).TagName) = "SELECT" Then '�ҵ�input��ǩ
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
        Debug.Print ("�ȴ���ҳ������ɵ��ǩ����ť")
        Exit Sub
    Else
        Timer5.Enabled = False
        dengdai = 0
        
        Dim str As String
        str = WebBrowser1.Document.documentelement.outerhtml

        Debug.Print ("��¼���ڼ��: " & CStr(InStr(str, "��ӭʹ�ö�ݸ��ѧԺ������֤��¼ϵͳ")))
        If InStr(str, "��ӭʹ�ö�ݸ��ѧԺ������֤��¼ϵͳ") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Timer5.Enabled = False
            Debug.Print ("��⵽��½����")
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
        
        Debug.Print ("���еı�ǩ����: " & CStr(vDoc.All.Length))
        For X = 0 To vDoc.All.Length - 1 '������б�ǩ
            timea = timea + 1
            
            If UCase(vDoc.All(X).TagName) = "INPUT" Then '�ҵ�input��ǩ
                Set VTag = vDoc.All(X)
                If VTag.Type = "button" And VTag.Value = "ǩ��" Then '�ҵ�ȷ����ť��
                    VTag.onclick
                    Debug.Print "�����ť���"
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
        Debug.Print ("�ȴ���ҳ������ɵ��ǩ�˰�ť")
        Exit Sub
    Else
        Timer6.Enabled = False
        dengdai = 0
        
        Dim str As String
        str = WebBrowser1.Document.documentelement.outerhtml

        Debug.Print ("��¼���ڼ��: " & CStr(InStr(str, "��ӭʹ�ö�ݸ��ѧԺ������֤��¼ϵͳ")))
        If InStr(str, "��ӭʹ�ö�ݸ��ѧԺ������֤��¼ϵͳ") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Timer5.Enabled = False
            Timer6.Enabled = False
            Debug.Print ("��⵽��½����")
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
        
        Debug.Print ("���еı�ǩ����: " & CStr(vDoc.All.Length))
        For X = 0 To vDoc.All.Length - 1 '������б�ǩ
            timea = timea + 1
            
            If UCase(vDoc.All(X).TagName) = "INPUT" Then '�ҵ�input��ǩ
                Set VTag = vDoc.All(X)
                If VTag.Type = "button" And VTag.Value = "ǩ��" Then '�ҵ�ȷ����ť��
                    VTag.onclick
                    Debug.Print "�����ť���"
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
        Debug.Print ("�ȴ���ҳ������ɼ���Ƿ�ǩ�����")
        Exit Sub
    Else
        Timer7.Enabled = False
        dengdai = 0
        
        Dim str As String
        str = WebBrowser1.Document.documentelement.outerhtml

        If InStr(str, "�ϸ�ǩ�����") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Timer5.Enabled = False
            Timer6.Enabled = False
            Debug.Print ("��⵽ǩ�����")
            If dijige = 0 Then dijige = 1
            ListView1.ListItems(dijige).SubItems(2) = "�ϸ�ǩ�����"
            WebBrowser1.Navigate "https://cas.dgut.edu.cn/user/logout?service=http://stu.dgut.edu.cn:80"
            'Timer2.Enabled = True
            'Timer2.Interval = 1000
            dengdai = 0
            
            Timer9.Enabled = True
            Timer9.Interval = 500
            Exit Sub
        ElseIf InStr(str, "�����ظ�ǩ��") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Timer5.Enabled = False
            Timer6.Enabled = False
            Debug.Print ("��⵽ǩ���ظ�")
            If dijige = 0 Then dijige = 1
            ListView1.ListItems(dijige).SubItems(2) = "ǩ�����ȷ��"
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
        Debug.Print ("�ȴ���ҳ������ɼ���Ƿ�ǩ�����")
        Exit Sub
    Else
        Timer8.Enabled = False
        dengdai = 0
        
        Dim str As String
        str = WebBrowser1.Document.documentelement.outerhtml

        If InStr(str, "�¸�ǩ�����") > 0 Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer2.Enabled = False
            Timer4.Enabled = False
            Timer5.Enabled = False
            Timer6.Enabled = False
            Debug.Print ("��⵽ǩ�����")
            If dijige = 0 Then dijige = 1
            ListView1.ListItems(dijige).SubItems(2) = "�¸�ǩ�����"
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
        Debug.Print ("�ȴ���ҳ������ɽ�����һ��ǩ��")
        Exit Sub
    Else
        Timer9.Enabled = False
        dengdai = 0
        
        WebBrowser1.Navigate "http://stu.dgut.edu.cn/student/partwork/attendance.jsp"
        
        Timer3.Enabled = True
        Timer3.Interval = 500

    End If
End Sub

