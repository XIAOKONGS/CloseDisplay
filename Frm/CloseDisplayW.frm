VERSION 5.00
Begin VB.Form CloseDisplayW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XK"
   ClientHeight    =   420
   ClientLeft      =   16650
   ClientTop       =   9285
   ClientWidth     =   1260
   Icon            =   "CloseDisplayW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   420
   ScaleWidth      =   1260
   Begin VB.CommandButton Command1 
      Caption         =   "关闭显示器"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "XIAOKONGS"
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
   Begin VB.Menu b 
      Caption         =   "XIAOKONGS"
      Visible         =   0   'False
      Begin VB.Menu xuanxiang2 
         Caption         =   "升级"
      End
      Begin VB.Menu xuanxiang3 
         Caption         =   "关于"
      End
      Begin VB.Menu xuanxiang1 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "CloseDisplayW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This version only supports Windows 7
'经过特别优化的界面
'版权所有 XIAOKONGS 2017

Private Declare Function SendScreenMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Const MONITOR_ON = -1&
Private Const MONITOR_LOWPOWER = 1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_CHILD = 5
Private Const GW_OWNER = 4
Private Const GW_MAX = 5
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SYSMENU = &H80000
Private Enum ESetWindowPosStyles
        SWP_SHOWWINDOW = &H40
        SWP_HIDEWINDOW = &H80
        SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
        SWP_NOACTIVATE = &H10
        SWP_NOCOPYBITS = &H100
        SWP_NOMOVE = &H2
        SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
        SWP_NOREDRAW = &H8
        SWP_NOREPOSITION = SWP_NOOWNERZORDER
        SWP_NOSIZE = &H1
        SWP_NOZORDER = &H4
        SWP_DRAWFRAME = SWP_FRAMECHANGED
        HWND_NOTOPMOST = -2
End Enum
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
 
Dim xa As Single, ya As Single
 
'显示、隐藏标题栏函数
Public Function ShowTitleBar(chenjl1031 As Form, ByVal bState As Boolean)
         Dim lStyle As Long
         Dim tR As RECT
         'Dim playscreen As Variant
         On Error Resume Next
         GetWindowRect chenjl1031.hwnd, tR
         lStyle = GetWindowLong(chenjl1031.hwnd, GWL_STYLE)
         If (bState) Then
            If chenjl1031.ControlBox Then
               lStyle = lStyle Or WS_SYSMENU
            End If
            If chenjl1031.MaxButton Then
               lStyle = lStyle Or WS_MAXIMIZEBOX
            End If
            If chenjl1031.MinButton Then
               lStyle = lStyle Or WS_MINIMIZEBOX
            End If
            If chenjl1031.Caption <> "" Then
               lStyle = lStyle Or WS_CAPTION
            End If
         Else
            lStyle = lStyle And Not WS_SYSMENU
            lStyle = lStyle And Not WS_MAXIMIZEBOX
            lStyle = lStyle And Not WS_MINIMIZEBOX
            lStyle = lStyle And Not WS_CAPTION
         End If
         SetWindowLong chenjl1031.hwnd, GWL_STYLE, lStyle
'         SetWindowPos chenjl1031.hwnd, 0, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
         chenjl1031.Refresh
End Function

'关闭 显示器
Public Function MonitorOff(Form As Form)
    
    Call SendScreenMessage(Form.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_OFF)

End Function

'开启显示器
Public Function MonitorOn(Form As Form)
    
    Call SendScreenMessage(Form.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_ON)

End Function

'关闭显示器电源 :)---深度睡眠
Public Function MonitorPowerDown(Form As Form)
    
    Call SendScreenMessage(Form.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_LOWPOWER)
    
End Function

'查询显示器状态'需要引用 Microsoft WMI Scipting V1.2 Library
Public Function WMIVideoControllerInfo() As Long
    Dim WMIObjSet As SWbemObjectSet
    Dim obj As SWbemObject
    Dim St As String
    
    Set WMIObjSet = GetObject("winmgmts:{impersonationLevel=impersonate}"). _
                        InstancesOf("Win32_VideoController")
    
    On Local Error Resume Next
    
    
    For Each obj In WMIObjSet
        WMIVideoControllerInfo = obj.Availability
        
        Select Case WMIVideoControllerInfo
        Case 1
           St = "其他"
        Case 2
           St = "未知"
        Case 3
           St = "运行"
        Case 4
           St = "警告"
        Case 5
           St = "试验"
        Case 6
           St = "不适用"
        Case 7
           St = "关闭电源"
        Case 8
           St = "离线"
        Case 9
           St = "下班"
        Case 10
           St = "退化"
        Case 11
           St = "未安装"
        Case 12
           St = "安装错误"
        Case 13
           St = "省电-未知" '该装置被称为是在省电模式，但其确切身份不明。
        Case 14
           St = "省电-低功耗" '该装置是在省电状态，但仍然运作，可能会出现退化的表现。
        Case 15
           St = "省电-待命" '该设备不能正常运行，但可以使全部力量迅速
        Case 16
           St = "动力循环"
        Case 17
           St = "省电警告" '该装置是在预警状态，虽然也处于省电模式。
        End Select
    Next
End Function
Private Sub Command1_Click()
MonitorOff Me
End Sub

Private Sub Form_Load()
Dim i
Dim a
i = WMIVideoControllerInfo
ShowTitleBar CloseDisplayW, False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
'利用PopupMenu方法
  If Button And vbRightButton Then
     CloseDisplayW.PopupMenu b, 0, X, Y '弹出菜单
  End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 Dim i As Long
 For i = 1 To Data.Files.Count '逐个读取文件路径
        Debug.Print Data.Files(i)
    Next
End Sub

Private Sub xuanxiang1_Click()
Unload Me
End Sub
