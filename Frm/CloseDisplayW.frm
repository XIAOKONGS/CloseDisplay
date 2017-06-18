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
      Caption         =   "�ر���ʾ��"
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
         Caption         =   "����"
      End
      Begin VB.Menu xuanxiang3 
         Caption         =   "����"
      End
      Begin VB.Menu xuanxiang1 
         Caption         =   "�˳�"
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
'�����ر��Ż��Ľ���
'��Ȩ���� XIAOKONGS 2017

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
 
'��ʾ�����ر���������
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

'�ر� ��ʾ��
Public Function MonitorOff(Form As Form)
    
    Call SendScreenMessage(Form.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_OFF)

End Function

'������ʾ��
Public Function MonitorOn(Form As Form)
    
    Call SendScreenMessage(Form.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_ON)

End Function

'�ر���ʾ����Դ :)---���˯��
Public Function MonitorPowerDown(Form As Form)
    
    Call SendScreenMessage(Form.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_LOWPOWER)
    
End Function

'��ѯ��ʾ��״̬'��Ҫ���� Microsoft WMI Scipting V1.2 Library
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
           St = "����"
        Case 2
           St = "δ֪"
        Case 3
           St = "����"
        Case 4
           St = "����"
        Case 5
           St = "����"
        Case 6
           St = "������"
        Case 7
           St = "�رյ�Դ"
        Case 8
           St = "����"
        Case 9
           St = "�°�"
        Case 10
           St = "�˻�"
        Case 11
           St = "δ��װ"
        Case 12
           St = "��װ����"
        Case 13
           St = "ʡ��-δ֪" '��װ�ñ���Ϊ����ʡ��ģʽ������ȷ����ݲ�����
        Case 14
           St = "ʡ��-�͹���" '��װ������ʡ��״̬������Ȼ���������ܻ�����˻��ı��֡�
        Case 15
           St = "ʡ��-����" '���豸�����������У�������ʹȫ������Ѹ��
        Case 16
           St = "����ѭ��"
        Case 17
           St = "ʡ�羯��" '��װ������Ԥ��״̬����ȻҲ����ʡ��ģʽ��
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
'����PopupMenu����
  If Button And vbRightButton Then
     CloseDisplayW.PopupMenu b, 0, X, Y '�����˵�
  End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 Dim i As Long
 For i = 1 To Data.Files.Count '�����ȡ�ļ�·��
        Debug.Print Data.Files(i)
    Next
End Sub

Private Sub xuanxiang1_Click()
Unload Me
End Sub
