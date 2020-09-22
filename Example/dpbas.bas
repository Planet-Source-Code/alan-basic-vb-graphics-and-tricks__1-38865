Attribute VB_Name = "AlanAKAdp"

'   _______     __           _______     ___    ____
'  /  __  \\   |  ||        /  __  \\   |  \\  /  //
' |  //_\  ||  |  ||       |  //_\  ||  |   \\/  //
' |   __   ||  |  ||____   |   __   ||  | |\    //
' |__|| |__||  |_______||  |__|| |__||  | ||\__//

'         _______     ___  ___    _______
'        /  __  \\   |  ||/ //   /  __  \\
'       |  //_\  ||  |  |/ //   |  //_\  ||
'       |   __   ||  |   _ \\   |   __   ||
'       |__|| |__||  |__||\_\\  |__|| |__||
  
'             _____       _____
'            | __  \\    |  __ \\
'            | || \ \\   | // \ \\
'            | || | ||   | \\_/ //
'            | ||/ //    |  ___//
'            |    //     | ||
'            |___//      |_||


Option Explicit
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Dim ArrayNum As Integer
Public Filename As String
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const conSwNormal = 1
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Public Const WM_COMMAND = &H111
Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0
Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_SETCURSEL = &H186
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5
Public Const VK_SPACE = &H20
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Sub MoveForm(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub

Sub Sleep(interval)
'Pause The Form
Dim atime
atime = Timer
Do While Timer - atime < Val(interval)
DoEvents
Loop
End Sub



Sub Form_ExitUp(Form As Form)
'Makes Ur Form Exit Upwards
Do Until Form.Top <= -5000
Form.Top = Trim(Str(Int(Form.Top) - 100))
Loop
Unload Form

End Sub

Sub Form_ExitDown(Form As Form)
'Makes Ur Form Exit Downwards
Do Until Form.Top >= 13000
Form.Top = Trim(Str(Int(Form.Top) + 100))
Loop
Unload Form
End
End Sub







Public Sub OpenLink(Form As Form, URL As String)
ShellExecute Form.hWnd, "Open", URL, "", "", 1
End Sub

Public Sub LoadList(dlgCommon As CommonDialog, List As ListBox)
Dim MyString$


    On Error GoTo OpenErr
    dlgCommon.Filter = "Text Files (*.txt)|*.txt" 'sets the file type
    dlgCommon.Filename = "" 'default filename
    dlgCommon.ShowOpen
    Open dlgCommon.Filename For Input As #1
   While Not EOF(1)
        Input #1, MyString$
        DoEvents
        List.AddItem (MyString$)
    Wend
    '-Close the file
    Close #1
    Close #1 'closes the fil
OpenErr:
End Sub

Public Sub SaveList(List1 As ListBox, dlgCommon As CommonDialog)

Dim SaveList As Long
    On Error GoTo SaveErr
    dlgCommon.FLAGS = cdlOFNOverwritePrompt + cdlOFNPathMustExist
    dlgCommon.Filter = "Text Files (*.txt)|*.txt"
    dlgCommon.ShowSave
    Open dlgCommon.Filename For Output As #1
     For SaveList& = 0 To List1.ListCount - 1
        Print #1, (List1.List(SaveList&))
    Next SaveList&
   
    Close #1
    
SaveErr:
End Sub
Public Sub ListCommand(List1 As ListBox, Text1 As TextBox)
Dim Scrll As Integer, Num As Integer, Str As String

Num% = 0
For Scrll% = 0 To List1.ListCount - 1
Str$ = List1.List(Scrll%)
If Num% >= 5 Then
Num% = 0
End If
Text1 = (Str)
Num% = Num% + 1
DoEvents
Next

End Sub
