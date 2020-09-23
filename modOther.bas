Attribute VB_Name = "modOther"
'* Programmed by : Vivek Patel                       *
'* Contact :  Email   => vivek_patel9@rediffmail.com *
'*            Website => www.VIVEKPATEL.cjb.net      *

'=====================================================
'*****************************************************
'* Vote For Me : If you really enjoy this utility or *
'                 helped by any of the functionality *
'                 than plz. reward us by your VOTE.  *
'*****************************************************
'=====================================================



Option Explicit

Public sUserName As String 'UserName String to display welcome msg
Public flagCloseAll As Boolean

'####################################################
'Declaration for making form as TOPMOST
' [ Start ]
'####################################################

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'####################################################
'Declaration for making form as TOPMOST
' [ End ]
'####################################################


'Code called when Logged Off Even Occured
Public Sub loggedOff()
    Call Main
    DoEvents
    frmLogin.Show
    DoEvents
End Sub

