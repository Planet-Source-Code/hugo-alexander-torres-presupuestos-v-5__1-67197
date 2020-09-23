VERSION 5.00
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clave de Acceso"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   FillStyle       =   0  'Solid
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox dbUserId 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmLogin.frx":1042
      Left            =   2040
      List            =   "frmLogin.frx":1044
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Seleccione un Nombre"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Clave Secreta"
      Top             =   2160
      Width           =   1695
   End
   Begin Presupuestos.chameleonButton cmdCancel 
      Height          =   405
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "Salir de la Aplicaciòn"
      Top             =   2760
      Width           =   1455
      _extentx        =   2566
      _extenty        =   714
      btype           =   3
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmLogin.frx":1046
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmLogin.frx":1072
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin Presupuestos.chameleonButton cmdLogin 
      Height          =   405
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
      _extentx        =   2566
      _extenty        =   714
      btype           =   3
      tx              =   "&Aceptar"
      enab            =   -1  'True
      font            =   "frmLogin.frx":1090
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmLogin.frx":10BC
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label lblPasswd 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BorderWidth     =   3
      FillColor       =   &H80000004&
      Height          =   2535
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   600
      Picture         =   "frmLogin.frx":10DA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Presup Ver 5.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   850
      Left            =   3480
      Picture         =   "frmLogin.frx":19A4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   750
   End
   Begin VB.Image imgBorder 
      Height          =   900
      Left            =   -240
      Picture         =   "frmLogin.frx":2177
      Top             =   0
      Width           =   4770
   End
   Begin VB.Label lblUserId 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   2940
      Left            =   0
      Picture         =   "frmLogin.frx":101C9
      Top             =   960
      Width           =   7965
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Srchflag As Boolean
      
Private Sub Form_Load()
DoEvents
    
           
    DoEvents
' Cambia resolucion de pantalla a 1024x768
ChangeScreenSettings 1024, 768, 32
On Error Resume Next
        
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    txtPasswd = ""
    
    Dim i As Integer
    
    For i = 0 To rsUSUARIO.RecordCount
        dbUserId.AddItem rsUSUARIO(0)
        rsUSUARIO.MoveNext
        If rsUSUARIO.EOF = True Then
             rsUSUARIO.MoveLast
             dbUserId.ListIndex = 0 'For displaying by default admin
             Exit Sub
        End If
    Next i
        
    Srchflag = False 'Initially Search is not found

End Sub
      
      
'Check for Valid Password Corresponding to user Name
Private Sub cmdLogin_Click()
On Error Resume Next
    rsUSUARIO.Close  'After the user_id are loaded into dbUserId close the
                  'Recordset for further usage
    rsUSUARIO.Open "Select password from usuario where usuar = '" & dbUserId.Text & "'", cn, adOpenDynamic, adLockOptimistic
    
    If rsUSUARIO.EOF <> True Then 'If Search is found
         If rsUSUARIO(0) = txtPasswd Then
             sUserName = UCase(dbUserId.Text)
             Form4.StatusBar1.Panels(2).Text = "Usuario Conectado : " & sUserName
             DoEvents
              
             Unload Me
             DoEvents
             Form4.Show
            ' DoEvents
        '     Srchflag = True
             Exit Sub
             rsUSUARIO.Close
         Else
             MsgBox "Password no Válido!!!" & vbCrLf & "Clave Errada", vbInformation, "HugoSoft 2006"
             txtPasswd.Text = ""
             txtPasswd.SetFocus
             Exit Sub
         End If
    End If
    
    'If Srchflag = False Then 'Display msg when search not found
    '     MsgBox "Invalid Password" & vbCrLf & "No Access!!!", vbCritical, "Invalid User"
    '     End
    'End If
End Sub

'Simply Quit the Loaded Stuff!!!
Private Sub cmdCancel_Click()
On Error Resume Next
    End
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        cmdLogin.SetFocus
    End If
End Sub
Private Sub form_Unload(Cancel As Integer)
     ' Liberar Memoria
      cn.Close
     
End Sub
