VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{75AACB43-0B9D-11D2-95B5-0000B43369D3}#1.2#0"; "ARFrmExt.ocx"
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Presupuestos.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      ToolTipText     =   "Saltar Introducción e ir al menú"
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cerrar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form5.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ARFormExtenderCtrl.ARFormExtender ARFormExtender1 
      Left            =   3120
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      BackgroundType  =   1
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _cx             =   10610
      _cy             =   5953
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ExactFit"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ResHorizontal As Integer
Public ResVertical As Integer

Private Sub chameleonButton1_Click()
    Unload Me
    Form4.Show
End Sub

Private Sub Form_Load()

'REPRODUCE VIDEO
ShockwaveFlash1.Movie = App.Path + "\PRESENTACION.swf"
ShockwaveFlash1.Loop = False
End Sub

