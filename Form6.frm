VERSION 5.00
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{7B72A3F4-FE91-11D3-917E-E5E1F9477021}#2.0#0"; "3DLine.ocx"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5445
   ClientLeft      =   4050
   ClientTop       =   2430
   ClientWidth     =   11115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   11115
   Begin VB.Data Data7 
      BackColor       =   &H00FFFFC0&
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Width           =   8580
   End
   Begin VB.Frame Frame1 
      Caption         =   "CAPITULOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   2
      Left            =   0
      TabIndex        =   22
      Top             =   3240
      Width           =   2415
      Begin VB.ListBox List3 
         BackColor       =   &H00FFFFC0&
         Height          =   1710
         ItemData        =   "Form6.frx":030A
         Left            =   50
         List            =   "Form6.frx":030C
         TabIndex        =   23
         Top             =   240
         Width           =   2310
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1575
      Left            =   2520
      TabIndex        =   18
      Top             =   3840
      Width           =   8505
      Begin Presupuestos.chameleonButton Eliminar 
         Height          =   615
         Left            =   6120
         TabIndex        =   27
         ToolTipText     =   "Borrar Insumo Seleccionado"
         Top             =   240
         Width           =   2200
         _extentx        =   3889
         _extenty        =   1085
         btype           =   3
         tx              =   "Eliminar Insumo"
         enab            =   -1  'True
         font            =   "Form6.frx":030E
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "Form6.frx":0336
         picn            =   "Form6.frx":0354
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VB.TextBox subtotal 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3600
         TabIndex        =   21
         Text            =   "0"
         Top             =   700
         Width           =   2340
      End
      Begin Presupuestos.chameleonButton chameleonButton1 
         Height          =   615
         Left            =   6120
         TabIndex        =   19
         ToolTipText     =   "Regresar al Menú Principal"
         Top             =   900
         Width           =   2200
         _extentx        =   4683
         _extenty        =   1085
         btype           =   3
         tx              =   "Regresar al Menú"
         enab            =   -1
         font            =   "Form6.frx":0670
         coltype         =   1
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "Form6.frx":0698
         picn            =   "Form6.frx":06B6
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1320
         Left            =   120
         Picture         =   "Form6.frx":09D2
         Stretch         =   -1  'True
         ToolTipText     =   "HugoSoft 2003"
         Top             =   160
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Costo Directo ..... :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   20
         Top             =   700
         Width           =   1695
      End
   End
   Begin VB.Data Data5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Data5"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Consulta3"
      Top             =   3480
      Visible         =   0   'False
      Width           =   5100
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ocultos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox clainsu 
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox codinsu 
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblAncho 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   60
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2520
      TabIndex        =   5
      Top             =   0
      Width           =   8535
      Begin Presupuestos.chameleonButton chameleonButton2 
         Height          =   495
         Left            =   6600
         TabIndex        =   35
         ToolTipText     =   "Grabar Nuevo APU"
         Top             =   1440
         Width           =   1800
         _extentx        =   2461
         _extenty        =   873
         btype           =   3
         tx              =   "Crear Otro APU"
         enab            =   0
         font            =   "Form6.frx":B06C
         coltype         =   1
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "Form6.frx":B094
         picn            =   "Form6.frx":B0B2
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin DLine.ThreeDLine ThreeDLine1 
         Height          =   1095
         Left            =   3960
         TabIndex        =   34
         Top             =   840
         Width           =   45
         _ExtentX        =   79
         _ExtentY        =   1931
         Orientation     =   1
      End
      Begin VB.TextBox caninsumo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   31
         ToolTipText     =   "Cantidad Insumo"
         Top             =   1560
         Width           =   3015
      End
      Begin Presupuestos.chameleonButton Incluir 
         Height          =   495
         Left            =   6600
         TabIndex        =   28
         ToolTipText     =   "Grabar Insumo al APU"
         Top             =   840
         Width           =   1800
         _extentx        =   2566
         _extenty        =   873
         btype           =   3
         tx              =   "Incluir Insumo"
         enab            =   -1
         font            =   "Form6.frx":B20E
         coltype         =   1
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "Form6.frx":B236
         picn            =   "Form6.frx":B254
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.TextBox capitulo 
         BackColor       =   &H00FFFFC0&
         Height          =   405
         Left            =   3960
         MultiLine       =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "Digite el nombre para el análisis a crear"
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox valinsumo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   14
         ToolTipText     =   "Ingrese la cantidad de APU"
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox uninsumo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox unidad 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5040
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox analisis 
         BackColor       =   &H00FFFFC0&
         Height          =   405
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         ToolTipText     =   "Digite el nombre para el análisis a crear"
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Análisis Unitario  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Capitulo  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   25
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "$ Insumo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "U.M Ins  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "U.M. APU :"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   9
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Nuevo APU :"
         ForeColor       =   &H8000000D&
         Height          =   555
         Index           =   0
         Left            =   4200
         TabIndex        =   7
         Top             =   900
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3840
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Data Data4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data4"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Consulta1"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data Data6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data6"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Consulta3"
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Data Data3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data3"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "UNITARIOS"
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "INSUMOS"
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CLAVES"
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   3000
      Top             =   6120
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButton     =   0
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.Frame Frame1 
      Caption         =   "INSUMOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFC0&
         Height          =   1710
         ItemData        =   "Form6.frx":B570
         Left            =   50
         List            =   "Form6.frx":B572
         TabIndex        =   3
         Top             =   240
         Width           =   2310
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CLAVE INSUMOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFC0&
         Height          =   720
         ItemData        =   "Form6.frx":B574
         Left            =   45
         List            =   "Form6.frx":B576
         TabIndex        =   1
         Top             =   240
         Width           =   2310
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form6.frx":B578
      Height          =   1455
      Left            =   2520
      OleObjectBlob   =   "Form6.frx":B58C
      TabIndex        =   32
      Top             =   2040
      Visible         =   0   'False
      Width           =   8535
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Height          =   1455
      Left            =   2520
      OleObjectBlob   =   "Form6.frx":CB33
      TabIndex        =   33
      Top             =   2040
      Width           =   8535
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Private Const LB_ITEMFROMPOINT = &H1A9

Private lngPuntoX As Long
Private lngPuntoY As Long
Private lngÍndice As Long

Private Sub chameleonButton2_Click()
'VISUALIZA EL CODIGO PARA EL NUEVO ANALISIS
Data3.DatabaseName = App.Path & ("\MI BASE.mdb")
Data3.RecordSource = "SELECT * FROM UNITARIOS order by COD_APU ASC"
Data3.Refresh
Dim refe As Integer
Dim cla As Integer
For refe = 1 To Data3.Recordset.RecordCount
    cla = Val(Data3.Recordset("COD_APU")) + 1
    Data3.Recordset.MoveNext
Next refe
codigo.Text = cla

'Borra Los datos visualizados en los cuadros luego de exportarlos al presupuesto
    Incluir.Enabled = False
    Eliminar.Enabled = True
    Data3.Refresh
    Data5.Refresh
    
    Data5.Visible = True
    Data7.Visible = False
    DBGrid1.Visible = True
    DBGrid1.Refresh
    DBGrid2.Visible = False
    valinsumo.Text = ""
    uninsumo.Text = ""
    caninsumo.Text = ""
    analisis.Text = ""
    capitulo.Text = ""
    chameleonButton2.Enabled = False
    
    Call Calculos
    
End Sub

Private Sub Incluir_Click()
    If Validar_Datos() = False Then Exit Sub
    'Graba en la Base de Datos el Insumo para el Analisis creado
    Data3.Recordset.AddNew
    Data3.Recordset.Fields("COD_APU") = codigo.Text
    Data3.Recordset.Fields("UN_APU") = unidad.Text
    Data3.Recordset.Fields("CAPITULO") = capitulo.Text
    Data3.Recordset.Fields("APU") = analisis.Text
    Data3.Recordset.Fields("COD_INSUMO") = codinsu.Text
    Data3.Recordset.Fields("CLAVE") = clainsu.Text
    Data3.Recordset.Fields("CANTIDAD_INSUMO") = caninsumo.Text
    Data3.Recordset.Update
    Dim anal As String
    anal = analisis.Text
    'Borra Los datos visualizados en los cuadros luego de exportarlos al presupuesto
    Incluir.Enabled = False
    Eliminar.Enabled = True
    Data3.Refresh
    Data5.Refresh
    
    Data5.Visible = True
    Data7.Visible = False
    DBGrid1.Visible = True
    DBGrid1.Refresh
    DBGrid2.Visible = False
    valinsumo.Text = ""
    uninsumo.Text = ""
    caninsumo.Text = ""
    
    Call Calculos
    
End Sub

Private Sub chameleonButton1_Click()
    Unload Me
    Form4.Show
End Sub

Private Sub Data5_Reposition()
    Data5.Caption = "     Insumos Incluidos en Análisis : " & (Data5.Recordset.AbsolutePosition + 1) & " de " & Data5.Recordset.RecordCount
End Sub


Private Sub Eliminar_Click()
    On Error GoTo DeleteRecordData_Err
    
    '//Elimina el registro actual con confirmación
    If Data5.Recordset.RecordCount = 0 Then
       MsgBox "No es permitido eliminar sin insumos", vbInformation, "Stop !"
       Eliminar.Enabled = False
    Else
       If MsgBox("¿ Confirma la eliminación del Insumo " & Data5.Recordset("INSUMO") & " de el Análisis ?", vbYesNo, "Advertencia Presup Ver 5.0") = vbYes Then
          Dim anal As String
          anal = analisis.Text
          Data3.DatabaseName = App.Path & ("\MI BASE.mdb")
          Data3.RecordSource = "SELECT * FROM UNITARIOS WHERE APU= '" & anal & "' And COD_INSUMO= VAL('" & Data5.Recordset("COD_INSUMO") & " ') "
          Data3.Refresh
          Data3.Recordset.Delete
          Data3.Refresh
          Call Calculos
          DBGrid1.Refresh
       End If
    End If
    Exit Sub
    
DeleteRecordData_Err:
    MsgBox Error$, vbInformation
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & ("\MI BASE.mdb")
Data2.DatabaseName = App.Path & ("\MI BASE.mdb")
Data3.DatabaseName = App.Path & ("\MI BASE.mdb")
Data4.DatabaseName = App.Path & ("\MI BASE.mdb")
Data5.DatabaseName = App.Path & ("\MI BASE.mdb")
Data6.DatabaseName = App.Path & ("\MI BASE.mdb")
Data7.DatabaseName = App.Path & ("\MI BASE.mdb")

Incluir.Enabled = False
Eliminar.Enabled = False
chameleonButton2.Enabled = False
'VISUALIZA EL CODIGO PARA EL NUEVO ANALISIS
Data3.DatabaseName = App.Path & ("\MI BASE.mdb")
Data3.RecordSource = "SELECT * FROM UNITARIOS order by COD_APU ASC"
Data3.Refresh
Dim refe As Integer
Dim cla As Integer
For refe = 1 To Data3.Recordset.RecordCount
    cla = Val(Data3.Recordset("COD_APU")) + 1
    Data3.Recordset.MoveNext
Next refe
codigo.Text = cla

Form6.Caption = "Presup Ver. 5.0  Módulo Creación de Análisis Unitarios"
Data1.DatabaseName = App.Path & ("\MI BASE.mdb")
Data1.RecordSource = "SELECT * FROM CLAVES order by ID ASC"
Data1.Refresh
Dim i As Integer
Do While Not Data1.Recordset.EOF
       List1.AddItem IIf(IsNull(Data1.Recordset("CLAVE")), "", Data1.Recordset("CLAVE")), i
       Data1.Recordset.MoveNext
       i = i + 1
Loop
Data4.DatabaseName = App.Path & ("\MI BASE.mdb")
Data4.RecordSource = "SELECT * FROM Consulta1 order by CAPITULO ASC"
Data4.Refresh
Dim H As Integer
Do While Not Data4.Recordset.EOF
       List3.AddItem IIf(IsNull(Data4.Recordset("CAPITULO")), "", Data4.Recordset("CAPITULO")), H
       Data4.Recordset.MoveNext
       H = H + 1
Loop

End Sub
Private Sub List1_click()
'Borra La lista de INSUMOS
Incluir.Enabled = False
List2.Clear
valinsumo.Text = ""
uninsumo.Text = ""
Dim cdname As String
cdname = List1.List(List1.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname '+ "*"
    End If
Data2.DatabaseName = App.Path & ("\MI BASE.mdb")
Data2.RecordSource = "SELECT * FROM INSUMOS  WHERE CLAVE like '" & cdname & "' order by INSUMO ASC"
Data2.Refresh
Dim a As Integer
Do While Not Data2.Recordset.EOF
       List2.AddItem IIf(IsNull(Data2.Recordset("INSUMO")), "", Data2.Recordset("INSUMO")), a
       Data2.Recordset.MoveNext
       a = a + 1
Loop
End If

End Sub

Private Sub List2_Click()
Dim cdname As String
cdname = List2.List(List2.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname '+ "*"
    End If
Data2.DatabaseName = App.Path & ("\MI BASE.mdb")
Data2.RecordSource = "SELECT * FROM INSUMOS  WHERE INSUMO like '" & cdname & "' order by CLAVE ASC"
Data2.Refresh
valinsumo.Text = IIf(IsNull(Data2.Recordset("VALOR")), "", Data2.Recordset("VALOR"))
valinsumo.Text = Format(valinsumo.Text, "$##,##0.00")
uninsumo.Text = IIf(IsNull(Data2.Recordset("UN_INSUMO")), "", Data2.Recordset("UN_INSUMO"))
codinsu.Text = IIf(IsNull(Data2.Recordset("COD_INSUMO")), "", Data2.Recordset("COD_INSUMO"))
clainsu.Text = IIf(IsNull(Data2.Recordset("CLAVE")), "", Data2.Recordset("CLAVE"))
analisis.SetFocus
Incluir.Enabled = True

End If
Incluir.Enabled = True
chameleonButton2.Enabled = True
End Sub
Private Sub List1_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)

  If Button = 0 Then
    lngPuntoX = CLng(X / Screen.TwipsPerPixelX)
    lngPuntoY = CLng(Y / Screen.TwipsPerPixelY)
    With List1
      lngÍndice = SendMessage(.hWnd, _
                              LB_ITEMFROMPOINT, _
                              0, _
                              ByVal ((lngPuntoY * 65536) + lngPuntoX))
      If lngÍndice < .ListCount Then
        lblAncho = .List(lngÍndice)
        If lblAncho.Width > List1.Width Then
          .ToolTipText = .List(lngÍndice)
         Else
          .ToolTipText = vbNullString
        End If
       Else
        .ToolTipText = vbNullString
      End If
    End With
  End If
  
End Sub

Private Sub List2_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)

  If Button = 0 Then
    lngPuntoX = CLng(X / Screen.TwipsPerPixelX)
    lngPuntoY = CLng(Y / Screen.TwipsPerPixelY)
    With List2
      lngÍndice = SendMessage(.hWnd, _
                              LB_ITEMFROMPOINT, _
                              0, _
                              ByVal ((lngPuntoY * 65536) + lngPuntoX))
      If lngÍndice < .ListCount Then
        lblAncho = .List(lngÍndice)
        If lblAncho.Width > List2.Width Then
          .ToolTipText = .List(lngÍndice)
         Else
          .ToolTipText = vbNullString
        End If
       Else
        .ToolTipText = vbNullString
      End If
    End With
  End If
  
End Sub
Function Validar_Datos() As Boolean
     If Trim(analisis.Text) = "" Then
          MsgBox "Debe Ingresar un nombre para el Análisis Unitario ..."
          Validar_Datos = False
          analisis.SetFocus
          Exit Function
     End If
     If Trim(caninsumo.Text) = "" Then
          MsgBox "Debe Ingresar la cantidad de insumo para el Análisis Unitario ..."
          Validar_Datos = False
          caninsumo.SetFocus
          Exit Function
     End If
     If Trim(unidad.Text) = "" Then
          MsgBox "Debe Ingresar la unidad de medida para el APU..."
          Validar_Datos = False
          unidad.SetFocus
          Exit Function
     End If
     If Trim(capitulo.Text) = "" Then
          If MsgBox("¿ Desea crear un Nuevo Capitulo ?", vbYesNo, "Advertencia Presup Ver 5.0") = vbYes Then
            capitulo.SetFocus
          End If
          MsgBox "Debe Seleccionar Un Capitulo de la Lista.."
          Validar_Datos = False
          capitulo.SetFocus
          Exit Function
     End If
     Validar_Datos = True
     
End Function
Private Sub Calculos()
    Dim anal As String
    anal = analisis.Text
    Data5.DatabaseName = App.Path & "\MI BASE.mdb"
    Data5.RecordSource = "SELECT * FROM Consulta3  WHERE APU like '" & anal & "' order by CLAVE ASC"
    Data5.Refresh
    Data6.DatabaseName = App.Path & "\MI BASE.mdb"
    Data6.RecordSource = "SELECT  sum(SUBTOTAL) As Total from Consulta3 WHERE APU= '" & anal & "' "
    Data6.Refresh
    If Data6.Recordset.RecordCount = 0 Then Exit Sub
    subtotal.Text = IIf(IsNull(Data6.Recordset!total), "0", Data6.Recordset!total)
    subtotal.Text = Format(subtotal.Text, "$##,##0.00")
    
End Sub

Private Sub List3_Click()
Dim cdname As String
cdname = List3.List(List3.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname '+ "*"
        End If
   End If
capitulo.Text = cdname
End Sub
