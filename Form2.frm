VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{184C2D02-FD61-11D0-8B58-000000000000}#1.0#0"; "fhMagicControlsB1.ocx"
Object = "{7B72A3F4-FE91-11D3-917E-E5E1F9477021}#2.0#0"; "3DLine.ocx"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   ClientHeight    =   9825
   ClientLeft      =   3405
   ClientTop       =   1155
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin vbskfree.Skinner Skinner1 
      Left            =   240
      Top             =   10560
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.Frame Frame6 
      Height          =   2655
      Left            =   2040
      TabIndex        =   35
      Top             =   7080
      Width           =   3765
      Begin DLine.ThreeDLine ThreeDLine1 
         Height          =   45
         Left            =   75
         TabIndex        =   46
         Top             =   840
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   79
      End
      Begin Presupuestos.chameleonButton exportar 
         Height          =   495
         Left            =   1920
         TabIndex        =   45
         ToolTipText     =   "Exportar Presupuesto a Project"
         Top             =   1200
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "Exportar Presupuesto a MS Project"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         MICON           =   "Form2.frx":030A
         PICN            =   "Form2.frx":0326
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Presupuestos.chameleonButton imprimir1 
         Height          =   495
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "Imprimir APU seleccionado"
         Top             =   1200
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "Imprimir APU actual"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
         MICON           =   "Form2.frx":0640
         PICN            =   "Form2.frx":065C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Presupuestos.chameleonButton imprimir 
         Height          =   500
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Imprimir reporte de Presupuesto"
         Top             =   140
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Imprimir Presupuesto"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         MICON           =   "Form2.frx":09F6
         PICN            =   "Form2.frx":0A12
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Presupuestos.chameleonButton chameleonButton1 
         Height          =   495
         Left            =   1950
         TabIndex        =   37
         ToolTipText     =   "Regresar al Menú Principal"
         Top             =   1920
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1138
         BTYPE           =   3
         TX              =   "Regresar al Menú"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         MICON           =   "Form2.frx":0B6C
         PICN            =   "Form2.frx":0B88
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Presupuestos.chameleonButton Eliminar 
         Height          =   495
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "Eliminar Análisis Activo"
         Top             =   1920
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1138
         BTYPE           =   3
         TX              =   "Borrar APU  Presupuesto"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         MICON           =   "Form2.frx":0EA2
         PICN            =   "Form2.frx":0EBE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSForms.OptionButton porcapitulos 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   44
         Top             =   360
         Width           =   1575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2778;450"
         Value           =   "0"
         Caption         =   "Por Capitulos"
         FontName        =   "Tahoma"
         FontHeight      =   135
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton detallado 
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   43
         Top             =   120
         Width           =   1575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2778;450"
         Value           =   "1"
         Caption         =   "Detallado"
         FontName        =   "Tahoma"
         FontHeight      =   135
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2655
      Left            =   5880
      TabIndex        =   22
      Top             =   7080
      Width           =   5895
      Begin VB.TextBox iva 
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3480
         TabIndex        =   34
         Text            =   "0"
         Top             =   1680
         Width           =   2100
      End
      Begin fhMagicControlsB1.MagicLabel MagicLabel1 
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   2170
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         Caption         =   "Total  Costos (Directos+Indirectos)  :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   16711680
         ForeColor1      =   16711680
         ForeColor2      =   0
      End
      Begin VB.TextBox totpres 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3480
         TabIndex        =   31
         Text            =   "0"
         Top             =   2150
         Width           =   2100
      End
      Begin VB.TextBox utilidades 
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3480
         TabIndex        =   30
         Text            =   "0"
         Top             =   1320
         Width           =   2100
      End
      Begin VB.TextBox imprevistos 
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3480
         TabIndex        =   28
         Text            =   "0"
         Top             =   960
         Width           =   2100
      End
      Begin VB.TextBox administracion 
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3480
         TabIndex        =   26
         Text            =   "0"
         Top             =   600
         Width           =   2100
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3480
         TabIndex        =   24
         Text            =   "0"
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label Label6 
         Caption         =   "label"
         Height          =   135
         Left            =   960
         TabIndex        =   33
         Top             =   1750
         Width           =   1905
      End
      Begin VB.Label Label5 
         Caption         =   "label"
         Height          =   135
         Left            =   960
         TabIndex        =   29
         Top             =   1450
         Width           =   1905
      End
      Begin VB.Label Label4 
         Caption         =   "label"
         Height          =   135
         Left            =   960
         TabIndex        =   27
         Top             =   1050
         Width           =   1905
      End
      Begin VB.Label Label3 
         Caption         =   "label"
         Height          =   135
         Left            =   960
         TabIndex        =   25
         Top             =   750
         Width           =   1905
      End
      Begin VB.Label Label2 
         Caption         =   "Costo Directo ....................... :"
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2175
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
      RecordSource    =   "PRESUPUESTO"
      Top             =   6720
      Width           =   9320
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "Form2.frx":1798
      Height          =   2295
      Left            =   2520
      OleObjectBlob   =   "Form2.frx":17AC
      TabIndex        =   19
      Top             =   4320
      Width           =   9315
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form2.frx":282F
      Height          =   2040
      Left            =   2520
      OleObjectBlob   =   "Form2.frx":2843
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   9315
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
      Height          =   855
      Left            =   960
      TabIndex        =   16
      Top             =   10440
      Visible         =   0   'False
      Width           =   1575
      Begin VB.TextBox incimano 
         Height          =   255
         Left            =   1080
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox inciequipos 
         Height          =   255
         Left            =   1200
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox incimaterial 
         Height          =   255
         Left            =   1320
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox mano 
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox equipos 
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox materiales 
         Height          =   255
         Left            =   480
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox capitulo 
         Height          =   255
         Left            =   1200
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox valtot 
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox obra 
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   240
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
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   60
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Análisis Unitario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2520
      TabIndex        =   5
      Top             =   0
      Width           =   9315
      Begin VB.TextBox vertotal 
         Alignment       =   1  'Right Justify
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
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5640
         TabIndex        =   55
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox capinuevo 
         BackColor       =   &H00FFFFC0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
         Height          =   600
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   40
         ToolTipText     =   "Digite aquí el Nombre para el Capitulo que tiene en Pliegos"
         Top             =   170
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.OptionButton cambiacapitulo 
         Caption         =   "Cambiar Nombre al Capitulo para unificar con pliegos"
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
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   2175
      End
      Begin Presupuestos.chameleonButton Cargar 
         Height          =   615
         Left            =   7680
         TabIndex        =   15
         ToolTipText     =   "Incluir APU en el presupuesto"
         Top             =   1440
         Width           =   1455
         _extentx        =   2566
         _extenty        =   873
         btype           =   3
         tx              =   "Agregar al Presupuesto"
         enab            =   0
         font            =   "Form2.frx":463E
         coltype         =   1
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "Form2.frx":466A
         picn            =   "Form2.frx":4688
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.TextBox cantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2880
         TabIndex        =   14
         ToolTipText     =   "Ingrese la cantidad de APU"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox valunit 
         Alignment       =   1  'Right Justify
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
         Left            =   2880
         TabIndex        =   12
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox unidad 
         Alignment       =   1  'Right Justify
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
         Left            =   720
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox codigo 
         Alignment       =   1  'Right Justify
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
         Left            =   720
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox analisis 
         BackColor       =   &H00FFFFC0&
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
         Height          =   525
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label Label1 
         Caption         =   "Valor Total Unitario   :"
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
         Height          =   615
         Index           =   5
         Left            =   4850
         TabIndex        =   56
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre APU Seleccionado :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad   :"
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
         Left            =   2000
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Valor Unit. :"
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
         Left            =   2000
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "U.M. :"
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
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. :"
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
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datas"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   10440
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Data Data11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data11"
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
         RecordSource    =   ""
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Data Data10 
         BackColor       =   &H80000002&
         Caption         =   "Data10"
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
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "RESUMEN_COSTOS_PRESUPUESTO"
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data Data9 
         BackColor       =   &H80000002&
         Caption         =   "Data9"
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
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Consulta8"
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data Data8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data8"
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
         RecordSource    =   "UNITARIOS"
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Data Data7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data7"
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
         RecordSource    =   "INSUMOS"
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
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
         RecordSource    =   "PRESUPUESTO"
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
         RecordSource    =   "Consulta3"
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data Data4 
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
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Consulta3"
         Top             =   240
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
         RecordSource    =   "Consulta2"
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
         RecordSource    =   "Consulta1"
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "APU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   2415
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFC0&
         Height          =   4185
         ItemData        =   "Form2.frx":4F64
         Left            =   45
         List            =   "Form2.frx":4F66
         TabIndex        =   3
         Top             =   240
         Width           =   2310
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CAPITULO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFC0&
         Height          =   2040
         ItemData        =   "Form2.frx":4F68
         Left            =   45
         List            =   "Form2.frx":4F6A
         TabIndex        =   1
         Top             =   240
         Width           =   2310
      End
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Height          =   2040
      Left            =   2520
      OleObjectBlob   =   "Form2.frx":4F6C
      TabIndex        =   38
      Top             =   2175
      Width           =   9255
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1905
      Left            =   45
      Picture         =   "Form2.frx":6D67
      Stretch         =   -1  'True
      ToolTipText     =   "CopyRigth HugoSoft 2006"
      Top             =   7560
      Width           =   1950
   End
End
Attribute VB_Name = "Form2"
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

Public obra11 As String
Public proponente11 As String
Public administracion11 As Double
Public imprevistos11 As Double
Public utilidades11 As Double
Public aiu11 As Double
Public iva11 As Double
Public desperdicio11 As Double
Public ajusmaterial11 As Double
Public ajusequipo11 As Double
Public ajusmano11 As Double
Public controlito2 As String




Private Sub cambiacapitulo_Click()
    If cambiacapitulo.Value = True Then
        capinuevo.Visible = True
        capinuevo.SetFocus
    End If
End Sub

Private Sub cantidad_Change()
    If cantidad.Text = "" Then
        cantidad.SetFocus
    Else
        vertotal.Text = valunit.Text * cantidad.Text
        vertotal.Text = Format(vertotal.Text, "$##,##0.00")
    End If
End Sub

Private Sub Cargar_Click()
    If Validar_Datos() = False Then Exit Sub
    'Graba en la Base de Datos el Presupuesto seleccionado
    Data5.DatabaseName = App.Path & ("\MI BASE.mdb")
    Data5.Recordset.AddNew
    obra.Text = obra11
    valunit.Text = Format(valunit.Text, "$##,##0.00")
    valtot.Text = Val(Str(valunit.Text * cantidad.Text))
    valtot.Text = Format(valtot.Text, "$##,##0.00")
    Data5.Recordset.Fields("OBRA") = obra.Text
    If cambiacapitulo.Value = False Then
        Data5.Recordset.Fields("CAPITULO") = capitulo.Text
        Else
        Data5.Recordset.Fields("CAPITULO") = capinuevo.Text
    End If
    Data5.Recordset.Fields("APU") = analisis.Text
    Data5.Recordset.Fields("UNIDAD") = unidad.Text
    Data5.Recordset.Fields("CANTIDAD") = cantidad.Text
    Data5.Recordset.Fields("VALOR UNIT") = valunit.Text
    Data5.Recordset.Fields("VALOR_TOTAL") = valtot.Text
    Data5.Recordset.Update
    
    Call CalculosUnitario
    
    'Borra Los datos visualizados en los cuadros luego de exportarlos al presupuesto
    Cargar.Enabled = False
    analisis.Text = ""
    codigo.Text = ""
    unidad.Text = ""
    valunit.Text = ""
    cantidad.Text = ""
    vertotal.Text = ""
    DBGrid1.Visible = False
    DBGrid2.Visible = True
    imprimir1.Enabled = False
    cambiacapitulo.Value = False
    cambiacapitulo.Enabled = False
    capinuevo.Text = ""
    capinuevo.Visible = False
    List2.SetFocus
    Call Calculos
    
End Sub

Private Sub chameleonButton1_Click()
    Unload Me
    Form4.Show
End Sub

Private Sub Data5_Reposition()
    Data5.Caption = "                                                        Análisis Incluidos en Presupuesto : " & (Data5.Recordset.AbsolutePosition + 1) & " de " & Data5.Recordset.RecordCount
End Sub


Private Sub Eliminar_Click()
    On Error GoTo DeleteRecordData_Err
    
    '//Elimina el registro actual con confirmación
    If Data5.Recordset.RecordCount = 0 Then
       MsgBox "No es permitido eliminar el único Análisis", vbInformation, "Stop !"
    Else
       If MsgBox("¿ Confirma la eliminación del Análisis " & Data5.Recordset("APU") & " de el Presupuesto ?", vbYesNo, "Advertencia Presup Ver 5.0") = vbYes Then
          Data5.Recordset.Delete
          Data5.Refresh
       End If
    End If
    If Data10.Recordset.RecordCount = 0 Then
       MsgBox "No es permitido eliminar el único Análisis", vbInformation, "Stop !"
    Else
       If MsgBox("¿ Confirma la eliminación del Análisis " & Data5.Recordset("APU") & " de el Presupuesto ?", vbYesNo, "Advertencia Presup Ver 5.0") = vbYes Then
          Data10.Recordset.Delete
          Data10.Refresh
       End If
    End If
    
    Call Calculos
    Exit Sub
    
DeleteRecordData_Err:
    MsgBox Error$, vbInformation
End Sub

Private Sub exportar_Click()
    exportar.Enabled = False
    Unload Me
    Form11.Show
End Sub

Private Sub Form_Load()

Data1.DatabaseName = App.Path & ("\MI BASE.mdb")
Data2.DatabaseName = App.Path & ("\MI BASE.mdb")
Data3.DatabaseName = App.Path & ("\MI BASE.mdb")
Data4.DatabaseName = App.Path & ("\MI BASE.mdb")
Data5.DatabaseName = App.Path & ("\MI BASE.mdb")
Data6.DatabaseName = App.Path & ("\MI BASE.mdb")
Data7.DatabaseName = App.Path & ("\MI BASE.mdb")
Data8.DatabaseName = App.Path & ("\MI BASE.mdb")
Data9.DatabaseName = App.Path & ("\MI BASE.mdb")
Data10.DatabaseName = App.Path & ("\MI BASE.mdb")
Data11.DatabaseName = App.Path & ("\MI BASE.mdb")

'Carga los datos que vienen del formulario 1
obra11 = Form1.obra11
controlito2 = Form1.controlito2
proponente11 = Form1.proponente11
administracion11 = Form1.administracion11
imprevistos11 = Form1.imprevistos11
utilidades11 = Form1.utilidades11
aiu11 = Val(Str(administracion11 + imprevistos11 + utilidades11))
iva11 = Form1.iva11
desperdicio11 = Form1.desperdicio11
ajusmaterial11 = Form1.ajusmaterial11
ajusequipo11 = Form1.ajusequipo11
ajusmano11 = Form1.ajusmano11
'Graba en la base temporal(Tabla) los datos traidos del formulario1
Dim AA As Integer
Data7.RecordSource = "SELECT * FROM INSUMOS  WHERE CLAVE='MAT' order by ID ASC"
Data7.Refresh
For AA = 1 To Data7.Recordset.RecordCount
    Data7.Recordset.Edit
    Data7.Recordset.Fields("AJUSTE_MAT") = ajusmaterial11
    Data7.Recordset.Fields("AJUSTE_EQP") = "0.0"
    Data7.Recordset.Fields("AJUSTE_MDO") = "0.0"
    Data7.Recordset.Update
    Data7.Recordset.MoveNext
Next AA
Dim BB As Integer
Data7.RecordSource = "SELECT * FROM INSUMOS  WHERE CLAVE='EQP' order by ID ASC"
Data7.Refresh
For BB = 1 To Data7.Recordset.RecordCount
    Data7.Recordset.Edit
    Data7.Recordset.Fields("AJUSTE_MAT") = "0.0"
    Data7.Recordset.Fields("AJUSTE_EQP") = ajusequipo11
    Data7.Recordset.Fields("AJUSTE_MDO") = "0.0"
    Data7.Recordset.Update
    Data7.Recordset.MoveNext
Next BB
Dim CC As Integer
Data7.RecordSource = "SELECT * FROM INSUMOS  WHERE CLAVE='MDO' order by ID ASC"
Data7.Refresh
For CC = 1 To Data7.Recordset.RecordCount
    Data7.Recordset.Edit
    Data7.Recordset.Fields("AJUSTE_MAT") = "0.0"
    Data7.Recordset.Fields("AJUSTE_EQP") = "0.0"
    Data7.Recordset.Fields("AJUSTE_MDO") = ajusmano11
    Data7.Recordset.Update
    Data7.Recordset.MoveNext
Next CC
Dim DD As Integer
Data8.RecordSource = "SELECT * FROM UNITARIOS WHERE CLAVE='MAT' order by ID ASC"
Data8.Refresh
For DD = 1 To Data8.Recordset.RecordCount
    Data8.Recordset.Edit
    Data8.Recordset.Fields("DESPERDICIO MAT") = desperdicio11
    Data8.Recordset.Update
    Data8.Recordset.MoveNext
Next DD
Dim EE As Integer
Data8.RecordSource = "SELECT * FROM UNITARIOS WHERE CLAVE='EQP' order by ID ASC"
Data8.Refresh
For EE = 1 To Data8.Recordset.RecordCount
    Data8.Recordset.Edit
    Data8.Recordset.Fields("DESPERDICIO MAT") = "0.0"
    Data8.Recordset.Update
    Data8.Recordset.MoveNext
Next EE
Dim FF As Integer
Data8.RecordSource = "SELECT * FROM UNITARIOS WHERE CLAVE='MDO' order by ID ASC"
Data8.Refresh
For FF = 1 To Data8.Recordset.RecordCount
    Data8.Recordset.Edit
    Data8.Recordset.Fields("DESPERDICIO MAT") = "0.0"
    Data8.Recordset.Update
    Data8.Recordset.MoveNext
Next FF

'Pasa al Formulario 2 datos traidos del formulario 1
Label3.Caption = "Administración " & administracion11 & "% :"
Label4.Caption = "Imprevistos " & imprevistos11 & "% :"
Label5.Caption = "Utilidades " & utilidades11 & "% :"
Label6.Caption = "IVA " & iva11 & "% :"
Call Calculos

'Form2.Caption = "Presup Ver. 5.0  Módulo Creación de Presupuestos    Obra : " & obra11

DBGrid3.Caption = "Presup Ver. 5.0  - Obra : " & obra11
Data1.DatabaseName = App.Path & ("\MI BASE.mdb")
Data1.RecordSource = "SELECT * FROM Consulta1  order by CAPITULO ASC"
Data1.Refresh
Dim i As Integer
Do While Not Data1.Recordset.EOF
       List1.AddItem IIf(IsNull(Data1.Recordset("CAPITULO")), "", Data1.Recordset("CAPITULO")), i
       Data1.Recordset.MoveNext
       i = i + 1
Loop


End Sub

Private Sub imprimir_Click()
    
    DataEnvironment1.Commands("Command1").Parameters("obra11") = obra11
    If controlito2 = "ok" Then
        Presupuesto.Sections("Sección5").Controls("Etiqueta19").Caption = subtotal.Text
        Presupuesto.Sections("Sección5").Controls("Etiqueta4").Caption = Label3.Caption
        Presupuesto.Sections("Sección5").Controls("Etiqueta14").Caption = administracion.Text
        Presupuesto.Sections("Sección5").Controls("Etiqueta10").Caption = Label4.Caption
        Presupuesto.Sections("Sección5").Controls("Etiqueta15").Caption = imprevistos.Text
        Presupuesto.Sections("Sección5").Controls("Etiqueta11").Caption = Label5.Caption
        Presupuesto.Sections("Sección5").Controls("Etiqueta16").Caption = utilidades.Text
        Presupuesto.Sections("Sección5").Controls("Etiqueta12").Caption = Label6.Caption
        Presupuesto.Sections("Sección5").Controls("Etiqueta17").Caption = iva.Text
        Presupuesto.Sections("Sección5").Controls("Etiqueta18").Caption = totpres.Text
        Presupuesto.Sections("Sección3").Controls("Etiqueta20").Caption = "Obra : " & Form2.obra11
        Presupuesto.Sections("Sección2").Controls("Etiqueta5").Caption = "PRESUPUESTO DE OBRA DETALLADO"
        imprimir.Enabled = False
        If detallado(0).Value = False Then Presupuesto.Sections("Sección1").Visible = False
        Presupuesto.Show
    End If
    If controlito2 = "no" Then
        Presupuesto2.Sections("Sección3").Controls("Etiqueta20").Caption = "Obra : " & Form2.obra11
        imprimir.Enabled = False
        If detallado(0).Value = False Then Presupuesto2.Sections("Sección1").Visible = False
        Presupuesto2.Show
    End If
End Sub

Private Sub imprimir1_Click()
    Dim apu11 As String
    Dim final11 As Currency
    apu11 = analisis.Text
    DataEnvironment1.Commands("Command3_Grouping").Parameters("apu11") = apu11
    If controlito2 = "ok" Then
        Unitarios.Sections("Sección2").Controls("Etiqueta23").Caption = analisis.Text
        Unitarios.Sections("Sección2").Controls("Etiqueta22").Caption = unidad.Text
        Unitarios.Sections("Sección5").Controls("Etiqueta14").Caption = valunit.Text
        Unitarios.Sections("Sección3").Controls("Etiqueta4").Caption = "Obra : " & obra11
        imprimir1.Enabled = False
        Unitarios.Show
    End If
    
    If controlito2 = "no" Then
        Data11.DatabaseName = App.Path & ("\MI BASE.mdb")
        Data11.RecordSource = "SELECT * FROM Consulta12  WHERE APU='" & apu11 & "' "
        Data11.Refresh
        Me.Data11.RecordSource = "Select SUM(FINAL) As Total from Consulta12 " & "WHERE APU like '" & apu11 & "' "
        Me.Data11.Refresh
        final11 = IIf(IsNull(Data11.Recordset!total), "0", Data11.Recordset!total)
        Unitarios2.Sections("Sección2").Controls("Etiqueta23").Caption = analisis.Text
        Unitarios2.Sections("Sección2").Controls("Etiqueta22").Caption = unidad.Text
        Unitarios2.Sections("Sección3").Controls("Etiqueta4").Caption = "Obra : " & obra11
        Unitarios2.Sections("Sección5").Controls("Etiqueta13").Caption = Format(final11, "$##,##0.00")
        If final11 = 0 Then Unitarios2.Sections("Sección3").Controls("Etiqueta14").Caption = "ANALISIS NO INCLUIDO EN PRESUPUESTO"
        imprimir1.Enabled = False
        Unitarios2.Show
    End If
End Sub

Private Sub List1_click()
'Borra La lista de APU y los cuadros de texto del APU
Cargar.Enabled = False
cambiacapitulo.Enabled = False
DBGrid1.Visible = False
DBGrid2.Visible = True
List2.Clear
analisis.Text = ""
codigo.Text = ""
unidad.Text = ""
valunit.Text = ""
cantidad.Text = ""
Dim cdname As String
cdname = List1.List(List1.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname '+ "*"
    End If
capitulo.Text = cdname
Data2.DatabaseName = App.Path & ("\MI BASE.mdb")
Data2.RecordSource = "SELECT * FROM Consulta2  WHERE CAPITULO like '" & cdname & "' order by APU ASC"
Data2.Refresh
Dim a As Integer
Do While Not Data2.Recordset.EOF
       List2.AddItem IIf(IsNull(Data2.Recordset("APU")), "", Data2.Recordset("APU")), a
       Data2.Recordset.MoveNext
       a = a + 1
Loop
End If

End Sub

Private Sub List2_Click()
DBGrid2.Visible = False
DBGrid1.Visible = True
Dim cdname As String
cdname = List2.List(List2.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname '+ "*"
    End If
Data2.DatabaseName = App.Path & ("\MI BASE.mdb")
Data2.RecordSource = "SELECT * FROM Consulta2  WHERE APU like '" & cdname & "' order by APU ASC"
Data2.Refresh
analisis.Text = IIf(IsNull(Data2.Recordset("APU")), "", Data2.Recordset("APU"))
codigo.Text = IIf(IsNull(Data2.Recordset("COD_APU")), "", Data2.Recordset("COD_APU"))
unidad.Text = IIf(IsNull(Data2.Recordset("UN_APU")), "", Data2.Recordset("UN_APU"))
cantidad.SetFocus
'VISUALIZA EN LA GRILLA EL ANALISIS UNITARIO GUARDADO EN LA BASE
Data3.DatabaseName = App.Path & ("\MI BASE.mdb")
Data3.RecordSource = "SELECT * FROM Consulta3  WHERE APU like '" & cdname & "' order by CLAVE ASC"
Data3.Refresh
Me.Data4.RecordSource = "Select SUM(SUBTOTAL) As Total from Consulta3 " & "WHERE APU like '" & cdname & "' "
Me.Data4.Refresh

valunit.Text = Format(Data4.Recordset!total, "$##,##0.00")
End If


Cargar.Enabled = True
imprimir.Enabled = True
imprimir1.Enabled = True
exportar.Enabled = True
cambiacapitulo.Value = False
cambiacapitulo.Enabled = True
capinuevo.Text = ""
capinuevo.Visible = False
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
     If cambiacapitulo.Value = True Then
        If Trim(capinuevo.Text) = "" Then
          MsgBox "Debe Ingresar el nombre para el nuevo Capitulo ..."
          Validar_Datos = False
          capinuevo.SetFocus
          Exit Function
        End If
     End If
     
     If Trim(cantidad.Text) = "" Then
          MsgBox "Debe Ingresar la cantidad ..."
          Validar_Datos = False
          cantidad.SetFocus
          Exit Function
          
     End If
     Validar_Datos = True
     
End Function
Private Sub Calculos()
    Data5.DatabaseName = App.Path & "\MI BASE.mdb"
    Data5.RecordSource = "SELECT * FROM PRESUPUESTO  WHERE OBRA like '" & obra11 & "' order by APU ASC"
    Data5.Refresh
    
    Data6.DatabaseName = App.Path & "\MI BASE.mdb"
    Data6.RecordSource = "SELECT  sum(VALOR_TOTAL) As Total from PRESUPUESTO WHERE OBRA= '" & obra11 & "' "
    Data6.Refresh
    If Data6.Recordset.RecordCount = 0 Then Exit Sub
    subtotal.Text = IIf(IsNull(Data6.Recordset!total), "0", Data6.Recordset!total)
    subtotal.Text = Format(subtotal.Text, "$##,##0.00")
    administracion.Text = Val(Str(subtotal.Text * (administracion11 / 100)))
    administracion.Text = Format(administracion.Text, "$##,##0.00")
    imprevistos.Text = Val(Str(subtotal.Text * (imprevistos11 / 100)))
    imprevistos.Text = Format(imprevistos.Text, "$##,##0.00")
    utilidades.Text = Val(Str(subtotal.Text * (utilidades11 / 100)))
    utilidades.Text = Format(utilidades.Text, "$##,##0.00")
    iva.Text = (Val(Str(subtotal.Text)) * (utilidades11 / 100) * (iva11 / 100)) / (1 + (aiu11 / 100))
    iva.Text = Format(iva.Text, "$##,##0.00")
    totpres.Text = Val(Str(subtotal.Text)) + Val(Str(administracion.Text)) + Val(Str(imprevistos.Text)) + Val(Str(utilidades.Text)) + Val(Str(iva.Text))
    totpres.Text = Format(totpres.Text, "$##,##0.00")
End Sub

Private Sub form_Unload(Cancel As Integer)
     Set DataEnvironment1 = Nothing
End Sub

Private Sub CalculosUnitario()
    Data9.DatabaseName = App.Path & "\MI BASE.mdb"
    Data9.RecordSource = "SELECT * FROM Consulta8 WHERE OBRA like '" & obra11 & "' AND CLAVE='MAT' "
    Data9.Refresh
    
    Dim Refe1 As String
    Dim Refe2 As String
    Refe1 = analisis.Text
    Refe2 = Val(Str(valunit.Text)) * Val(Str(cantidad.Text))
    
    'REALIZA SUMA DE MATERIALES,EQUIPOS Y MANO DE OBRA DE EL APU
    Me.Data9.RecordSource = "Select SUM(SUBT) As Total from Consulta8 " & "WHERE OBRA like '" & obra11 & "' AND APU='" & Refe1 & "' AND CLAVE='MAT' "
    Me.Data9.Refresh
    materiales.Text = IIf(IsNull(Data9.Recordset!total), "0", Data9.Recordset!total)
    materiales.Text = Format(materiales.Text, "$##,##0.00")
    incimaterial.Text = (Val(Str(materiales.Text)) / Refe2) * 100
    incimaterial.Text = Format(incimaterial.Text, "##,##0.00")
    If Val(Str(incimaterial.Text)) > 100 Then MsgBox "Ojo error en la base ..."
    Me.Data9.RecordSource = "Select SUM(SUBT) As Total from Consulta8 " & "WHERE OBRA like '" & obra11 & "' AND APU='" & Refe1 & "' AND CLAVE='EQP' "
    Me.Data9.Refresh
    equipos.Text = IIf(IsNull(Data9.Recordset!total), "0", Data9.Recordset!total)
    equipos.Text = Format(equipos.Text, "$##,##0.00")
    inciequipos.Text = (Val(Str(equipos.Text)) / Refe2) * 100
    inciequipos.Text = Format(inciequipos.Text, "##,##0.00")
    If Val(Str(inciequipos.Text)) > 100 Then MsgBox "Ojo error en la base ..."
    Me.Data9.RecordSource = "Select SUM(SUBT) As Total from Consulta8 " & "WHERE OBRA like '" & obra11 & "' AND APU='" & Refe1 & "' AND CLAVE='MDO' "
    Me.Data9.Refresh
    mano.Text = IIf(IsNull(Data9.Recordset!total), "0", Data9.Recordset!total)
    mano.Text = Format(mano.Text, "$##,##0.00")
    incimano.Text = (Val(Str(mano.Text)) / Refe2) * 100
    incimano.Text = Format(incimano.Text, "##,##0.00")
    If Val(Str(incimano.Text)) > 100 Then MsgBox "Ojo error en la base ..."
    'GRABA EN LA TABLA RESUMEN LOS DATOS
    Data10.DatabaseName = App.Path & "\MI BASE.mdb"
    Data10.RecordSource = "SELECT * FROM RESUMEN_COSTOS_PRESUPUESTO WHERE OBRA like '" & obra11 & "' "
    Data10.Refresh
    Data10.Recordset.AddNew
    obra.Text = obra11
    Data10.Recordset.Fields("OBRA") = obra.Text
    Data10.Recordset.Fields("APU") = analisis.Text
    Data10.Recordset.Fields("MATERIALES") = materiales.Text
    Data10.Recordset.Fields("INCIMATERIAL") = incimaterial.Text
    Data10.Recordset.Fields("EQUIPOS") = equipos.Text
    Data10.Recordset.Fields("INCIEQUIPO") = inciequipos.Text
    Data10.Recordset.Fields("MANO") = mano.Text
    Data10.Recordset.Fields("INCIMANO") = incimano.Text
    Data10.Recordset.Update
End Sub

