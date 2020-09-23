VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{7B72A3F4-FE91-11D3-917E-E5E1F9477021}#2.0#0"; "3DLine.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presup Ver. 5.0   Información de Usuario"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   5865
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5865
   StartUpPosition =   1  'CenterOwner
   Begin Presupuestos.chameleonButton REGRESAR 
      Height          =   495
      Left            =   1800
      TabIndex        =   37
      ToolTipText     =   "Regresa al menú Principal"
      Top             =   3720
      Width           =   1575
      _extentx        =   2778
      _extenty        =   873
      btype           =   3
      tx              =   "Regresar al Menú"
      enab            =   -1
      font            =   "Form1.frx":014A
      coltype         =   1
      focusr          =   -1
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Form1.frx":0176
      picn            =   "Form1.frx":0194
      umcol           =   -1
      soft            =   0
      picpos          =   0
      ngrey           =   0
      fx              =   0
      hand            =   0
      check           =   0
      value           =   0
   End
   Begin Presupuestos.chameleonButton CmdConectar 
      Height          =   495
      Left            =   50
      TabIndex        =   36
      ToolTipText     =   "Conecta con la Base Principal"
      Top             =   3720
      Width           =   1695
      _extentx        =   2990
      _extenty        =   873
      btype           =   3
      tx              =   "Conectar Base de Datos"
      enab            =   -1
      font            =   "Form1.frx":04B0
      coltype         =   1
      focusr          =   -1
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Form1.frx":04DC
      picn            =   "Form1.frx":04FA
      umcol           =   -1
      soft            =   0
      picpos          =   0
      ngrey           =   0
      fx              =   0
      hand            =   0
      check           =   0
      value           =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5741
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "   Clasificación"
      TabPicture(0)   =   "Form1.frx":0816
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtproyecto"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtlicitacion"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtproponente"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtcliente"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtlugar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fecFecha"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtcodigo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "  Indirectos"
      TabPicture(1)   =   "Form1.frx":0B30
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "aiu"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "OptionButton1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtiva"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "caliva"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ThreeDLine2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "  Ajustes"
      TabPicture(2)   =   "Form1.frx":0E4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ThreeDLine1"
      Tab(2).Control(1)=   "txtajusmano"
      Tab(2).Control(2)=   "txtajusequipo"
      Tab(2).Control(3)=   "txtajusmaterial"
      Tab(2).Control(4)=   "txtdesperdicio"
      Tab(2).Control(5)=   "Label9"
      Tab(2).Control(6)=   "Label8"
      Tab(2).Control(7)=   "Label7"
      Tab(2).Control(8)=   "Label5"
      Tab(2).ControlCount=   9
      Begin DLine.ThreeDLine ThreeDLine2 
         Height          =   2475
         Left            =   -71640
         TabIndex        =   44
         Top             =   600
         Width           =   45
         _ExtentX        =   79
         _ExtentY        =   4366
         Orientation     =   1
      End
      Begin DLine.ThreeDLine ThreeDLine1 
         Height          =   45
         Left            =   -74760
         TabIndex        =   43
         Top             =   1200
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   79
      End
      Begin VB.CheckBox caliva 
         Caption         =   "Calcular IVA"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -71400
         TabIndex        =   40
         Tag             =   "S"
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   39
         Tag             =   "S"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtajusmano 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71760
         TabIndex        =   35
         Tag             =   "S"
         Text            =   "0.00"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtajusequipo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71760
         TabIndex        =   34
         Tag             =   "S"
         Text            =   "0.00"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtajusmaterial 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71760
         TabIndex        =   33
         Tag             =   "S"
         Text            =   "0.00"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtdesperdicio 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71760
         TabIndex        =   29
         Tag             =   "S"
         Text            =   "0.00"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtiva 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70320
         TabIndex        =   27
         Tag             =   "S"
         Top             =   1920
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "A I U"
         Height          =   1455
         Left            =   -74520
         TabIndex        =   19
         Top             =   1440
         Width           =   2655
         Begin VB.TextBox txtutilidades 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   25
            Tag             =   "S"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtimprevistos 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   24
            Tag             =   "S"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtadministracion 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   21
            Tag             =   "S"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Utilidades        (%) "
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Imprevistos      (%) "
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Administración (%) "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1455
         End
      End
      Begin MSComCtl2.DTPicker fecFecha 
         Height          =   375
         Left            =   1500
         TabIndex        =   17
         Tag             =   "S"
         Top             =   2820
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   21233665
         CurrentDate     =   37890
      End
      Begin VB.TextBox txtlugar 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   15
         Tag             =   "S"
         ToolTipText     =   "Ubicación del Proyecto"
         Top             =   2460
         Width           =   4200
      End
      Begin VB.TextBox txtcliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   13
         Tag             =   "S"
         ToolTipText     =   "Nombre del Cliente"
         Top             =   2100
         Width           =   4200
      End
      Begin VB.TextBox txtproponente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   11
         Tag             =   "S"
         ToolTipText     =   "Nombre del Proponente"
         Top             =   1740
         Width           =   4200
      End
      Begin VB.TextBox txtlicitacion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   9
         Tag             =   "S"
         ToolTipText     =   "Referencia Licitación"
         Top             =   1380
         Width           =   4200
      End
      Begin VB.TextBox txtproyecto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   7
         Tag             =   "S"
         ToolTipText     =   "Nombre de la Obra"
         Top             =   1050
         Width           =   4200
      End
      Begin MSForms.OptionButton OptionButton1 
         Height          =   375
         Left            =   -74520
         TabIndex        =   42
         Tag             =   "S"
         Top             =   1080
         Width           =   2775
         VariousPropertyBits=   746588185
         BackColor       =   -2147483633
         ForeColor       =   -2147483635
         DisplayStyle    =   5
         Size            =   "4895;661"
         Value           =   "0"
         Caption         =   "Calcular AIU en Unitarios"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label10 
         Caption         =   "Código"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Ajuste a Mano de Obra (%)"
         Height          =   255
         Left            =   -73920
         TabIndex        =   32
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Ajuste a Equipos/Herr.  (%)"
         Height          =   255
         Left            =   -73920
         TabIndex        =   31
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Ajuste a Materiales        (%)"
         Height          =   375
         Left            =   -73920
         TabIndex        =   30
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Desperdicio                   (%)"
         Height          =   255
         Left            =   -73920
         TabIndex        =   28
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "IVA (%)"
         Height          =   255
         Left            =   -71400
         TabIndex        =   26
         Top             =   1920
         Width           =   1095
      End
      Begin MSForms.OptionButton aiu 
         Height          =   375
         Left            =   -74520
         TabIndex        =   18
         Tag             =   "S"
         Top             =   720
         Width           =   2775
         VariousPropertyBits=   746588185
         BackColor       =   -2147483633
         ForeColor       =   -2147483635
         DisplayStyle    =   5
         Size            =   "4895;661"
         Value           =   "0"
         Caption         =   "Calcular AIU en Presupuesto"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2820
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Lugar"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   2100
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Proponente"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Licitación"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Proyecto"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1020
         Width           =   1455
      End
   End
   Begin VB.Frame FrameBuscar 
      Height          =   600
      Left            =   3450
      TabIndex        =   1
      Top             =   3650
      Visible         =   0   'False
      Width           =   2400
      Begin VB.CommandButton CmdBuscar 
         Height          =   330
         Left            =   1575
         Picture         =   "Form1.frx":1164
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Búsqueda "
         Top             =   180
         Width           =   735
      End
      Begin MSForms.Label Label6 
         Height          =   240
         Left            =   90
         TabIndex        =   4
         Top             =   225
         Width           =   825
         VariousPropertyBits=   8388627
         Caption         =   "Buscar ID:"
         Size            =   "1455;423"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TxtIDbuscar 
         Height          =   315
         Left            =   855
         TabIndex        =   3
         Tag             =   "S"
         Top             =   180
         Width           =   690
         VariousPropertyBits=   746604571
         MaxLength       =   4
         Size            =   "1217;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":140E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":156A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1822
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":197E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":24B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3596
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "nuevo"
            Object.ToolTipText     =   "Grabar Nueva Obra"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Editar Obra Grabada"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar Obra"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Guardar Datos"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "inicio"
            Object.ToolTipText     =   "Registro Inicial"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "anterior"
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "siguiente"
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "final"
            Object.ToolTipText     =   "Registro Final"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "buscar"
            Object.ToolTipText     =   "Buscar Obra"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Registros Grabados"
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "imprimir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   4305
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
            Picture         =   "Form1.frx":36F2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gConexion   As ADODB.Connection
Dim gRsPROYECTOS As New ADODB.Recordset
Dim sModo       As String
Public obra11 As String
Public proponente11 As String
Public administracion11 As Double
Public imprevistos11 As Double
Public utilidades11 As Double
Public iva11 As Double
Public desperdicio11 As Double
Public ajusmaterial11 As Double
Public ajusequipo11 As Double
Public ajusmano11 As Double
Public controlito As String
Public controlito2 As String


Private Sub CmdBuscar_Click()
     Dim FilaRetorno As Variant
     
     If TxtIDbuscar.Text <> "" Then
     Else
          FrameBuscar.Visible = False
          Set FrmListado.gRecordset = gRsPROYECTOS
          FrmListado.Show vbModal
          FilaRetorno = FrmListado.StrColListado
          If Not IsEmpty(FilaRetorno) Then
               TxtIDbuscar.Text = FilaRetorno(0)
          Else
               TxtIDbuscar.Text = 0
          End If
     End If
     ' ------------------------------------------------------
     ' ------> Empezar la busqueda desde el inicio <---------
     ' ------------------------------------------------------
     gRsPROYECTOS.MoveFirst
     ' Buscar el ID igual al indicado, hacia adelante(adSearchForward)
     gRsPROYECTOS.Find "ID=" & TxtIDbuscar.Text, , adSearchForward
     ' Si es EOF -> No encontró
     If gRsPROYECTOS.EOF Then
          MsgBox "Código Proyecto No encontrado!", , Me.Caption
     Else
          ' Si encontró, refrescar campos
          Call CargarDatos
          FrameBuscar.Visible = False
     End If

End Sub

Private Sub CmdConectar_Click()
     ' Conectarse uTilizando Access con Cursor de Cliente
     If AbreConexion(gConexion, tAccess, adUseClient) Then
          Call AccionBotones(Toolbar1, tMover)
          
          Set gRsPROYECTOS = New ADODB.Recordset
          Screen.MousePointer = vbHourglass
          ' Abrir el recordset utilizado a lo largo de la ejecución...
          gRsPROYECTOS.Open "SELECT * FROM PROYECTOS ORDER BY id", gConexion, adOpenStatic, adLockOptimistic
          Screen.MousePointer = vbDefault
          If gRsPROYECTOS.RecordCount > 0 Then
               gRsPROYECTOS.MoveFirst
               Call CargarDatos
               StatusBar1.Panels(1).Text = " Proyecto en uso : " & txtcodigo.Text & " de " & gRsPROYECTOS.RecordCount
          Else
               Call AccionBotones(Toolbar1, tNoRegistros)
               MsgBox "No Existen PROYECTOS!", , Me.Caption
          End If
     Else
          MsgBox "Errores al Conectar!"
          Call AccionBotones(Toolbar1, tDesHabilitar)
     End If
     If aiu.Value = False Then OptionButton1.Value = True
     CmdConectar.Enabled = False
End Sub
Private Sub Form_Load()
     
     CmdConectar.Enabled = True
     sModo = "Ver"
     
End Sub

Private Sub form_Unload(Cancel As Integer)
     ' Liberar Memoria
     Set gConexion = Nothing
     Set gRsPROYECTOS = Nothing
     Set DataEnvironment1 = Nothing
End Sub

Private Sub REGRESAR_Click()
    'Carga valores en memoria que iran al formulario2 antes de salir al menú
    If CmdConectar.Enabled = False Then
        If aiu.Value = True Then
            controlito2 = "ok"
        Else
            controlito2 = "no"
        End If
        controlito = "ok"
        obra11 = txtproyecto.Text
        proponente11 = txtproponente.Text
        administracion11 = txtadministracion.Text
        imprevistos11 = txtimprevistos.Text
        utilidades11 = txtutilidades.Text
        iva11 = txtiva.Text
        desperdicio11 = txtdesperdicio.Text
        ajusmaterial11 = txtajusmaterial.Text
        ajusequipo11 = txtajusequipo.Text
        ajusmano11 = txtajusmano.Text
        Unload Me
        Form4.Show
    End If
    If CmdConectar.Enabled = True Then
        Unload Me
        Form4.Show
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RsAux As New ADODB.Recordset
Dim sSQl As String
     Select Case Button.key
          Case "nuevo":
               Call AccionBotones(Toolbar1, tEditando)
               Call ActivarControles(Me, True)
               Call EncerarControles(Me)
               
               ' Recuperar el Siguiente ID de cliente
               RsAux.Open "SELECT  MAX(ID)+1  FROM PROYECTOS", gConexion
               txtcodigo.Text = IIf(IsNull(RsAux.Fields(0).Value), 1, RsAux.Fields(0).Value)
               sModo = "Nuevo"
          Case "modificar":
               Call AccionBotones(Toolbar1, tEditando)
               Call ActivarControles(Me, True)
               sModo = "Modificar"
          Case "eliminar":
               ' Borrar el Registro Actual
               sSQl = "DELETE FROM PROYECTOS WHERE id = " & txtcodigo.Text
               RsAux.Open sSQl, gConexion
               
               ' Volver a Consultar la tabla
               gRsPROYECTOS.Requery
               Call CargarDatos
               sModo = "Ver"
          Case "grabar":
               If Not Validar_Datos() Then
                    Exit Sub
               End If
               ' Si es un registro Nuevo, Agregar usando SQL (INSERT INTO tabla1 VALUES (...))
               If sModo = "Nuevo" Then
                    sSQl = "INSERT INTO PROYECTOS(ID, PROYECTO, LICITACION, PROPONENTE, CLIENTE, LUGAR, FECHA, AIU_EN_UNITARIOS, ADMINISTRACION, IMPREVISTOS, UTILIDADES, CALCULAR_IVA, VALOR_IVA, DESPERDICIO, INCREMATERIAL, INCREQUIPO, INCREMANO) " & _
                           "VALUES " & _
                           "(" & txtcodigo.Text & ",'" & txtproyecto.Text & "','" & txtlicitacion.Text & "','" & txtproponente.Text & "','" & txtcliente.Text & "','" & txtlugar.Text & "',#" & fecFecha.Month & "/" & fecFecha.Day & "/" & fecFecha.Year & "#,'" & IIf(aiu.Value, "S", "N") & "','" & txtadministracion.Text & "','" & txtimprevistos.Text & "','" & txtutilidades.Text & "','" & IIf(caliva.Value = 1, "S", "N") & "','" & txtiva.Text & "','" & txtdesperdicio.Text & "','" & txtajusmaterial.Text & "','" & txtajusequipo.Text & "','" & txtajusmano.Text & "')"
               Else
                    ' Si es un registro existente, Modificar usando SQL(UPDATE tabla SET campo1 = expr1, ...)
                    sSQl = "UPDATE PROYECTOS " & _
                           "SET " & _
                           "    PROYECTO   = '" & txtproyecto.Text & "'," & _
                           "    LICITACION = '" & txtlicitacion.Text & "'," & _
                           "    PROPONENTE = '" & txtproponente.Text & "'," & _
                           "    CLIENTE    = '" & txtcliente.Text & "'," & _
                           "    LUGAR      = '" & txtlugar.Text & "'," & _
                           "    FECHA      = #" & fecFecha.Month & "/" & fecFecha.Day & "/" & fecFecha.Year & "#," & _
                           "    AIU_EN_UNITARIOS = '" & IIf(aiu.Value, "S", "N") & "'," & _
                           "    ADMINISTRACION   = '" & txtadministracion.Text & "'," & _
                           "    IMPREVISTOS      = '" & txtimprevistos.Text & "'," & _
                           "    UTILIDADES       = '" & txtutilidades.Text & "'," & _
                           "    CALCULAR_IVA     = '" & IIf(caliva = 1, "S", "N") & "', " & _
                           "    VALOR_IVA        = '" & txtiva.Text & "'," & _
                           "    DESPERDICIO      = '" & txtdesperdicio.Text & "'," & _
                           "    INCREMATERIAL    = '" & txtajusmaterial.Text & "'," & _
                           "    INCREQUIPO       = '" & txtajusequipo.Text & "'," & _
                           "    INCREMANO        = '" & txtajusmano.Text & " '" & _
                           "WHERE ID = " & txtcodigo.Text

               End If
               RsAux.Open sSQl, gConexion
               Call AccionBotones(Toolbar1, tMover)
               Call ActivarControles(Me, False)
               gRsPROYECTOS.Requery
               Call CargarDatos
               sModo = "Ver"
          Case "cancelar":
               Call AccionBotones(Toolbar1, tMover)
               Call ActivarControles(Me, False)
               gRsPROYECTOS.Requery
               Call CargarDatos
               sModo = "Ver"
        '--------->
        Case "anterior"
            If Not gRsPROYECTOS.BOF Then
                gRsPROYECTOS.MovePrevious
                If gRsPROYECTOS.BOF Then
                    gRsPROYECTOS.MoveFirst
                End If
                Call CargarDatos
                StatusBar1.Panels(1).Text = " Proyecto en uso : " & txtcodigo.Text & " de " & gRsPROYECTOS.RecordCount
            End If
        Case "siguiente"
            If Not gRsPROYECTOS.EOF Then
                gRsPROYECTOS.MoveNext
                If gRsPROYECTOS.EOF Then
                    gRsPROYECTOS.MoveLast
                End If
                Call CargarDatos
                StatusBar1.Panels(1).Text = " Proyecto en uso : " & txtcodigo.Text & " de " & gRsPROYECTOS.RecordCount
            End If
        Case "inicio"
            If Not gRsPROYECTOS.BOF Then
                gRsPROYECTOS.MoveFirst
                Call CargarDatos
                StatusBar1.Panels(1).Text = " Proyecto en uso : " & txtcodigo.Text & " de " & gRsPROYECTOS.RecordCount
            End If
        Case "final"
            If Not gRsPROYECTOS.EOF Then
                gRsPROYECTOS.MoveLast
                Call CargarDatos
                StatusBar1.Panels(1).Text = " Proyecto en uso : " & txtcodigo.Text & " de " & gRsPROYECTOS.RecordCount
            End If
        Case "buscar"
               ' Mostrar el frame de Busqueda
               FrameBuscar.Visible = True
               TxtIDbuscar.Enabled = True
               TxtIDbuscar.Text = ""
               TxtIDbuscar.SetFocus
        Case "imprimir"
               ' Asignar los valores a los controles de la seccion
               ' DETALLE del DataReport directamente sin enlace a datos.
               RptProyecto.Sections("Detalle").Controls("LblId").Caption = txtcodigo.Text
               RptProyecto.Sections("Detalle").Controls("LblFecha1").Caption = fecFecha.Value
               RptProyecto.Sections("Detalle").Controls("Lblproyecto").Caption = txtproyecto.Text
               RptProyecto.Sections("Detalle").Controls("Lblicitacion").Caption = txtlicitacion.Text
               RptProyecto.Sections("Detalle").Controls("Lblproponente").Caption = txtproponente.Text
               RptProyecto.Sections("Detalle").Controls("Lblugar").Caption = txtlugar.Text
               RptProyecto.Sections("Detalle").Controls("Lblcliente").Caption = txtcliente.Text
               RptProyecto.Sections("Detalle").Controls("Lbliva").Caption = txtiva.Text
               RptProyecto.Sections("Detalle").Controls("Lbladministracion").Caption = txtadministracion.Text
               RptProyecto.Sections("Detalle").Controls("Lblimprevistos").Caption = txtimprevistos.Text
               RptProyecto.Sections("Detalle").Controls("Lblutilidades").Caption = txtutilidades.Text
               RptProyecto.Sections("Pie").Controls("LblFecha").Caption = "Impreso el " & Format(Date, "Long Date")
               RptProyecto.Show
        '>-----------
     
     End Select
     Set RsAux = Nothing
End Sub


Private Sub CargarDatos()
     ' Si el Recordset Tiene Datos...
     If Not gRsPROYECTOS.EOF Then
          ' En todos los campos validar que no hayan NULOS en los mismos.
          txtcodigo.Text = IIf(IsNull(gRsPROYECTOS!ID), 0, gRsPROYECTOS!ID)
          txtproyecto.Text = IIf(IsNull(gRsPROYECTOS!proyecto), "", gRsPROYECTOS!proyecto)
          txtlicitacion.Text = IIf(IsNull(gRsPROYECTOS!licitacion), "", gRsPROYECTOS!licitacion)
          txtproponente.Text = IIf(IsNull(gRsPROYECTOS!proponente), "", gRsPROYECTOS!proponente)
          txtcliente.Text = IIf(IsNull(gRsPROYECTOS!cliente), "", gRsPROYECTOS!cliente)
          txtlugar.Text = IIf(IsNull(gRsPROYECTOS!lugar), "", gRsPROYECTOS!lugar)
          fecFecha.Value = IIf(IsNull(gRsPROYECTOS!Fecha), "01/01/2003", gRsPROYECTOS!Fecha)
          If IsNull(gRsPROYECTOS!aiu_en_unitarios) Then
               aiu.Value = False
               
          Else
               If gRsPROYECTOS!aiu_en_unitarios = "S" Then
                    aiu.Value = True
               End If
          End If
          txtadministracion.Text = IIf(IsNull(gRsPROYECTOS!administracion), 0, gRsPROYECTOS!administracion)
          txtimprevistos.Text = IIf(IsNull(gRsPROYECTOS!imprevistos), 0, gRsPROYECTOS!imprevistos)
          txtutilidades.Text = IIf(IsNull(gRsPROYECTOS!utilidades), 0, gRsPROYECTOS!utilidades)
          caliva.Value = IIf(gRsPROYECTOS!calcular_iva = "S", 1, 0)
          txtiva.Text = IIf(IsNull(gRsPROYECTOS!valor_iva), 0, gRsPROYECTOS!valor_iva)
          txtdesperdicio.Text = IIf(IsNull(gRsPROYECTOS!desperdicio), 0, gRsPROYECTOS!desperdicio)
          txtajusmaterial.Text = IIf(IsNull(gRsPROYECTOS!incrematerial), 0, gRsPROYECTOS!incrematerial)
          txtajusequipo.Text = IIf(IsNull(gRsPROYECTOS!increquipo), 0, gRsPROYECTOS!increquipo)
          txtajusmano.Text = IIf(IsNull(gRsPROYECTOS!incremano), 0, gRsPROYECTOS!incremano)
          
          
     Else
          ' Si no Hay datos, Solo mostrar el boton de Agregar
          Call EncerarControles(Me)
          Call AccionBotones(Toolbar1, tNoRegistros)
          MsgBox "No Existen PROYECTOS!", , Me.Caption
     End If
End Sub
Function Validar_Datos() As Boolean
     If Trim(txtproyecto.Text) = "" Then
          MsgBox "Debe Ingresar nombre del Proyecto..."
          Validar_Datos = False
          Exit Function
     End If
     If Trim(txtlicitacion.Text) = "" Then
          MsgBox "Debe Ingresar número o referencia de la licitación..."
          Validar_Datos = False
          Exit Function
     End If
     If Trim(txtproponente.Text) = "" Then
          MsgBox "Debe Ingresar el nombre de el Proponente..."
          Validar_Datos = False
          Exit Function
     End If
     If Trim(txtcliente.Text) = "" Then
          MsgBox "Debe Ingresar el nombre de el Cliente..."
          Validar_Datos = False
          Exit Function
     End If
     If Trim(txtlugar.Text) = "" Then
          MsgBox "Debe Ingresar el nombre de el Sitio de Ejecución..."
          Validar_Datos = False
          Exit Function
     End If
          
     Validar_Datos = True
End Function






