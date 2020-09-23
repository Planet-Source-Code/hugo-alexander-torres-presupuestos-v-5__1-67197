VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{75AACB43-0B9D-11D2-95B5-0000B43369D3}#1.2#0"; "ARFrmExt.ocx"
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.1#0"; "cPopMenu6.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2A4A8B91-0052-11D1-9CCD-444553540000}#2.0#0"; "sText.ocx"
Object = "{C532C895-BA31-4362-B986-569ABD68920E}#17.0#0"; "MenuXP.ocx"
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presup Ver. 5.0   CopyRigth HugoSoft 2006"
   ClientHeight    =   6975
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11760
   ForeColor       =   &H00400040&
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form4.frx":030A
   MousePointer    =   99  'Custom
   Picture         =   "Form4.frx":0614
   ScaleHeight     =   6975
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MenuXP.XPsidemenu XPsidemenu1 
      Height          =   10000
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   17648
      Speed           =   0
   End
   Begin stext1.sText sText1 
      Height          =   220
      Left            =   6480
      TabIndex        =   2
      Top             =   6720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      BackStyle       =   0
      Caption         =   "                                        Presup Ver. 5.0 -  ® - HugoSoft 2006"
      Interval        =   150
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList6"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   25
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "dos"
            Object.ToolTipText     =   "Crear/Editar Proyecto"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tres"
            Object.ToolTipText     =   "Crear/Editar Presupuesto"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cinco"
            Object.ToolTipText     =   "Creación de Análisis Unitario"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "seis"
            Object.ToolTipText     =   "Editar Análisis Unitario Existente"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ocho"
            Object.ToolTipText     =   "Creación/Edición de Materiales"
            ImageIndex      =   40
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nueve"
            Object.ToolTipText     =   "Creación/Edición de Equipos"
            ImageIndex      =   41
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "diez"
            Object.ToolTipText     =   "Creación/Edición Mano de Obra"
            ImageIndex      =   42
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "doce"
            Object.ToolTipText     =   "Presupuesto Detallado"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "trece"
            Object.ToolTipText     =   "Presupuesto Por Capitulos"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "catorce"
            Object.ToolTipText     =   "Análisis Unitarios"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "diezyseis"
            Object.ToolTipText     =   "Exportar Presupuesto a MS Project"
            ImageIndex      =   43
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "diezysiete"
            Object.ToolTipText     =   "Resumen Costo Presupuesto Por Análisis"
            ImageIndex      =   34
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "diezyocho"
            Object.ToolTipText     =   "Resumen Costos Presupuesto Por Insumos"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "veinte"
            Object.ToolTipText     =   "Calculadora"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ventiuno"
            Object.ToolTipText     =   "Compactar Base de Datos"
            ImageIndex      =   44
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ventidos"
            Object.ToolTipText     =   "Calendario"
            ImageIndex      =   45
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "venticuatro"
            Object.ToolTipText     =   "Acerca de.."
            ImageIndex      =   26
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "venticinco"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
      Begin VB.Frame Frame1 
         Height          =   375
         Left            =   12500
         TabIndex        =   8
         Top             =   0
         Width           =   2295
         Begin VB.Label Label1 
            Caption         =   "HugoSoft Multimedia 2006"
            ForeColor       =   &H80000010&
            Height          =   240
            Left            =   180
            TabIndex        =   9
            Top             =   105
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      Height          =   4695
      Left            =   4200
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
      Begin VB.PictureBox picBackground 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   1920
         Picture         =   "Form4.frx":0956
         ScaleHeight     =   660
         ScaleWidth      =   645
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Timer tmr_pos 
         Interval        =   1
         Left            =   960
         Top             =   1200
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   360
         Top             =   1200
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   720
         Top             =   1920
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":0FF3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":130D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":1627
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":1941
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":1C5B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   1920
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":2535
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":268F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":27E9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":2B3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":2C95
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":2DEF
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":2F49
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":30A3
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ARFormExtenderCtrl.ARFormExtender ARFormExtender1 
         Left            =   240
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         BackgroundType  =   1
         Picture         =   "Form4.frx":37F5
         GradientIniColor=   -2147483645
         GradientEndColor=   16711680
         AppName         =   "hugo"
      End
      Begin cPopMenu6.PopMenu ctlMenu 
         Left            =   1440
         Top             =   360
         _ExtentX        =   1058
         _ExtentY        =   1058
         HighlightCheckedItems=   0   'False
         TickIconIndex   =   0
         OfficeXpStyle   =   -1  'True
         MenuBackgroundColor=   12632256
         BackgroundPicture=   "Form4.frx":3B533F
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   120
         Top             =   2520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   27
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3B59EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3B5D86
               Key             =   "mnusale(1)"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3B60A0
               Key             =   "mnuequipos(1)"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3B7022
               Key             =   "mnuresumencosto2(4)"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3B7FA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3B8F26
               Key             =   "mnumanobra(2)"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3B9EA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3BAE2A
               Key             =   "mnumateriales(0)"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3BB26C
               Key             =   "mnucreacion(0)"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3BBB46
               Key             =   "mnucrea(0)"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3BC420
               Key             =   "mnureparar(2)"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3BC9BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3C25DC
               Key             =   "mnucreaunitario(0)"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3C9ADE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3CFD78
               Key             =   "mnudetallado(0)"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D0112
               Key             =   "mnucapitulos(1)"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D04AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D0846
               Key             =   "mnueditaunitario(1)"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D0BE0
               Key             =   "mnutodo(0)"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D0F7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D1314
               Key             =   "mnuresumencosto(3)"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D16AE
               Key             =   "mnuexporta(2)"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D1A48
               Key             =   "mnucalculadora(0)"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D1FE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D257C
               Key             =   "mnucalendario(1)"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D2916
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D2CB0
               Key             =   "mnuacerca(0)"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList4 
         Left            =   720
         Top             =   2520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   20
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D304A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D3924
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D41FE
               Key             =   "mnusale"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D4650
               Key             =   "mnucreaproy"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D4AA2
               Key             =   "mnucreacion"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D4EF4
               Key             =   "mnucreaunitario"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D5346
               Key             =   "mnueditaunitario"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D5798
               Key             =   "mnumateriales"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D5BEA
               Key             =   "mnuequipos"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D603C
               Key             =   "mnumanobra"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D648E
               Key             =   "mnudetallado"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D68E0
               Key             =   "mnucapitulos"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D6D32
               Key             =   "mnutodo"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D7184
               Key             =   "mnuexporta"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D75D6
               Key             =   "mnuresumencosto"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D7A28
               Key             =   "mnuresumencosto2"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D7E7A
               Key             =   "mnucalculadora"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D82CC
               Key             =   "mnucalendario"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D871E
               Key             =   "mnureparar"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D8878
               Key             =   "mnuacerca"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList6 
         Left            =   720
         Top             =   3120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   45
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D8B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D928C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3D9B66
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DA260
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DAB3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DB414
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DBCEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DC5C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DCEA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DD77C
               Key             =   "btn11"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DE056
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DE930
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DEECA
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DF464
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3DFD3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E0618
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E0EF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E15EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E1746
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E1CE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E25BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E2CB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E324E
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E37E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E40C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E499C
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E4F36
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E54D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E5DAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E6684
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E6F5E
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E7838
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E8112
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E89EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E92C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E9860
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E99BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3E9F54
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3EA0AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3EA208
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3EB18A
               Key             =   ""
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3EC10C
               Key             =   ""
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3ED08E
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3ED3A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3ED942
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList7 
         Left            =   120
         Top             =   3720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   19
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3EDA9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3EF22E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3F0F38
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3F26CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3F43D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3F5B66
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3F72F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3F9002
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3FAD0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3FCA16
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3FE1A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":3FFEB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":401644
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":40334E
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":404AE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":4057BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":406F4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":408C56
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form4.frx":40A960
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ImageList ImageList5 
         Left            =   120
         Top             =   3120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   11
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form4.frx":40C0F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form4.frx":40D074
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form4.frx":40DFF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form4.frx":40EF78
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form4.frx":40FEFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form4.frx":410214
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form4.frx":411196
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form4.frx":412118
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form4.frx":4122F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form4.frx":412B0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form4.frx":412CE6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5280
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":413000
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin stext1.sText sText2 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      BackStyle       =   0
      Caption         =   "                                        Presup Ver. 3.0 -  ® - HugoSoft 2004"
      Interval        =   150
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6675
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   441
            MinWidth        =   441
            Picture         =   "Form4.frx":413112
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "Usuario Actualmente conectado"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   441
            MinWidth        =   441
            Picture         =   "Form4.frx":413464
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "04:46 p.m."
            Object.ToolTipText     =   "Hora del Sistema"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "Form4.frx":4137B6
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1676
            MinWidth        =   1676
            TextSave        =   "23/11/2006"
            Object.ToolTipText     =   "Fecha del Sistema"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5821
            MinWidth        =   5821
            Picture         =   "Form4.frx":413910
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4762
            MinWidth        =   4762
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "Base de Datos Actual : MI BASE.mdb"
            TextSave        =   "Base de Datos Actual : MI BASE.mdb"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label clip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000004&
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   1
      Top             =   4800
      Width           =   135
   End
   Begin VB.Menu mnuproyectoTOP 
      Caption         =   "&Proyecto"
      Begin VB.Menu mnucreaproy 
         Caption         =   "Creacion/Edicion de proyectos"
      End
      Begin VB.Menu mnusale 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnupresupuestoTOP 
      Caption         =   "P&resupuesto"
      Begin VB.Menu mnucreacion 
         Caption         =   "Creacion/Edición del presupuesto"
      End
   End
   Begin VB.Menu mnuedicionTOP 
      Caption         =   "Analisis &Unitarios"
      Begin VB.Menu mnucreaunitario 
         Caption         =   "Creación de Análisis Unitarios"
      End
      Begin VB.Menu mnueditaunitario 
         Caption         =   "Edición de Análisis Unitarios Existentes"
      End
   End
   Begin VB.Menu mnuinsumosTOP 
      Caption         =   "&Insumos"
      Begin VB.Menu mnumateriales 
         Caption         =   "Creación/Edición de Materiales"
      End
      Begin VB.Menu mnuequipos 
         Caption         =   "Creación/Edición de Equipos y herramientas"
      End
      Begin VB.Menu mnumanobra 
         Caption         =   "Creación/Edición Mano de Obra"
      End
   End
   Begin VB.Menu mnuimpresionTOP 
      Caption         =   "I&nformes"
      Begin VB.Menu mnuimprimepresupuesto 
         Caption         =   "Presupuesto"
         Begin VB.Menu mnudetallado 
            Caption         =   "Detallado"
         End
         Begin VB.Menu mnucapitulos 
            Caption         =   "Por Capitulos"
         End
      End
      Begin VB.Menu mnuimprimeapu 
         Caption         =   "Análisis Unitarios"
         Begin VB.Menu mnutodo 
            Caption         =   "Seleccionar de el Presupuesto"
         End
      End
      Begin VB.Menu mnuguion2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexporta 
         Caption         =   "Exportar Presupuesto a MS Project"
      End
      Begin VB.Menu mnuresumencosto 
         Caption         =   "Resumen Costos de Presupuesto por Análisis"
      End
      Begin VB.Menu mnuresumencosto2 
         Caption         =   "Resumen Costos de Presupuesto por Insumos"
      End
   End
   Begin VB.Menu mnuherramientasTOP 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnucalculadora 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu mnucalendario 
         Caption         =   "Calendario"
      End
      Begin VB.Menu mnuguion 
         Caption         =   "-"
      End
      Begin VB.Menu mnureparar 
         Caption         =   "Compactar y Reparar Base de Datos"
      End
   End
   Begin VB.Menu mnuayudaTOP 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuacerca 
         Caption         =   "Acerca de.."
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = 1024
Public controlito As String
Public controlito2 As String
Public ResHorizontal As Integer
Public ResVertical As Integer
Public obra11 As String
Dim sh As New Shell 'This is to acces shell automation


Private Sub Form_Load()
Call cargamenu
obra11 = Form1.obra11
'CREA MENU VERTICAL XP

With Me.XPsidemenu1
Set .Pimagelist = Me.ImageList2
Set .hImageList = Me.ImageList1
    'set speed to 0 befor building menu speeds up process
    .Speed = 0

    .Addpanel "P1", "PROYECTO ", Opened, False, 1
    .Addpanel "P2", "PRESUPUESTO ", Opened, False, 2
    .Addpanel "P3", "ANALISIS UNITARIOS ", Opened, True, 3, , App.Path & "\smallMultiProjectLogo.bmp"
    .Addpanel "P4", "INSUMOS ", Opened, False, 4, Picture1
    .Addpanel "P5", "CREDITOS ", Opened, True, 5, , App.Path & "\MI LOGO.bmp"
   
    .AddHyper "H1", "P1", "Información", True, Hyperlink, 1, "This is tooltip 1"
    '.AddHyper "H2", "P1", "La punta de la verga", True, Hyperlink, 5, "This is tooltip 2"
    
    .AddHyper "H2", "P2", "Creación", True, Hyperlink, 2, "This is tooltip 3"
    '.AddHyper "H4", "P2", "Hyperlink 2", True, Hyperlink, 3, "This is tooltip 4"
    '.AddHyper "H5", "P2", "Hyperlink 3", True, Hyperlink, 6, "This is tooltip 5"
    
    .AddHyper "H3", "P3", "Crear APU Nuevo", True, Hyperlink, 3, "This is tooltip 6"
    .AddHyper "H4", "P3", "Editar APU Existente", True, Hyperlink, 4, "This is tooltip 7"
    
    .AddHyper "H5", "P4", "Materiales", True, Hyperlink, 5, "This is tooltip 8"
    .AddHyper "H6", "P4", "Equipos", True, Hyperlink, 6, "This is tooltip 9"
    .AddHyper "H7", "P4", "Mano de Obra", True, Hyperlink, 7, "This is tooltip 10"
    '.AddHyper "H11", "P4", "Hyperlink 4", True, Hyperlink, 11, "This is tooltip 11"
    
    .AddHyper "H8", "P5", "Acerca de..", True, Hyperlink, 8, "This is tooltip 10"
    '.AddHyper "H12", "P5", "Hyperlink 1", True, Hyperlink, 12, "This is tooltip 9"
    '.AddHyper "H13", "P5", "Hyperlink 2", True, Hyperlink, 13, "This is tooltip 10"
    'set speed to desired value higher means faster
    .Speed = 50
End With
' FIN MENU VERTICAL



'AVERIGUA RESOLUCION DE PANTALLA DONDE SE EJECUTA
TWidth% = Screen.Width \ Screen.TwipsPerPixelX
THeight% = Screen.Height \ Screen.TwipsPerPixelY
ResHorizontal = Str$(TWidth%)
ResVertical = Str$(THeight%)
'TEXTO SCROLL EN BARRA ESTADO
If ResHorizontal = 800 Then
        sText1.Top = 8010
    End If
    If ResHorizontal = 1024 Then
        sText1.Top = 10570
End If

StatusBar1.Panels(8).Text = "Resolución Actual : " & ResHorizontal & " x " & ResVertical
  'Desactiva Boton X en formulario
  Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
       lHndSysMenu = GetSystemMenu(Form4.hWnd, 0)
        lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
  
  'Gradient1.GradientForm Me
  
       
     ' Indicar la ruta del datareport1
     
     
     
End Sub

Private Sub form_Unload(Cancel As Integer)
     Set DataEnvironment1 = Nothing
End Sub

Private Sub mnuacerca_Click()
    Form5.Show
End Sub


Private Sub mnucalculadora_Click()
    Shell "calc.exe", vbNormalFocus
End Sub



Private Sub mnucalendario_Click()
    Form15.Show
    Form15.SetFocus
    Form15.WindowState = 0
End Sub

Private Sub mnucapitulos_Click()
    controlito = Form1.controlito
    controlito2 = Form1.controlito2
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        Form2.Show
        Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para imprimir informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
            DataEnvironment1.Commands("Command1").Parameters("obra11") = Form2.obra11
            If controlito2 = "ok" Then
                Presupuesto.Sections("Sección5").Controls("Etiqueta19").Caption = Form2.subtotal.Text
                Presupuesto.Sections("Sección5").Controls("Etiqueta4").Caption = Form2.Label3.Caption
                Presupuesto.Sections("Sección5").Controls("Etiqueta14").Caption = Form2.administracion.Text
                Presupuesto.Sections("Sección5").Controls("Etiqueta10").Caption = Form2.Label4.Caption
                Presupuesto.Sections("Sección5").Controls("Etiqueta15").Caption = Form2.imprevistos.Text
                Presupuesto.Sections("Sección5").Controls("Etiqueta11").Caption = Form2.Label5.Caption
                Presupuesto.Sections("Sección5").Controls("Etiqueta16").Caption = Form2.utilidades.Text
                Presupuesto.Sections("Sección5").Controls("Etiqueta12").Caption = Form2.Label6.Caption
                Presupuesto.Sections("Sección5").Controls("Etiqueta17").Caption = Form2.iva.Text
                Presupuesto.Sections("Sección5").Controls("Etiqueta18").Caption = Form2.totpres.Text
                Presupuesto.Sections("Sección3").Controls("Etiqueta20").Caption = "Obra : " & Form2.obra11
                Presupuesto.Sections("Sección1").Visible = False
                Presupuesto.Show
            End If
            If controlito2 = "no" Then
                Presupuesto2.Sections("Sección3").Controls("Etiqueta20").Caption = "Obra : " & Form2.obra11
                Presupuesto2.Sections("Sección1").Visible = False
                Presupuesto2.Show
            End If
        End If
    End If
End Sub





Private Sub mnucreacion_Click()
    controlito = Form1.controlito
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then Form2.Show
End Sub



Private Sub mnucreaproy_Click()
    Form1.Show
End Sub

Private Sub mnucreaunitario_Click()
    Form6.Show
End Sub

Private Sub mnudetallado_Click()
    controlito = Form1.controlito
    controlito2 = Form1.controlito2
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        Form2.Show
        Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para imprimir informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
            DataEnvironment1.Commands("Command1").Parameters("obra11") = Form2.obra11
            If controlito2 = "ok" Then
                Presupuesto.Sections("Sección5").Controls("Etiqueta19").Caption = Form2.subtotal.Text
                Presupuesto.Sections("Sección5").Controls("Etiqueta4").Caption = Form2.Label3.Caption
                Presupuesto.Sections("Sección5").Controls("Etiqueta14").Caption = Form2.administracion.Text
                Presupuesto.Sections("Sección5").Controls("Etiqueta10").Caption = Form2.Label4.Caption
                Presupuesto.Sections("Sección5").Controls("Etiqueta15").Caption = Form2.imprevistos.Text
                Presupuesto.Sections("Sección5").Controls("Etiqueta11").Caption = Form2.Label5.Caption
                Presupuesto.Sections("Sección5").Controls("Etiqueta16").Caption = Form2.utilidades.Text
                Presupuesto.Sections("Sección5").Controls("Etiqueta12").Caption = Form2.Label6.Caption
                Presupuesto.Sections("Sección5").Controls("Etiqueta17").Caption = Form2.iva.Text
                Presupuesto.Sections("Sección5").Controls("Etiqueta18").Caption = Form2.totpres.Text
                Presupuesto.Sections("Sección3").Controls("Etiqueta20").Caption = "Obra : " & Form2.obra11
                Presupuesto.Show
            End If
            If controlito2 = "no" Then
                Presupuesto2.Sections("Sección3").Controls("Etiqueta20").Caption = "Obra : " & Form2.obra11
                Presupuesto2.Show
            End If
        End If
    End If
End Sub

Private Sub mnueditaunitario_Click()
    Form10.Show
End Sub

Private Sub mnuequipos_Click()
    Form8.Show
End Sub

Private Sub mnuexporta_Click()
    controlito = Form1.controlito
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        'Form2.Show
        'Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para exportar informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
            
            Form11.Show
        End If
    End If
End Sub

Private Sub mnumanobra_Click()
    Form9.Show
End Sub

Private Sub mnumateriales_Click()
    Form7.Show
End Sub

Private Sub mnureparar_Click()
    Set cn = Nothing
    CompactDatabase (App.Path & "\MI BASE.mdb")
End Sub

Private Sub mnuresumencosto_Click()
    controlito = Form1.controlito
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        Form2.Show
        Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para imprimir informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
            Form13.Show
        End If
    End If
End Sub

Private Sub mnuresumencosto2_Click()
    controlito = Form1.controlito
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        Form2.Show
        Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para imprimir informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
            Form14.Show
        End If
    End If
End Sub

Private Sub mnusale_Click()
    Set DataEnvironment1 = Nothing
    Unload Form2
    Unload Me
End Sub

Private Sub mnutodo_Click()
    controlito = Form1.controlito
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        Form2.Show
        Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para imprimir informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
            Form12.Show
        End If
    End If
    
End Sub

Private Sub XPsidemenu1_HyperClick(key As String)
'Label1.Caption = "You have Clicked on Hyperlink " & key
If key = "H1" Then
    Form1.Show
End If

If key = "H2" Then
    controlito = Form1.controlito
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then Form2.Show
End If

If key = "H3" Then
    Form6.Show
End If

If key = "H4" Then
    Form10.Show
End If

If key = "H5" Then
    Form7.Show
End If

If key = "H6" Then
    Form8.Show
End If

If key = "H7" Then
    Form9.Show
End If

If key = "H8" Then
    Form5.Show
End If

End Sub

Private Sub XPsidemenu1_PictureClick(key As String)
 'Label1.Caption = "You have Clicked on Picture " & key
End Sub
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
        Case Is = 2:  mnucre
        Case Is = 3:  mnucreacio
        Case Is = 5:  mnucreaunitari
        Case Is = 6:  mnueditaunitari
        Case Is = 8:  mnumateriale
        Case Is = 9:  mnuequipo
        Case Is = 10: mnumanobr
        Case Is = 12: mnudetallad
        Case Is = 13: mnucapitulo
        Case Is = 14: mnutod
        Case Is = 16: mnuexport
        Case Is = 17: mnuresumencost
        Case Is = 18: mnuresumencos
        Case Is = 20: mnucalculador
        Case Is = 21: mnurepara
        Case Is = 22: mnucalendari
        Case Is = 24: mnuacerc
        Case Is = 25: mnusal
End Select
      
        
End Sub
Private Sub mnuacerc()
    Form5.Show
End Sub
Private Sub mnucalculador()
    Shell "calc.exe", vbNormalFocus
End Sub



Private Sub mnucalendari()
    Form15.Show
    Form15.SetFocus
    Form15.WindowState = 0
End Sub

Private Sub mnucapitulo()
    controlito = Form1.controlito
    controlito2 = Form1.controlito2
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        Form2.Show
        Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para imprimir informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
        DataEnvironment1.Commands("Command1").Parameters("obra11") = obra11
        If controlito2 = "ok" Then
        Presupuesto.Sections("Sección5").Controls("Etiqueta19").Caption = Form2.subtotal.Text
        Presupuesto.Sections("Sección5").Controls("Etiqueta4").Caption = Form2.Label3.Caption
        Presupuesto.Sections("Sección5").Controls("Etiqueta14").Caption = Form2.administracion.Text
        Presupuesto.Sections("Sección5").Controls("Etiqueta10").Caption = Form2.Label4.Caption
        Presupuesto.Sections("Sección5").Controls("Etiqueta15").Caption = Form2.imprevistos.Text
        Presupuesto.Sections("Sección5").Controls("Etiqueta11").Caption = Form2.Label5.Caption
        Presupuesto.Sections("Sección5").Controls("Etiqueta16").Caption = Form2.utilidades.Text
        Presupuesto.Sections("Sección5").Controls("Etiqueta12").Caption = Form2.Label6.Caption
        Presupuesto.Sections("Sección5").Controls("Etiqueta17").Caption = Form2.iva.Text
        Presupuesto.Sections("Sección5").Controls("Etiqueta18").Caption = Form2.totpres.Text
        Presupuesto.Sections("Sección3").Controls("Etiqueta20").Caption = "Obra : " & Form2.obra11
        Presupuesto.Sections("Sección2").Controls("Etiqueta5").Caption = "PRESUPUESTO DE OBRA POR CAPITULOS"
        Presupuesto.Sections("Sección1").Visible = False
        Presupuesto.Show
        End If
        If controlito2 = "no" Then
            Presupuesto2.Sections("Sección3").Controls("Etiqueta20").Caption = "Obra : " & Form2.obra11
            Presupuesto2.Sections("Sección2").Controls("Etiqueta5").Caption = "PRESUPUESTO DE OBRA POR CAPITULOS"
            Presupuesto2.Sections("Sección1").Visible = False
            Presupuesto2.Show
        End If
        End If
    End If
End Sub

Private Sub mnucre()
    Form1.Show
End Sub



Private Sub mnucreacio()
    controlito = Form1.controlito
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then Form2.Show
End Sub



Private Sub mnucreaunitari()
    Form6.Show
End Sub

Private Sub mnudetallad()
    controlito = Form1.controlito
    controlito2 = Form1.controlito2
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        Form2.Show
        Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para imprimir informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
            DataEnvironment1.Commands("Command1").Parameters("obra11") = Form2.obra11
            If controlito2 = "ok" Then
            Presupuesto.Sections("Sección5").Controls("Etiqueta19").Caption = Form2.subtotal.Text
            Presupuesto.Sections("Sección5").Controls("Etiqueta4").Caption = Form2.Label3.Caption
            Presupuesto.Sections("Sección5").Controls("Etiqueta14").Caption = Form2.administracion.Text
            Presupuesto.Sections("Sección5").Controls("Etiqueta10").Caption = Form2.Label4.Caption
            Presupuesto.Sections("Sección5").Controls("Etiqueta15").Caption = Form2.imprevistos.Text
            Presupuesto.Sections("Sección5").Controls("Etiqueta11").Caption = Form2.Label5.Caption
            Presupuesto.Sections("Sección5").Controls("Etiqueta16").Caption = Form2.utilidades.Text
            Presupuesto.Sections("Sección5").Controls("Etiqueta12").Caption = Form2.Label6.Caption
            Presupuesto.Sections("Sección5").Controls("Etiqueta17").Caption = Form2.iva.Text
            Presupuesto.Sections("Sección5").Controls("Etiqueta18").Caption = Form2.totpres.Text
            Presupuesto.Sections("Sección3").Controls("Etiqueta20").Caption = "Obra : " & Form2.obra11
            Presupuesto.Sections("Sección2").Controls("Etiqueta5").Caption = "PRESUPUESTO DE OBRA DETALLADO"
            Presupuesto.Show
            End If
            If controlito2 = "no" Then
            Presupuesto2.Sections("Sección3").Controls("Etiqueta20").Caption = "Obra : " & Form2.obra11
            Presupuesto2.Sections("Sección2").Controls("Etiqueta5").Caption = "PRESUPUESTO DE OBRA DETALLADO"
            Presupuesto2.Show
            End If
        End If
    End If
End Sub

Private Sub mnueditaunitari()
    Form10.Show
End Sub

Private Sub mnuequipo()
    Form8.Show
End Sub

Private Sub mnuexport()
    controlito = Form1.controlito
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        'Form2.Show
        'Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para exportar informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
            
            Form11.Show
        End If
    End If
End Sub

Private Sub mnumanobr()
    Form9.Show
End Sub

Private Sub mnumateriale()
    Form7.Show
End Sub

Private Sub mnurepara()
    Set cn = Nothing
    CompactDatabase (App.Path & "\MI BASE.mdb")
End Sub

Private Sub mnuresumencost()
    controlito = Form1.controlito
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        Form2.Show
        Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para imprimir informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
            Form13.Show
        End If
    End If
End Sub

Private Sub mnuresumencos()
    controlito = Form1.controlito
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        Form2.Show
        Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para imprimir informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
            Form14.Show
        End If
    End If
End Sub

Private Sub mnusal()
    Set DataEnvironment1 = Nothing
    Unload Form2
    Unload Me
    
End Sub

Private Sub mnutod()
    controlito = Form1.controlito
    If controlito <> "ok" Then MsgBox "Primero debe Elegir la obra a presupuestar", vbCritical, "Advertencia Presup !"
    If controlito = "ok" Then
        Form2.Show
        Form2.Visible = False
        If Form2.Data5.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Primero debe Crear el Presupuesto para imprimir informes ", vbCritical, "Advertencia Presup !"
            Exit Sub
        Else
            Form12.Show
        End If
    End If
    
End Sub
Sub cargamenu()
' Menu con iconos
  Dim c As Control
  Set ctlMenu.BackgroundPicture = picBackground.Picture
  With ctlMenu
         .ImageList = ImageList4
         .SubClassMenu Me
         .HighlightStyle = cspHighlightGradient
  End With
  ' Asociar la opciones del menu con las imagenes del ImageList4
  
  For Each c In Me.Controls
      If TypeName(c) = "Menu" Then
       On Error Resume Next
       ctlMenu.ItemIcon(c.Name) = ImageList4.ListImages(c.Name).Index - 1
      End If
  Next
End Sub
