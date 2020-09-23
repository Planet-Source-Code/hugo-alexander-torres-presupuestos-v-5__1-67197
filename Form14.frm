VERSION 5.00
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{7B72A3F4-FE91-11D3-917E-E5E1F9477021}#2.0#0"; "3DLine.ocx"
Begin VB.Form Form14 
   Caption         =   "Presup Ver. 5.0   Módulo Resumen Presupuesto"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4665
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin Presupuestos.chameleonButton chameleonButton2 
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   2800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ver Informe / Imprimir"
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form14.frx":0442
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Propietario\Mis documentos\PRESUPUESTO VB\MI BASE.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Consulta8"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin Presupuestos.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   2800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Regresar al Menú"
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
      MICON           =   "Form14.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   1860
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   2520
      Top             =   3840
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin DLine.ThreeDLine ThreeDLine1 
      Height          =   45
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   79
   End
   Begin VB.TextBox incimano 
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
      Left            =   3600
      TabIndex        =   12
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox inciequipo 
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
      Left            =   3600
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox valmano 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   2160
      Width           =   1450
   End
   Begin VB.TextBox valequipo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   1680
      Width           =   1450
   End
   Begin VB.TextBox incimaterial 
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
      Left            =   3600
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox valmaterial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   1450
   End
   Begin VB.TextBox valor 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox obra 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Incidencia % :"
      Height          =   255
      Index           =   7
      Left            =   2580
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Incidencia % :"
      Height          =   255
      Index           =   6
      Left            =   2580
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "M de Obra :"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Equipos  :"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Incidencia % :"
      Height          =   255
      Index           =   3
      Left            =   2580
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Materiales :"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Valor :"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Obra :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public obra11 As String

Private Sub chameleonButton2_Click()
    DataEnvironment1.Commands("Command6_Grouping").Parameters("obra11") = obra11
    RepCostosPresup2.Sections("Sección2").Controls("Etiqueta11").Caption = obra11
    RepCostosPresup2.Sections("Sección2").Controls("Etiqueta13").Caption = valor.Text
    RepCostosPresup2.Sections("Sección2").Controls("Etiqueta16").Caption = valmaterial.Text
    RepCostosPresup2.Sections("Sección2").Controls("Etiqueta18").Caption = valequipo.Text
    RepCostosPresup2.Sections("Sección2").Controls("Etiqueta22").Caption = valmano.Text
    RepCostosPresup2.Show
End Sub

Private Sub Form_Load()
    Data1.DatabaseName = App.Path & ("\MI BASE.mdb")
    Data2.DatabaseName = App.Path & ("\MI BASE.mdb")
    Data4.DatabaseName = App.Path & ("\MI BASE.mdb")

    obra11 = Form1.obra11
    Data1.DatabaseName = App.Path & ("\MI BASE.mdb")
    Data1.RecordSource = "SELECT * FROM Consulta6  WHERE OBRA='" & obra11 & "' "
    Data1.Refresh
    obra.Text = obra11
    Me.Data1.RecordSource = "Select SUM(VALOR_TOTAL) As Total from Consulta6 " & "WHERE OBRA like '" & obra11 & "' "
    Me.Data1.Refresh
    valor.Text = Format(Data1.Recordset!total, "$##,##0.00")
    Data2.DatabaseName = App.Path & ("\MI BASE.mdb")
    Data2.RecordSource = "SELECT * FROM Consulta8  WHERE OBRA='" & obra11 & "' "
    Data2.Refresh
    Me.Data2.RecordSource = "Select SUM(SUBT) As Total from Consulta8 " & "WHERE OBRA like '" & obra11 & "' AND CLAVE='MAT' "
    Me.Data2.Refresh
    valmaterial.Text = Format(Data2.Recordset!total, "$##,##0.00")
    incimaterial.Text = (Val(Str(valmaterial.Text)) / Val(Str(valor.Text))) * 100
    incimaterial.Text = Format(incimaterial.Text, "##,##0.00")
    Me.Data2.RecordSource = "Select SUM(SUBT) As Total from Consulta8 " & "WHERE OBRA like '" & obra11 & "' AND CLAVE='EQP' "
    Me.Data2.Refresh
    valequipo.Text = Format(Data2.Recordset!total, "$##,##0.00")
    inciequipo.Text = (Val(Str(valequipo.Text)) / Val(Str(valor.Text))) * 100
    inciequipo.Text = Format(inciequipo.Text, "##,##0.00")
    Me.Data2.RecordSource = "Select SUM(SUBT) As Total from Consulta8 " & "WHERE OBRA like '" & obra11 & "' AND CLAVE='MDO' "
    Me.Data2.Refresh
    valmano.Text = Format(Data2.Recordset!total, "$##,##0.00")
    incimano.Text = (Val(Str(valmano.Text)) / Val(Str(valor.Text))) * 100
    incimano.Text = Format(incimano.Text, "##,##0.00")
    
    
    
End Sub
Private Sub chameleonButton1_Click()
    Unload Me
    Form4.Show
End Sub



