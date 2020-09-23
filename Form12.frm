VERSION 5.00
Object = "{75AACB43-0B9D-11D2-95B5-0000B43369D3}#1.2#0"; "ARFrmExt.ocx"
Begin VB.Form Form12 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4695
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Propietario\Mis documentos\PRESUPUESTO VB\MI BASE.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Consulta7"
      Top             =   3600
      Width           =   1815
   End
   Begin ARFormExtenderCtrl.ARFormExtender ARFormExtender1 
      Left            =   3240
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      BackgroundType  =   1
      GradientIniColor=   12648447
   End
   Begin Presupuestos.chameleonButton chameleonButton1 
      Height          =   495
      Left            =   1500
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Salir al Menú"
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
      MICON           =   "Form12.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Propietario\Mis documentos\PRESUPUESTO VB\MI BASE.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Consulta7"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Propietario\Mis documentos\PRESUPUESTO VB\MI BASE.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Consulta6"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   " Análisis Incluidos en el Presupuesto  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFC0&
         Height          =   1620
         ItemData        =   "Form12.frx":001C
         Left            =   45
         List            =   "Form12.frx":001E
         TabIndex        =   1
         Top             =   240
         Width           =   4590
      End
   End
   Begin VB.Label lblAncho 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   3360
      Width           =   60
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'DECLARACIONES PARA EL TOOLTYPETEXT
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
Public apu11 As String
Public unidad11 As String
Public valor11 As String
Public controlito2 As String
Public final11 As Currency

Private Sub chameleonButton1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & ("\MI BASE.mdb")
Data2.DatabaseName = App.Path & ("\MI BASE.mdb")
Data3.DatabaseName = App.Path & ("\MI BASE.mdb")

'Carga los datos que vienen del formulario 1
obra11 = Form1.obra11
controlito2 = Form1.controlito2
Form12.Caption = "Presup Ver. 5.0  Módulo Impresión de Análisis Unitarios"
Data1.DatabaseName = App.Path & ("\MI BASE.mdb")
Data1.RecordSource = "SELECT * FROM Consulta6  WHERE OBRA='" & obra11 & "' order by APU ASC"
Data1.Refresh
Dim i As Integer
Do While Not Data1.Recordset.EOF
       List1.AddItem IIf(IsNull(Data1.Recordset("APU")), "", Data1.Recordset("APU")), i
       Data1.Recordset.MoveNext
       i = i + 1
Loop


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

Private Sub List1_click()

Dim cdname As String
cdname = List1.List(List1.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname '+ "*"
    End If
    apu11 = cdname
    Data2.DatabaseName = App.Path & ("\MI BASE.mdb")
    Data2.RecordSource = "SELECT * FROM Consulta7  WHERE APU='" & cdname & "' "
    Data2.Refresh
    unidad11 = Data2.Recordset("UNIDAD")
    Me.Data2.RecordSource = "Select SUM(SUBTOTAL) As Total from Consulta7 " & "WHERE APU like '" & cdname & "' "
    Me.Data2.Refresh
    valor11 = Format(Data2.Recordset!total, "$##,##0.00")
        
    If controlito2 = "ok" Then
        DataEnvironment1.Commands("Command3_Grouping").Parameters("apu11") = apu11
        Unitarios.Sections("Sección2").Controls("Etiqueta23").Caption = apu11
        Unitarios.Sections("Sección2").Controls("Etiqueta22").Caption = unidad11
        Unitarios.Sections("Sección5").Controls("Etiqueta14").Caption = valor11
        Unitarios.Sections("Sección3").Controls("Etiqueta4").Caption = "Obra : " & obra11
        Unitarios.Show
    End If
    Data3.DatabaseName = App.Path & ("\MI BASE.mdb")
    Data3.RecordSource = "SELECT * FROM Consulta12  WHERE APU='" & cdname & "' "
    Data3.Refresh
    Me.Data3.RecordSource = "Select SUM(FINAL) As Total from Consulta12 " & "WHERE APU like '" & cdname & "' "
    Me.Data3.Refresh
    final11 = Format(Data3.Recordset!total, "$##,##0.00")
    If controlito2 = "no" Then
        DataEnvironment1.Commands("Command3_Grouping").Parameters("apu11") = apu11
        Unitarios2.Sections("Sección2").Controls("Etiqueta23").Caption = apu11
        Unitarios2.Sections("Sección2").Controls("Etiqueta22").Caption = unidad11
        Unitarios2.Sections("Sección3").Controls("Etiqueta4").Caption = "Obra : " & obra11
        Unitarios2.Sections("Sección5").Controls("Etiqueta13").Caption = Format(final11, "$##,##0.00")
        Unitarios2.Show
    End If
End If

End Sub
