VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{75AACB43-0B9D-11D2-95B5-0000B43369D3}#1.2#0"; "ARFrmExt.ocx"
Begin VB.Form Form11 
   AutoRedraw      =   -1  'True
   Caption         =   "Presup Ver. 5.0 - Advertencia..."
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin ARFormExtenderCtrl.ARFormExtender ARFormExtender1 
      Left            =   2160
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      BackgroundType  =   1
      GradientIniColor=   16777152
   End
   Begin VB.Frame Frame1 
      Caption         =   "MS PROJECT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   4695
      Begin Presupuestos.chameleonButton chameleonButton2 
         Height          =   500
         Left            =   2520
         TabIndex        =   4
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Exportar"
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
         MICON           =   "Form11.frx":030A
         PICN            =   "Form11.frx":0326
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Presupuestos.chameleonButton chameleonButton1 
         Height          =   500
         Left            =   480
         TabIndex        =   3
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
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
         MICON           =   "Form11.frx":0640
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSForms.CheckBox CheckBox1 
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   960
         Width           =   3855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "6800;661"
         Value           =   "0"
         Caption         =   "Abrir MS Project y editar la Programacion"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Para Poder Exportar el Presupuesto deberá tener instalado en su equipo el Software MS Project de Microsoft"
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public obra11 As String
Public proponente11 As String
Public administracion11 As Double
Public imprevistos11 As Double
Public utilidades11 As Double
Public iva11 As Double
Dim cnn As New ADODB.Connection
Dim comando As ADODB.Command
Dim query As ADODB.Recordset
Dim QUERYACT As ADODB.Recordset
Dim P As Object
Dim task As Integer
Dim BPar As Boolean
Private Sub chameleonButton1_Click()
    Unload Me
    Form4.Show
End Sub

Private Sub chameleonButton2_Click()

On Error GoTo ErrorProject
BPar = False
Set P = CreateObject("MSProject.Application")
If CheckBox1.Value = False Then
    P.Visible = False
    Else
    P.Visible = True
End If
P.FileNew
'INSERTA NUEVAS COLUMNAS
TableEdit Name:="Entrada", TaskTable:=True, _
        NewFieldName:="Texto1", Title:="U.M.", Width:=8, _
        ShowInMenu:=True, DateFormat:=pjDate_mm_dd_yy, _
        ColumnPosition:=2
 TableEdit Name:="Entrada", TaskTable:=True, _
        NewFieldName:="Número1", Title:="Cant.", Width:=12, _
        ShowInMenu:=True, DateFormat:=pjDate_mm_dd_yy, _
        ColumnPosition:=3
 TableEdit Name:="Entrada", TaskTable:=True, _
        NewFieldName:="Costo", Title:="Costo Total                          IVA Incluido", Width:=18, _
        ShowInMenu:=True, DateFormat:=pjDate_mm_dd_yy, _
        ColumnPosition:=4
TableApply "Entrada"
Call Conexion
Call GProject
Unload Me
Exit Sub

ErrorProject:
Dim Msj
' Comprueba el error, después muestra un mensaje.
If Err.Number <> 0 Then
   Msj = "El Error No. " & Str(Err.Number) & " fué generado por " _
         & Err.Source & Chr(13) & Err.Description & ", inténtelo después nuevamente - Gracias.."
   MsgBox Msj, , "Presup Ver. 5.0 - Error", Err.HelpFile, Err.HelpContext
End If

Unload Me

End Sub
Private Sub Conexion()

With cnn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Data Source") = App.Path & "\MI BASE.mdb"
    '.Properties("Jet OLEDB:Database Password") = "2822con"
    .CursorLocation = adUseClient
    .Open
End With

Set comando = New ADODB.Command
Set query = New ADODB.Recordset

comando.ActiveConnection = cnn
comando.CommandType = adCmdText

End Sub

Private Sub GProject()
'DATOS TRAIDOS DEL FORMULARIO 1
obra11 = Form1.obra11
proponente11 = Form1.proponente11
Dim Refe3 As Double
Refe3 = Val(Form1.administracion11) + Val(Form1.imprevistos11) + Val(Form1.utilidades11)
iva11 = Form1.iva11
utilidades11 = Form1.utilidades11

comando.CommandText = "SELECT CAPITULO FROM Consulta5 WHERE OBRA='" & obra11 & "' "
Set QUERYACT = comando.Execute()
Do While QUERYACT.EOF = False
    P.ActiveProject.Tasks.Add Name:="" & Trim(QUERYACT!capitulo)
    If BPar = True Then
        P.ActiveProject.Tasks(Trim(QUERYACT!capitulo)).OutlineOutdent
        BPar = False
    End If
    comando.CommandText = "SELECT APU, UNIDAD, CANTIDAD, VALOR_TOTAL FROM PRESUPUESTO WHERE CAPITULO='" & QUERYACT!capitulo & "' "
    Set query = comando.Execute()
    Do While query.EOF = False
        P.ActiveProject.Tasks.Add Name:="" & query!APU
        P.ActiveProject.Tasks(Trim(query!APU)).Text1 = "" & query!unidad
        P.ActiveProject.Tasks(Trim(query!APU)).Number1 = "" & query!cantidad
        P.ActiveProject.Tasks(Trim(query!APU)).Cost = "" & ((query!VALOR_TOTAL) * (1 + (Refe3 / 100))) + (((query!VALOR_TOTAL) * (utilidades11 / 100) * (iva11 / 100)) / (1 + (Refe3 / 100)))
        If l = 0 Then
            P.ActiveProject.Tasks(Trim(query!APU)).OutlineIndent
            l = 1
        End If
    query.MoveNext
    Loop
    l = 0
    BPar = True
    QUERYACT.MoveNext
Loop
'ELIMINA COLUMNA
P.ColumnDelete
'AJUSTA EL ANCHO DE LA COLUMNA
P.ColumnBestFit
'ENCABEZADO Y PIE DE PAGINA
P.FilePageSetupHeader , pjCenter, "OBRA: '" & obra11 & "' "
P.FilePageSetupFooter , pjLeft, "PROPONENTE: '" & proponente11 & "'  "
'PROPIEDADES DE VISTA
P.FilePageSetupView , False, , False, False, False
'PROPIEDADES DE LEGENDA
P.FilePageSetupLegend , , pjNoLegend
If CheckBox1.Value = False Then P.Quit
'LIBERA MEMORIA
cnn.Close
Set cnn = Nothing
Set comando = Nothing
Set query = Nothing
Set QUERYACT = Nothing
End Sub

Private Sub form_Unload(Cancel As Integer)
     ' Liberar Memoria
     Set cnn = Nothing
     Set comando = Nothing
     Set query = Nothing
     Set QUERYACT = Nothing
     
End Sub

