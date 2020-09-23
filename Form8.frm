VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presup Ver. 5.0   Edición/Creación de Equipos & Herr.."
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   Icon            =   "Form8.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5865
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   5895
      Begin VB.TextBox txtobservacion 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         Tag             =   "S"
         Top             =   2880
         Width           =   3975
      End
      Begin VB.TextBox txtproveedor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Tag             =   "S"
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox txtvalor 
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
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Tag             =   "S"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtunidad 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Tag             =   "S"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtnombre 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Tag             =   "S"
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox txtClave 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Tag             =   "S"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtidentificador 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Tag             =   "S"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Tag             =   "S"
         Top             =   720
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   3840
         Picture         =   "Form8.frx":0442
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Observaciones"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Proveedor"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Valor"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Unidad de medida"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Nombre Eqp/Herr"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Clave"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Código"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Identificador"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin Presupuestos.chameleonButton REGRESAR 
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      ToolTipText     =   "Regresa al menú Principal"
      Top             =   3720
      Width           =   1575
      _extentx        =   2778
      _extenty        =   873
      btype           =   3
      tx              =   "Regresar al Menú"
      enab            =   -1
      font            =   "Form8.frx":13B4
      coltype         =   1
      focusr          =   -1
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Form8.frx":13E0
      picn            =   "Form8.frx":13FE
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
      TabIndex        =   5
      ToolTipText     =   "Conecta con la Base Materiales"
      Top             =   3720
      Width           =   1695
      _extentx        =   2990
      _extenty        =   873
      btype           =   3
      tx              =   "Conectar Base de Datos"
      enab            =   -1
      font            =   "Form8.frx":171A
      coltype         =   1
      focusr          =   -1
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Form8.frx":1746
      picn            =   "Form8.frx":1764
      umcol           =   -1
      soft            =   0
      picpos          =   0
      ngrey           =   0
      fx              =   0
      hand            =   0
      check           =   0
      value           =   0
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
         Picture         =   "Form8.frx":1A80
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
            Picture         =   "Form8.frx":1BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form8.frx":1D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form8.frx":1E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form8.frx":1FE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form8.frx":213E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form8.frx":229A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form8.frx":2836
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form8.frx":2DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form8.frx":336E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form8.frx":390A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form8.frx":3EB2
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
            Object.ToolTipText     =   "Grabar Nuevo Insumo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Editar Insumo Grabado"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar Insumo"
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
            Object.ToolTipText     =   "Buscar Insumo"
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
      TabIndex        =   24
      Top             =   4275
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gConexion   As ADODB.Connection
Dim gRsINSUMOS As New ADODB.Recordset
Dim sModo       As String
Dim AJMAT As Double
Dim AJEQP As Double
Dim AJMDO As Double

Private Sub CmdBuscar_Click()
     Dim FilaRetorno As Variant
     
     If TxtIDbuscar.Text <> "" Then
     Else
          FrameBuscar.Visible = False
          Set FrmListado.gRecordset = gRsINSUMOS
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
     gRsINSUMOS.MoveFirst
     ' Buscar el ID igual al indicado, hacia adelante(adSearchForward)
     gRsINSUMOS.Find "ID=" & TxtIDbuscar.Text, , adSearchForward
     ' Si es EOF -> No encontró
     If gRsINSUMOS.EOF Then
          MsgBox "Código Insumo No encontrado!", , Me.Caption
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
          
          Set gRsINSUMOS = New ADODB.Recordset
          Screen.MousePointer = vbHourglass
          ' Abrir el recordset utilizado a lo largo de la ejecución...
          gRsINSUMOS.Open "SELECT * FROM INSUMOS WHERE CLAVE= 'EQP'ORDER BY id", gConexion, adOpenStatic, adLockOptimistic
          Screen.MousePointer = vbDefault
          If gRsINSUMOS.RecordCount > 0 Then
               gRsINSUMOS.MoveFirst
               StatusBar1.Panels(1).Text = " Identificador de Equipos/Herram. en uso : " & gRsINSUMOS.RecordCount
               Call CargarDatos
          Else
               Call AccionBotones(Toolbar1, tNoRegistros)
               MsgBox "No Existen Insumos!", , Me.Caption
          End If
     Else
          MsgBox "Errores al Conectar!"
          Call AccionBotones(Toolbar1, tDesHabilitar)
     End If
     CmdConectar.Enabled = False
End Sub
Private Sub Form_Load()
     CmdConectar.Enabled = True
     sModo = "Ver"
End Sub

Private Sub form_Unload(Cancel As Integer)
     ' Liberar Memoria
     Set gConexion = Nothing
     Set gRsINSUMOS = Nothing
     Set DataEnvironment1 = Nothing
End Sub

Private Sub REGRESAR_Click()
    Unload Me
    Form4.Show
    
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
               RsAux.Open "SELECT  MAX(ID)+1  FROM INSUMOS", gConexion
               txtidentificador.Text = IIf(IsNull(RsAux.Fields(0).Value), 1, RsAux.Fields(0).Value)
               txtcodigo.Text = Val(Str(txtidentificador.Text)) + 3309
               txtClave.Text = "EQP"
               txtnombre.SetFocus
               sModo = "Nuevo"
          Case "modificar":
               Call AccionBotones(Toolbar1, tEditando)
               Call ActivarControles(Me, True)
               sModo = "Modificar"
          Case "eliminar":
               ' Borrar el Registro Actual
               sSQl = "DELETE FROM INSUMOS WHERE id = " & txtidentificador.Text
               RsAux.Open sSQl, gConexion
               
               ' Volver a Consultar la tabla
               gRsINSUMOS.Requery
               Call CargarDatos
               sModo = "Ver"
          Case "grabar":
               AJMAT = 0
               AJEQP = 0
               AJMDO = 0
               If Not Validar_Datos() Then
                    Exit Sub
               End If
               ' Si es un registro Nuevo, Agregar usando SQL (INSERT INTO tabla1 VALUES (...))
               If sModo = "Nuevo" Then
                    
                    sSQl = "INSERT INTO INSUMOS(ID, COD_INSUMO, CLAVE, INSUMO, UN_INSUMO, AJUSTE_MAT, AJUSTE_EQP, AJUSTE_MDO, VALOR, PROVEEDOR, OBSERVACIONES) " & _
                           "VALUES " & _
                           "(" & txtidentificador.Text & ",'" & txtcodigo.Text & "','" & txtClave.Text & "','" & txtnombre.Text & "','" & txtunidad.Text & "','" & txtvalor.Text & "','" & txtproveedor.Text & "','" & txtobservacion.Text & "')"
                    
               Else
                    ' Si es un registro existente, Modificar usando SQL(UPDATE tabla SET campo1 = expr1, ...)
                    sSQl = "UPDATE PROYECTOS " & _
                           "SET " & _
                           "    COD_INSUMO = '" & txtcodigo.Text & "'," & _
                           "    CLAVE      = '" & txtClave.Text & "'," & _
                           "    INSUMO     = '" & txtnombre.Text & "'," & _
                           "    UN_INSUMO  = '" & txtunidad.Text & "'," & _
                           "    VALOR      = '" & txtvalor.Text & "'," & _
                           "    PROVEEDOR  = '" & txtproveedor.Text & "'," & _
                           "    OBSERVACIONES    = '" & txtobservacion.Text & " '" & _
                           "WHERE ID = " & txtidentificador.Text

               End If
               RsAux.Open sSQl, gConexion
               Call AccionBotones(Toolbar1, tMover)
               Call ActivarControles(Me, False)
               gRsINSUMOS.Requery
               Call CargarDatos
               sModo = "Ver"
          Case "cancelar":
               Call AccionBotones(Toolbar1, tMover)
               Call ActivarControles(Me, False)
               gRsINSUMOS.Requery
               Call CargarDatos
               sModo = "Ver"
        '--------->
        Case "anterior"
            If Not gRsINSUMOS.BOF Then
                gRsINSUMOS.MovePrevious
                If gRsINSUMOS.BOF Then
                    gRsINSUMOS.MoveFirst
                End If
                Call CargarDatos
            End If
        Case "siguiente"
            If Not gRsINSUMOS.EOF Then
                gRsINSUMOS.MoveNext
                If gRsINSUMOS.EOF Then
                    gRsINSUMOS.MoveLast
                End If
                Call CargarDatos
            End If
        Case "inicio"
            If Not gRsINSUMOS.BOF Then
                gRsINSUMOS.MoveFirst
                Call CargarDatos
            End If
        Case "final"
            If Not gRsINSUMOS.EOF Then
                gRsINSUMOS.MoveLast
                Call CargarDatos
            End If
        Case "buscar"
               ' Mostrar el frame de Busqueda
               FrameBuscar.Visible = True
               TxtIDbuscar.Enabled = True
               TxtIDbuscar.Text = ""
               TxtIDbuscar.SetFocus
        Case "imprimir"
               ' Asignar los valores a los controles de la seccion
               DataEnvironment1.Commands("Command2").CommandText = "INSUMOS WHERE CLAVE= 'EQP'ORDER BY id"
               equipos.Show
        '>-----------
     
     End Select
     Set RsAux = Nothing
End Sub


Private Sub CargarDatos()
     ' Si el Recordset Tiene Datos...
     If Not gRsINSUMOS.EOF Then
          ' En todos los campos validar que no hayan NULOS en los mismos.
          txtidentificador.Text = IIf(IsNull(gRsINSUMOS!ID), 0, gRsINSUMOS!ID)
          txtcodigo.Text = IIf(IsNull(gRsINSUMOS!cod_insumo), 0, gRsINSUMOS!cod_insumo)
          txtClave.Text = IIf(IsNull(gRsINSUMOS!Clave), "", gRsINSUMOS!Clave)
          txtnombre.Text = IIf(IsNull(gRsINSUMOS!insumo), "", gRsINSUMOS!insumo)
          txtunidad.Text = IIf(IsNull(gRsINSUMOS!un_insumo), "", gRsINSUMOS!un_insumo)
          txtvalor.Text = IIf(IsNull(gRsINSUMOS!valor), "", gRsINSUMOS!valor)
          txtvalor.Text = Format(txtvalor.Text, "$##,##0.00")
          txtproveedor.Text = IIf(IsNull(gRsINSUMOS!proveedor), "", gRsINSUMOS!proveedor)
          txtobservacion.Text = IIf(IsNull(gRsINSUMOS!observaciones), "", gRsINSUMOS!observaciones)
                  
          
     Else
          ' Si no Hay datos, Solo mostrar el boton de Agregar
          Call EncerarControles(Me)
          Call AccionBotones(Toolbar1, tNoRegistros)
          MsgBox "No Existen Insumos!", , Me.Caption
     End If
End Sub
Function Validar_Datos() As Boolean
     If Trim(txtnombre.Text) = "" Then
          MsgBox "Debe Ingresar nombre del Insumo..."
          Validar_Datos = False
          Exit Function
     End If
     If Trim(txtunidad.Text) = "" Then
          MsgBox "Debe Ingresar la unidad de medida..."
          Validar_Datos = False
          Exit Function
     End If
     If Trim(txtvalor.Text) = "" Then
          MsgBox "Debe Ingresar el valor de el Insumo..."
          Validar_Datos = False
          Exit Function
     End If
     If Trim(txtproveedor.Text) = "" Then
          MsgBox "Debe Ingresar el nombre de el Proveedor..."
          Validar_Datos = False
          Exit Function
     End If
       
     Validar_Datos = True
End Function







