VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form FrmListado 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Archivos Disponibles ..."
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5175
   Icon            =   "FrmListado.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbalIml6.vbalImageList ilsIcons 
      Left            =   900
      Top             =   180
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   17860
      Images          =   "FrmListado.frx":014A
      KeyCount        =   19
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin vbAcceleratorGrid6.vbalGrid vbalGrid1 
      Height          =   5010
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   8837
      BackgroundPicture=   "FrmListado.frx":472E
      BackgroundPictureHeight=   40
      BackgroundPictureWidth=   40
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      DisableIcons    =   -1  'True
   End
End
Attribute VB_Name = "FrmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StrColListado As Variant
Public gRecordset As ADODB.Recordset

Private Sub Form_Load()
Dim i As Integer
     With vbalGrid1
          .Clear False
          .Redraw = False
          .RowMode = True
          .HeaderFlat = True
          .ImageList = FrmListado.ilsIcons
          
          '-> Agregar Columnas <-
          .AddColumn "col1", gRecordset.Fields(0).Name, , 1
          .AddColumn "col2", gRecordset.Fields(1).Name
          .AddColumn "col3", gRecordset.Fields(2).Name
          .AddColumn "col4", gRecordset.Fields(3).Name
          .KeySearchColumn = 1
          .SetHeaders
          
          gRecordset.MoveFirst
          i = 1
          While Not gRecordset.EOF
               .CellDetails i, 1, gRecordset.Fields(0).Value
               .CellDetails i, 2, gRecordset.Fields(1).Value
               .CellDetails i, 3, gRecordset.Fields(2).Value
               .CellDetails i, 4, gRecordset.Fields(3).Value
               i = i + 1
               gRecordset.MoveNext
          Wend
          .Redraw = True
     End With
End Sub

Private Sub vbalGrid1_ColumnClick(ByVal lCol As Long)
Dim iCol As Long
   With vbalGrid1.SortObject
      .Clear
      .SortColumn(1) = lCol
      If (vbalGrid1.ColumnSortOrder(lCol) = CCLOrderNone) Or (vbalGrid1.ColumnSortOrder(lCol) = CCLOrderDescending) Then
         .SortOrder(1) = CCLOrderAscending
      Else
         .SortOrder(1) = CCLOrderDescending
      End If
      vbalGrid1.ColumnSortOrder(lCol) = .SortOrder(1)
      .SortType(1) = vbalGrid1.ColumnSortType(lCol)
      vbalGrid1.KeySearchColumn = lCol
      For iCol = 1 To vbalGrid1.Columns
         If (iCol <> lCol) Then
            If vbalGrid1.ColumnImage(iCol) >= 1 And vbalGrid1.ColumnImage(iCol) <= 2 Then
               vbalGrid1.ColumnImage(iCol) = 0
            End If
         ElseIf vbalGrid1.ColumnHeader(iCol) <> "" Then
            vbalGrid1.ColumnImageOnRight(iCol) = True
            If (.SortOrder(1) = CCLOrderAscending) Then
               vbalGrid1.ColumnImage(iCol) = 1
            Else
               vbalGrid1.ColumnImage(iCol) = 2
            End If
         End If
      Next iCol
   End With
   Screen.MousePointer = vbHourglass
   vbalGrid1.Sort
   Screen.MousePointer = vbDefault
End Sub

Private Sub vbalGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
Dim i As Long
    If lRow > 0 Then
      ReDim StrColListado(vbalGrid1.Columns - 1)
      For i = 1 To vbalGrid1.Columns
         StrColListado(i - 1) = StrColListado(i - 1) + vbalGrid1.CellText(lRow, i)
      Next
      Unload FrmListado
    Else
      MsgBox "No se ha seleccionado ningún archivo", vbCritical + vbOKOnly, "SFC - Reutilizables"
    End If
End Sub
