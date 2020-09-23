Attribute VB_Name = "Module1"
Option Explicit
Enum eTiposBases
     tSQLserver
     tAccess
End Enum

Enum AccionesBotones
     tHabilitar     ' Habilitar los botones
     tDesHabilitar  ' Deshabilitar los botones
     tEditando      ' Solo Botones Grabar y Deshacer Habilitados.
     tMover         ' Solo Botones Grabar y Deshacer DesHabilitados.
     tNoRegistros   ' Solo Agregar Habilitador
End Enum


Public Function AbreConexion(xbase As ADODB.Connection, TipoBasedeDatos As eTiposBases, vCursorLocation As ADODB.CursorLocationEnum) As Boolean
    On Error GoTo ErrorConexion
    Dim str_con As String
    
    ' Para SQL server, Hay que indicar el Nombre de la BD y el Nombre del Servidor.
    If TipoBasedeDatos = tSQLserver Then
          str_con = "Provider=SQLOLEDB.1;Persist Security Info=False;" & _
                   "Data Source=\\PC_SERVIDOR" & _
                   ";Initial Catalog=OTbases" & _
                   ";User Id=sa;Password=; "
    End If
    ' Para Microsoft Access hay que indicar el nombre del archivo .MDB
    If TipoBasedeDatos = tAccess Then
               str_con = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & _
                   App.Path & "\MI BASE.mdb;Mode=Read|Write"
    End If
    Set xbase = New ADODB.Connection
    
    'Cadena de conexion
    xbase.ConnectionString = str_con
    xbase.CursorLocation = vCursorLocation
    xbase.ConnectionTimeout = 60
    xbase.Open
    AbreConexion = True
    Exit Function
ErrorConexion:
    AbreConexion = False
    MsgBox "Error#:" & Err.Number & vbCrLf & _
           "Descripción:" & Err.Description & vbCrLf & _
           "Origen:" & Err.Source & vbCrLf, vbCritical, "Error al Abrir Conexión...(OTbases)"
End Function


Public Sub AccionBotones(ByVal Toolbar1 As Toolbar, ByVal tipoAccion As AccionesBotones)
Dim i As Integer
Dim bEstado As Boolean
Dim bEstadoSaveUndo As Boolean

     If tipoAccion = tHabilitar Or tipoAccion = tMover Then
          bEstado = True
          bEstadoSaveUndo = False
     End If
     If tipoAccion = tDesHabilitar Or tipoAccion = tEditando Then
          bEstado = False
          bEstadoSaveUndo = True
     End If
          
     For i = 1 To Toolbar1.Buttons.Count
          ' Si está Editando ACTIVAR los botones de Grabar y Deshacer.
          ' Si está en Modo Ver DESACTIVAR los botones de Grabar y Deshacer.
          If tipoAccion = tNoRegistros Then
               If Toolbar1.Buttons(i).key = "nuevo" Then
                    Toolbar1.Buttons(i).Enabled = True
               Else
                    Toolbar1.Buttons(i).Enabled = False
               End If
          Else
               If tipoAccion = tEditando Or tipoAccion = tMover Then
                    If Toolbar1.Buttons(i).key = "grabar" _
                         Or Toolbar1.Buttons(i).key = "cancelar" Then
                         Toolbar1.Buttons(i).Enabled = bEstadoSaveUndo
                    Else
                         Toolbar1.Buttons(i).Enabled = bEstado
                    End If
               Else
                    Toolbar1.Buttons(i).Enabled = bEstado
               End If
          End If
     Next i
End Sub

' Activar los controles con la propiedad TAG = "S"
Public Sub ActivarControles(miForm As Form, bEstado As Boolean)
Dim i As Integer
     For i = 0 To miForm.Controls.Count - 1
          If UCase(miForm.Controls(i).Tag) = "S" Then
               miForm.Controls(i).Enabled = bEstado
          End If
     Next i
End Sub

' A los controles TextBox, Label Y DTpicker (identificados con el prefijo TXT, LBL y FEC)
' les asigno "S" en la propiedad auxiliar TAG para saber que se
' debe encerar mediante este procedimiento.
Public Sub EncerarControles(miForm As Form)
Dim i As Integer

     'Rastrear la coleccion controls del formulario
     For i = 0 To miForm.Controls.Count - 1
          If UCase(miForm.Controls(i).Tag) = "S" Then
               If Left(UCase(miForm.Controls(i).Name), 3) = "TXT" Then
                    miForm.Controls(i).Text = ""
               End If
               If Left(UCase(miForm.Controls(i).Name), 3) = "LBL" Then
                    miForm.Controls(i).Caption = ""
               End If
               If Left(UCase(miForm.Controls(i).Name), 3) = "FEC" Then
                    miForm.Controls(i).Value = "1/1/2006"
               End If
          End If
     Next i
End Sub


