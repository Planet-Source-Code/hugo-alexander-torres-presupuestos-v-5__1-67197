Attribute VB_Name = "modDatabase"
Option Explicit
    
'Add Reference for Microsoft ADO 2.1 Library
'otherwise it won't work if u will try to copy & paste
'this code to your project.

'Declaration

Public cn As Connection
Public rsUSUARIO As Recordset 'Recordset for users


'It creates connection and on successful connection
'allows to proceed other work
Public Sub Main()
    'Error Handling (Suppress Any Error during Program execution)
    On Error Resume Next
    
    
    
    'Database Connection Code
    Set cn = New Connection
    cn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;data source=" & App.Path & "\MI BASE.mdb;"
    cn.CursorLocation = adUseClient
    cn.Open
    
    If cn.State = adStateOpen Then
        'when no database connection error occurs
        Set rsUSUARIO = New Recordset
        rsUSUARIO.CursorLocation = adUseClient
        rsUSUARIO.Open "Select usuar from usuario", cn, adOpenKeyset, adLockPessimistic
        
        frmLogin.Show
                
    Else
        'when database connection error occurs
        MsgBox "Error de conexión", vbCritical, "Presupuestos"
        End
    End If
    
End Sub



