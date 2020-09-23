VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function rsCustomer(sConn As String) As ADODB.Recordset
On Error GoTo rsCustomerErr
    Dim objRecordset As New ADODB.Recordset
    Set objRecordset = New ADODB.Recordset
    obra11 = Form1.obra11
    With objRecordset
        'Run Query
        .ActiveConnection = sConn
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Source = "Select * from Consulta6 WHERE OBRA='" & obra11 & "' order by APU ASC"
        'Get Data
        .Open
        
        'set connection to nothing
        Set .ActiveConnection = Nothing
        
        'Return the Records
        Set rsCustomer = objRecordset
    End With

rsCustomerExit:
    Exit Function


rsCustomerErr:
    Debug.Print Err.Description
    MsgBox Err.Description, vbCritical
    Resume rsCustomerExit

End Function
Public Sub Auto_Search(KeyAscii As Integer, obj1 As Object, rs As ADODB.Recordset, sStr As String)
'Purpose : This routine allows the user to Automatically search from the List(Data Combo/Data List) when keypress is made in Data Combo Box

Dim iLen, iStart, iSelLength As Integer
Dim sCriteria As String
Dim tempstr As String

    'If control was locked then exit
    If obj1.Locked = True Then Exit Sub

    If KeyAscii = 8 Then
       tempstr = obj1.Text
       iStart = obj1.SelStart
        iSelLength = obj1.SelLength
        obj1.Text = tempstr
        If iStart > 0 Then
            obj1.SelStart = iStart - 1
            obj1.SelLength = iSelLength + 1
           Else
            obj1.SelStart = iStart
            obj1.SelLength = iSelLength
        End If
        
        KeyAscii = 0
        Exit Sub
    End If


    If KeyAscii = 27 Then KeyAscii = 0: Exit Sub
    If Chr(KeyAscii) = "'" Then KeyAscii = Asc("`")
    If Not printable(KeyAscii) Then Exit Sub



    iStart = obj1.SelStart + 1
    obj1.SelText = Chr(KeyAscii)
    KeyAscii = 0
    
    
    
    tempstr = sStr
    
    If rs.RecordCount = 0 Then
            obj1.Text = tempstr
            iStart = obj1.SelStart + 1
            iSelLength = obj1.SelLength
            obj1.SelStart = iStart - 1
            obj1.SelLength = iSelLength
            KeyAscii = 0
            Exit Sub
        End If
    iLen = obj1.SelStart + 1
    iSelLength = obj1.SelLength
    
    
    If KeyAscii <> 1 Then obj1.SelText = Chr(KeyAscii)
            With rs
              .MoveFirst
              sCriteria = obj1.ListField & " like '" & obj1.Text & "%'"
              .Find sCriteria
              If Not .EOF Then
                 obj1.Text = Trim(.Fields(obj1.ListField) & " ")
                 obj1.SelStart = iStart
                 iLen = Len(obj1.Text)
                 If iLen = 0 Then Exit Sub
                 obj1.SelLength = Len(obj1.Text) - iStart + 1
              Else
                  If tempstr <> "" Then
                      .MoveFirst
                      sCriteria = obj1.ListField & " like '" & obj1.Text & "%'"
                      .Find sCriteria
                  End If
                  If .AbsolutePosition > 0 Then
                      obj1.Text = Trim(.Fields(obj1.ListField) & " ")
                  Else
                      obj1.Text = tempstr
                  End If
                 obj1.SelStart = iStart - 1
                 iLen = Len(obj1.Text)
                 obj1.SelLength = Len(obj1.Text) - iStart + 1
              End If
      End With
End Sub

Private Function printable(ch As Integer) As Boolean
Dim chrs As String
  chrs = " ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890`~!@#$%^&*()_+-={}[]:;""',./\<>?|"
  If InStr(chrs, UCase(Chr(ch))) > 0 Then printable = True Else printable = False
End Function
