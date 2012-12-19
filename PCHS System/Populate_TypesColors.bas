Attribute VB_Name = "Populate_TypesColors"
'********************************************************************************************
'* Runs when the form is loaded.  Populates the animal type combo box, and the colors combo box.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-09-2002
'********************************************************************************************

Dim types() As combo_info
Dim colors() As combo_info
Dim looper As Integer

Dim rstType As ADODB.Recordset
Dim objConnection As ADODB.Connection

Set objConnection = New ADODB.Connection
Set rstType = New ADODB.Recordset

objConnection.ConnectionString = frmPCHS_Main.strConnectionString
objConnection.Open
looper = 0

Dim strSQL As String
'Populates the recordset
strSQL = "SELECT TYPE_NUMBER, TYPE_NAME FROM ANIMAL_TYPES"

Set rstType = objConnection.Execute(strSQL)

With rstType
    rstType.MoveFirst
    Do While Not rstType.EOF
        ReDim Preserve types(looper)
        If Not IsNull(![TYPE_NUMBER]) Then
            types(looper).Number = (![TYPE_NUMBER])
        End If
        If Not IsNull(![TYPE_NAME]) Then
            types(looper).name = (![TYPE_NAME])
        End If
    rstType.MoveNext
    looper = looper + 1
    Loop
End With
'Populates the combo box
For looper = 0 To UBound(types)
    cboAnimalType.AddItem (types(looper).name)
Next looper

'Populates the color recordset
Set rstType = Nothing
looper = 0

strSQL = "SELECT COLOR_NUMBER, COLOR_NAME FROM COLOR"

Set rstType = objConnection.Execute(strSQL)

With rstType
    rstType.MoveFirst
    Do While Not rstType.EOF
        ReDim Preserve colors(looper)
        If Not IsNull(![COLOR_NUMBER]) Then
            colors(looper).Number = (![COLOR_NUMBER])
        End If
        If Not IsNull(![COLOR_NAME]) Then
            colors(looper).name = (![COLOR_NAME])
        End If
    rstType.MoveNext
    looper = looper + 1
    Loop
End With
'Populates the combo box

For looper = 0 To UBound(colors)
   cboColor.AddItem (colors(looper).name)
Next looper

'Closes the connection
rstType.Close
objConnection.Close
Set objConnection = Nothing
Set rstType = Nothing

End Sub
