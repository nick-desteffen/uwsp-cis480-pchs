Attribute VB_Name = "Get_Colors"
Option Explicit

'******************************************************************
'* This function will return the number of the color of the animal,
'* if the color is not found it will insert it into the database.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-20-2002
'******************************************************************
Function Get_Colors(strColorName As String) As Integer

Dim rstInsert As ADODB.Recordset    'Used for interfacing with the database
Dim strSQL As String                'SQL Statement
Dim bolTypeFound As Boolean         'Boolean used to determine if a match is found
Dim intMsgBox As Integer            'Used for messageboxes

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset
strColorName = Replace(strColorName, "'", "")

bolTypeFound = False

While bolTypeFound = False

    strSQL = "SELECT COLOR_NUMBER FROM COLOR WHERE COLOR_NAME = '" & strColorName & "'"
    Set rstInsert = Open_Recordsets.objConnection.Execute(strSQL)
    
    If rstInsert.EOF = False Then
        With rstInsert
            rstInsert.MoveFirst
            Do While Not rstInsert.EOF
            
                If Not IsNull(![COLOR_NUMBER]) Then
                    Get_Colors = (![COLOR_NUMBER])
                    bolTypeFound = True
                End If
                
                rstInsert.MoveNext
            Loop
        End With
    End If

    If bolTypeFound = False Then
        strSQL = "INSERT INTO COLOR (COLOR_NAME) VALUES ('" & strColorName & "')"
        Open_Recordsets.objConnection.Execute (strSQL)
    End If
Wend

rstInsert.Close
Set rstInsert = Nothing
Exit Function
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End

End Function
