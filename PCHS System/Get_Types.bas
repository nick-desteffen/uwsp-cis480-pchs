Attribute VB_Name = "Get_Types"
Option Explicit

'******************************************************************
'* This function will return the number of the type of the animal,
'* if the type is not found it will insert it into the database.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-20-2002
'******************************************************************
Function Get_Types(strTypeName As String) As Integer

Dim rstInsert As ADODB.Recordset    'Used for interfacing with the database
Dim bolTypeFound As Boolean         'Boolean used to determine if a match is found
Dim strSQL As String                'SQL Statement
Dim intMsgBox As Integer            'Used for messageboxes

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset

strTypeName = Replace(strTypeName, "'", "")

bolTypeFound = False

While bolTypeFound = False
    
        strSQL = "SELECT TYPE_NUMBER FROM ANIMAL_TYPES WHERE TYPE_NAME = '" & strTypeName & "'"
        Set rstInsert = Open_Recordsets.objConnection.Execute(strSQL)
    
        If rstInsert.EOF = False Then
            With rstInsert
                rstInsert.MoveFirst
                Do While Not rstInsert.EOF
                
                    If Not IsNull(![TYPE_NUMBER]) Then
                        Get_Types = (![TYPE_NUMBER])
                        bolTypeFound = True
                    End If
                    rstInsert.MoveNext
                Loop
            End With
        End If

    If bolTypeFound = False Then
        strSQL = "INSERT INTO ANIMAL_TYPES (TYPE_NAME) VALUES ('" & strTypeName & "')"
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
