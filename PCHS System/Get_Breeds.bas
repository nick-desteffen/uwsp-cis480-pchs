Attribute VB_Name = "Get_Breeds"
Option Explicit

'******************************************************************
'* This function will return the number of the breed of the animal,
'* if the breed is not found it will insert it into the database.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-20-2002
'******************************************************************
Function Get_Breeds(intType As Integer, strBreed As String) As Integer

Dim rstInsert As ADODB.Recordset     'Recordset used for interfacing with the database
Dim bolTypeFound As Boolean          'Boolean used to determine if a match is found
Dim strSQL As String                 'SQL Statement
Dim intMsgBox As Integer             'Used for messageboxes

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset
strBreed = Replace(strBreed, "'", "''")

bolTypeFound = False

While bolTypeFound = False

If (intType = 1 Or intType = 2) Then
    If intType = 1 Then
        strSQL = "SELECT BREED_NUMBER FROM DOG_BREEDS WHERE BREED_NAME = '" & strBreed & "'"
    ElseIf intType = 2 Then
        strSQL = "SELECT BREED_NUMBER FROM CAT_BREEDS WHERE BREED_NAME = '" & strBreed & "'"
    End If

    Set rstInsert = Open_Recordsets.objConnection.Execute(strSQL)

    If rstInsert.EOF = False Then
        With rstInsert
            rstInsert.MoveFirst
                Do While Not rstInsert.EOF
    
                If Not IsNull(![BREED_NUMBER]) Then
                    Get_Breeds = (![BREED_NUMBER])
                    bolTypeFound = True
                End If
                
                rstInsert.MoveNext
            Loop
        End With
    End If
End If

If bolTypeFound = False Then
    If intType = 1 Then
        strSQL = "INSERT INTO DOG_BREEDS (BREED_NAME) VALUES ('" & strBreed & "')"
    ElseIf intType = 2 Then
        strSQL = "INSERT INTO CAT_BREEDS (BREED_NAME) VALUES ('" & strBreed & "')"
    End If
    
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
