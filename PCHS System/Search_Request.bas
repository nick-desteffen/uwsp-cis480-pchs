Attribute VB_Name = "Search_Request"
Option Explicit

'************************************************************************
'* This function searches through the request table to see if the new animal
'* entered into the system is similar to an animal that has been requested.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-22-2002
'*************************************************************************

Public Sub Search_Requests(intType, intBreed, intColor, strAge, strSex)

Dim rstRequest As ADODB.Recordset   'Used for interfacing with the database
Dim strSQL As String                'SQL Statement
Dim intRequest As Integer           'Request number if found
Dim intMsgBox As Integer            'Used for messageboxes

On Error GoTo ErrorHandler

Set rstRequest = New ADODB.Recordset

strSQL = "SELECT REQUEST_NUMBER FROM REQUESTS WHERE REQUEST_TYPE = " & intType
strSQL = strSQL & " AND REQUEST_BREED = " & intBreed
strSQL = strSQL & " AND (REQUEST_COLOR = " & intColor & " OR REQUEST_COLOR = 0)"
strSQL = strSQL & " AND (REQUEST_AGE = '" & strAge & "' OR REQUEST_AGE = 'UNSPECIFIED')"
strSQL = strSQL & " AND (REQUEST_SEX = '" & strSex & "' OR REQUEST_SEX = 'U')"

Set rstRequest = Open_Recordsets.objConnection.Execute(strSQL)

If rstRequest.EOF = False Then
    With rstRequest
        rstRequest.MoveFirst
    
        Do While Not rstRequest.EOF
            
            If Not IsNull(![REQUEST_NUMBER]) Then
                intRequest = ![REQUEST_NUMBER]
                frmRequestFound.intRequest = intRequest
                frmRequestFound.Show
            End If
        rstRequest.MoveNext
        Loop
    End With
End If

   
rstRequest.Close
Set rstRequest = Nothing
Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

