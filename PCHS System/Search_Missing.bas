Attribute VB_Name = "Search_Missing"
Option Explicit

'************************************************************************
'* This function searches through the missing table to see if the new animal
'* entered into the system is similar to an animal that has been reported
'* missing.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-22-2002
'*************************************************************************

Public Sub Search_Missing(intType, intBreed, intColor, strAge, strSex)

Dim rstMissing As ADODB.Recordset   'Recordset used for interfacing with the database
Dim strSQL As String                'SQL Statement
Dim intMissing As Integer           'Missing number if found
Dim intMsgBox As Integer            'Used for messageboxes

On Error GoTo ErrorHandler

Set rstMissing = New ADODB.Recordset

strSQL = "SELECT MISSING_NUMBER FROM MISSING WHERE MISSING_TYPE = " & intType
strSQL = strSQL & " AND MISSING_BREED = " & intBreed
strSQL = strSQL & " AND MISSING_COLOR = " & intColor
strSQL = strSQL & " AND MISSING_AGE = '" & strAge & "'"
strSQL = strSQL & " AND MISSING_SEX = '" & strSex & "'"

Set rstMissing = Open_Recordsets.objConnection.Execute(strSQL)
If rstMissing.EOF = False Then
    With rstMissing
        rstMissing.MoveFirst
        Do While Not rstMissing.EOF
    
            If Not IsNull(![MISSING_NUMBER]) Then
                intMissing = ![MISSING_NUMBER]
                frmMissingFound.intMissing = intMissing
                frmMissingFound.Show
            End If
            rstMissing.MoveNext
    
        Loop
    End With
End If

rstMissing.Close
Set rstMissing = Nothing

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End

End Sub

