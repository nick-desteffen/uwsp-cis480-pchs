Attribute VB_Name = "Search_Person"
Option Explicit

'************************************************************************
'* This function searches through the person table to see if the person
'* already has their info stored there.  It searches for first and last names
'* and returns true if found, false if not.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-22-2002
'*************************************************************************
Public Sub Search_Person(ByRef strFname As String, ByRef strLname As String, _
                         ByRef strAddress As String, ByRef strCity As String, _
                         ByRef strState As String, ByRef strZip As String, _
                         ByRef strPhone As String, ByRef strEmail As String, _
                         ByRef strLicense As String, ByRef dteDOB As Date, _
                         ByVal intPersonNum As Integer)
                         
                         
Dim rstSearchPerson As ADODB.Recordset  'Recordset used for interfacing with the database
Dim strSQL As String                    'SQL Statement
Dim intMsgBox As Integer                'Used for messageboxes

On Error GoTo ErrorHandler

Set rstSearchPerson = New ADODB.Recordset

strSQL = "SELECT * FROM PERSON WHERE PERSON_NUMBER = " & intPersonNum
Set rstSearchPerson = Open_Recordsets.objConnection.Execute(strSQL)


If rstSearchPerson.EOF = False Then
    With rstSearchPerson
        rstSearchPerson.MoveFirst

        Do While Not rstSearchPerson.EOF
                
                
                If Not IsNull(![PERSON_FNAME]) Then
                    strFname = ![PERSON_FNAME]
                Else
                    strFname = ""
                End If
                
                If Not IsNull(![PERSON_LNAME]) Then
                    strLname = ![PERSON_LNAME]
                Else
                    strLname = ""
                End If
                
                If Not IsNull(![person_address]) Then
                    strAddress = ![person_address]
                Else
                    strAddress = ""
                End If
                
                If Not IsNull(![person_city]) Then
                    strCity = ![person_city]
                Else
                    strCity = ""
                End If
                
                If Not IsNull(![person_state]) Then
                    strState = ![person_state]
                Else
                    strState = ""
                End If
                
                If Not IsNull(![person_zip]) Then
                    strZip = ![person_zip]
                Else
                    strZip = ""
                End If
                
                If Not IsNull(![person_telephone]) Then
                    strPhone = ![person_telephone]
                Else
                    strPhone = ""
                End If
                
                If Not IsNull(![person_email]) Then
                    strEmail = ![person_email]
                Else
                    strEmail = ""
                End If
                
                If Not IsNull(![person_license]) Then
                    strLicense = ![person_license]
                Else
                    strLicense = ""
                End If
                
                If Not IsNull(![person_dob]) Then
                    dteDOB = ![person_dob]
                Else
                    dteDOB = dteDOB
                End If
                
                If Not IsNull(![PERSON_NUMBER]) Then
                    intPersonNum = ![PERSON_NUMBER]
                Else
                    intPersonNum = 0
                End If

            rstSearchPerson.MoveNext
        Loop
    rstSearchPerson.Close
    Set rstSearchPerson = Nothing
    End With
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub
