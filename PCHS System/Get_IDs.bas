Attribute VB_Name = "Get_IDs"
Option Explicit

'**************************************************************************
'* This function will return the number of the animal that is automatically
'* generated by the access database.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-26-2002
'**************************************************************************

Public Function Get_Animal_Num(strAnimalName As String, intType As Integer, intBreed As Integer, strSex As String, intColor As Integer, strAge As String, strTemperment As String, intNeutor As Integer, strAcquired As String, intMix As Integer) As Integer

Dim rstNum As New ADODB.Recordset       'Used for interfacing with the database
Dim strSQL As String                    'SQL Statement
Dim intMsgBox As Integer                'Used for messageboxes

On Error GoTo ErrorHandler

Set rstNum = New ADODB.Recordset

strSQL = "SELECT ANIMALS.ANIMAL_NUMBER FROM ANIMALS WHERE"
strSQL = strSQL & " ANIMAL_NAME = '" & strAnimalName
strSQL = strSQL & "' AND ANIMAL_TYPE = " & intType
strSQL = strSQL & " AND ANIMAL_BREED = " & intBreed
strSQL = strSQL & " AND ANIMAL_SEX = '" & strSex
strSQL = strSQL & "' AND ANIMAL_COLOR = " & intColor
strSQL = strSQL & " AND ANIMAL_AGE = '" & strAge
strSQL = strSQL & "' AND ANIMAL_TEMPERMENT = '" & strTemperment
strSQL = strSQL & "' AND ANIMAL_SPAY_NEUTER = " & intNeutor
strSQL = strSQL & " AND ANIMAL_ACQUIRED = '" & strAcquired
strSQL = strSQL & "' AND ANIMAL_MIX = " & intMix

Set rstNum = Open_Recordsets.objConnection.Execute(strSQL)

Do While rstNum.EOF = False
    With rstNum
        rstNum.MoveFirst
        Do While Not rstNum.EOF
            Get_Animal_Num = ![animal_number]
            rstNum.MoveNext
        Loop
    End With
Loop

Exit Function
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Function