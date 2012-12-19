Attribute VB_Name = "Verify_Data"
Option Explicit

'**************************************************************
'* Checks to make sure that the zip code is valid or not.
'* Accepts zip codes in the following format XXXXX or XXXXX-XXXX
'*
'* Written by: Nick DeSteffen
'* Written on: 11-02-2002
'**************************************************************

Function Check_Zip(strZip As String) As Boolean
    Dim intMsgBox As Integer        'Used for messageboxes
    
    On Error GoTo ErrorHandler
    Check_Zip = True
    
    Select Case Len(strZip)
        Case 5
            If Not IsNumeric(strZip) Then Check_Zip = False
        Case 10
        If Not IsNumeric(Left$(strZip, 5)) Or Not IsNumeric(Right$(strZip, 4)) Or Mid$(strZip, 6, 1) <> "-" Then Check_Zip = False
        Case Else
            Check_Zip = False
    End Select
Exit Function
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End

End Function

'**************************************************************
'* Checks to make sure that the phone number is valid or not.
'* Accepts Phone numbers in the following format XXXXXXXXXX or
'* XXX-XXX-XXXX
'*
'* Written by: Nick DeSteffen
'* Written on: 11-05-2002
'**************************************************************

Function Check_Phone(strPhone As String) As Boolean
    Dim strLeft As String       'Left digits
    Dim strMid As String        'Middle digits
    Dim strRight As String      'Right digits
    Dim strDash1 As String      'First Dash
    Dim strDash2 As String      '2nd Dash
    Dim intMsgBox As Integer    'Used for messageboxes
        
    On Error GoTo ErrorHandler
    Check_Phone = True
    Select Case Len(strPhone)
        Case 10
            If Not IsNumeric(strPhone) Then Check_Phone = False
        Case 12
            strLeft = Left$(strPhone, 3)
            strMid = Mid$(strPhone, 5, 3)
            strRight = Right$(strPhone, 4)
            strDash1 = Mid$(strPhone, 4, 1)
            strDash2 = Mid$(strPhone, 8, 1)
            If Not IsNumeric(strLeft) Or Not IsNumeric(strRight) Or Not IsNumeric(strMid) Or strDash1 <> "-" Or strDash2 <> "-" Then Check_Phone = False
        Case Else
            Check_Phone = False
    End Select
    Exit Function
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End

End Function
