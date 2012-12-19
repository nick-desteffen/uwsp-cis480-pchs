VERSION 5.00
Begin VB.Form frmSearchLicense 
   Caption         =   "Search License Numbers"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3975
   Icon            =   "frmSearchLicense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtLicense 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmSearchLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'************************************************************************
'* This form is used when a stray animal is brought in with a license
'* tag.  The user enters the license number in the box and click's search
'* if the license is found, the license report is displayed.
'*
'* Written by: Nick DeSteffen
'* Written on: 12-01-2002
'************************************************************************

Private Sub cmdSearch_Click()
Dim intLicense As Integer           'License number
Dim intMsgBox As Integer            'Used for messageboxes
Dim rstSearch As ADODB.Recordset    'Recordset used for interfacing with the database

On Error GoTo ErrorHandler

Set rstSearch = New ADODB.Recordset

If IsNumeric(txtLicense.Text) = True And txtLicense.Text <> "" Then
    intLicense = txtLicense.Text
Else
    intMsgBox = MsgBox("Please enter a valid license number!", vbOKOnly, "Invalid Number")
    Exit Sub
End If

Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT * FROM LICENSE WHERE LICENSE_NUMBER = " & intLicense)

If rstSearch.EOF = False Then
    With rstSearch
        rstSearch.MoveFirst
        Do While Not rstSearch.EOF
            If Not IsNull(![LICENSE_NUMBER]) Then
                frmShowLicense.intLicense = intLicense
                frmShowLicense.Show
            End If
            rstSearch.MoveNext
        Loop
    End With
Else
    intMsgBox = MsgBox("No matches found.", vbOKOnly, "None found")
    Unload Me
End If
Unload Me

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub
