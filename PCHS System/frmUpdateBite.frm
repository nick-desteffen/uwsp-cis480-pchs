VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUpdateBite 
   Caption         =   "Update Bite Report"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4140
   Icon            =   "frmUpdateBite.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtPreparedBy 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtEuthanizeBy 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtpLabDate 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   47251457
      CurrentDate     =   37582
   End
   Begin MSComCtl2.DTPicker dtpEuthanizeDate 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   47251457
      CurrentDate     =   37582
   End
   Begin VB.Label lblLabDate 
      Caption         =   "Date sent to State Lab:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblSpecimen 
      Caption         =   "Specimen prepared by:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblEuthanizedBy 
      Caption         =   "Euthanized by:"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblDateEuthanized 
      Caption         =   "Date died or euthanized:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmUpdateBite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************
'* This form is used to update the information on a bite report.
'* If the animal is later euthanized the fields in the database can
'* be updated via this form.
'*
'* Written by: Nick DeSteffen
'* Written on: 12-04-2002
'***********************************************************************

Dim intBiteNum As Integer           'Bite number
Private Sub cmdUpdate_Click()

Dim strEuthanizePerson As String    'Person performing the euthanization
Dim strSpecimen As String           'Person handling the speciman
Dim intMsgBox As Integer            'Used for messageboxes
Dim strSQL As String                'SQL Statement

On Error GoTo ErrorHandler

strEuthanizePerson = txtEuthanizeBy.Text
strSpecimen = txtPreparedBy.Text

strSQL = "UPDATE BITE SET BITE_EUTHANIZER = '" & strEuthanizePerson
strSQL = strSQL & "', BITE_PREPARED_BY = '" & strSpecimen
strSQL = strSQL & "', BITE_EUTHANIZED_DATE = " & Check_Date(dtpEuthanizeDate.Value)
strSQL = strSQL & ", BITE_LAB_DATE = " & Check_Date(dtpLabDate.Value)
strSQL = strSQL & " WHERE BITE_NUMBER = " & intBiteNum

intMsgBox = MsgBox("Update bite report?", vbYesNo, "Update?")
If intMsgBox = 6 Then
    Open_Recordsets.objConnection.Execute (strSQL)
ElseIf intMsgBox = 7 Then
    Exit Sub
End If
Unload Me
Exit Sub

ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

Private Sub Form_Load()

Dim rstBite As ADODB.Recordset          'Recordset used for interfacing with the database
Dim strSQL As String                    'SQL Statement
Dim intMsgBox As Integer                'Used for messageboxes
Dim strBiteNum As String                'Number entered from inputbox

'On Error GoTo ErrorHandler

Set rstBite = New ADODB.Recordset
strBiteNum = InputBox("What is the bite report number?", "Bite Report Number")
intBiteNum = CInt(strBiteNum)

strSQL = "SELECT BITE_EUTHANIZER, BITE_PREPARED_BY, BITE_EUTHANIZED_DATE, BITE_LAB_DATE FROM BITE WHERE BITE_NUMBER = " & intBiteNum

Set rstBite = Open_Recordsets.objConnection.Execute(strSQL)

If rstBite.EOF = False Then
    With rstBite
        rstBite.MoveFirst
        Do While Not rstBite.EOF
            If Not IsNull(![BITE_LAB_DATE]) Then
                dtpLabDate.Value = ![BITE_LAB_DATE]
            Else
                dtpLabDate.Value = Date
            End If
            If Not IsNull(![BITE_EUTHANIZED_DATE]) Then
                dtpEuthanizeDate.Value = ![BITE_EUTHANIZED_DATE]
            Else
                dtpEuthanizeDate.Value = Date
            End If
            txtEuthanizeBy.Text = ![BITE_EUTHANIZER]
            txtPreparedBy.Text = ![BITE_PREPARED_BY]
            rstBite.MoveNext
        Loop
    End With
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub
Function Check_Date(dteCheck) As String

If IsNull(dteCheck) = True Then
    Check_Date = "Null"
Else
    Check_Date = "'" & dteCheck & "'"
End If

End Function

Private Sub cmdCancel_Click()
Unload Me
frmPCHS_Main.Show
End Sub
