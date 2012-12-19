VERSION 5.00
Begin VB.Form frmNewReciept 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Reciept"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   ControlBox      =   0   'False
   Icon            =   "frmNewReciept.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save and Print Reciept"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtCheckNum 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox chkTax 
      Caption         =   "Tax"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtComments 
      Height          =   525
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4695
   End
   Begin VB.ComboBox cboReason 
      Height          =   315
      ItemData        =   "frmNewReciept.frx":08CA
      Left            =   360
      List            =   "frmNewReciept.frx":08DD
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblComments 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblCheckNumj 
      Caption         =   "Check Number"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblAmount 
      Caption         =   "Transaction Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewReciept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'******************************************************************************
'* This form is displayed each time there is a transaction and a receipt is
'* needed to be generated.  The original form passes a number to determine
'* which table to update the reciept information to and which type of transaction
'* is taking place.
'*
'* Written by: Nick DeSteffen
'* Written on: 11-25-2002
'******************************************************************************

Public intPersonNum As Integer          'Person who is issued the receipt
Public intType As Integer               'Form type for the transaction
Public intNumber As Integer             'Number of the service
Public intReceiptNum As Integer         'Number of the receipt

Private Sub cmdSave_Click()
Dim dblAmount As Double         'Amount of the transaction
Dim dblTax As Double            'Tax for the transaction
Dim intCheck As Integer         'Check number
Dim strComments As String       'Comments on the check
Dim strReason As String         'Reason of the transaction

Dim intMsgBox As Integer            'Used for messageboxes
Dim strSQL As String                'SQL statement
Dim rstReceipt As ADODB.Recordset   'Recordset used for interfacing with the database

On Error GoTo ErrorHandler

Set rstReceipt = New ADODB.Recordset

'Checks to see if all required fields are valid

If IsNumeric(txtAmount.Text) = True Then
    dblAmount = txtAmount.Text
Else
    intMsgBox = MsgBox("Please enter a valid amount for the transaction.", vbOKOnly, "Error!")
    Exit Sub
End If

If (IsNumeric(txtCheckNum.Text) = True Or txtCheckNum.Text = "") Then
    If txtCheckNum.Text = "" Then
        intCheck = 0
    Else
        intCheck = txtCheckNum.Text
    End If
Else
    intMsgBox = MsgBox("Please inter a valid check number.", vbOKOnly, "Error!")
    Exit Sub
End If

'Calculates the tax

If chkTax.Value = 1 Then
    dblTax = dblAmount * 0.055
End If

strReason = cboReason.Text
strComments = Replace(txtComments.Text, "'", "''")

Set rstReceipt = objConnection.Execute("SELECT MAX(RECEIPT_NUMBER) AS RECEIPT_NUM FROM RECEIPT")

With rstReceipt
    If IsNull(![RECEIPT_NUM]) = True Then
        intReceiptNum = 1
    Else
        intReceiptNum = ![RECEIPT_NUM]
        intReceiptNum = intReceiptNum + 1
    End If
End With

'Inserts new reciept into the receipt table

strSQL = "INSERT INTO RECEIPT (RECEIPT_NUMBER, RECEIPT_PERSON, RECEIPT_TOTAL, RECEIPT_REASON, RECEIPT_CHECK_NUM, "
strSQL = strSQL & "RECEIPT_TAX, RECEIPT_AMOUNT, RECEIPT_COMMENTS, RECEIPT_SERVICE_NUM) VALUES (" & intReceiptNum & ", "
strSQL = strSQL & intPersonNum & ", " & dblAmount + dblTax & ", '" & strReason & "', " & intCheck & ", "
strSQL = strSQL & dblTax & ", " & dblAmount & ", '" & strComments & "', " & intNumber & ")"

Open_Recordsets.objConnection.Execute (strSQL)

'Inserts the reciept number into the appropriate table

Select Case intType
    Case 1: Open_Recordsets.objConnection.Execute ("UPDATE ADOPTION SET ADOPTION_RECEIPT = " & intReceiptNum & " WHERE ADOPTION_NUMBER = " & intNumber)
    Case 2: Open_Recordsets.objConnection.Execute ("UPDATE DONATION SET DONATION_RECEIPT = " & intReceiptNum & ", DONATION_AMOUNT = " & dblAmount & " WHERE DONATION_NUMBER = " & intNumber)
    Case 3: Open_Recordsets.objConnection.Execute ("UPDATE LICENSE SET LICENSE_RECEIPT = " & intReceiptNum & " WHERE LICENSE_NUMBER = " & intNumber)
    Case 4: Open_Recordsets.objConnection.Execute ("UPDATE ANIMALS SET ANIMAL_STATUS = 'C' WHERE ANIMAL_NUMBER = " & intNumber)
    Case 5: Open_Recordsets.objConnection.Execute ("SELECT * FROM PERSON")
End Select

Set rstReceipt = Nothing

'Displays the generated receipt

frmShowReciept.intReceiptNum = intReceiptNum
frmShowReciept.Show

frmPCHS_Main.Enabled = True
Unload Me

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

Private Sub Form_Load()
frmPCHS_Main.Enabled = False
End Sub
