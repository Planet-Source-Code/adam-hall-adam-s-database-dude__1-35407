VERSION 5.00
Begin VB.Form frmEdit 
   Caption         =   "Edit Record"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2423
      TabIndex        =   3
      Top             =   2730
      Width           =   1125
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Default         =   -1  'True
      Height          =   345
      Left            =   1133
      TabIndex        =   2
      Top             =   2730
      Width           =   1125
   End
   Begin VB.TextBox txtFieldName 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Field Name:"
      Top             =   90
      Width           =   1245
   End
   Begin VB.TextBox txtFieldValue 
      Height          =   375
      Index           =   0
      Left            =   1980
      TabIndex        =   0
      Text            =   "Field Value"
      Top             =   90
      Width           =   2415
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/***************************** '
' *         DBADaMIN          * '
' *       by Adam Hall        * '
' *    psc@ahall.cjb.net      * '
' ***************************** '
' * You can freely use any of * '
' * the code here as long as  * '
' * you give me credit.       * '
' *****************************/'
Option Explicit

Private Sub cmdEdit_Click()
    Dim X As Integer
    
    '// for each field
    For X = txtFieldName.lBound To txtFieldName.UBound
        '// set its value, if its changed
        If txtFieldValue(X).Text <> txtFieldValue(X).Tag Then
            objRS.Fields(txtFieldName(X).Text).Value = txtFieldValue(X).Text
        End If
    Next X
    '// save it
    objRS.Update
    '// close the recordset
    objRS.Close
    
    '// open the existing recordset
    objRS.Open "SELECT * FROM [" & sTable & "]", objConn, adOpenStatic

    '// unload this form
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    '// unload form without making any changes
    Unload Me
End Sub

Private Sub Form_Load()
    Dim fld, X As Integer
       
    '// close the recordset
    objRS.Close
    '// open a new one
    objRS.Open "SELECT * FROM [" & sTable & "] WHERE " & sEditWhere, objConn, adOpenStatic, adLockPessimistic

    '// list all the fields
    For Each fld In objRS.Fields
        '// if the objects are not loaded
        If X > 0 Then
            '// load them
            Load txtFieldName(X)
            '// make them visible
            txtFieldName(X).Visible = True
            '// move them
            txtFieldName(X).Move txtFieldName(X - 1).Left, txtFieldName(X - 1).Height + txtFieldName(X - 1).Top
            
            '// load them
            Load txtFieldValue(X)
            '// make them visible
            txtFieldValue(X).Visible = True
            '// move them
            txtFieldValue(X).Move txtFieldValue(X - 1).Left, txtFieldValue(X - 1).Height + txtFieldValue(X - 1).Top
        End If
        '// setup the field name
        txtFieldName(X).Text = fld.Name
        '// select it
        txtFieldName(X).SelStart = 0
        txtFieldName(X).SelLength = Len(txtFieldName(X).Text)
        '// change the taborder
        txtFieldName(X).TabIndex = X * 2
        '// update the field value
        txtFieldValue(X).Text = fld.Value
        txtFieldValue(X).Tag = fld.Value
        '// change the taborder
        txtFieldValue(X).TabIndex = X * 2 + 1
        
        '// goto the next field
        X = X + 1
    Next fld
    
    '// resize the form to fit all the fields
    Me.Height = txtFieldValue(X - 1).Top + txtFieldValue(X - 1).Height + 500 + cmdEdit.Height
    '// move the buttons
    cmdEdit.Top = Me.ScaleHeight - cmdEdit.Height
    cmdCancel.Top = cmdEdit.Top
    

End Sub
