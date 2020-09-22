VERSION 5.00
Begin VB.Form frmAdd 
   Caption         =   "Add Record"
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
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
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
Attribute VB_Name = "frmAdd"
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

Private Sub cmdAdd_Click()
    Dim X As Integer
    
    '// free the recordset
    objRS.Close
    '// open the table
    objRS.Open "[" & sTable & "]", objConn, adOpenStatic, adLockPessimistic, adCmdTable
    '// add a new record
    objRS.AddNew
    '// for each field
    For X = txtFieldName.lBound To txtFieldName.UBound
        '// set its value
        If txtFieldValue(X).Text <> vbNullString Then
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
        '// empty the field value
        txtFieldValue(X).Text = ""
        '// change the taborder
        txtFieldValue(X).TabIndex = X * 2 + 1
        
        '// goto the next field
        X = X + 1
    Next fld
    
    '// resize the form to fit all the fields
    Me.Height = txtFieldValue(X - 1).Top + txtFieldValue(X - 1).Height + 500 + cmdAdd.Height
    '// move the buttons
    cmdAdd.Top = Me.ScaleHeight - cmdAdd.Height
    cmdCancel.Top = cmdAdd.Top
    
End Sub
