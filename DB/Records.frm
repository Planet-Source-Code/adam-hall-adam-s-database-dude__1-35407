VERSION 5.00
Begin VB.Form frmRecords 
   Caption         =   "Records"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNav 
      Height          =   345
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   2310
      Width           =   500
   End
   Begin VB.TextBox txtFieldValue 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Text            =   "Field Value"
      Top             =   150
      Width           =   2415
   End
   Begin VB.TextBox txtFieldName 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Field Name:"
      Top             =   150
      Width           =   1245
   End
   Begin VB.Label lblRecord 
      Caption         =   "Record X of Y"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   2790
      Width           =   4335
   End
End
Attribute VB_Name = "frmRecords"
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
Private iPage As Integer

Private Sub InitNav()
    Dim X As Integer

    '// for each unloaded button
    For X = 1 To 4
        '// load the button
        Load cmdNav(X)
        '// make it visible
        cmdNav(X).Visible = True
        '// move it
        cmdNav(X).Move cmdNav(X - 1).Left + cmdNav(X - 1).Width
    Next X
    '// setup the captions
    cmdNav(0).Caption = "1st"
    cmdNav(1).Caption = "<"
    cmdNav(2).Caption = ">"
    cmdNav(3).Caption = "last"
    cmdNav(4).Caption = "add"
End Sub

Private Sub cmdNav_Click(Index As Integer)
    If Index = 4 Then
        frmAdd.Show 1
        Exit Sub
    End If
    Dim iPageCount As Integer, X As Integer, fld
      
    '// we only show one item per page
    objRS.PageSize = 1
    objRS.CacheSize = 1
    '// get the number of pages (or records in this case)
    iPageCount = objRS.PageCount
    
    '// move the page the user requested
    Select Case Index
        Case 0
            iPage = 1
        Case 1
            iPage = iPage - 1
        Case 2
            iPage = iPage + 1
        Case 3
            iPage = iPageCount
    End Select
    
    If iPage < 1 Then iPage = 1
    If objRS.RecordCount > 0 Then objRS.AbsolutePage = iPage
    

    '// enable the buttons that need it
    cmdNav(0).Enabled = (iPage <> 1)
    cmdNav(1).Enabled = (iPage > 1)
    cmdNav(2).Enabled = (iPage < iPageCount)
    cmdNav(3).Enabled = (iPage <> iPageCount)

    '// update the record info
    lblRecord.Caption = "Page (record) " & iPage & " of " & iPageCount
    
    '// for all the fields
    For Each fld In objRS.Fields
        '// if there are no records
        If objRS.EOF Then
            '// prevent errors and alert user
            txtFieldValue(X).Text = "<EOF>"
        Else
            '// prevent Null errors
            On Error Resume Next
            '// update the value
            txtFieldValue(X).Text = fld.Value
            '// select it
            txtFieldValue(X).SelStart = 0
            txtFieldValue(X).SelLength = Len(txtFieldValue(X).Text)
        End If
        
        '// goto the next field
        X = X + 1
    Next fld
End Sub

Private Sub Form_Load()
    Dim fld, X As Integer
    
    '// build navigation bar
    InitNav
    '// open the connection to the database
    objRS.CursorLocation = adUseServer
    objRS.Open "SELECT * FROM [" & sTable & "]", objConn, adOpenStatic
    
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
    Me.Height = txtFieldValue(X - 1).Top + txtFieldValue(X - 1).Height + lblRecord.Height + cmdNav(0).Height + 500
    
    '// open the first record
    cmdNav_Click 0
End Sub

Private Sub Form_Resize()
    Dim X As Integer
    '// move the record info just below the edge of the form
    lblRecord.Move lblRecord.Left, Me.ScaleHeight - lblRecord.Height
    '// for each nav button
    For X = 0 To 4
        '// move the navigation buttons just above the record info
        cmdNav(X).Move cmdNav(X).Left, lblRecord.Top - cmdNav(X).Height
    Next X
End Sub
