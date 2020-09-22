VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRec32 
   Caption         =   "Records"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   ScaleHeight     =   3645
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdButton 
      Caption         =   "Edit"
      Height          =   345
      Index           =   1
      Left            =   1260
      TabIndex        =   2
      Top             =   3180
      Width           =   1125
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   345
      Index           =   3
      Left            =   3600
      TabIndex        =   4
      Top             =   3180
      Width           =   1125
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Remove"
      Height          =   345
      Index           =   2
      Left            =   2430
      TabIndex        =   3
      Top             =   3180
      Width           =   1125
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Add"
      Height          =   345
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   3180
      Width           =   1125
   End
   Begin MSComctlLib.ListView lstRecords 
      Height          =   3045
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   5371
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmRec32"
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

Private Sub cmdButton_Click(Index As Integer)
    Select Case Index
        Case 0
            '// show the add form
            frmAdd.Show 1
        Case 1
            '// pass this var onto the other form
            sEditWhere = lstRecords.ColumnHeaders(1).Text & "=" & lstRecords.SelectedItem.Text
            '// show the edit form
            frmEdit.Show 1
        Case 2
            '// delete the record
            objConn.Execute "DELETE FROM [" & sTable & "] WHERE " & lstRecords.ColumnHeaders(1).Text & "=" & lstRecords.SelectedItem.Text
        Case 3
            '// unload the form
            Unload Me
            Exit Sub
    End Select
    
    '// close the recordset
    objRS.Close
    '// refresh the list
    Form_Load
End Sub

Private Sub Form_Load()
    Dim fld, X As Integer
    
    '// clear the list box
    lstRecords.ColumnHeaders.Clear
    lstRecords.ListItems.Clear
    
    '// open the connection to the database
    objRS.CursorLocation = adUseServer
    objRS.Open "SELECT * FROM [" & sTable & "]", objConn, adOpenStatic
    
    '// list all the fields
    For Each fld In objRS.Fields
        '// add the header
        lstRecords.ColumnHeaders.Add , , fld.Name
    Next fld
    
    Dim lstItem As ListItem
    
    '// for each record
    Do Until objRS.EOF
        '// add the item to the list
        Set lstItem = lstRecords.ListItems.Add(, , objRS.Fields(0).Value)
        '// reset the field
        X = 0
        '// for each field
        For Each fld In objRS.Fields
            '// if it is a proper header
            If X > 0 Then
                '// ignore Null errors
                On Error Resume Next
                '// update the value
                lstItem.SubItems(X) = fld.Value
            End If
            '// goto the next field
            X = X + 1
        Next fld
        
        '// goto the next record
        objRS.MoveNext
    Loop
End Sub

Private Sub Form_Resize()
    Dim X As Integer
    '// move the records list box to occupy most of the form
    lstRecords.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - (cmdButton(0).Height + 15)
    '// for each button
    For X = cmdButton.UBound To cmdButton.lBound Step -1
        '// move it just above the edge of the form
        If X = cmdButton.UBound Then
            cmdButton(X).Move Me.ScaleWidth - cmdButton(X).Width, Me.ScaleHeight - cmdButton(X).Height
        Else
            cmdButton(X).Move cmdButton(X + 1).Left - cmdButton(X).Width, Me.ScaleHeight - cmdButton(X).Height
        End If
    Next X
End Sub
