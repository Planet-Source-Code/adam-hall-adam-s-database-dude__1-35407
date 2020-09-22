VERSION 5.00
Begin VB.Form frmTables 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tables"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTables 
      Height          =   2985
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   4485
   End
End
Attribute VB_Name = "frmTables"
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

Private Sub Form_Load()
    Dim cTables As Collection, X As Integer
    
    DBTables_List cTables
    
    For X = 1 To cTables.Count
        lstTables.AddItem cTables.Item(X)
    Next X
End Sub

Private Sub Form_Resize()
    lstTables.Move 30, 30, Me.ScaleWidth - 60, Me.ScaleHeight - 60
End Sub

Private Sub lstTables_DblClick()
    sTable = lstTables.List(lstTables.ListIndex)
    Unload Me
End Sub
