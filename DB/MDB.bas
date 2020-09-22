Attribute VB_Name = "MDB"
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

Public sDBName As String
Public sTable As String
Public sEditWhere As String
Public objConn As ADODB.Connection
Public objRS As ADODB.Recordset

'/* gets the name of a database to open etc. */
Public Sub Main()
    '// get the filename of the database from the user
    If VBGetOpenFileName(sDBName, , True, False, False, True, "Microsoft Access Databases (*.mdb)|*.mdb|All Files (*.*)|*.*", , , "Open Database", "mdb", , OFN_EXPLORER) Then
        '// if they oked the box then
        '// open the database
        DBConnection_Open sDBName
        '// get the table to use
        frmTables.Show 1
        '// if they chose a table
        If sTable <> "" Then
            '// show the records in it
            frmRec32.Show 1
        End If
        '// close the connection
        DBConnection_Close
    End If
    '// end of program
End Sub

'/* creates a new instance of the object and connects to the database */
Public Sub DBConnection_Open(sDBName As String)
    Set objConn = New ADODB.Connection
    '// apparently the Jet engine is much faster than the Access driver.
    objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & sDBName & ";"  '& "Jet OLEDB:Database Password=" & sPassword & ";"
End Sub

'/* closes the connection to the database and nullifies the object */
Public Sub DBConnection_Close()
    objConn.Close
    Set objConn = Nothing
End Sub

'/* if the sPath does not end in a slash ("\") then one is added */
Private Function AddASlash(sPath As String) As String
    AddASlash = sPath
    If Right(AddASlash, 1) <> "\" Then
        AddASlash = AddASlash & "\"
    End If
End Function

'/* creates the table sTablename in the opened db */
Public Sub DBTable_Create(sTablename As String)
    objConn.Execute "CREATE TABLE [" & sTablename & "]"
End Sub

'/* deletes the table and all fields in it. must be empty i think */
Public Sub DBTable_Remove(sTablename As String)
    objConn.Execute "DROP TABLE [" & sTablename & "]"
End Sub

'/* removes the field, and also the primarykey if it is indexed */
Public Sub DBField_Remove(sTablename As String, sFieldname As String)
    Dim sIndex As String
    sIndex = DBIndex_IsIndexed(sTablename, sFieldname)
    If sIndex <> vbNullString Then
        DBIndex_Remove sTablename, sIndex
    End If
    objConn.Execute "ALTER TABLE [" & sTablename & "] DROP COLUMN [" & sFieldname & "]"
End Sub

'/* removes the index from the table */
Public Sub DBIndex_Remove(sTablename As String, sIndex As String)
    objConn.Execute "ALTER TABLE [" & sTablename & "] DROP CONSTRAINT [" & sIndex & "]"
End Sub

'/* checks if the field is indexed */
Public Function DBIndex_IsIndexed(sTablename As String, sFieldname As String) As String
    DBIndex_IsIndexed = vbNullString
    Set objRS = objConn.OpenSchema(adSchemaIndexes)
    Do While Not objRS.EOF
        If objRS("table_name") = sTablename Then
            If (objRS("column_name") = sFieldname) Then
                DBIndex_IsIndexed = objRS("index_name")
            End If
        End If
        objRS.MoveNext
    Loop
    DBRecordSet_Close
End Function

'/* closes the recordset. does not nullify */
Public Sub DBRecordSet_Close()
    objRS.Close
End Sub

'/* invokes a search on the field in the table */
Public Sub DBRecordSet_Search(sTablename As String, sFieldname As String, sText As String, Optional bWholeWord As Boolean = False)
    Dim sQuery As String
    sQuery = sText
    If bWholeWord Then
        sQuery = " " & sQuery & " "
    End If
    sQuery = "%" & sQuery & "%"
    Set objRS = objConn.Execute("SELECT * FROM [" & sTablename & "] WHERE " & sFieldname & " LIKE '" & sQuery & "'")
End Sub

'/* returns all the tables in the db as a collection */
Public Sub DBTables_List(cTables As Collection)
    Set cTables = New Collection
    Set objRS = objConn.OpenSchema(adSchemaTables)
    
    Do While Not objRS.EOF
        If objRS("table_type") = "TABLE" Then
            cTables.Add objRS.Fields("table_name").Value
        End If
        objRS.MoveNext
    Loop
    objRS.Close
End Sub


