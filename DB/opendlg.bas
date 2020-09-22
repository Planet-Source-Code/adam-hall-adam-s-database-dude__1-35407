Attribute VB_Name = "MCommonDialog"
'/* I think this one was by Bruce McKinney of the book Hardcore VB */
'   You can find it online at http://www.mvps.org/vb/hardcore/
Option Explicit
Private Const cMaxPath = 255, cMaxFile = 255

Public Enum EErrorCommonDialog
    eeBaseCommonDialog = 13450  ' CommonDialog
End Enum

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hwndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    Flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" Alias "GetFileTitleA" (ByVal szFile As String, ByVal szTitle As String, ByVal cbBuf As Long) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

' Common dialog errors
Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF

    CDERR_GENERALCODES = &H0
    CDERR_STRUCTSIZE = &H1
    CDERR_INITIALIZATION = &H2
    CDERR_NOTEMPLATE = &H3
    CDERR_NOHINSTANCE = &H4
    CDERR_LOADSTRFAILURE = &H5
    CDERR_FINDRESFAILURE = &H6
    CDERR_LOADRESFAILURE = &H7
    CDERR_LOCKRESFAILURE = &H8
    CDERR_MEMALLOCFAILURE = &H9
    CDERR_MEMLOCKFAILURE = &HA
    CDERR_NOHOOK = &HB
    CDERR_REGISTERMSGFAIL = &HC

    PDERR_PRINTERCODES = &H1000
    PDERR_SETUPFAILURE = &H1001
    PDERR_PARSEFAILURE = &H1002
    PDERR_RETDEFFAILURE = &H1003
    PDERR_LOADDRVFAILURE = &H1004
    PDERR_GETDEVMODEFAIL = &H1005
    PDERR_INITFAILURE = &H1006
    PDERR_NODEVICES = &H1007
    PDERR_NODEFAULTPRN = &H1008
    PDERR_DNDMMISMATCH = &H1009
    PDERR_CREATEICFAILURE = &H100A
    PDERR_PRINTERNOTFOUND = &H100B
    PDERR_DEFAULTDIFFERENT = &H100C

    CFERR_CHOOSEFONTCODES = &H2000
    CFERR_NOFONTS = &H2001
    CFERR_MAXLESSTHANMIN = &H2002

    FNERR_FILENAMECODES = &H3000
    FNERR_SUBCLASSFAILURE = &H3001
    FNERR_INVALIDFILENAME = &H3002
    FNERR_BUFFERTOOSMALL = &H3003

    CCERR_CHOOSECOLORCODES = &H5000
End Enum

Function VBGetOpenFileName(Filename As String, _
                           Optional FileTitle As String, _
                           Optional FileMustExist As Boolean = True, _
                           Optional MultiSelect As Boolean = False, _
                           Optional ReadOnly As Boolean = False, _
                           Optional HideReadOnly As Boolean = False, _
                           Optional filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional Flags As Long = 0) As Boolean

    Dim opfile As OPENFILENAME, s As String, afFlags As Long
With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    .Flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
             (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-ReadOnly * OFN_READONLY) Or _
             (-HideReadOnly * OFN_HIDEREADONLY) Or _
             (Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hwndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' To make Windows-style filter, replace | and : with nulls
    Dim ch As String, i As Integer
    For i = 1 To Len(filter)
        ch = Mid(filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = Filename & String(cMaxPath - Len(Filename), 0)
    .lpstrFile = s
    .nMaxFile = cMaxPath
    s = FileTitle & String(cMaxFile - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = cMaxFile
    ' All other fields set to zero
    
    If GetOpenFileName(opfile) Then
        VBGetOpenFileName = True
        
        Filename = .lpstrFile
        FileTitle = .lpstrFileTitle
        If Not MultiSelect Then
            Filename = TrimNull(Filename)
            FileTitle = TrimNull(FileTitle)
        End If
        Flags = .Flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        filter = FilterLookup(.lpstrFilter, FilterIndex)
        If (.Flags And OFN_READONLY) Then ReadOnly = True
    Else
        VBGetOpenFileName = False
        Filename = vbNullString
        FileTitle = vbNullString
        Flags = 0
        FilterIndex = -1
        filter = vbNullString
    End If
End With
End Function

Function VBGetSaveFileName(Filename As String, _
                           Optional FileTitle As String, _
                           Optional OverWritePrompt As Boolean = True, _
                           Optional filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional Flags As Long) As Boolean
            
    Dim opfile As OPENFILENAME, s As String
With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    .Flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
             OFN_HIDEREADONLY Or _
             (Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hwndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' Make new filter with bars (|) replacing nulls and double null at end
    Dim ch As String, i As Integer
    For i = 1 To Len(filter)
        ch = Mid(filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = Filename & String(cMaxPath - Len(Filename), 0)
    .lpstrFile = s
    .nMaxFile = cMaxPath
    s = FileTitle & String(cMaxFile - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = cMaxFile
    ' All other fields zero
    
    If GetSaveFileName(opfile) Then
        VBGetSaveFileName = True
        Filename = .lpstrFile
        FileTitle = .lpstrFileTitle
        Flags = .Flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        filter = FilterLookup(.lpstrFilter, FilterIndex)
    Else
        VBGetSaveFileName = False
        Filename = vbNullString
        FileTitle = vbNullString
        Flags = 0
        FilterIndex = 0
        filter = vbNullString
    End If
End With
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long, iEnd As Long, s As String
    iStart = 1
    If sFilters = vbNullString Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function

Function VBGetFileTitle(ByVal sFile As String) As String
    Dim sFileTitle As String, cFileTitle As Integer

    cFileTitle = cMaxPath
    sFileTitle = String(cMaxPath, 0)
    cFileTitle = GetFileTitle(sFile, sFileTitle, cMaxPath)
    If cFileTitle Then
        VBGetFileTitle = vbNullString
    Else
        VBGetFileTitle = Left(sFileTitle, InStr(sFileTitle, vbNullChar) - 1)
    End If
End Function

Private Sub ErrRaise(ByVal e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".CommonDialog"
    Else
        ' Raise standard Visual Basic error
        sSource = App.EXEName
    End If
    Err.Raise e, sSource
End Sub

Public Function ExtractFiles(ByVal FileString As String, FileTable() As String) As Integer
    Dim nEntries As Integer
    Dim nNull As Integer
    Do While Asc(FileString)
        ReDim Preserve FileTable(0 To nEntries)
        nNull = InStr(FileString, vbNullChar)
        FileTable(nEntries) = Left(FileString, nNull - 1)
        FileString = Mid(FileString, nNull + 1)
        nEntries = nEntries + 1
    Loop
    ExtractFiles = nEntries - 1
End Function

Public Function GetPathOnly(ByVal MyFile As String) As String
    Dim X As Integer
    For X = Len(MyFile) To 1 Step -1
        If Mid(MyFile, X, 1) = "\" Then GoTo OK
    Next X
    GetPathOnly = MyFile
    Exit Function
OK:
    GetPathOnly = Left(MyFile, X - 1)
End Function

Public Function GetNameOnly(ByVal fullname As String, ByVal nBaseOnly As Boolean) As String
    Dim nCnt As Integer, sSpecOut As String, nDot As Integer
    On Local Error Resume Next
    
    If InStr(fullname, "\") Then
        For nCnt = Len(fullname) To 1 Step -1
            If Mid(fullname, nCnt, 1) = "\" Then
                sSpecOut = Mid(fullname, nCnt + 1)
                Exit For
            End If
        Next nCnt
    
    ElseIf InStr(fullname, ":") = 2 Then
        sSpecOut = Mid(fullname, 3)
      
    Else
        sSpecOut = fullname
    End If
      
    If nBaseOnly Then
        nDot = InStr(sSpecOut, ".")
        If nDot Then
            sSpecOut = Left(sSpecOut, nDot - 1)
        End If
    End If
    
    GetNameOnly = sSpecOut
End Function

Public Function TrimNull(sString As String) As String
    TrimNull = Left(sString, InStr(1, sString, vbNullChar) - 1)
End Function
