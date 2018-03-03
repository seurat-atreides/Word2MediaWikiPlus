Attribute VB_Name = "modWord2MediaWikiPlusGlobal"
' Word2MediaWikiPlus
' Converts Microsoft Word documents to MediaWiki.
'
' Copyright 2006, 2007 Gunter Schmidt.
'
' Website: http://www.mediawiki.org/wiki/Extension:Word2MediaWikiPlus
' Project site: http://sourceforge.net/projects/word2mediawikip
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Public Declare Function GetUserDefaultLangID Lib "kernel32.dll" () As Long

Private Declare Function GetFileAttributesA Lib "kernel32" (ByVal lpFileName As String) As Long
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10&
Private Const FILE_ATTRIBUTE_INVALID   As Long = -1&  ' = &HFFFFFFFF&

'Used for Color-Conversion
Public Declare Function OleTranslateColor Lib "oleaut32.dll" _
    (ByVal lOleColor As Long, ByVal lHPalette As Long, _
    ByRef lColorRef As Long) As Long

'Special folder location
Private Const CSIDL_DESKTOP = &H0                 '{desktop}
Private Const CSIDL_INTERNET = &H1                'Internet Explorer (icon on desktop)
Private Const CSIDL_PROGRAMS = &H2                'Start Menu\Programs
Private Const CSIDL_CONTROLS = &H3                'My Computer\Control Panel
Private Const CSIDL_PRINTERS = &H4                'My Computer\Printers
Private Const CSIDL_PERSONAL = &H5                'My Documents
Private Const CSIDL_FAVORITES = &H6               '{user}\Favourites
Private Const CSIDL_STARTUP = &H7                 'Start Menu\Programs\Startup
Private Const CSIDL_RECENT = &H8                  '{user}\Recent
Private Const CSIDL_SENDTO = &H9                  '{user}\SendTo
Private Const CSIDL_BITBUCKET = &HA               '{desktop}\Recycle Bin
Private Const CSIDL_STARTMENU = &HB               '{user}\Start Menu
Private Const CSIDL_DESKTOPDIRECTORY = &H10       '{user}\Desktop
Private Const CSIDL_DRIVES = &H11                 'My Computer
Private Const CSIDL_NETWORK = &H12                'Network Neighbourhood
Private Const CSIDL_NETHOOD = &H13                '{user}\nethood
Private Const CSIDL_FONTS = &H14                  'windows\fonts
Private Const CSIDL_TEMPLATES = &H15
Private Const CSIDL_COMMON_STARTMENU = &H16       'All Users\Start Menu
Private Const CSIDL_COMMON_PROGRAMS = &H17        'All Users\Programs
Private Const CSIDL_COMMON_STARTUP = &H18         'All Users\Startup
Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19 'All Users\Desktop
Private Const CSIDL_APPDATA = &H1A                '{user}\Application Data
Private Const CSIDL_PRINTHOOD = &H1B              '{user}\PrintHood
Private Const CSIDL_LOCAL_APPDATA = &H1C          '{user}\Local Settings _
                                                 '\Application Data (non roaming)
Private Const CSIDL_ALTSTARTUP = &H1D             'non localized startup
Private Const CSIDL_COMMON_ALTSTARTUP = &H1E      'non localized common startup
Private Const CSIDL_COMMON_FAVORITES = &H1F
Private Const CSIDL_INTERNET_CACHE = &H20
Private Const CSIDL_COOKIES = &H21
Private Const CSIDL_HISTORY = &H22
Private Const CSIDL_COMMON_APPDATA = &H23          'All Users\Application Data
Private Const CSIDL_WINDOWS = &H24                 'GetWindowsDirectory()
Private Const CSIDL_SYSTEM = &H25                  'GetSystemDirectory()
Private Const CSIDL_PROGRAM_FILES = &H26           'C:\Program Files
Public Const CSIDL_MYPICTURES = &H27               'C:\Program Files\My Pictures
Private Const CSIDL_PROFILE = &H28                 'USERPROFILE
Private Const CSIDL_SYSTEMX86 = &H29               'x86 system directory on RISC
Private Const CSIDL_PROGRAM_FILESX86 = &H2A        'x86 C:\Program Files on RISC
Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B    'C:\Program Files\Common
Private Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C 'x86 Program Files\Common on RISC
Private Const CSIDL_COMMON_TEMPLATES = &H2D        'All Users\Templates
Private Const CSIDL_COMMON_DOCUMENTS = &H2E        'All Users\Documents
Private Const CSIDL_COMMON_ADMINTOOLS = &H2F       'All Users\Start Menu\Programs _
                                                  '\Administrative Tools
Private Const CSIDL_ADMINTOOLS = &H30              '{user}\Start Menu\Programs _
                                                  '\Administrative Tools
Private Const CSIDL_FLAG_CREATE = &H8000&          'combine with CSIDL_ value to force
                                                  'create on SHGetSpecialFolderLocation()
Private Const CSIDL_FLAG_DONT_VERIFY = &H4000      'combine with CSIDL_ value to force
                                                  'create on SHGetSpecialFolderLocation()
Private Const CSIDL_FLAG_MASK = &HFF00             'mask for all possible flag values

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

'Delete to reycle bin
Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type
Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long




'Window change size
Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'Public Enum ShwWindow
Private Const SW_SHOWNORMAL& = 1
Private Const SW_MAXIMIZE& = 3
Private Const SW_NoAction& = -1
'End Enum

'Routinen, um auch unter Win2000 das Fenster nach vorne zu bringen
'Functions, to bring a window in Win2000/XP on top
Public Declare Function apiAttachThreadInput Lib "user32" Alias _
    "AttachThreadInput" (ByVal idAttach As Long, ByVal idAttachTo As Long, _
        ByVal fAttach As Long) As Long
Public Declare Function apiGetForegroundWindow Lib "user32" _
    Alias "GetForegroundWindow" () As Long

Public Declare Function apiGetWindowThreadProcessId Lib "user32" Alias _
    "GetWindowThreadProcessId" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Public Declare Function apiSetForegroundWindow Lib "user32" _
    Alias "SetForegroundWindow" (ByVal hWnd As Long) As Long

'Fenster nach oben bringen (darf nicht minimiert sein)
Declare Function apiBringWindowToTop Lib "user32.dll" Alias "BringWindowToTop" (ByVal hWnd As Long) As Long

Declare Function apiGetWindow Lib "user32" Alias _
                "GetWindow" (ByVal hWnd As Long, _
                ByVal wCmd As Long) As Long

Declare Function apiGetDesktopWindow Lib "user32" Alias _
                "GetDesktopWindow" () As Long

Declare Function apiGetWindowTextLength Lib "user32" Alias _
    "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Public Const GW_OWNER = 4
Private Const WM_CLOSE = &H10

Public Declare Function apiFindWindow Lib "user32" Alias _
  "FindWindowA" (ByVal lpClassName As String, _
  ByVal lpWindowName As String) As Long
  
'Window open?
Public Declare Function PostMessage& Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)
Public Const mcGWLSTYLE = (-16)
Public Const mcWSVISIBLE = &H10000000
Public Const mconMAXLEN = 255

Declare Function apiGetWindowText Lib "user32" Alias _
                "GetWindowTextA" (ByVal hWnd As Long, ByVal _
                lpString As String, ByVal aint As Long) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As _
    String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hWnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long

Private Declare Function GetShortPathName _
    Lib "kernel32.dll" _
    Alias "GetShortPathNameA" _
( _
    ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal lBuffer As Long _
) As Long

Private Const MAX_PATH As Long = 260&

Public Function GetShortPath(ByVal FileName As String) As String
    'returns short names of directory
    Dim n As Long, Temp As String
    Temp = Space$(MAX_PATH)
    n = GetShortPathName(FileName, Temp, MAX_PATH)
    
    ' Es ist ein Fehler aufgetreten./error
    If n = 0& Then
        GetShortPath = FileName
    Else
        GetShortPath = Left$(Temp, n)
    End If
End Function
Public Sub IExplorer(url)
'Opens URL (website) in standard browser
On Error Resume Next
    Dim lRet As Long
    lRet = ShellExecute(0, "open", url, "", "", vbNormalFocus) 'NoTag
    DoEvents
End Sub

Public Function AppActivatePlus(ByVal Titel As String, ByVal exact As Boolean, Optional ShowWindowCmd& = SW_NoAction, Optional ExcludeString = "") As Boolean
'Activates an application in all windows versions
'released Nov 5, 2006
On Error GoTo Err_AppActivatePlus

    Dim MyHandle As Long, FGHandle As Long, MyWindowThreadID As Long, lForegroundWindowThreadID As Long
    
    AppActivatePlus = False
    
    MyHandle = GetWindowTitleHandle(Titel, exact, ExcludeString)
    FGHandle = apiGetForegroundWindow()
    If MyHandle = 0 Or FGHandle = 0 Then Exit Function
    
    If MyHandle = FGHandle Then GoTo ExitTrue_AppActivatePlus:
    
    MyWindowThreadID = apiGetWindowThreadProcessId(MyHandle, ByVal 0&)
    lForegroundWindowThreadID = apiGetWindowThreadProcessId(FGHandle, ByVal 0&)
    
    If (MyWindowThreadID <> lForegroundWindowThreadID) Then
        Call apiAttachThreadInput(lForegroundWindowThreadID, _
            MyWindowThreadID, 1)    'Attach thread
    
        Call apiSetForegroundWindow(MyHandle)
    
        Call apiAttachThreadInput(lForegroundWindowThreadID, _
            MyWindowThreadID, 0)    'Detach thread
    End If

    If (apiGetForegroundWindow() <> MyHandle) Then Call apiSetForegroundWindow(MyHandle)

    'windows maximieren, normal, minimieren
    If ShowWindowCmd <> SW_NoAction Then
        apiShowWindow MyHandle, ShowWindowCmd
    End If
    
ExitTrue_AppActivatePlus:
    AppActivatePlus = True
    'Otherwise the window is aktiv, but not on top
    apiBringWindowToTop (MyHandle)
    
Exit_AppActivatePlus:
    Exit Function

Err_AppActivatePlus:
    DisplayError "AppActivatePlus"
    Resume Exit_AppActivatePlus
End Function

Public Sub DisplayError(ProzedureName As String)
'Zeigt Programmfehler an
'On Error Resume Next 'Nein, sonst geht die Fehlernummer verloren

    Application.Activate
    
    If Err.Number = 4605 Then
        MsgBox ProzedureName & vbCrLf & "ErrNo: " & Err.Number & " " & Err.Description & vbCr & vbCr & "Document is too complex, split in parts and convert each part separatly!", vbCritical, ConverterPrgTitle
        If DebugMode = False Then End
    End If

    'DisplayErrorL Prozedurname, Err.Number, Err.Description
    MsgBox ProzedureName & vbCrLf & "ErrNo: " & Err.Number & " " & Err.Description, vbExclamation, ConverterPrgTitle

End Sub

Private Function fGetCaption(hWnd As Long) As String
    Dim strBuffer As String
    Dim intCount As Integer

    strBuffer = String$(mconMAXLEN - 1, 0)
    intCount = apiGetWindowText(hWnd, strBuffer, mconMAXLEN)
    If intCount > 0 Then
        fGetCaption = Left$(strBuffer, intCount)
    End If
End Function

Public Function GetWindowTitleHandle(ByVal WindowTitle As String, Optional useBinaryCompare As Boolean, Optional ByVal ExcludeString$ = "", Optional PartOnly As Boolean = False) As Long
' -------------------------------------------------------------------
' Function: retrieves window handle for a not so specific title
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: WindowTitle to search for (begins with title), exclusion words
'
' returns: Number of Window
'
' released: Nov 12, 2006
' -------------------------------------------------------------------

    Dim lngx As Long, lngLen As Long
    Dim lngStyle As Long, strCaption As String
    Dim titleLength As Integer, MyCompare As Integer
    
    If useBinaryCompare Then
        MyCompare = vbBinaryCompare
    Else
        MyCompare = vbTextCompare
    End If
    
    GetWindowTitleHandle = 0
    
    titleLength = Len(WindowTitle)
    If titleLength = 0 Then Exit Function
    
    lngx = apiGetDesktopWindow()
    'Return the first child to Desktop
    lngx = apiGetWindow(lngx, GW_CHILD)
    
    If useBinaryCompare Then
        GetWindowTitleHandle = apiFindWindow(vbNullString, WindowTitle)
    Else
        Do While Not lngx = 0
            strCaption = Left(fGetCaption(lngx), titleLength)
            If strCaption <> "" Then
                'Debug.Print strCaption
                If StrComp(strCaption, WindowTitle, MyCompare) = 0 Then
                    If ExcludeString = "" Then
                        GetWindowTitleHandle = lngx
                        Exit Do
                    ElseIf InStr(1, fGetCaption(lngx), ExcludeString, vbTextCompare) = 0 Then
                        GetWindowTitleHandle = lngx
                        Exit Do
                    End If
                End If
            End If
            lngx = apiGetWindow(lngx, GW_HWNDNEXT)
        Loop
    End If
    
    'nothing found
    If PartOnly And GetWindowTitleHandle = 0 Then
        'Find first matching window
        lngx = apiGetDesktopWindow()
        'Return the first child to Desktop
        lngx = apiGetWindow(lngx, GW_CHILD)
    
        Do While Not lngx = 0
            strCaption = fGetCaption(lngx)
            If strCaption <> "" Then
                Debug.Print strCaption
                If InStr(1, strCaption, WindowTitle, MyCompare) > 0 Then
                    If ExcludeString = "" Then
                        GetWindowTitleHandle = lngx
                        Exit Do
                    ElseIf InStr(1, fGetCaption(lngx), ExcludeString, vbTextCompare) = 0 Then
                        GetWindowTitleHandle = lngx
                        Exit Do
                    End If
                End If
            End If
            lngx = apiGetWindow(lngx, GW_HWNDNEXT)
        Loop
    End If
    
End Function

'Here is some VB Code I used to Find a WinWord 6.0 Document window (DOC2.DOC), and Close it with a PostMessage.
'
'
'Sub Command1_Click()
'    Dim hWndApp%, hWndChild%, r%, text As String * 80
'
'
'    'Fill text with spaces
'    text = Space$(80)   'For safety.
'
'
'    'Find the App Window.
'    hWndApp% = FindWindow(0&, "Microsoft Word")
'
'
'    'Find the Winword Desktop's Top Window
'    hWndChild% = GetTopWindow(hWndApp%)
'    While hWndChild%
'        r% = GetClassName(hWndChild%, text, 80)
'        If Trim(text) = "OpusDesk" & Chr$(0) Then
'            'Find the Child MDI Window.
'            hWndChild% = GetTopWindow(hWndChild%)
'            While hWndChild%
'                r% = GetWindowText(hWndChild%, text, 80)
'                If Trim(text) = "DOC2.DOC" & Chr$(0) Then
'                    r% = PostMessage(hWndChild%, WM_CLOSE, 0, 0&)
'                    hWndChild% = 0&
'                Else
'                    hWndChild% = GetNextWindow(hWndChild%, GW_HWNDNEXT)
'                End If
'            Wend
'        End If
'        hWndChild% = GetNextWindow(hWndChild%, GW_HWNDNEXT)
'    Wend
'End Sub
'
'
'Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
'Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
'Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
'Declare Function GetNextWindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
'Declare Function GetTopWindow Lib "User" (ByVal hWnd As Integer) As Integer
'Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
'Declare Function PostMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Integer
'
'
'Global Const GW_HWNDFIRST = 0   'Returns the first sibling window for a child window; otherwise, it returns the first top-level window in the list.
'Global Const GW_HWNDLAST = 1    'Returns the last sibling window for a child window; otherwise, it returns the last top-level window in the list.
'Global Const GW_HWNDNEXT = 2    'Returns the sibling window that follows the given window in the window manager's list.
'Global Const GW_HWNDPREV = 3    'Returns the previous sibling window in the window manager's list.
'Global Const GW_OWNER = 4       'Identifies the window's owner.
'Global Const GW_CHILD = 5       'Identifies the window's first child window.
'
'
'Global Const WM_CLOSE = &H10
'
'

Public Function GetSpecialfolder(CSIDL As Long) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    Const NOERROR& = 0
    Dim Path$
    
    'Get the special folder
    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If r = NOERROR Then
        'Create a buffer
        Path$ = Space$(512)
        'Get the path from the IDList
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        'Remove the unnecessary chr$(0)'s
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

Public Function DirExists(sPathName) As Boolean
' -------------------------------------------------------------------
' Funktion: Prüft, ob Verzeichnis existiert
'
' Parameter: Path, trailing \ does not matter
'
' Rückgabewerte: wahr, wenn existent
'
' Aufgerufene Prozeduren: GetFileAttributesA
'
' letzte Änderung: 26.05.2002
' -------------------------------------------------------------------
    Dim attr As Long

    attr = GetFileAttributesA(sPathName)

    DirExists = Not (attr = FILE_ATTRIBUTE_INVALID)

End Function

Public Function FileExists(ByVal sPathName As String) As Boolean
' -------------------------------------------------------------------
' Funktion: Prüft Existenz von Datei, schneller als Dir
'   Since we only want to return TRUE for a file in this case, we only need to
'   check for a set '1' in the directory flag position.
'
' Parameter: keine
'
' Rückgabewerte: wahr, wenn vorhanden
'
' Aufgerufene Prozeduren: GetFileAttributesA
'
' letzte Änderung: 26.05.2002
' -------------------------------------------------------------------
    Dim attr As Long

    attr = GetFileAttributesA(sPathName)

    ' The directory bit is set if the path is a directory
    ' or if it does not exist (in which case attr will be -1,
    ' which includes a set directory bit).
    FileExists = ((attr And FILE_ATTRIBUTE_DIRECTORY) = 0)
End Function

Public Function GetFileName(ByVal Path$) As String
'retrieves filename of path & file
On Error Resume Next
    
    Dim p&
    
    p = InStrRev(Path, "\")
    If p > 0 Then
        GetFileName = Mid$(Path, p + 1)
    Else
        GetFileName = Path
    End If
    
End Function

Public Function GetFilePath(ByVal Pfad As String) As String
'ermittelt aus Verzeichnis & Datei das Verzeichnis
'erstellt 12.09.00
On Error Resume Next
    
    Dim p As Integer
    
    p = 0
    Do
        p = InStr(p + 1, Pfad, "\")
        If p > 0 Then GetFilePath = Left(Pfad, p) Else Exit Do
    Loop
    
End Function

Function FormatPath(ByVal Path As Variant, Optional LastChar$ = "\") As String
' -------------------------------------------------------------------
' Function: makes sure a \ is at the end of the path
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: Path
'
' returns: formatted path
'
' released: June 04, 2006
' changed:
' -------------------------------------------------------------------
On Error Resume Next
    
    FormatPath = IIf(right$(Path, 1) = LastChar, Path, Path & LastChar)

End Function

Function ReplaceStr(TextIn, SearchStr, Replacement, Optional CompMode As Integer = vbTextCompare)
' Replaces the SearchStr string with Replacement string in the TextIn string.
' Uses CompMode to determine comparison mode
' Aus der Neatcd97.mdb Microsoft
' for Word97 (it does not have the replace function)
' not used, because of other problems with word, like colors, pictures...
'
Dim WorkText As String, Pointer As Integer
    If IsNull(TextIn) Then
        ReplaceStr = Null
    Else
        WorkText = TextIn
        Pointer = InStr(1, WorkText, SearchStr, CompMode)
        Do While Pointer > 0
            WorkText = Left(WorkText, Pointer - 1) & Replacement & Mid(WorkText, Pointer + Len(SearchStr))
            Pointer = InStr(Pointer + Len(Replacement), WorkText, SearchStr, CompMode)
        Loop
        ReplaceStr = WorkText
    End If
End Function

Public Function RGB2HTML(ByVal RGBColor As Long) As String
'http://www.aboutvb.de/khw/artikel/khwrgbhtml.htm
    Dim nRGBHex As String
    
    nRGBHex = right$("000000" & Hex$(OleConvertColor(RGBColor)), 6)
    RGB2HTML = "#" & right$(nRGBHex, 2) & Mid$(nRGBHex, 3, 2) & Left$(nRGBHex, 2)

End Function

Public Sub KillToBin(FilePath$)
'Delete to recycle bin
'http://www.allapi.net/tips/tip050.shtml
'http://vbnet.mvps.org/index.html?code/shell/shfileopadv.htm

    Dim SHop As SHFILEOPSTRUCT
    Dim strFile As String
    Const FOF_NOCONFIRMATION As Long = &H10     'don't prompt the user.

    If FileExists(FilePath) Then
        With SHop
            .wFunc = FO_DELETE
            .pFrom = FilePath
            .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
        End With
        
        SHFileOperation SHop
        DoEvents
    End If

End Sub

Public Function OleConvertColor(ByVal color As Long) As Long
  Dim nColor As Long
  
  OleTranslateColor color, 0&, nColor
  OleConvertColor = nColor
End Function

Public Function DirExistsCreate(ByVal Path$, Optional ErrMsg As Boolean = True) As Boolean
' -------------------------------------------------------------------
' Function: creates path of any depth
'
' returns: true, if exists or could be created
'
' Released: 06-MAY-2005
' -------------------------------------------------------------------
On Error Resume Next

    If DirExists(Path) Then DirExistsCreate = True: Exit Function
    
    Path = FormatPath(Path)
    If Not DirExists(Path) Then SHCreateDirectoryEx 0, Path, ByVal 0&
    If DirExists(Path) Then
        DirExistsCreate = True
    ElseIf ErrMsg Then
        'MsgBox "Der Path: " & Path & vbCrLf & "konnte nicht angelegt werden. Ändern Sie den Path.", vbExclamation, App.Title
        MsgBox "The directory: " & Path & vbCrLf & "could not be created.", vbExclamation
    End If

End Function

Public Sub AnonymizeText()
On Error Resume Next 'GoTo 0

    Dim myDoc As Document
    Dim St As Object, sh As Shape
    Dim pg As Paragraph
    Dim fd As Field
    Dim pc&, i&
       
    ActiveDocument.TrackRevisions = False
    WordBasic.AcceptAllChangesInDoc
       
    Set myDoc = ActiveDocument
    For Each St In myDoc.StoryRanges
        For Each fd In St.Fields
            If fd.Type = wdFieldTOC Then
                fd.Delete
            Else
                fd.Select
                AnonymizeWords fd.Result
            End If
        Next
        For Each pg In St.Paragraphs
            i = i + 1
            If i Mod 50 = 0 Then pg.Range.Select
            AnonymizeWords pg.Range
        Next
    Next
    For Each St In myDoc.InlineShapes
        St.Select
        DoEvents
        AnonymizeWords St.Range
    Next
    For Each St In myDoc.Shapes
        St.Select
        DoEvents
        Select Case St.Type
            Case msoGroup
                For Each sh In St.GroupItems
                    If sh.TextFrame.HasText Then AnonymizeWords sh.TextFrame.TextRange
                Next
            Case Else
                If St.TextFrame.HasText Then AnonymizeWords St.TextFrame.TextRange
        End Select
    Next

    Selection.HomeKey Unit:=wdStory
    ActiveDocument.SaveAs FileName:=myDoc.Path & "\" & myDoc.Name & "_Anonymized.doc", _
         FileFormat:=wdFormatDocument, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False
    MsgBox "Document anomynized!"
   
End Sub

Private Sub AnonymizeWords(WordObject As Object)

    Dim wd As Object
    Dim l&, c&, iWd&
    Dim c1&, c2&
       
    c1 = WordObject.Words.Count
    For iWd = 1 To WordObject.Words.Count
        If WordObject.Words.Count <> c1 Then Stop
        Set wd = WordObject.Words(iWd)
        c = wd.Characters.Count
        If iWd Mod 50 = 0 Then wd.Select
        If wd.Fields.Count = 0 Then
            If c > 1 Then
                l = Len(wd.Text)
                'If iWd = 17 Then Stop
                If c <> l And Asc(wd.Characters.First) <> 13 Then Stop
                Select Case Asc(wd.Characters.First)
                    Case 13
                        'we bypass this word
                    Case Is < 48, 58 To 64, 91 To 96, 123 To 127, 148
                        'keep first character
                        Select Case Asc(wd.Characters.Last)
                            Case 32
                                wd.Text = wd.Characters.First & String(l - 2, "x") & wd.Characters.Last
                            Case Else
                                wd.Text = wd.Characters.First & String(l - 1, "x")
                        End Select
                    Case Else
                        Select Case Asc(wd.Characters.Last)
                            Case 32
                                wd.Text = String(l - 1, "x") & wd.Characters.Last
                            Case Else
                                wd.Text = String(l, "x")
                        End Select
                End Select
            Else
                Select Case Asc(wd.Characters.First)
                    Case 48 To 57, 65 To 90, 97 To 122
                        wd.Text = "x"
                End Select
            End If
        End If
    Next
   
End Sub
