Attribute VB_Name = "modW2MWP_FileDialog"
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

'#MGe# - Modul erstellt auf Basis von
'Quelle: http://www.mvps.org/access/api/api0001.htm
'Datum: 18.08.06 - Manfred Gerwing

' API: Call the standard Windows File Open/Save dialog box
' Author(s) Ken Getz
'
'  This can be done by either using the Common Dialog Control in Access 97 or by using the
'  APIs defined for this purpose.
'  To call the actual dialog from your code, see the included function TestIt() within the
'  module or use the following example as a guideline and ...

'Dim strFilter As String
'Dim strInputFileName As String
'
'strFilter = ahtAddFilterItem(strFilter, "Excel Files (*.XLS)", "*.XLS")
'strInputFileName = ahtCommonFileOpenSave( _
'                Filter:=strFilter, OpenFile:=True, _
'                DialogTitle:="Please select an input file...", _
'                Flags:=ahtOFN_HIDEREADONLY)

'  ... note that in order to call the Save As dialog box, you can use the same wrapper function
'  by just setting the OpenFile option as False. For example,

'Ask for SaveFileName
'strFilter = ahtAddFilterItem(myStrFilter, "Excel Files (*.xls)", "*.xls")
'strSaveFileName = ahtCommonFileOpenSave( _
'                                    OpenFile:=False, _
'                                    Filter:=strFilter, _
'                    Flags:=ahtOFN_OVERWRITEPROMPT Or ahtOFN_READONLY)

'***************** Code Start **************
'This code was originally written by Ken Getz.
'It is not to be altered or distributed, except as part of an application.
'You are free to use it in any application, provided the copyright notice is left unchanged.
'
' Code courtesy of:
' Microsoft Access 95 How-To
' Ken Getz and Paul Litwin
' Waite Group Press, 1996

Option Explicit

Type tagOPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    strFilter As String
    strCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    strFile As String
    nMaxFile As Long
    strFileTitle As String
    nMaxFileTitle As Long
    strInitialDir As String
    strTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    strDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Declare Function aht_apiGetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" (OFN As tagOPENFILENAME) As Boolean

Declare Function aht_apiGetSaveFileName Lib "comdlg32.dll" _
    Alias "GetSaveFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Global Const ahtOFN_READONLY = &H1
Global Const ahtOFN_OVERWRITEPROMPT = &H2
Global Const ahtOFN_HIDEREADONLY = &H4
Global Const ahtOFN_NOCHANGEDIR = &H8
Global Const ahtOFN_SHOWHELP = &H10
' You won't use these.
'Global Const ahtOFN_ENABLEHOOK = &H20
'Global Const ahtOFN_ENABLETEMPLATE = &H40
'Global Const ahtOFN_ENABLETEMPLATEHANDLE = &H80
Global Const ahtOFN_NOVALIDATE = &H100
Global Const ahtOFN_ALLOWMULTISELECT = &H200
Global Const ahtOFN_EXTENSIONDIFFERENT = &H400
Global Const ahtOFN_PATHMUSTEXIST = &H800
Global Const ahtOFN_FILEMUSTEXIST = &H1000
Global Const ahtOFN_CREATEPROMPT = &H2000
Global Const ahtOFN_SHAREAWARE = &H4000
Global Const ahtOFN_NOREADONLYRETURN = &H8000
Global Const ahtOFN_NOTESTFILECREATE = &H10000
Global Const ahtOFN_NONETWORKBUTTON = &H20000
Global Const ahtOFN_NOLONGNAMES = &H40000
' New for Windows 95
Global Const ahtOFN_EXPLORER = &H80000
Global Const ahtOFN_NODEREFERENCELINKS = &H100000
Global Const ahtOFN_LONGNAMES = &H200000

Function TestIt()
    Dim strFilter As String
    Dim lngFlags As Long
    strFilter = ahtAddFilterItem(strFilter, "Access Files (*.mda, *.mdb)", _
                    "*.MDA;*.MDB")
    strFilter = ahtAddFilterItem(strFilter, "dBASE Files (*.dbf)", "*.DBF")
    strFilter = ahtAddFilterItem(strFilter, "Text Files (*.txt)", "*.TXT")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    MsgBox "You selected: " & ahtCommonFileOpenSave(InitialDir:="C:\", _
        Filter:=strFilter, FilterIndex:=3, Flags:=lngFlags, _
        DialogTitle:="Hello! Open Me!")
    ' Since you passed in a variable for lngFlags,
    ' the function places the output flags value in the variable.
    Debug.Print Hex(lngFlags)
End Function

Function GetOpenFile(Optional varDirectory As Variant, _
    Optional varTitleForDialog As Variant) As Variant
' Here's an example that gets an Access database name.
Dim strFilter As String
Dim lngFlags As Long
Dim varFileName As Variant
' Specify that the chosen file must already exist,
' don't change directories when you're done
' Also, don't bother displaying
' the read-only box. It'll only confuse people.
    lngFlags = ahtOFN_FILEMUSTEXIST Or _
                ahtOFN_HIDEREADONLY Or ahtOFN_NOCHANGEDIR
    If IsMissing(varDirectory) Then
        varDirectory = ""
    End If
    If IsMissing(varTitleForDialog) Then
        varTitleForDialog = ""
    End If

    ' Define the filter string and allocate space in the "c"
    ' string Duplicate this line with changes as necessary for
    ' more file templates.
    strFilter = ahtAddFilterItem(strFilter, _
                "Access (*.mdb)", "*.MDB;*.MDA")
    ' Now actually call to get the file name.
    varFileName = ahtCommonFileOpenSave( _
                    OpenFile:=True, _
                    InitialDir:=varDirectory, _
                    Filter:=strFilter, _
                    Flags:=lngFlags, _
                    DialogTitle:=varTitleForDialog)
    If Not IsNull(varFileName) Then
        varFileName = TrimNull(varFileName)
    End If
    GetOpenFile = varFileName
End Function

Function ahtCommonFileOpenSave( _
            Optional ByRef Flags As Variant, _
            Optional ByVal InitialDir As Variant, _
            Optional ByVal Filter As Variant, _
            Optional ByVal FilterIndex As Variant, _
            Optional ByVal DefaultExt As Variant, _
            Optional ByVal FileName As Variant, _
            Optional ByVal DialogTitle As Variant, _
            Optional ByVal hWnd As Variant, _
            Optional ByVal OpenFile As Variant) As Variant
' This is the entry point you'll use to call the common
' file open/save dialog. The parameters are listed
' below, and all are optional.
'
' In:
' Flags: one or more of the ahtOFN_* constants, OR'd together.
' InitialDir: the directory in which to first look
' Filter: a set of file filters, set up by calling
' AddFilterItem. See examples.
' FilterIndex: 1-based integer indicating which filter
' set to use, by default (1 if unspecified)
' DefaultExt: Extension to use if the user doesn't enter one.
' Only useful on file saves.
' FileName: Default value for the file name text box.
' DialogTitle: Title for the dialog.
' hWnd: parent window handle
' OpenFile: Boolean(True=Open File/False=Save As)
' Out:
' Return Value: Either Null or the selected filename

'#MGe#: eingef¸gt wg Mehrfachauswahl + dort ersetzt, wo vorher 256 stand
Const lngMaxFilenameBuffer = &H7FFF     'entspricht 32 KB

Dim OFN As tagOPENFILENAME
Dim strFileName As String
Dim strFileTitle As String
Dim fResult As Boolean
    ' Give the dialog a caption title.
    If IsMissing(InitialDir) Then InitialDir = CurDir
    If IsMissing(Filter) Then Filter = ""
    If IsMissing(FilterIndex) Then FilterIndex = 1
    If IsMissing(Flags) Then Flags = 0&
    If IsMissing(DefaultExt) Then DefaultExt = ""
    If IsMissing(FileName) Then FileName = ""
    If IsMissing(DialogTitle) Then DialogTitle = ""
'#MGe#
'    Kann nat¸rlich in einem NICHT-Access-Umfeld nicht funktionieren. Ggf. durch API-Call ersetzen.
'    If IsMissing(hwnd) Then hwnd = Application.hWndAccessApp
'    If IsMissing(hwnd) Then hwnd = GetWindowTitleHandle(Application.ActiveDocument.Name)
    If IsMissing(hWnd) Then hWnd = 0
    If IsMissing(OpenFile) Then OpenFile = True
    ' Allocate string space for the returned strings.
    strFileName = Left(FileName & String(lngMaxFilenameBuffer, 0), lngMaxFilenameBuffer)
    strFileTitle = String(lngMaxFilenameBuffer, 0)
    ' Set up the data structure before you call the function
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = hWnd
        .strFilter = Filter
        .nFilterIndex = FilterIndex
        .strFile = strFileName
        .nMaxFile = Len(strFileName)
        .strFileTitle = strFileTitle
        .nMaxFileTitle = Len(strFileTitle)
        .strTitle = DialogTitle
        .Flags = Flags
        .strDefExt = DefaultExt
        .strInitialDir = InitialDir
        ' Didn't think most people would want to deal with
        ' these options.
        .hInstance = 0
        '.strCustomFilter = ""
        '.nMaxCustFilter = 0
        .lpfnHook = 0
        'New for NT 4.0
        .strCustomFilter = String(255, 0)
        .nMaxCustFilter = 255
    End With
    ' This will pass the desired data structure to the
    ' Windows API, which will in turn it uses to display
    ' the Open/Save As Dialog.
    If OpenFile Then
        fResult = aht_apiGetOpenFileName(OFN)
    Else
        fResult = aht_apiGetSaveFileName(OFN)
    End If

    ' The function call filled in the strFileTitle member
    ' of the structure. You'll have to write special code
    ' to retrieve that if you're interested.
    If fResult Then
        ' You might care to check the Flags member of the
        ' structure to get information about the chosen file.
        ' In this example, if you bothered to pass in a
        ' value for Flags, we'll fill it in with the outgoing
        ' Flags value.
        If Not IsMissing(Flags) Then Flags = OFN.Flags

'#MGe#
'   Das Null-Zeichen zwischen mehreren ausgew‰hlten Dateien muss erhalten bleiben.
'        ahtCommonFileOpenSave = TrimNull(OFN.strFile)
        ahtCommonFileOpenSave = TrimTrailingNull(OFN.strFile)
    Else
        ahtCommonFileOpenSave = vbNullString
    End If
End Function

Function ahtAddFilterItem(strFilter As String, _
    strDescription As String, Optional varItem As Variant) As String
' Tack a new chunk onto the file filter.
' That is, take the old value, stick onto it the description,
' (like "Databases"), a null character, the skeleton
' (like "*.mdb;*.mda") and a final null character.

    If IsMissing(varItem) Then varItem = "*.*"
    ahtAddFilterItem = strFilter & _
                strDescription & vbNullChar & _
                varItem & vbNullChar
End Function

Private Function TrimNull(ByVal strItem As String) As String
Dim intPos As Integer
    intPos = InStr(strItem, vbNullChar)
    If intPos > 0 Then
        TrimNull = Left(strItem, intPos - 1)
    Else
        TrimNull = strItem
    End If
End Function
'************** Code End *****************

'#MGe#
'   Neue Funktion zum Abschneiden von abschlieﬂenden Null-Zeichen.
Private Function TrimTrailingNull(ByVal strItem As String) As String
Dim intPos As Integer

    While right(strItem, 1) = vbNullChar
        intPos = Len(strItem)
        strItem = Left(strItem, intPos - 1)
    Wend

    TrimTrailingNull = strItem
End Function


