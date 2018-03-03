Attribute VB_Name = "modWord2MediaWikiPlus"
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

Private Declare Function MessageBoxW Lib "user32.dll" _
    (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long

Private Declare Function MessageBoxA Lib "user32.dll" _
    (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long) As Long

Private Const MB_ICONINFORMATION As Long = &H40&
Private Const MB_TASKMODAL As Long = &H2000&

Public Const WMPVersion As String = "0.7.0.11 Beta"
Public Const DebugMode As Boolean = False

'Usage
'* Make sure the macro files are in your project NORMAL
'1. Save your Word-Document before you convert!
'   The macro will not save or overwrite the original word file. But for safety reasons do not work with your original.
'2. Run the macro, a configuration dialog will appear.
'3. If everything works out, the data will be in your clipboard: Just paste the text into your wiki editor.

'Elaborate usage
'* Change the CONST settings of your language and your environment
'* Check the const settings for the Image Upload Customizing

'See CONST value description for customizing!
'const with a capital R at the end are stored in registry. This value will be used as a default; do not change unless you customize the defaults for your company.
'See GetRegValidate for default values if none is set

'-- Declarations --

'-- Declarations --
'Public Const WM_KEYDOWN = &H100
'Public Const WM_KEYUP = &H101
'Public Const WM_CHAR = &H102
'Public Const WM_SYSKEYDOWN = &H104
'Public Const WM_SYSKEYUP = &H105
'Private Const WM_SYSCHAR = &H106
'Public Const VK_MENU = &H12
'Private Const WM_COPY = &H301
'Private Const WM_PASTE = &H302
'Private Const EM_SETSEL = &HB1
Private Const WM_CLOSE = &H10
'Private Const WM_GETTEXT = &HD
'Public Const WM_SETTEXT = &HC
'Public Const WM_KEYPRESS = &HC
'
'Const KEYEVENTF_EXTENDEDKEY = &H1
'Const KEYEVENTF_KEYUP = &H2
'Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

'-- Declarations End --

'Global Const
Const autoLoadDocument As Boolean = True 'Will load a document if none or the macro document itself is open (this way the user does not have to copy the macro to Normal.dotm)
Const WikiOpenPage As Boolean = True 'If true, the wiki-page will be opened after converting
Const ProjectHome$ = "http://meta.wikimedia.org/wiki/Word2MediaWikiPlus"
Const NoWikiOn = "$nowiki$"
Const NoWikiOff = "$qnowiki$"
 
'Text converting customization
Const convertFirstLineIndent As Boolean = False 'Set to true and line indention will be marked with :
Const HeaderFirstLevel$ = "==" 'Use "=" if you like, but not recommended by MediaWiki
Const ReplacePageBreakWithLine As Boolean = True 'If true, there will be a horizontal line in the wike, if false, page breaks are deleted only
'Paragraph spacing: What is a new line? Both variants have some flaws, I prefer: false
Const NewParagraphWithBR As Boolean = False 'false: Make two Paragraphs, true: use <br> for line break (not MediaWiki-style)

'manual Tabs should be replaced by tables
Const convertTableWithTabs As Boolean = True  'Manual Tables, build with tabs, will be converted into real tables
Const TabBlanksNo& = 4 'Tabs will be replaced with x spaces; this will not be displayed in wiki; set to 0 to disable replacement

'Table Customizing
Const TableTemplate$ = "border=""2"" cellspacing=""0"" cellpadding=""4"""
'Const TableTemplate$ = "{{Prettytable}}" 'Wikipedia uses this template
'Const TableTemplate$ = "{{Tabellenkopf}}"
Const TableTemplateNoFrame$ = "border=""0"" cellspacing=""2"""
Const TableTemplateParagraphFrame$ = "cellspacing=""0"" cellpadding = ""10"" style=""border-style:solid; border-color:black; border-width:1px;"""

Const TableTemplateParagraphNoFrame$ = "border=""0"" cellspacing=""0"" cellpadding = ""0"""
Const Cell_justify As Boolean = False 'if set to true, cells can be justified. As of MediaWiki 1.8, this is not interpreted

'Photo Editor Information
Const EditorTitle$ = "Microsoft Photo Editor" 'The Photo Editor, I assume you have this Program, if you have MS-Word!
Const EditorPrgPath$ = "" 'give full path to your program or leave empty, it will try the standard position
Const EditorPathPE_R$ = "" 'give full path to MS Photo Editor or leave empty, it will try the standard position
Const EditorPathPM_R$ = "" 'give full path to MS Picture Manager or leave empty, it will try the standard position
Const UsePhotoEditor As Boolean = False
            
            
'Wiki Image Upload Customizing
'This is a bit problematic as I use SendKeys. Diffent languages, different keys.
'You need to customize first.
'To set the quality settings, convert any picture in Microsoft Photo Editor to jpg or png. The Quality settings will be stored and used afterwards.
'Set ImageExtraction to true first an check coding
'then set ImageUploadAuto to true
'ImageExtraction will always be made to the folder of your document
Const ImageUploadTabToDescription& = 3 'Number of tabs to press to get to the file description field
Const ImageUploadTabToEnter& = 3 'Number of tabs to press to get to the file Upload button
Const ImageUploadWaitTime& = 4000 'Milliseconds to wait for wiki website to respond
Const ImageMaxUploadSize = 120000 'Byte; macro will not upload to wiki if filesize is greater
Const ImageIconOnlyType As Boolean = False 'Will save only the type of an icon, not the individual icon. _
    If set to true, you will loose the file name, which is only on the icon, but save unnessecary upload of less usefull pictures.
'Note: Images will always be saved with the correct format extension
Const ImageTagFormat$ = "png" 'Used for the image-link within MediaWiki, should be corresponding to your final export format
Const ImageSaveJPG As Boolean = False 'Always creates a jpg-file (for easy comparison)
Const ImageSaveGIF As Boolean = False 'Always creates a gif-file (for easy comparison)
Const ImageSavePNG As Boolean = True 'Always creates a png-file (for easy comparison), best for screenshots and most charts used in word pictures
Const ImageSaveBMP As Boolean = False 'Always creates a bmp-file (for easy comparison)


'--- Do not change below unless you know what you are doing ---

'Registry Defaults:
'Do not change for an individual computer. Unless you want diffent defaults for other users, there is no need to change, as these settings will be set with dialogs and stored.
'If you change these after your registry is filled, nothing will happen!
Public Const GlobalRegNoSave = False 'If true registry will not be written in any case -> image upload will not work in german!
Const isCustomizedR = False ' True: Shuts off info about not having customized
'Your Environment
Const WikiAddressRootTestR$ = "http://scratchpad.wikia.com/index.php?title=" 'The URL before the Article-name
Const WikiAddressRootProdR$ = "http://localhost/wiki/" 'The URL before the Article-name
Const ImagePathR$ = "" 'Leave empty and the program uses the My-Pictures\wiki path
Const ImageNamePreFixR$ = "" 'uploads of images will get this PreFix & ArticleName
Const ImageUploadTabToFileNameR& = 2 'usally 2 for Internet Explorer; Number of tabs to press to get to the field where the upload filename is entered
Const CategoryArticleR$ = "uncategorized" '
Const CategoryImagesR$ = "uncategorized" '
Const insertTitlePageIfNeededR As Boolean = True
Const convertPageHeadersR As Boolean = False 'PageHeader
Const convertPageFootersR As Boolean = False 'PageFooter
Const deleteHiddenCharsR As Boolean = True 'will delete all text which is marked as hidden
Const ImageExtractionR As Boolean = False 'if false, the image conversion and upload routine is disabled. Problematic routine, details see in MediaWikiExtract_Images
Const ImageUploadAutoR As Boolean = False 'if false, the Upload function is disabled
Const ImageConverterR = "MSPhotoEditor" '"SnagIt" and "MSPhotoEditor" supported
    'you may use SnagIt to convert the images. This is done internally, without using SendKeys.
    'Unfortunatly, it can not interpret the clipboard as good as the MS Photo Editor does.
Const ImageConvertCheckFileExistsR As Boolean = True 'will not convert again, if first file is found on disk!
Const ImageMaxPixelSizeR = 630 'If a picture is wider than this, this pixelsize will be added. Good for printing of documents.

'Global Variables
Private Type DocInfoType
    'stores some global variables
    DocName As String 'Name of converted word document
    DocNameNoExt As String 'Name of converted word document without extension .doc
    ArticleName As String ' Name of wiki article (target name)
    ImageMaxWidth As Long ' Resize information, max. width
    ImageResizeOption As Long
      '=1: use full size image, do not resize at all
      '=2: use full size image, resize only if bigger then ImageMaxWidth
      '=3: use display size: resize only, if bigger then ImageMaxWidth
    ImagePath As String 'Path of the imagefiles
    PowerpointStarted As Boolean ' true, if Powerpoint was actually used
End Type
Public DocInfo As DocInfoType
Public Const ConverterPrgTitle$ = "MediaWiki2WordPlus Converter"
Public Const g_RegTitle = "Word2MediaWikiPlus"
Public isInitialized As Boolean
Public languageArr()
Public ImageArr() As String 'Stores filenames (incl. path) of converted images, counts from 1
Public convertImagesOnly As Boolean
Public EditorPathPE$, EditorPathPM$

Dim DefaultFontSize& 'use to find big and small text elements
Dim DefaultIndent As Single
Public EditorPath$ 'Path of the Photo Editor, do not change
Dim Word97 As Boolean 'do not use the image converter
Dim TableInfoArr() As TableInfoType
Dim WordParagraph$, WordNewLine$, WordManualPageBreak$, WordForceBlank$
Dim OptionSmart As Boolean, OptionWords As Boolean
Dim HeaderCount&
Dim WikiAddressRoot$
Dim NeedsFootnoteReference As Boolean

'Dim CurrWorkArea As Object 'usually ActiveDocument, PageHeader etc; not yet fully supported

'Language specific Texts
'definition look in MW_LanguageTexts
'many more, but not stored in variables, but registry
Dim Msg_Upload_Info$, Msg_Finished$, Msg_NoDocumentLoaded$, Msg_LoadDocument$, Msg_CloseAll$

'used for image conversion
Private Type ImageInfoType
    DisplayWidth As Long 'Display width in Word document in pixels
    hasFrame As Boolean 'true if wiki frame will be added
    Name As String 'Name of the image with extension
    NameDisplay As String 'Name of the image with extension, display size
    NameNoExt As String 'Name of the image without extension
    PositionText As String 'Header, Footer or nothing 'not used for now
    PPTUsed As Boolean ' true, if Powerpoint-Conversion took place for single image
    Resized As Boolean 'True, if image was resized in word
    ScaleHeightReal As Double 'ScaleHeight in %
    ScaleWidthReal As Double 'ScaleWidth in %
    Width As Long 'unscaled width
    Height As Long 'unscaled height
    '--- Test only ---
    ImageNo As Long
    StoryType As Long
    TextStart As Long
    Top As Long
    Left As Long
    IsInlineShape As Boolean
    Type As Long
End Type

'used for table conversion
Private Type TableInfoType
    tableWidth As Double 'in points
    preferredWidth As Long
    ParentCellWidth As Double
End Type
Private Type CellInfoType
    cIndex As Single
    toDelete As Boolean
    mergeRight As Long
    mergeDown As Long
    
    cWidth As Single
    cHeight As Single
    cBackgroundcolor As Long 'Cell.Range.Shading.BackgroundPatternColor
    cAlignment As String 'Cell.Range.Paragraphs(1).Alignment
    cText As String
End Type

'Credits
'* The first version of this converter seems from Swythan at [http://tikiwiki.org/tiki-index.php?page=WordToWiki_swythan TikiWiki]
'* A slightly modified version is [http://www.infpro.com/downloads/downloads/wordmedia.htm Word2Wiki] from [http://www.infpro.com/ InfPro IT-Solutions].
'* Also this converter derives from the above mentioned at this point it is hard to find a single line of code still in use, but they led me in the right direction!

'* For the numerous api descriptions I thank [http://www.allapi.net/ The KPD-Team].
'* And of course: A lot of work of me: [[Benutzer:GunterS|Gunter Schmidt]]. Have a look at my [http://www.beadsoft.de website].
   

'Changes V0.3:
'- general: added some const to customize this
'- general: added hourglass and statustext
'- text: added text color
'- tables: added blank space in empty cells
'- tables: added alignment of text
'- tables: added tableformat string, const TableTemplate
'- hyperlinks: redesign: changed html and file-links, others will not be converted
'- images: added function to save all pictures of the document as .bmp and replace with Image-Tag
'- paragraph spacing: added manual line break and MediaWiki-like paragraphs
'- cleanup-function

'Changes V0.4
'- Image conversion and upload
'- styles localized
'- tables with background colors
'- tables with line breaks
'- tables with merged cells
'- fixed bug with combined formats
'- added simple fontsize
'- added simple indention of paragraphs
'- Word97 disabled with message
'- cleaned up text format functions

'Changes V0.5
'* Feature: Nested tables
'* Feature: CheckWikiUpload test
'* Feature: ImagePasteInEditor which works better (now we might go with Word97 again (but color problem remains)
'* Feature: "manual tables" conversion, converts tables that are made with tab stops
'* Feature: PageHeaders and PageFooters (optional)
'* Feature: delete hidden text (optional)
'* Feature: localization for ENG and GER, other languages can be added
'* Feature: Macro can be started from a special macro document (Name must include the term Word2MediaWikiPlus; must not include the term Demo)
'* Changed Upload function, customizing the keystrokes, use CheckWikiUpload to test
'* Changed Picture Export to Paste with MS Photo Editor (works a lot better)
'* Changed: no need for Default Paragraph Style anymore
'* Changed: code reorganisation and optimization
'* BugFix: Macro crashed if Icon returned error
'* BugFix: FontSize could go into endless loop
'* BugFix: Escape characters did not work correctly, if document contained "*"
'* BugFix: some tables did not merge correctly
'* BugFix: superscript and subscript
'* BugFix: text colors in tables
'* BugFix: some bugs with combined format like colored and bold
'* BugFix: exclude internal hyperlinks within document
'* BugFix: some minor changes and bugfixes

'Changes V0.6
'* Feature: Dialog for Article name and categories
'* Feature: Using registry to store user values
'* Feature: Installation routine (in separate document)
'* Feature: Separate Upload Image File Dialog
'* Feature: Converts footnotes
'* Feature: table width and alignment
'* Feature: framed and colored paragraphs
'* Feature: Icons of applications now will be saved separatly to prevent double uploading
'* Feature: User Headings based on the build in Headings will be correctly recognized
'* Feature: Word 2002 uses internal PNG copy function for better image extraction
'* Feature: Tables cells with dark background will get white font
'* Feature: Textboxes will be converted to framed tables
'* Feature: Added PreFix for Imagefiles
'* Feature: Manually numbered lists -> Word numbering will be used
'* Feature: Optional use SnagIt for converting screenshots
'* Feature: Image upload dialog
'* Changed: User language will be detected
'* Changed: Manually created tab tables do not have a frame anymore
'* Changed: Const values for categories
'* Changed: Discontinued option to save to BMP and convert with MS Photo Editor. Now it will always paste in MS Photo Editor and then save the desired format.
'* Changed: Font size from big and small to font 1 to 7
'* Changed: Discontinued internal BMP conversion
'* BugFix: Formatting in tables, needed a workaround for a word bug
'* BugFix: Merged Cells did not work all the times

'Changes V0.6c
'* BugFix: Now allows converting without MS Photo Editor (without pictures)
'* BugFix: SearchPage did not work on systems that needed title =
'* BugFix: remembers last used system for upload (test / prod)
'* BugFix: some line shapes did not convert
'* Changed: Image upload function now does not rely on vb components (Manfred Gerwing)
'* Changed: FontColor converter a bit faster (Hal Eden)

'Changes V0.7a
'* BugFix: Turn of hyphenation before converting
'* BugFix: lists in tables work now
'* BugFix: headers in tables work now
'* BugFix: fixed error with msoCanvas pictures
'* BugFix: fixed error Confic Dialog, did not find MS Photo Editor
'* Feature: Images will now be exported with html-converting (needed for Word 2003)
'* Feature: Images will now be copied to special upload directory when uploaded to wiki
'* Feature: Images will now be saved in separate folders for each document (HTML-Export only)
'* Feature: recognize "forced blanks" and replace with &nbsp;
'* Feature: converts form fields into text
'* Feature: comments are ignored
'* Feature: cell_justify switch implemented. Since MediaWiki will not support the justify tag, it is now removed. Search for "cell_justify".
'* Feature: numbered lists which use symbols will be converted to bullet lists
'* Feature: Help links in user dialog forms
'* Changed: Replacement function now works with a range object, which is faster and more reliable
'* Changed: Replaced some SendKeys with SendMessage
'* Changed: FontFormat now uses more sizes then big and small
'* Changed: Category handling
'* Deprecated: Allow/disallow empty categories
'* Deprecated: SnagIt Image conversion

'ToDo:
'- parameter for picture resize: display size or max size
'  - Formularänderung und Registry, DocInfo benutzen

'- HTML Sonderzeichen (entities)
'- User Language (=office language) and Wiki Language for test and prod system
'- Choose path button
'- Save as temporary file with different name.
'- Word2MediaWiki-Category entfernen
'- Objektorientierte Funktionen
'- Namespace for articles
'- take a look at the document structure and make a difference between one and two paragraph spacing (Bsp. Abschied)
'- HTML-Background
'- Formular-Felder entfernen

'ToDo images
'- inlineshape = shape
'- Textfelder
'- do not delete all lines
'- Suchtext für Bilder eingeben
'- Upload Images Category
'- Image Category frei einstellbar, Haken vor Kategorie
'- check upload/reload image functions
'- Grafiken mit Position
'- überlagernde Grafiken
'- gruppierte grafiken erst gruppiert und dann ungruppiert als extra datei speichern und über Powerpoint
'- MD5-hash of images to identify identical images

'ToDo tables
'- reprogramming of table function
'- TabTables: Use TabStop as width info, add width

Public Sub Word2MediaWikiPlus()
' -------------------------------------------------------------------
' Function: Converts a word document to the wiki syntax
' Main Procedure: Run this
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: June 17, 2006
' -------------------------------------------------------------------
On Error GoTo Err_Word2MediaWikiPlus
   
    Dim Answer&, p&
    
    MW_LanguageTexts
    
    'no document open --> load
    If Documents.Count = 0 Then
        MsgBox Msg_LoadDocument, vbInformation, ConverterPrgTitle
        Application.Dialogs(wdDialogFileOpen).Show
        DoEvents
        If Documents.Count = 0 Then
            'User did not open a document
            MsgBox Msg_NoDocumentLoaded, vbInformation, ConverterPrgTitle
            Exit Sub
        End If
    End If
        
    'Determine which document to convert, usally the active document
    'check if the active document is not the macro document itself
    If autoLoadDocument Then
        'use active document
        If Documents.Count = 1 Then
            If InStr(1, Documents(1).Name, "Word2MediaWikiPlus", vbTextCompare) > 0 And InStr(1, ActiveDocument.Name, "Demo", vbTextCompare) = 0 Then
                'but not if it is the macro document
                MsgBox Msg_LoadDocument, vbInformation, ConverterPrgTitle
                Application.Dialogs(wdDialogFileOpen).Show
                DoEvents
                If InStr(1, ActiveDocument.Name, "Word2MediaWikiPlus", vbTextCompare) > 0 Then
                    'User did not open a document
                    MsgBox Msg_NoDocumentLoaded, vbInformation, ConverterPrgTitle
                    Exit Sub
                End If
            End If
        End If
        
        'Active Document not macro document, then ok.
        'In case we have more than one document, we do not load but use the open document
        If InStr(1, ActiveDocument.Name, "Word2MediaWikiPlus", vbTextCompare) > 0 And InStr(1, ActiveDocument.Name, "Demo", vbTextCompare) = 0 Then
            If Documents.Count > 2 Then
                'we can not identify which of the open documents the user wants to convert
                MsgBox Msg_CloseAll, vbInformation, ConverterPrgTitle
                Exit Sub
            End If
            
            'use other document then the macro document
            Dim myDoc As Document
            For Each myDoc In Documents
                With myDoc
                    If InStr(1, myDoc.Name, "Word2MediaWikiPlus", vbTextCompare) = 0 Then
                        myDoc.Activate
                        DoEvents
                        Exit For
                    End If
                End With
            Next myDoc
        End If
    End If
    DocInfo.DocName = ActiveDocument.Name
    DocInfo.DocNameNoExt = DocInfo.DocName
    p = InStrRev(DocInfo.DocName, ".")
    If p > 0 Then DocInfo.DocNameNoExt = Left$(DocInfo.DocName, p - 1)
        
'---- user dialog on config settings for this conversion run ----
    
    If Not MW_Initialize Then Exit Sub
    frmW2MWP_Doc_Config.Show
    If Not isInitialized Then Exit Sub
    
'---- start converting ----
    
    'Delete the image directory
    If GetReg("ImageExtraction") Then
        Dim txt$
        txt = MW_GetImagePath
        If txt <> "" Then RemoveDir (txt)
    End If
    
    'Replace standard values
    With Options
        OptionWords = .AutoWordSelection
        'OptionSmart = .SmartParaSelection 'Word 2002
    End With
   
    'All conversions
    
    'Set CurrWorkArea = ActiveDocument
    Erase ImageArr
    
    If convertImagesOnly Then
        MW_InsertPageHeaders
        MediaWikiExtract_Images
    Else
        MediaWikiConvert_Prepare
        MediaWikiConvert_Fields
        MediaWikiReplaceQuotes
        MediaWikiConvert_Comments
        MediaWikiConvert_EscapeChars
        
        MediaWikiConvert_FormFields
        MediaWikiConvert_Hyperlinks
        MediaWikiConvert_Headings
        MediaWikiConvert_FootNotes
        MediaWikiConvert_IndentionTab
        MediaWikiConvert_TextFormat
        MediaWikiConvert_Lists
        MediaWikiConvert_Tables
        MediaWikiConvert_Paragraphs
        MediaWikiExtract_Images
        'MediaWikiConvert_Indention 'not wanted
        MediaWikiConvert_CleanUp
    End If
    
    'Replace standard values
    With Options
        .AutoWordSelection = OptionWords
        '.SmartParaSelection = OptionSmart 'Word 2002
    End With
    
    'Upload images
    'MediaWikiImageUpload

    ActiveDocument.Content.Copy ' Copy to clipboard
    Application.ScreenUpdating = True
    system.Cursor = wdCursorNormal
    StatusBar = "Converting to MediaWiki finished! Now paste in your wiki!"

    'Close Photo Editor
    'MW_CloseProgramm EditorTitle ' commented out by E. Lorenz becuase the moduels hangs here foreever
    'If AppActivatePlus(EditorTitle, False) Then DoEvents: SendKeys "%{F4}", True 'End Photo Editor
    DoEvents
                    
    'Open Page in MediaWiki
    'MediaWikiOpen
    
    'Finished - do not enter code here
    
Exit_Word2MediaWikiPlus:
    Exit Sub

Err_Word2MediaWikiPlus:
    DisplayError "Word2MediaWikiPlus"
    Resume Exit_Word2MediaWikiPlus
End Sub

Public Function MW_CloseProgramm(WindowTitle$, Optional useCompareBinary As Boolean = False) As Boolean

    Dim hWnd&, lRet&

    hWnd = GetWindowTitleHandle(WindowTitle, useCompareBinary)
    If hWnd = 0 Then Exit Function
    lRet = SendMessage(hWnd, WM_CLOSE, 0, 0)
    If lRet = 0 Then MW_CloseProgramm = True

End Function


Public Sub Word2MediaWikiPlus_Config()
'Calls the config dialog

    MW_LanguageTexts
    isInitialized = True
    frmW2MWP_Config.Show
    isInitialized = False

End Sub

Public Sub Word2MediaWikiPlus_Upload()
    frmW2MWP_UploadImages.Show
End Sub

Private Sub MediaWikiConvert_CleanUp()
' -------------------------------------------------------------------
' Function: Final steps on converting the document
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: Nov. 28, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_CleanUp

    Dim pg As Paragraph
    Dim rg As Range
    Dim txt$
    Dim c&, p&
    
    If WordParagraph = "" Then WordParagraph = "^p" 'Only for testing
    
    'font type is irrelevant after conversion, but symbol types sometimes mess up format strings
    ActiveDocument.Range.Style = wdStyleNormal
    With ActiveDocument.Range.Font
        .Name = "Arial"
        .Size = 10
        .Bold = False
        .Italic = False
    End With
    
    'replace "forced blanks" (geschützte Leerzeichen)
    MW_ReplaceString WordForceBlank, "&nbsp;"
    
    'remove TOC if not needed
    If HeaderCount < 4 And Not GetReg("AllowWiki") Then
        MW_ReplaceString "__TOC__" & WordParagraph & WordParagraph & "----", ""
        MW_ReplaceString "__TOC__" & WordParagraph & "----", ""
        MW_ReplaceString "__TOC__", ""
    End If
    
    If HeaderCount = 0 Then
        MW_ReplaceString GetReg("txt_TitlePage") & WordParagraph & WordParagraph & "----", ""
        MW_ReplaceString GetReg("txt_TitlePage") & WordParagraph & "----", ""
    End If
    
    'remove empty paragraphs at begin of document
    Set pg = ActiveDocument.Paragraphs(1)
    Do While Not pg.Next Is Nothing
        If pg.Range.Text = vbCr Then
            pg.Range.Delete
        Else
            Exit Do
        End If
    Loop
    
    'no double big
    MW_ReplaceString "</big><big>", ""

    'remove blanks at begin of each paragraph
    'maybe there is a faster method?
    For Each pg In ActiveDocument.Paragraphs
        'blanks at the beginning
        Do While Left$(pg.Range.Text, 1) = " "
            pg.Range.Text = Mid$(pg.Range.Text, 2)
'            Set rg = pg.Range
'            rg.Collapse
'            rg.MoveEnd
'            rg.Delete
        Loop
    Next

    'remove all empty paragraphs at end of document
    Selection.EndKey Unit:=wdStory
    Do
        Selection.MoveLeft wdCharacter, 1, wdExtend
        If Selection.Text = vbCr Then
            Selection.Delete
        Else
            Exit Do
        End If
    Loop Until ActiveDocument.Paragraphs.Count = 1
    
    'do not allow more than two empty paragraphs in a row
    MW_ReplaceString WordParagraph & WordParagraph & WordParagraph, WordParagraph & WordParagraph, True, , 20
    
    'Clean up Tab Indention
    If 1 = 2 Then 'not used
    For Each pg In ActiveDocument.Paragraphs
        Do
            p = InStr(1, pg.Range.Text, "##TAB##", vbBinaryCompare)
            If p > 1 Then
                'we need to move this!
                pg.Range.Select
                Selection.Collapse
                Selection.MoveRight , p - 1
                Selection.MoveRight , 7, True
                Selection.Delete
                pg.Range.InsertBefore ":"
            End If
        Loop Until p <= 1
    Next
    MW_ReplaceString "##TAB##", ":"
    End If

    'non convertable images
    MW_ReplaceString "#$$$#<", " "
    
    'Tables even with a blank line are to close together, separate with two blank lines
    'might be a wiki bug
    MW_ReplaceString "|}" & WordParagraph & WordParagraph & "{|", "|}" & WordParagraph & WordParagraph & WordParagraph & "{|"
    
    'Footnote reference
    If NeedsFootnoteReference Then
        Selection.EndKey wdStory
        If HeaderCount > 10 And 1 = 2 Then
            Selection.InsertAfter vbCr & vbCr & HeaderFirstLevel & GetReg("txt_Footnote") & HeaderFirstLevel & vbCr & vbCr & "<references/>" & vbCr
        Else
            Selection.InsertAfter vbCr & vbCr & "----" & vbCr & vbCr & "<references/>" & vbCr
        End If
    End If
    
    'Finally add the categories
    txt = MW_FormatCategoryString(GetReg("CategoryArticle"))
    If GetReg("CategoryArticleUse") And txt <> "" Then
        Selection.EndKey wdStory
        Selection.InsertAfter vbCr & vbCr & txt & vbCr
    End If
    
    If MW_WordVersion = 2003 Then MW_SetOptions_2003 False
    SetReg "Z_finished", True

Exit_MediaWikiConvert_CleanUp:
    MW_ChangeView 4
    Exit Sub

Err_MediaWikiConvert_CleanUp:
    DisplayError "MediaWikiConvert_CleanUp"
    Resume Exit_MediaWikiConvert_CleanUp
End Sub

Public Sub MediaWikiConvert_Comments()
' -------------------------------------------------------------------
' Function: converts typical Comments in wiki style: Nop, will only delete them
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: Oct. 01, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_Comments

    Dim MyComment As Comment
    Dim i&
    Dim addr$

    For i = 1 To ActiveDocument.Comments.Count
       
        Set MyComment = ActiveDocument.Comments(1)
        With MyComment 'must be 1, since the delete changes count and position

           .Delete
          
        End With

    Next i

Exit_MediaWikiConvert_Comments:
    Exit Sub

Err_MediaWikiConvert_Comments:
    DisplayError "MediaWikiConvert_Comments"
    Resume Exit_MediaWikiConvert_Comments
End Sub

Private Sub MediaWikiConvert_EscapeChars()

    MW_ReplaceSpecialCharactersFirst
    
    'MW_ReplaceCharacter "_"
    'MW_ReplaceCharacter "-"
    MW_ReplaceCharacter "+"
    'MW_ReplaceCharacter "/"
    MW_ReplaceCharacter "{"
    MW_ReplaceCharacter "}"
    MW_ReplaceCharacter "["
    MW_ReplaceCharacter "]"
    MW_ReplaceCharacter "~"
    MW_ReplaceCharacter "^^"
    MW_ReplaceCharacter "|"
    MW_ReplaceCharacter "'"
    
    'HTML-Chars
    MW_ReplaceCharacter "<"
    MW_ReplaceCharacter ">"
    
    MW_ReplaceString NoWikiOff & NoWikiOn, "" 'save some space if several characters in a row
    MW_ReplaceString NoWikiOn, "<nowiki>"
    MW_ReplaceString NoWikiOff, "</nowiki>"
    
End Sub

Private Sub MediaWikiConvert_Fields()
' -------------------------------------------------------------------
' Function: converts Word fields
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: Nov 20, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_Fields
    
    Dim MyField As Field
    Dim c&, i&
    
    MW_Statusbar True, "converting fields..."
    c = ActiveDocument.Fields.Count
    'ActiveDocument.Fields.Update
    Do While c > 0
        Set MyField = ActiveDocument.Fields(c)
        With MyField
            '.Select 'only for testing
            Select Case .Type
                Case wdFieldTOC
                    .Delete 'TOC will be deleted
                Case wdFieldRef
                    '.Select 'for debug only
                    '.Update 'no, we do not know what really happens
                    '#ToDo# set reference to [[#..]] if heading
                    'do not know how
                    .Unlink
'                    If Asc(.Result.Text) <> 1 Then
'                        'MyField.Result.InsertBefore .Result.Text 'inserts in field and will be deleted
'                        MyField.Select
'                        Selection.InsertBefore .Result.Text
'                    End If
'                    'check for pictures, jup it happens
'                    If .Result.InlineShapes.Count = 0 Then
'                        'If .Result.ShapeRange.Count = 0 Then 'may result in error
'                        .Delete
'                    End If
                'Case wdFieldHyperlink ' do nothing
            End Select
            
        End With
        c = c - 1
    Loop

Exit_MediaWikiConvert_Fields:
    Exit Sub

Err_MediaWikiConvert_Fields:
    DisplayError "MediaWikiConvert_Fields"
    If DebugMode Then
        If Not MyField Is Nothing Then MyField.Select
        Stop
        Resume Next
    End If
    Resume Exit_MediaWikiConvert_Fields
End Sub

Private Sub MediaWikiConvert_FontColors()
' -------------------------------------------------------------------
' Function: Converts FontColors in HTML-colors
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input:
'
' returns: nothing
'
' released: Nov 06, 2006
' adapted to swiki haleden 10/28/06, try word level scanning before reverting to character level
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_FontColors

    Dim CurColor& 'Current Color, indicates change
    Dim OpenColor& 'Color the font was opened with
    Dim pgColor&, wordColor&
    Dim cNo& 'Number of characters
    Dim txt$
    Dim FontOpen As Boolean, NeedsClose As Boolean

    Dim pg As Paragraph, MyWd As Object
    
    'First check, if the paragraphs have different colors
    'seems Word gives 9999999 if more than one color!

    'Does not work if color is set to wdAuto and background is dark

    MW_Statusbar True, "converting font colors"

    For Each pg In ActiveDocument.Paragraphs
        'blanks at the beginning
        If pgColor <> pg.Range.Font.color And (Asc(pg.Range.Text) <> 13 Or FontOpen) Then
            pgColor = pg.Range.Font.color
            If pgColor = "9999999" Then 'different colors in paragraph
                'Check each letter in paragraph
                'I found no other possibility other then to check each letter
                'Dead slow
                'trying to check each word instead, only chars if mixed in word
                With pg.Range
                    For Each MyWd In pg.Range.Words
                        wordColor = MyWd.Font.color
                        If wordColor <> "9999999" Then  'not more than one font in word
                            If FontOpen = False Then 'currently no  <font> open
                                CurColor = wordColor
                                If RGB2HTML(CurColor) <> "#000000" Then ' not automatic font
                                    If Asc(MyWd.Characters.First.Text) > 13 Then 'not only the paragraph itself!
                                        OpenColor = wordColor
                                        txt = "<font color=""" & RGB2HTML(OpenColor) & """>"
                                        MyWd.InsertBefore txt
                                        FontOpen = True
                                    End If
                                End If
                            Else 'have <font> open
                                If RGB2HTML(wordColor) <> "#000000" Then ' not automatic font
                                    If OpenColor <> wordColor Then 'switching fonts
                                        'close font
                                        OpenColor = wordColor
                                        txt = "</font><font color=""" & RGB2HTML(OpenColor) & """>"
                                        MyWd.InsertBefore txt
                                    End If
                                Else 'switching back to automatic font color
                                    CurColor = 0
                                    OpenColor = 0
                                    txt = "</font>"
                                    MyWd.InsertBefore txt
                                    FontOpen = False
                                End If 'automatic or not
                            End If ' <font> open or not
                        Else 'multiple colors in word, need to handle char by char
                            cNo = 0
                            With MyWd
                            Do While cNo < .Characters.Count
                                cNo = cNo + 1
                                'Debug.Print cNo, .Characters(cNo)
                                If cNo Mod 20 = 0 Then DoEvents
                                    If CurColor <> .Characters(cNo).Font.color Then
                                        If FontOpen = False Then
                                            'open font
                                            CurColor = .Characters(cNo).Font.color
                                            If RGB2HTML(CurColor) <> "#000000" Then
                                                OpenColor = .Characters(cNo).Font.color
                                                txt = "<font color=""" & RGB2HTML(OpenColor) & """>"
                                                .Characters(cNo).InsertBefore txt
                                                FontOpen = True
                                                cNo = cNo + Len(txt) - 1
                                            End If
                                        Else
                                            'close font
                                            CurColor = 0
                                            OpenColor = 0
                                            txt = "</font>"
                                            .Characters(cNo).InsertBefore txt
                                            FontOpen = False
                                            cNo = cNo + Len(txt) - 1
                                            pgColor = 0
                                        End If
                                    End If
                                Loop
                            End With
                        End If
                    Next MyWd
                End With

            'whole paragraph
            ElseIf FontOpen = False Then
                    'open font
                    pgColor = pg.Range.Font.color
                    'pg.Range.Select 'just for testing
                    If RGB2HTML(pgColor) <> "#000000" Then
                        OpenColor = pg.Range.Font.color
                        txt = "<font color=""" & RGB2HTML(OpenColor) & """>"
                        pg.Range.InsertBefore txt
                        FontOpen = True
                        cNo = cNo + Len(txt) - 1
                        NeedsClose = False
                        If pg.Range.Tables.Count > 0 Then
                            NeedsClose = True
                        ElseIf pg.Next Is Nothing Then
                            NeedsClose = True
                        ElseIf pg.Next.Range.Font.color <> pg.Range.Font.color Then
                            NeedsClose = True
                        End If
                        If NeedsClose Then
                            'In tables we need to close within the cell
                            'pg.Range.Characters.Count
                            CurColor = 0
                            OpenColor = 0
                            txt = "</font>"
                            pg.Range.Characters(pg.Range.Characters.Count - 1).InsertAfter txt
                            FontOpen = False
                            pgColor = 0
                            cNo = cNo + Len(txt)
                        End If
                    End If
                Else
                    'close font
                    If pgColor <> OpenColor Then
                        CurColor = 0
                        OpenColor = 0
                        txt = "</font>"
                        pg.Range.InsertBefore txt
                        FontOpen = False
                        cNo = cNo + Len(txt) - 1
                    End If
                'End If
            End If
        End If
    Next
    
Exit_MediaWikiConvert_FontColors:
    MW_Statusbar False
    Exit Sub

Err_MediaWikiConvert_FontColors:
    DisplayError "MediaWikiConvert_FontColors"
    Resume Exit_MediaWikiConvert_FontColors
End Sub

Private Sub MediaWikiConvert_FootNotes()
' -------------------------------------------------------------------
' Function: Converts footnotes for the cite extention
'           http://meta.wikimedia.org/wiki/Cite/Cite.php
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: June 17, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_FootNotes

    Dim rg As Range, fn As Footnote
      
    NeedsFootnoteReference = False
      
    If ActiveDocument.Footnotes.Count = 0 Then Exit Sub
    
    NeedsFootnoteReference = True
    
    For Each fn In ActiveDocument.Footnotes
        fn.Reference.InsertBefore "<ref>" & fn.Range.Text & "</ref>"
        fn.Delete
    Next
    
Exit_MediaWikiConvert_FootNotes:
    Exit Sub

Err_MediaWikiConvert_FootNotes:
    DisplayError "MediaWikiConvert_FootNotes"
    Resume Exit_MediaWikiConvert_FootNotes
End Sub

Private Sub MediaWikiConvert_FormFields()
' -------------------------------------------------------------------
' Function: converts form fields in normal text (what else?)
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input:
'
' returns: nothing
'
' released: September 24, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_FormFields

    Dim MyField As FormField
    Dim i&, j&, Max&
    Dim txt$
    
    For i = 1 To ActiveDocument.FormFields.Count
        
        Set MyField = ActiveDocument.FormFields(1) 'must be 1, since the delete changes count and position
        MyField.Select
        Select Case MyField.Type
        
            Case wdFieldFormTextInput
                If Len(Selection.Text) > 0 Then
                    txt = "<u>"
                    For j = 1 To Len(Selection.Text)
                        txt = txt & "&nbsp;"
                    Next j
                    Selection.InsertBefore txt & "</u>"
                End If
                MyField.Delete
                
            Case wdFieldFormDropDown
                If MyField.DropDown.ListEntries.Count > 0 Then
                    txt = "<u>"
                    Max = 0
                    For j = 1 To MyField.DropDown.ListEntries.Count
                        If Max < Len(MyField.DropDown.ListEntries(j).Name) Then Max = Len(MyField.DropDown.ListEntries(j).Name)
                    Next j
                    For j = 1 To Max
                        txt = txt & "&nbsp;"
                    Next j
                    Selection.InsertBefore txt & "</u>"
                End If
                MyField.Delete
                
            Case wdFieldFormCheckBox
                Selection.InsertBefore "<u>&nbsp;&nbsp;</u>"
                MyField.Delete
        
        End Select

    Next i

Exit_MediaWikiConvert_FormFields:
    Exit Sub

Err_MediaWikiConvert_FormFields:
    DisplayError "MediaWikiConvert_FormFields"
    Resume Exit_MediaWikiConvert_FormFields
End Sub

Private Sub MediaWikiConvert_Headings()
' -------------------------------------------------------------------
' Function: convert paragraphs with style Header to wiki syntax
'           also remove/replace page breaks
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: June 12, 2006
' -------------------------------------------------------------------
On Error Resume Next

    Dim pg As Paragraph, rg As Range

    HeaderCount = 0
    
    MW_Statusbar True, "converting headers..."
    
    Set pg = ActiveDocument.Paragraphs.First
    Do While Not pg Is Nothing
        'pg.Range.Select 'just for testing
        
        Select Case pg.Style
            Case ActiveDocument.Styles(wdStyleHeading1)
                MW_SurroundHeader pg.Range, 1
            Case ActiveDocument.Styles(wdStyleHeading2)
                MW_SurroundHeader pg.Range, 2
            Case ActiveDocument.Styles(wdStyleHeading3)
                MW_SurroundHeader pg.Range, 3
            Case ActiveDocument.Styles(wdStyleHeading4)
                MW_SurroundHeader pg.Range, 4
            Case ActiveDocument.Styles(wdStyleHeading5)
                MW_SurroundHeader pg.Range, 5
            
            Case Else
            'check if style has a header base style
                Select Case pg.Style.BaseStyle
                    Case ActiveDocument.Styles(wdStyleHeading1)
                        MW_SurroundHeader pg.Range, 1
                    Case ActiveDocument.Styles(wdStyleHeading2)
                        MW_SurroundHeader pg.Range, 2
                    Case ActiveDocument.Styles(wdStyleHeading3)
                        MW_SurroundHeader pg.Range, 3
                    Case ActiveDocument.Styles(wdStyleHeading4)
                        MW_SurroundHeader pg.Range, 4
                    Case ActiveDocument.Styles(wdStyleHeading5)
                        MW_SurroundHeader pg.Range, 5
                End Select
        End Select
        
        'remove page breaks
        'after heading, because headings do not get a ---- line
        If pg.Range.ParagraphFormat.PageBreakBefore Then
            pg.Range.ParagraphFormat.PageBreakBefore = False
            pg.Range.InsertBefore Chr(12)
        End If
        If InStr(1, pg.Range.Text, Chr(12)) > 0 Then 'faster
            Set rg = pg.Range
            rg.MoveStartUntil Chr(12)
            rg.Collapse
            rg.Delete
            rg.InsertAfter vbCr & "----" & vbCr
            rg.Style = wdStyleNormal
'            If rg.Style <> wdStyleNormal Then
'                rg.Select
'                'Selection.ClearFormatting 'not in Word 2000
'                If DebugMode Then Stop
'            End If
            Set pg = rg.Next
        End If
        
        Set pg = pg.Next
    Loop
    
End Sub

Private Sub MediaWikiConvert_HTMLChars()
' -------------------------------------------------------------------
' Function: converts special symbols in HTML-Entities
' http://de.selfhtml.org/html/referenz/zeichen.htm
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input:
'
' returns: nothing
'
' released: Nov. 17, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_HTMLChars

    'Prozedur
    
Exit_MediaWikiConvert_HTMLChars:
    Exit Sub

Err_MediaWikiConvert_HTMLChars:
    DisplayError "MediaWikiConvert_HTMLChars"
    If DebugMode Then Stop: Resume Next
    Resume Exit_MediaWikiConvert_HTMLChars
End Sub

Private Sub MediaWikiConvert_Hyperlinks()
' -------------------------------------------------------------------
' Function: converts typical hyperlinks in wiki style
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input:
'
' returns: nothing
'
' released: June 04, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_Hyperlinks

    Dim HyLink As Hyperlink
    Dim i&
    Dim addr$

    For i = 1 To ActiveDocument.Hyperlinks.Count
        
        Set HyLink = ActiveDocument.Hyperlinks(1) 'must be 1, since the delete changes count and position
        With HyLink

            addr = .Address
            If Trim$(addr) = "" Then addr = "no hyperlink found"
            'title = .Range.Text
           
           .Range.Select 'only for testing
           
            If .SubAddress <> "" Then
                'Link within the document
                'we can not do this yet
                .Delete 'hyperlink
                GoTo MediaWikiConvert_Hyperlinks_Next
            End If
            
            If .Range.InlineShapes.Count > 0 Then
                'We have a clickable object
                'we can not do this yet
                'so we insert new text and move the hyperlink
                .Range.InsertAfter GetReg("ClickChartText")
                .Range.Select
                Selection.Collapse wdCollapseEnd
                Selection.MoveRight , Len(GetReg("ClickChartText")), True
                Selection.Hyperlinks.Add Selection.Range, HyLink.Address
                .Delete 'hyperlink
                GoTo MediaWikiConvert_Hyperlinks_Next
            End If
           
            'http, ftp
            If LCase(Left$(addr, 4)) = "http" Or LCase(Left$(addr, 3)) = "ftp" Then
                .Delete 'hyperlink
                .Range.InsertBefore "[" & addr & " "
                .Range.InsertAfter "]"
               
                GoTo MediaWikiConvert_Hyperlinks_Next
            End If
           
            'mailto:
            If LCase(Left$(addr, 7)) = "mailto:" Then
                .Delete 'hyperlink
                .Range.InsertBefore "[" & addr & " "
                .Range.InsertAfter "]"
               
                GoTo MediaWikiConvert_Hyperlinks_Next
            End If
           
            'file guess
            If Len(addr) > 4 Then 'the reason for not nice goto
                If Mid$(addr, Len(addr) - 3, 1) = "." Then
                    .Delete
                    If GetFilePath(addr) = "" Then addr = FormatPath(ActiveDocument.Path) & addr
                    .Range.InsertBefore "[file://" & Replace(addr, " ", "_") & " "
                    .Range.InsertAfter "]"
                   
                    GoTo MediaWikiConvert_Hyperlinks_Next
                End If
            End If
           
            'unidentified
            .Delete
            .Range.InsertBefore GetReg("UnableToConvertMarker") & "[" & addr & " "
            .Range.InsertAfter "]"

MediaWikiConvert_Hyperlinks_Next:
        End With

    Next i

Exit_MediaWikiConvert_Hyperlinks:
    Exit Sub

Err_MediaWikiConvert_Hyperlinks:
    DisplayError "MediaWikiConvert_Hyperlinks"
    Resume Exit_MediaWikiConvert_Hyperlinks
End Sub

Private Sub MW_ImageInfoReset(ImageInfo As ImageInfoType)
    
    Dim ImageInfoEmpty As ImageInfoType
    
    ImageInfo = ImageInfoEmpty
    
    'set defaults
    With ImageInfo
        .hasFrame = True
    End With

End Sub

Public Sub MediaWikiExtract_Images(Optional DocArea& = 0, Optional SectionNo&, Optional HeaderNo&)

    'change view to layout
    MW_ChangeView 4
    
    'SetReg "ImageExtractionPE", False
    If GetReg("ImageExtractionPE") Then
        MediaWikiExtract_ImagesPhotoEditor DocArea, SectionNo, HeaderNo
    Else
        'default
        MediaWikiExtract_ImagesHtml DocArea, SectionNo, HeaderNo
    End If

End Sub

Public Sub MediaWikiExtract_ImagesHtml(Optional DocArea& = 0, Optional SectionNo&, Optional HeaderNo&)
' -------------------------------------------------------------------
' Function: Extracts and saves all images to disk
' Some words to the extraction prozess
' All pictures are copied to a new document, which is then saved as webpage (html).
' There will be two pictures created, the original picture and another one in a common graphic format, just the size of display
' Depending on the settings and the format one of the two will be used.
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: Nov. 20, 2006
' -------------------------------------------------------------------
' #ToDo#
' floating shapes may be converted separatly
' use Anchor information
' redesign, so shape and inlineshape will work similar

On Error GoTo Err_MediaWikiExtract_ImagesHtml
    
    Dim s As Selection
    Dim MyWorkArea As Object
    Dim MyIS As InlineShape, MyS As Shape
    Dim ImageNameBase$, PicType$, IconClassType$, OLabel$, IconPath$, cTxt$
    Dim ConvertShape As Boolean, RepeatLoop As Boolean, FrameCreated As Boolean
    Dim PicNo&, Crop&
    Dim ImageInfo As ImageInfoType
    
    Const txtNotConvertExample$ = vbCr & "#$$$#< " & ConverterPrgTitle & " found a non convertable object. Please send example to developer." & vbCr & " #$$$#< " & ProjectHome & vbCr
    Const txtNotConvert$ = vbCr & "#$$$#< " & ConverterPrgTitle & " found a non convertable object." & vbCr
    
    If Not isInitialized Then MW_Initialize
    DocInfo.ImagePath = MW_GetImagePath
    
    'Erase ImageArr 'no, because it may be called twice (header and document)
    
    MW_ImageInfoReset ImageInfo
    
    'ImageNameBase = GetReg("ImageNamePreFix") & DocInfo.ArticleName & IIf(ImageInfo.PositionText <> "", "_" & ImageInfo.PositionText, "") & "_"
    ImageNameBase = GetReg("ImageNamePreFix") & DocInfo.ArticleName & "_"
    
    Select Case DocArea
        Case 0
            Set MyWorkArea = Documents(DocInfo.DocName)
            ImageInfo.PositionText = ""
        Case 1
            Set MyWorkArea = Documents(DocInfo.DocName).Sections(SectionNo).Headers(HeaderNo)
            ImageInfo.PositionText = "Header"
            MW_ChangeView 1
        Case 2
            Set MyWorkArea = Documents(DocInfo.DocName).Sections(SectionNo).Footers(HeaderNo)
            ImageInfo.PositionText = "Footer"
            MW_ChangeView 1
    End Select
    
    MW_Statusbar False, "converting images"
    
    'Get current pic number
    On Error Resume Next
    PicNo = UBound(ImageArr)
    If Err.Number <> 0 Then PicNo = 0: Err.Clear
    On Error GoTo Err_MediaWikiExtract_ImagesHtml
    
    'Different count for header and footer
'    Do While PicNo > 0
'        If Len(ImageArr(PicNo)) > Len(ImageNameBase) Then
'            If Left$(ImageArr(PicNo), Len(ImageNameBase)) = ImageNameBase Then
'                PicNo = Val(Mid$(ImageArr(PicNo), Len(ImageNameBase) + 1))
'                Exit Do
'            End If
'        End If
'    Loop
    
    'Find all shapes and convert to InlineShapes, so we know its position
    'Well, the outcome is unpredictable, messing up the picture position
    'For Each myS In MyWorkArea.Shapes
    '    If myS.Type = msoGroup Then myS.Ungroup
    'Next myS
    Do
    RepeatLoop = False
    For Each MyS In MyWorkArea.Shapes
    
        MW_ImageInfoReset ImageInfo
        
        MyS.Select
        Set s = Selection
        DoEvents
        ConvertShape = False
        'On Error Resume Next 'Sometimes it can not be converted, but no crash
        'Debug.Print MyS.Anchor.Start, MyS.Left, MyS.Width, MyS.top, MyS.Height
        Select Case MyS.Type
            Case msoGroup
                If MW_WordVersion >= 2002 Then
                    'can not convert to InlineShape, so we just export it
                    ConvertShape = True
                Else
'                    'Word can copy an png, otherwise it will be safed as gif with less quality.
                    Selection.Copy
                    MyS.Delete
                    Selection.MoveRight
                    'Paste as png inline
                    DoEvents
                    Selection.PasteSpecial Link:=False, DataType:=14, Placement:=wdFloatOverText, DisplayAsIcon:=False
                    RepeatLoop = True
                End If
                
            Case msoLine
                MyS.Select
                If MyS.Height > 2 Then '(not a horizontal line)
                    ConvertShape = True
                Else
                    'convert to wiki divider line, could be wrong direction... #ToDo#
                    MyS.Delete
                    Selection.InsertAfter vbCr & "----" & vbCr 'wiki line
                    Selection.Range.Style = wdStyleNormal
                End If
                
            Case msoTextBox, msoAutoShape
                'convert to framed paragraph
                'MyS.ConvertToInlineShape
                If MW_WordVersion >= 2002 Then
                    'Again, a word bug: Font will be Times New Roman
                    'in this case we copy the text and make a table
                    With MyS
                        If .TextFrame.HasText Then
                            '#ToDo# convert to wiki markup
                            cTxt = .TextFrame.TextRange.Text
                            cTxt = Replace(cTxt, vbCr, "<br>" & vbCr)
                            'for now we just add a wiki table
                            '#ToDo# Line type and frame size
                            cTxt = vbCr & "{|" & TableTemplateParagraphFrame & vbCr & "|" & cTxt & "|}" & vbCr & vbCr
                            If Selection.Start = 0 Then
                                'we need to unselect the shape, so we go to the beginning
                                Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:="1"
                                'and move back
                                Selection.Start = MyS.Anchor.Start
                            Else
                                Selection.MoveLeft
                            End If
                            Selection.InsertAfter cTxt
                        End If
                        .Select
                    End With
                    ImageInfo.hasFrame = False
                    ConvertShape = True
                Else
                    MyS.ConvertToFrame
                    FrameCreated = True
                    Selection.InsertBefore txtNotConvert
                    '#ToDo#: text and test
                    If MyS.AlternativeText <> "" Then
                        MyS.ConvertToFrame
                        FrameCreated = True
                    End If
                End If
                
            Case msoPicture, msoFreeform
                'if converted to InlineShape, we get two shapes and a lot of problems, so we convert directly, which we now can with html export
                ConvertShape = True

            Case 20 'msoCanvas; does not exist in Word 2000
                'can not convert to InlineShape, so we just export it
                'ConvertShape = True
                'Extract Image
                PicNo = PicNo + 1
                ImageInfo.NameNoExt = MW_CheckFileName(ImageNameBase & Format(PicNo, "00"))
                ImageInfo.Name = ""
'                If MyS.Fill.Visible = msoFalse Then
'                    MyS.Fill.Solid
'                    MyS.Fill.ForeColor.RGB = RGB(255, 255, 255)
'                    MyS.Fill.Transparency = 0#
'                End If
                DoEvents
                MW_ImageExtract MyS, ImageInfo
                'if conversion is off, we do not have the image format #ToDo#
                If ImageInfo.Name = "" Then ImageInfo.Name = MW_GetImageNameFromFile(ImageInfo.NameNoExt, DocInfo.ImagePath)
                
                'prevent compiler error in Word 2000
                If MW_WordVersion >= 2002 Then MediaWikiExtract_ImagesHtml2002 MyS, ImageInfo
                
                MyS.Select
                Set s = Selection
                If Selection.Start = 0 And DebugMode Then Stop
                MyS.Delete
                Selection.MoveRight , , True
                If Selection.InlineShapes.Count > 0 Then Selection.Delete Else Selection.MoveLeft
                'Selection.MoveLeft
                Selection.InsertAfter vbCr & "[[Image:" & ImageInfo.Name & IIf(ImageInfo.hasFrame, "|framed|none", "") & "]]" & vbCr
            
            Case Else
                'On Error Resume Next 'Sometimes it can not be converted, but should not crash
                'MyS.ConvertToInlineShape
                'If Err.Number <> 0 Then
                    If DebugMode Then Stop
                    'MyS.Delete
                    ConvertShape = True
                    Selection.InsertBefore txtNotConvertExample & " FormType = " & MyS.Type
                    'Err.Clear
                'End If
                'On Error GoTo Err_MediaWikiExtract_ImagesHtml
        End Select
        If ConvertShape Then
            'Extract Image
            PicNo = PicNo + 1
            ImageInfo.NameNoExt = MW_CheckFileName(ImageNameBase & Format(PicNo, "00"))
            ImageInfo.Name = ""
            If MyS.Type = msoPicture Then
                If MW_WordVersion >= 2002 Then
                    'Check if picture is cut, otherwise we would extract the whole picture, but we only want the visible area
                    On Error Resume Next
                    Crop = 0
                    Crop = MyIS.PictureFormat.CropBottom + MyIS.PictureFormat.CropLeft + MyIS.PictureFormat.CropRight + MyIS.PictureFormat.CropTop
                    On Error GoTo Err_MediaWikiExtract_ImagesHtml
                    If Crop > 10 Then
                        MW_ImageExtract MyS, ImageInfo, True
                    Else
                        MW_ImageExtract MyS, ImageInfo
                    End If
                Else
                    MW_ImageExtract MyS, ImageInfo
                End If
            Else
'                If MyS.Fill.Visible = msoFalse Then
'                    MyS.Fill.Solid
'                    MyS.Fill.ForeColor.RGB = RGB(255, 255, 255)
'                    MyS.Fill.Transparency = 0#
'                End If
                DoEvents
                MW_ImageExtract MyS, ImageInfo
            End If
            'if conversion is off, we do not have the image format #ToDo#
            If ImageInfo.Name = "" Then ImageInfo.Name = MW_GetImageNameFromFile(ImageInfo.NameNoExt, DocInfo.ImagePath)
            'MyS.Select
            Set s = Selection
            If Selection.Start = 0 Then
                'we need to unselect the shape, so we go to the beginning
                Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:="1"
                'and move back
                Selection.Start = MyS.Anchor.Start
            End If
            MyS.Delete
            'Selection.Delete
                'not sure if this deletes consequtive pictures
                'Selection.MoveRight , , True
'                If Selection.InlineShapes.Count > 0 Then Selection.Delete Else Selection.MoveLeft
            Selection.InsertAfter vbCr & "[[Image:" & ImageInfo.Name & IIf(ImageInfo.hasFrame, "|framed|none", "") & "]]" & vbCr
        End If
        Selection.InsertAfter vbCr & vbCr
        DoEvents
        On Error GoTo Err_MediaWikiExtract_ImagesHtml
        
    Next MyS
    Loop Until RepeatLoop = False
    DoEvents

    'if header or footer only convert to inline shapes
    Select Case DocArea
        Case 1, 2
            MW_ChangeView 4
            Exit Sub
    End Select
    
    'Textboxes will get a nice frame at least
    If FrameCreated And convertImagesOnly = False Then MediaWikiConvert_Tables
    
    'Convert all InlineShapes
    For Each MyIS In MyWorkArea.InlineShapes
    
        MW_ImageInfoReset ImageInfo
        
        'identify Image-Type, it could be an inline-document
        MyIS.Select
        IconPath = ""
        IconClassType = ""
        OLabel = ""
        Select Case MyIS.Type
            Case wdInlineShapePicture
                PicType = "Pic"
                
            Case wdInlineShapeHorizontalLine
                PicType = "Line"
                MyIS.Select
                MyIS.Range.InsertAfter vbCr & "----" & vbCr 'wiki line
                MyIS.Range.Style = wdStyleNormal
                MyIS.Delete
            
            Case wdInlineShapeOLEControlObject
                PicType = "do not convert"
                MyIS.Select
                If DebugMode Then Stop
                MyIS.Range.InsertAfter txtNotConvert
                MyIS.Range.Style = wdStyleNormal
                MyIS.Delete
                
            Case Else
                PicType = "Pic" 'default or unidentified
                
                If Not MyIS.OLEFormat Is Nothing Then
                    'in some cases the icon does not have this information
                    IconClassType = Replace(MyIS.OLEFormat.ClassType, ".", " ")
                    If MyIS.OLEFormat.DisplayAsIcon Then
                        PicType = "Icon"
                        'IconPath = myIS.OLEFormat.IconPath
                        OLabel = MyIS.OLEFormat.IconLabel
                        If OLabel = "" Then
                            'Word bug, does not return the IconLabel first time
                            DoEvents
                            OLabel = MyIS.OLEFormat.IconLabel
                        End If
                        If OLabel = "" Then OLabel = IconClassType & " Object"
                    End If
                    OLabel = Replace(OLabel, """", "")
                    If OLabel = "" Then OLabel = "no label"
                    
                End If
        End Select
        
        
        Select Case PicType
        
            Case "Pic", "Icon"
    
                'Insert Image-Tag
                
                PicNo = PicNo + 1
                If PicType = "Icon" Then
                    'IconNo = IconNo + 1
                    If ImageIconOnlyType Then
                        'General usage
                        ImageInfo.NameNoExt = MW_CheckFileName("Icon " & IconClassType)
                    Else
                        'Document specific usage
                        'ImageInfo.NameNoExt = MW_CheckFileName("Icon " & ImageNameBase & Format(IconNo, "00"))
                        ImageInfo.NameNoExt = MW_CheckFileName(ImageNameBase & Format(PicNo, "00"))
                    End If
                Else
                    ImageInfo.NameNoExt = MW_CheckFileName(ImageNameBase & Format(PicNo, "00"))
                End If
                
                'Calculate Original Size (displayed size)
                ImageInfo.DisplayWidth = PointsToPixels(MyIS.Width, False)
                'Fill white or some backgrounds will be black
                If MyIS.Fill.Visible = msoFalse Then
                    MyIS.Fill.Solid
                    MyIS.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    MyIS.Fill.Transparency = 0#
                End If
                DoEvents
                
                'Calculate real ScaleHeight and ScaleWidth
                MW_GetScaleIS MyIS, ImageInfo.ScaleWidthReal, ImageInfo.ScaleHeightReal
            
                'If we extract a picture, there are usualle two files
                '1) The original picture, saved in original size and format
                '2) The displayed picture, saved as gif or jpg in its displayed size,
                '   so we loose colors (quality) and size
                'The macro will always use the full size picture and use wiki resize function, unless it can not be used:
                '(so it can be viewed full or by deleting the size and thumb parameter in the image-link included as full)
                ' a) The format is emz or wmz which can not be displayed
                ' b) The picture is cropped. In this case the original picture has different content as it displays more.
                'In any case, both files are stored, if usable:
                ' 1) ArticleName_Number.ext for the original
                ' 2) ArticleName_Number-display.ext for the displayed image
                'The user can then change the imagelink if he likes to use the other file better.
                
                'To maximize quality all displayed images are maximized to 100% or at least to 800x600 (max. which word can export) and resized in wiki.
                'If we need the display picture, we copy as png, which has some quality loss (word bug), but is better than jpg.
                
                'Check size, maximize to 800x600 for better results
                ImageInfo.Resized = False
                If ImageInfo.ScaleWidthReal < 99 And ImageInfo.ScaleHeightReal < 99 Then 'smaller than 100%
                    MW_ScaleMax MyIS
                    ImageInfo.Resized = True
                End If
                
                'Extract Image
                MyIS.Select
                ImageInfo.Name = ""
                On Error Resume Next
                Crop = 0
                Crop = MyIS.PictureFormat.CropBottom + MyIS.PictureFormat.CropLeft + MyIS.PictureFormat.CropRight + MyIS.PictureFormat.CropTop
                On Error GoTo Err_MediaWikiExtract_ImagesHtml
                If MW_WordVersion >= 2002 Then
                    'Check if picture is cropped, otherwise we would extract the whole picture, but we only want the visible area
                    'We could use the display picture, but that is jpg. If we copy as png, we have better quality
                    If Crop > 10 Then
                        'First extrakt the full image, the user can edit that one later
                        MW_ImageExtract MyIS, ImageInfo, True
                    Else
                        MW_ImageExtract MyIS, ImageInfo
                    End If
                Else
                    'Check if picture is cropped, otherwise we would extract the whole picture, but we only want the visible area
                    'We could use the display picture, but that is jpg. If we copy as png, we have better quality
                    If Crop > 10 Then
                        'First extrakt the full image, the user can edit that one later
                        '#ToDo# Powerpoint
                        MW_ImageExtract MyIS, ImageInfo ', True
                    Else
                        MW_ImageExtract MyIS, ImageInfo
                    End If
                End If
                'if conversion is off, we do not have the image format
                If ImageInfo.Name = "" Then ImageInfo.Name = MW_GetImageNameFromFile(ImageInfo.NameNoExt, DocInfo.ImagePath)
                
                'Delete picture from document (replace by link)
                MyIS.Select
                'Insert Wiki Picture Link [[Image:xxx]]
                'If it is an Icon, then add a Wiki-Link
                If PicType = "Icon" Then MyIS.Range.InsertAfter vbCr & "[[" & ImageNameBase & OLabel & "]]" & vbCr
                
                'resize takes place, if
                'IF is not useFullSize
                'AND (useDisplayWidth 'Info: NOT useDisplayWidth: resize only if greater than ImageMaxWidth
                ' OR DisplayWidth > ImageMaxWidth)
                'AND (image is smaller than 90% of it's original size OR image width is greater than ImageMaxWidth)
                'IF resizing, then check size:
                ImageInfo.Width = PointsToPixels(MyIS.Width)
                'DocInfo.ImageResizeOption = 1
                Select Case DocInfo.ImageResizeOption
                    Case 1 'use full size image, do not resize at all
                        MyIS.Range.InsertAfter "[[Image:" & ImageInfo.Name & "|framed|none" & "]]" & vbCr
                    
                    Case 2 'resize only if bigger then ImageMaxWidth
                        If ImageInfo.Width > GetReg("ImageMaxWidth") Then
                            MyIS.Range.InsertAfter "[[Image:" & ImageInfo.Name & "|thumb|none|" & GetReg("ImageMaxWidth") & "px]]" & vbCr
                        Else
                            MyIS.Range.InsertAfter "[[Image:" & ImageInfo.Name & "|framed|none" & "]]" & vbCr
                        End If
                        
                    Case 3 'use display size: resize only, if bigger then ImageMaxWidth
                        If ImageInfo.DisplayWidth > GetReg("ImageMaxWidth") Then
                            MyIS.Range.InsertAfter "[[Image:" & ImageInfo.Name & "|thumb|none|" & GetReg("ImageMaxWidth") & "px]]" & vbCr
                        ElseIf (ImageInfo.ScaleWidthReal < 95 Or ImageInfo.ScaleWidthReal > 105) Then
                            MyIS.Range.InsertAfter "[[Image:" & ImageInfo.Name & "|thumb|none|" & ImageInfo.DisplayWidth & "px]]" & vbCr
                        Else
                            MyIS.Range.InsertAfter "[[Image:" & ImageInfo.Name & "|framed|none" & "]]" & vbCr
                        End If
                        
                End Select
                
                '#ToDo# if myIS has text flow, do not know how to check
                MyIS.Range.InsertAfter vbCr
                
                MyIS.Delete
                Selection.MoveRight wdCharacter, 1, wdExtend
                If Selection.Text = " " Then Selection.Collapse: Selection.Delete
                
        End Select
        
Err_Continue_5825:
        
    Next MyIS

Exit_MediaWikiExtract_ImagesHtml:
    MW_ChangeView 0
    MW_PowerpointQuit
    Exit Sub

Err_MediaWikiExtract_ImagesHtml:
    Select Case Err.Number
        Case 4605
            MsgBox "You encountered a MS-Word Bug. Delete the marked image from the document and try again", vbCritical, ConverterPrgTitle
            If DebugMode Then Stop
            Resume Next
            End
        Case 5825 'object was deleted; by word, no clue
            Err.Clear
            Selection.InsertAfter "[[Image:## Error Converting ##]]"
            Resume Err_Continue_5825
        Case Else
            DisplayError "MediaWikiExtract_ImagesHtml"
            If DebugMode Then Stop: Resume Next
            Resume Exit_MediaWikiExtract_ImagesHtml
    End Select
End Sub


Private Sub MediaWikiExtract_ImagesPhotoEditor(Optional DocArea& = 0, Optional SectionNo&, Optional HeaderNo&)
' -------------------------------------------------------------------
' Function: Extracts and saves all images to disk
' Some words to the extraction prozess
' A picture can be extracted easily, if it was inserted as bmp, like a screenshot
' and its size is less then 800x600px.
' Problem: picture is bigger than 800x600
'  Solution: MS Photo Editor: paste as new.
'            MS Picture Manager: reduce size to 800x600 and paste.
' Problem: picture is not a bmp, e.g. powerpoint slide
'  Solution: MS Photo Editor: reduce size to 800x600, make new picture in PE with the word size and paste
'          : MS Picture Manager: reduce size to 800x600 and paste.
' Problem: picture is a group of pictures (word group function)
'  Solution: MS Photo Editor: no solution, leave out
'          : MS Picture Manager: reduce size to 800x600 and paste.
'
' Depending on the problem both programs are usefull,
' but Picture Manager limits size to 800x600 and can not extract big screenshots
' I strongly complain the loss of the MS Photo Editor.
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: Nov 12, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiExtract_ImagesPhotoEditor
    
    Dim MyIS As InlineShape
    Dim MyS As Shape
    Dim ImageNameBase$, ImagePathName$, PicType$, IconClassType$, OLabel$, IconPath$, ImageTxt$ 'IconPath not used anymore
    Dim ImageExtractionInt As Boolean, RepeatLoop As Boolean, FrameCreated As Boolean
    Dim p&
    Dim MyWorkArea As Object
    Dim PicNo&, IconNo&
    Const txtNotConvert$ = vbCr & "#$$$#< " & ConverterPrgTitle & " found a non convertable object." & vbCr
    
    Erase ImageArr
    
    If Not isInitialized Then MW_Initialize
    
    ImageNameBase = GetReg("ImageNamePreFix") & DocInfo.ArticleName & "_"
    
    Select Case DocArea
        Case 0
            Set MyWorkArea = ActiveDocument
        Case 1
            Set MyWorkArea = ActiveDocument.Sections(SectionNo).Headers(HeaderNo)
            MW_ChangeView 1
        Case 2
            Set MyWorkArea = ActiveDocument.Sections(SectionNo).Footers(HeaderNo)
            MW_ChangeView 1
    End Select
    
    MW_Statusbar False, "converting images"
    
    'Find all shapes and convert to InlineShapes, so we know its position
    'Well, the outcome is unpredictable, messing up the picture position
    'For Each myS In MyWorkArea.Shapes
    '    If myS.Type = msoGroup Then myS.Ungroup
    'Next myS
    p = 0
    Do
    RepeatLoop = False
    For Each MyS In MyWorkArea.Shapes
    
        MyS.Select
        DoEvents
        'On Error Resume Next 'Sometimes it can not be converted, but no crash
        Select Case MyS.Type
            Case msoGroup
                'can not convert
                If MW_WordVersion <= 2000 Or 1 = 1 Then
                    'We do not even have a picture number
                    Selection.Move
                    p = p + 1
                    ImageTxt = MW_CheckFileName(ImageNameBase & "Group_" & Format(p, "00"))
                    Selection.InsertBefore vbCr & "[[Image:" & ImageTxt & "." & ImageTagFormat & "]]" & vbCr
                    'Selection.InsertAfter vbCr & GetReg("UnableToConvertMarker") & ": Grouped pictures can not be converted." & vbCr
                    RepeatLoop = True
                    MyS.Delete
                    Exit For
                Else
                    'Try to copy picture to png
                    'Word 2002 can copy as png without problem, Word 2000 will copy only one image of the group.
                    Application.ScreenRefresh
                    Selection.Copy
                    'Documents.Add.Content.PasteSpecial Link:=False, DataType:=14, DisplayAsIcon:=False   'DataType 14 = PNG
                    MyS.Delete
                    'Paste as PNG
                    Selection.Move
                    'This only works in Debug - no clue
                    Selection.PasteSpecial Link:=False, DataType:=14, DisplayAsIcon:=False   'DataType 14 = PNG
                    Debug.Print "ok!"
                    RepeatLoop = True
                    Exit For
                    'Later on we get an non critical error as word pastes two pictures, seems to be a word bug
                    
                End If
                
            Case msoAutoShape, 20 'msoCanvas
                'can not convert
                Selection.InsertBefore txtNotConvert
                p = p + 1
                ImageTxt = MW_CheckFileName(ImageNameBase & "Shape_" & Format(p, "00"))
                Selection.InsertAfter vbCr & "[[Image:" & ImageTxt & "." & ImageTagFormat & "]]" & vbCr
                If MyS.AlternativeText <> "" Then
                    MyS.ConvertToFrame
                    FrameCreated = True
                End If
                
            Case msoLine
                'convert to wiki divider line
                MyS.Select
                MyS.Delete
                Selection.InsertAfter vbCr & "----" & vbCr 'wiki line
                Selection.Range.Style = wdStyleNormal
                
            Case msoTextBox
                'convert to framed paragraph
                MyS.ConvertToFrame
                FrameCreated = True
                
            Case Else
                MyS.ConvertToInlineShape
        End Select
        If Err.Number <> 0 Then
            Debug.Print "Images: " & Err.Number, Err.Description
            Err.Clear
        End If
        Selection.InsertAfter vbCr & vbCr
        DoEvents
        On Error GoTo Err_MediaWikiExtract_ImagesPhotoEditor
        
    Next MyS
    Loop Until RepeatLoop = False
    DoEvents

    'if header or footer only convert to inline shapes
    Select Case DocArea
        Case 1, 2
            MW_ChangeView 4
            Exit Sub
    End Select
    
    'Textboxes will get a nice frame at least
    If FrameCreated And convertImagesOnly = False Then MediaWikiConvert_Tables
    
    'Convert all InlineShapes
    For Each MyIS In MyWorkArea.InlineShapes
    
        'identify Image-Type, it could be an inline-document
        MyIS.Select
        IconPath = ""
        IconClassType = ""
        OLabel = ""
        Select Case MyIS.Type
            Case wdInlineShapePicture
                PicType = "Pic" 'default or unidentified
                If 1 = 2 Then 'Test only
                    MyIS.Select
                    MyIS.ScaleHeight = 100
                    MyIS.ScaleWidth = 100
                    MyIS.Range.Copy
                    'Selection.Copy 'to clipboard
                    'savefile
                    Shell "E:\Projekte\Diverse Projekte\Clipboard2BMP\Clipboard2BMP.exe " & "E:\Temp\Test_" & Format(Time(), "hhmmss") & ".bmp", vbHide
                    DoEvents
                    'SavePicture , "e:\temp\Test_" & Format(Time(), "hhmmss") & ".bmp"
                End If
                
            Case wdInlineShapeHorizontalLine
                PicType = "Line"
                'and ?
                MyIS.Select
                MyIS.Range.InsertAfter vbCr & "----" & vbCr 'wiki line
                MyIS.Range.Style = wdStyleNormal
                MyIS.Delete
            
            
            Case wdInlineShapeOLEControlObject
                If DebugMode Then Stop
                PicType = "do not convert"
                MyIS.Select
                MyIS.Range.InsertAfter vbCr & "#$$$#< " & ConverterPrgTitle & " found a non convertable object." & vbCr
                MyIS.Range.Style = wdStyleNormal
                MyIS.Delete
                
            Case Else
                PicType = "Pic" 'default or unidentified
                
                If Not MyIS.OLEFormat Is Nothing Then
                    'in some cases the icon does not have this information
                    IconClassType = Replace(MyIS.OLEFormat.ClassType, ".", " ")
                    If MyIS.OLEFormat.DisplayAsIcon Then
                        PicType = "Icon"
                        'IconPath = myIS.OLEFormat.IconPath
                        OLabel = MyIS.OLEFormat.IconLabel
                        If OLabel = "" Then OLabel = IconClassType & " Object"
                    End If
                    OLabel = Replace(OLabel, """", "")
                    If OLabel = "" Then OLabel = "no label"
                    
                    'discontinued usage V0.6
                    'If InStr(1, IconClassType, "picture", vbTextCompare) = 0 Then 'if picture then it is ok
                        'Other filetypes must be added
                        'If InStr(1, IconClassType, "sheet", vbTextCompare) > 0 Then
                        '    PicType = "File"
                        'ElseIf InStr(1, IconClassType, "document", vbTextCompare) > 0 Then
                        '    PicType = "File"
                        'End If
                    'End If
                End If
        End Select
        
        
        Select Case PicType
        
            Case "File"
                'not in use anymore
            
                'Insert Wiki-Link
                MyIS.Select
                MyIS.Range.InsertAfter vbCr & "[[" & ImageNameBase & IconClassType & "]]"
            
                'Delete picture from document
                Selection.Delete
                Selection.MoveRight wdCharacter, 1, wdExtend
                If Selection.Text = " " Then Selection.Collapse: Selection.Delete
                
            Case "Pic", "Icon"
    
                'Insert Image-Tag
                
                PicNo = PicNo + 1
                ImagePathName = IIf(GetReg("ImagePath") <> "", GetReg("ImagePath"), ActiveDocument.Path)
                If PicType = "Icon" Then
                    IconNo = IconNo + 1
                    If ImageIconOnlyType Then
                        'General usage
                        ImageTxt = MW_CheckFileName("Icon " & IconClassType)
                    Else
                        'Document specific usage
                        ImageTxt = MW_CheckFileName("Icon " & ImageNameBase & Format(IconNo, "00"))
                    End If
                Else
                    ImageTxt = MW_CheckFileName(ImageNameBase & Format(PicNo - IconNo, "00"))
                End If
                ImagePathName = FormatPath(ImagePathName) & ImageTxt & ".bmp"
                
                'Copy to ClipBoard and save as bitmap
                Dim OrgSize#
                MyIS.Select
                OrgSize = PointsToPixels(MyIS.Width, False)
                'Fill white or some backgrounds will be black
                If MyIS.Fill.Visible = msoFalse Then
                    MyIS.Fill.Solid
                    MyIS.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    MyIS.Fill.Transparency = 0#
                End If
                DoEvents
                
                'Insert [[Image:xxx]]
                'If it is an Icon, then add a Wiki-Link
                If PicType = "Icon" Then MyIS.Range.InsertAfter vbCr & "[[" & ImageNameBase & OLabel & "]]" & vbCr
                Select Case MyIS.Type '= wdInlineShapeLinkedOLEObject
                    'Case wdInlineShapeEmbeddedOLEObject
                        'no size information --> full size
                        'word bug
                        'myIS.Range.InsertAfter "[[Image:" & ImageTxt & "." & ImageTagFormat & "]]" & vbCr
                   Case Else
                        If ((MyIS.ScaleWidth < 95 Or MyIS.ScaleWidth > 105) And GetReg("ImagePixelSize")) Or (GetReg("ImageMaxPixel") And MW_ScaleMaxOK(MyIS, GetReg("ImageMaxPixelSize")) = False) Then
                            'give pixels
                            MyIS.Range.InsertAfter "[[Image:" & ImageTxt & "." & ImageTagFormat & "|" & Round(OrgSize) & "px]]" & vbCr
                        Else
                            'no size information --> full size
                            MyIS.Range.InsertAfter "[[Image:" & ImageTxt & "." & ImageTagFormat & "]]" & vbCr
                        End If
                End Select
                'if myIS has text flox, do not know how to check
                MyIS.Range.InsertAfter vbCr
                
                'Check if file already exist
                ImageExtractionInt = GetReg("ImageExtraction")
                If GetReg("ImageConvertCheckFileExists") And ImageExtractionInt Then
                    If ImageSavePNG Then
                        ImageExtractionInt = Not FileExists(Left$(ImagePathName, Len(ImagePathName) - 3) & "png")
                    ElseIf ImageSaveGIF Then
                        ImageExtractionInt = Not FileExists(Left$(ImagePathName, Len(ImagePathName) - 3) & "gif")
                    ElseIf ImageSaveJPG Then
                        ImageExtractionInt = Not FileExists(Left$(ImagePathName, Len(ImagePathName) - 3) & "jpg")
                    ElseIf ImageSaveBMP Then
                        ImageExtractionInt = Not FileExists(Left$(ImagePathName, Len(ImagePathName) - 3) & "bmp")
                    End If
                End If
                
                'Image conversion
                If ImageExtractionInt Then
                    'Put image in clipboard
                    'If we have Word 2002 and above, we can copy and paste in Word as PNG. _
                        This allows us to copy even problematic images for sure in PhotoEditor and use SnagIt
                    DoEvents
                    
                    If MW_WordVersion >= 2002 And MW_ScaleMaxOK(MyIS) Then
                        'convert to png to make save copy, but only if picture can be copied, because it is small enough
                        'otherwise a paste as will work better, as it can copy all sizes
                        MW_ScaleMax MyIS
                        MyIS.Range.Copy
                        Documents.Add DocumentType:=wdNewBlankDocument
                        DoEvents
                        'Paste as PNG
                        Selection.PasteSpecial Link:=False, DataType:=14, DisplayAsIcon:=False   'DataType 14 = PNG
                        DoEvents
                        If ActiveDocument.Shapes.Count > 0 Then
                            ActiveDocument.Shapes(1).Select
                            Selection.Copy
                        ElseIf ActiveDocument.InlineShapes.Count > 0 Then
                            ActiveDocument.InlineShapes(1).Range.Copy
                        End If
                        DoEvents
                        ActiveDocument.Close wdDoNotSaveChanges
                    Else
                        MyIS.Range.Copy 'AsPicture '?
                        DoEvents
                    End If
                    
                    'save clipboard image as file
                    Select Case GetReg("ImageConverter")
                    
                        Case "MSPhotoEditor"
                        
                            If GetReg("ImagePastePixel") = False Then MW_PhotoEditor_Convert "Paste", ImagePathName
                            
                            'Strange World, some pictures can't be pasted in photo editor
                            'e.g. PowerPoint slides
                            'so we check the file
                            
                            'check if image was successfully saved -> if not save as BMP
                            If ImageSavePNG Or ImageSaveJPG Then
                                If Not (FileExists(MW_ImagePathName(ImagePathName, "png")) Or FileExists(MW_ImagePathName(ImagePathName, "jpg")) Or GetReg("ImageConvertCheckFileExists") = False) Then
                                    'make it otherwise, better then nothing
                                    If MW_WordVersion >= 2002 Then
                                        MW_ScaleMax MyIS
                                        MyIS.Range.Copy
                                        Documents.Add DocumentType:=wdNewBlankDocument
                                        'usually we would copy as PNG straight, but again a Word bug
                                        'Selection.PasteAndFormat (wdPasteDefault)
                                        'DoEvents
                                        'ActiveDocument.InlineShapes(1).Select
                                        'Selection.Copy
                                        'Selection.ShapeRange.Delete
                                        DoEvents
                                        'Paste as PNG
                                        Selection.PasteSpecial Link:=False, DataType:=14, DisplayAsIcon:=False   'DataType 14 = PNG
                                        DoEvents
                                        If ActiveDocument.Shapes.Count > 0 Then
                                            ActiveDocument.Shapes(1).Select
                                            Selection.Copy
                                        ElseIf ActiveDocument.InlineShapes.Count > 0 Then
                                            ActiveDocument.InlineShapes(1).Range.Copy
                                        End If
                                        DoEvents
                                        ActiveDocument.Close wdDoNotSaveChanges
                                    Else
                                        MW_ScaleMax MyIS
                                        MyIS.Range.Copy 'AsPicture '?
                                        DoEvents
                                    End If
                                    
                                    MW_PhotoEditor_Convert "PastePixel", ImagePathName, , PointsToPixels(MyIS.Width), PointsToPixels(MyIS.Height)
                                End If
                            End If
                            
                        'Case "SnagIt"
                            'MW_SnagIt_Clipboard_to_File ImagePathName
                        
                            
                        'Case Else
                            'User Converter
                            'Const ConverterPath$ = "E:\Projekte\Word2Wiki\b2p152w\BMP2PNG.exe"
                            'Shell ConverterPath & " " & GetShortPath(imageArr(i))
                            'convert in MediaWiki_ConvertImages
                    End Select
                End If
                
                If ImageExtractionInt Or GetReg("ImageReload") Then
                    'Save Imagename in ImageArray for MediaWikiExtract_ImagesPhotoEditor
                    On Error Resume Next
                    p = UBound(ImageArr)
                    If Err.Number <> 0 Then ReDim ImageArr(1 To 1): p = 1: Err.Clear
                    On Error GoTo Err_MediaWikiExtract_ImagesPhotoEditor
                    If ImageArr(p) <> "" Then
                        p = p + 1
                        ReDim Preserve ImageArr(1 To p)
                    End If
                    
                    ImageArr(p) = Left$(ImagePathName, Len(ImagePathName) - 3) & ImageTagFormat
                End If
                
                'Delete picture from document (replaced by link)
                MyIS.Select
                Selection.Delete
                Selection.MoveRight wdCharacter, 1, wdExtend
                If Selection.Text = " " Then Selection.Collapse: Selection.Delete
                
        End Select
        
Err_Continue_5825:
        
    Next MyIS

Exit_MediaWikiExtract_ImagesPhotoEditor:
    MW_ChangeView 0
    'Close Photo Editor
    MW_CloseProgramm EditorTitle
    'If AppActivatePlus(EditorTitle, False) Then DoEvents: SendKeys "%{F4}", True 'End Photo Editor
    DoEvents
    Exit Sub

Err_MediaWikiExtract_ImagesPhotoEditor:
    Select Case Err.Number
        Case 4605
            MsgBox "You encountered a MS-Word Bug. Delete the marked image from the document and try again", vbCritical, ConverterPrgTitle
            End
        Case 5825 'object was deleted; by word, no clue
            Err.Clear
            Selection.InsertAfter "[[Image:## Error Converting ##]]"
            Resume Err_Continue_5825
        Case Else
            DisplayError "MediaWikiExtract_ImagesPhotoEditor"
            Resume Exit_MediaWikiExtract_ImagesPhotoEditor
            'Resume Next
    End Select
End Sub

Private Sub MediaWikiConvert_Indention()
' -------------------------------------------------------------------
' Function: indents text if first line of paragraph is indented
'           usally strange results
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: June 10, 2006
' -------------------------------------------------------------------

    Dim c&
    Dim pg As Paragraph

    'look for FirstLineIndent
    If convertFirstLineIndent Then
        For Each pg In ActiveDocument.Paragraphs
            If pg.Range.ParagraphFormat.FirstLineIndent > DefaultIndent Then
                pg.Range.Select
                If pg.Range.Characters.Count > 1 Then pg.Range.InsertBefore ":"
            End If
        Next
    End If

End Sub

Private Sub MediaWikiConvert_IndentionTab()
' -------------------------------------------------------------------
' Function: indents text if first char is a tab
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: June 10, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_IndentionTab

    Dim pg As Paragraph
    Dim TabInLineBefore&, LastTabAlign&, p&
    Dim IsIndention As Boolean
    
    'Must run before text formatting took place
        
    'Delete all tabs at end of manual line break (or word messes up the table)
    MW_ReplaceString "^t" & WordNewLine, WordNewLine, True
    'Delete all tabs if only char in paragraph
    MW_ReplaceString WordParagraph & "^t" & WordParagraph, WordParagraph & WordParagraph, True
    
    MW_Statusbar True, "looking for line indention"
    'Paragraph must have a blank at the beginning then
    For Each pg In ActiveDocument.Paragraphs
        IsIndention = False
        TabInLineBefore = TabInLineBefore - 1
        p = InStr(1, pg.Range.Text, vbTab, vbBinaryCompare)
        If p > 0 Then
            'only indention?
            If p = 1 Then IsIndention = InStr(2, pg.Range.Text, vbTab, vbBinaryCompare) = 0
            If IsIndention = False Or TabInLineBefore = 1 Then 'do we continue a table?
                TabInLineBefore = 2
                'Exit Do
            Else
                pg.Range.Characters.First = ":"
                'MW_ClearFormatting pg.Range.Characters.First, True
                pg.Range.Characters.First.Font.Reset
            End If
        End If
    Next
        
    Exit Sub
    
    'look for tabs at the beginning of a paragraph
    Dim c&
    Dim rg As Range
    For Each pg In ActiveDocument.Paragraphs
        c = 0
        Do While Left$(pg.Range.Text, 1) = vbTab
            Set rg = pg.Range
            rg.Collapse
            rg.MoveEnd
            rg.Delete
            c = c + 1
        Loop
        Do While c > 0
            pg.Range.InsertBefore "##TAB##"
            c = c - 1
        Loop
        'we need to to this, because the formatting will move the : later
    Next

Exit_MediaWikiConvert_IndentionTab:
    Exit Sub

Err_MediaWikiConvert_IndentionTab:
    DisplayError "MediaWikiConvert_IndentionTab"
    Resume Exit_MediaWikiConvert_IndentionTab
End Sub

Private Sub MediaWikiConvert_Lists()
' -------------------------------------------------------------------
' Function: converts lists and numbering (Aufzählungen)
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
    'Lists are a bit problematik in wiki, if continued after a blank line
    'see http://meta.wikimedia.org/wiki/Help:List
    'Therefore we need to work with html-tags
    'ToDo: Will not resume numbers if line break inbetween
    'ToDo: Will not work correctly if list in list

' Input:
'
' returns: nothing
'
' released: Nov. 06, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_Lists

    Dim i&, c&, cL&, Level&, p&
    Dim pg As Paragraph
    Dim rg As Range
   
    'replace paragraphs with HTML-breaks to allow continued lists
    cL = ActiveDocument.ListParagraphs.Count
    For Each pg In ActiveDocument.ListParagraphs
        With pg.Range
           
            c = c + 1
            If c Mod 10 = 0 Then MW_Statusbar True, "converting lists " & c & " of " & cL, False
           
            '.Select 'only for testing
            'Debug.Print .ListFormat.ListLevelNumber, .ListFormat.ListValue, Left$(.Text, 20)
           
            .InsertBefore " " 'no need, but looks better in MediaWiki editor
           
            'replace manual page breaks
            Set rg = pg.Range
            '.Select
            'MW_ReplaceString WordNewLine, "<br>", , wdFindStop
            p = InStr(1, rg, Chr(11)) 'fast check
            If p > 0 Then
                p = rg.MoveStartUntil(Chr(11)) 'slow, but correct position
                Do While p > 0 And 1 = 2
                    rg.Collapse
                    rg.MoveEnd , 1
                    rg.Delete
                    rg.InsertAfter "<br>"
                    p = rg.MoveStartUntil(Chr(11))
                Loop
            End If
            Level = .ListFormat.ListLevelNumber
            If GetReg("ListNumbersManual") Then Level = 1
            For i = 1 To Level
                'Debug.Print "i:" & i, .ListFormat.CountNumberedItems
                If .ListFormat.ListType = wdListBullet Then
                    .InsertBefore "*"
                Else
                    If GetReg("ListNumbersManual") Then
                        'Some people make lists with symbols, we try to detect them
                        'Could not find a way to see, if ListString is in type Symbol or WingDings
                        If Val(.ListFormat.ListString) > 0 Then
                            .InsertBefore .ListFormat.ListString
                        Else
                            'insert bullet if not numerich
                            .InsertBefore "*"
                        End If
                    Else
                        .InsertBefore "#"
                    End If
                End If
            Next i
            .ListFormat.RemoveNumbers

        End With
    Next pg

Exit_MediaWikiConvert_Lists:
    Exit Sub

Err_MediaWikiConvert_Lists:
    DisplayError "MediaWikiConvert_Lists"
    Resume Exit_MediaWikiConvert_Lists
End Sub

Private Sub MediaWikiConvert_Paragraphs()
' -------------------------------------------------------------------
' Function: converts Paragraphs for better reading in MediaWiki. Otherwise it will resume within the line.
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: June 12, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_Paragraphs

    Dim txt$
    Dim pg As Paragraph
    Dim lH&, jump&, p&
    Dim InTable&
   
    lH = Len(HeaderFirstLevel)
   
    If NewParagraphWithBR Then
    
        'code not tested!!!
        
        'add <br> to all manual line breaks
        MW_ReplaceString WordNewLine, "<br>"
        'add <br> to all paragraphs
        MW_ReplaceString WordParagraph, "<br>" & WordParagraph
       
        'That is too much, so now eliminate all wrong <br>
       
        'Headers
        MW_ReplaceString HeaderFirstLevel & "<br>" & WordParagraph, HeaderFirstLevel & WordParagraph
       
        'Double <br> will be recognized correctly as new line
        MW_ReplaceString "<br>" & WordParagraph & "<br>" & WordParagraph, WordParagraph & WordParagraph
        MW_ReplaceString "<br>" & WordParagraph & "<br>" & WordParagraph, WordParagraph & WordParagraph
        MW_ReplaceString "<br>" & WordParagraph & "<br>" & WordParagraph, WordParagraph & WordParagraph
       
       
        'Further unused coding to clean up
        For Each pg In ActiveDocument.Paragraphs
            With pg
           
                txt = .Range.Text
           
            End With
        Next
       
    Else
        'use two lines
       
        'add <br> to all manual line breaks
        MW_ReplaceString WordNewLine, "<br>"
       
        'Add empty line at document end to prevent error
        Dim rg As Range
        Set rg = ActiveDocument.Range
        rg.InsertAfter vbCr
       
        For Each pg In ActiveDocument.Paragraphs
            With pg
           
                If jump = 0 Then
                    If InStr(1, .Range.Text, "{|") > 0 Then InTable = InTable + 1
                    p = InStr(1, .Range.Text, "|}")
                    If InStr(1, .Range.Text, "|}") > 0 Then
                        InTable = InTable - 1
                        If InTable > 0 And 1 = 2 Then
                            'clean up an error in nested tables
                            .Range.Select
                            Selection.Collapse wdCollapseEnd
                            Do While Selection.Text = vbCr
                                Selection.MoveRight
                            Loop
                            If Left$(Selection.Text, 1) <> "|" Then
                                Selection.InsertBefore "|"
                            End If
                        End If
                    End If
               
                    If InTable = 0 Then
                        If Asc(.Range.Text) = 13 Then
                            'Paragraph empty?
                            'nothing
                            'goto next paragraph
                        ElseIf Left$(.Range.Text, 1) = "*" Or Left$(.Range.Text, 1) = "#" Then
                            'List?
                            'nothing
                            'goto next paragraph
                        ElseIf Left$(.Range.Text, lH) = HeaderFirstLevel Then
                            'Header?
                            'nothing
                            'jump = 1
                            'goto next paragraph
                        ElseIf Asc(.Next.Range.Text) = 13 Then
                            'Next Paragraph empty?
                            'nothing
                            'goto next paragraph
                        ElseIf right$(.Range.Text, 5) = "<br>" & vbCr Then
                            'manual line break?
                            'nothing
                            'goto next paragraph
                        Else
                            .Range.InsertAfter vbCr
                            txt = .Range.Text 'Debug Info
                        End If
                    End If
               
                Else
                    jump = jump - 1
                End If
           
            End With
        Next
    End If 'NewParagraphWithBR
    
Exit_MediaWikiConvert_Paragraphs:
    Exit Sub

Err_MediaWikiConvert_Paragraphs:
    DisplayError "MediaWikiConvert_Paragraphs"
    Resume Exit_MediaWikiConvert_Paragraphs
End Sub

Private Sub MediaWikiConvert_Prepare()
' -------------------------------------------------------------------
' Function: document and word preparation before converting
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: June 12, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_Prepare
'On Error Resume Next
    
    Dim pg As Paragraph
    Dim rg As Range
    Dim c&, p&
    
    ' history of changes must be turned off
    ActiveDocument.TrackRevisions = False
    ActiveDocument.AcceptAllRevisions
    'turn off automatic hyphenation (Silbentrennung)
    'prevents unwanted "-" within words
    ActiveDocument.AutoHyphenation = False
    Err.Clear

    'Store user options
    If MW_WordVersion = 2003 Then MW_SetOptions_2003 True
    SetReg "Z_finished", False
    
    MW_ChangeView 0

    'Clear the normal style
    With ActiveDocument.Styles(wdStyleNormal).Font
        '.Font.Reset 'does not work here
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .Strikethrough = False
        .Subscript = False
        .Superscript = False
        .ColorIndex = wdAuto
    End With
    
    'Get some standard values
    With ActiveDocument.Styles(wdStyleNormal)
        DefaultFontSize = .Font.Size
        DefaultIndent = .ParagraphFormat.LeftIndent
        DefaultIndent = .ParagraphFormat.FirstLineIndent
    End With
    
'--- Start working in document ---
    
    'remove empty paragraphs at begin of document
    Set pg = ActiveDocument.Paragraphs(1)
    Do While Not pg.Next Is Nothing
        If pg.Range.Text = vbCr Then
            If pg.Range.InlineShapes.Count > 0 Then Exit Do 'ups, a picture
            On Error Resume Next
            If pg.Range.ShapeRange.Count > 0 Then
                If Err.Number = 0 Then Exit Do               'ups, a picture
            End If
            Err.Clear
            On Error GoTo Err_MediaWikiConvert_Prepare
            pg.Range.Delete
        Else
            Exit Do
        End If
    Loop
    'Now, if we might have some problems, if we are in a table
    pg.Range.Select
    If Selection.Information(wdWithInTable) Then Selection.SplitTable
    
    MW_InsertPageHeaders
    MW_ChangeView 0
    
    ' Delete all manual pagebreaks, must be at beginning of macro (otherwise problems with headers)
    ' Done in Headers

    'remove blanks at the end of paragraph
    MW_Statusbar True, "removing blanks at end of line..."
    MW_ReplaceString " " & WordParagraph, WordParagraph, True, , 5 'True: In some cases replacement can not be done.
    
    'Now find the missed blanks
    c = 0
    Selection.GoTo wdParagraph, 1
    Do
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = " " & WordParagraph
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop 'wdFindContinue means whole document
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            p = .Execute 'Replace:=wdReplaceAll 'In Word 2003 the active document will change!!! Bug.
            If p Then
                c = c + 1
                If Selection.Characters.First = " " Then
                    'Selection.Characters.First = "-"
                    Selection.Characters.First.Delete
                End If
            End If
            
        End With
    Loop Until p = False Or c = 1000
    
    'Store table information before converting changes sizes
    MW_TableInfo
    
Exit_MediaWikiConvert_Prepare:
    MW_Statusbar False
    Exit Sub

Err_MediaWikiConvert_Prepare:
    DisplayError "MediaWikiConvert_Prepare"
    Resume Exit_MediaWikiConvert_Prepare
End Sub

Private Sub MediaWikiConvert_Tables()
' -------------------------------------------------------------------
' Function: converts all tables of the document, even nested (table in cell of other table)
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: June 16, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_Tables
    
    Const DeleteCellMarker$ = "$$DeleteCell$$"
    Dim thisTable As Table
    Dim pg As Paragraph, rg As Range, aCell As Cell
    Dim sStart&, sEnd&, i&, pc&, CellBGcolor&
   
    If Not isInitialized Then MW_Initialize: MW_TableInfo
       
    MW_Statusbar True, "converting tables..."
       
    'convert all normal tables to wiki syntax
    For i = ActiveDocument.Tables.Count To 1 Step -1
        MW_Convert_Table ActiveDocument.Tables(i), , i
        'DoEvents
    Next
    
    'convert framed paragraphs to tables
    Set pg = ActiveDocument.Paragraphs.First
    i = 0
    Do While Not pg Is Nothing 'i < ActiveDocument.Paragraphs.Count
        i = i + 1
        If i Mod 100 = 0 Then MW_Statusbar True, "checking for framed paragraphs " & i & " of " & ActiveDocument.Paragraphs.Count
        'pg.Range.Select 'just for testing
        If pg.Borders.Enable <> 0 Then
            If sStart = 0 Then If pg.Range.Information(wdWithInTable) = False Then _
                    sStart = i: CellBGcolor = pg.Shading.BackgroundPatternColor
        End If
        If sStart > 0 Then
            If pg.Borders.Enable = 0 Or pg.Range.Information(wdWithInTable) Or pg.Shading.BackgroundPatternColor <> CellBGcolor Then
                If sStart > 0 Then
                    sEnd = i - 1
                    Set rg = ActiveDocument.Range( _
                        Start:=ActiveDocument.Paragraphs(sStart).Range.Start, _
                        End:=ActiveDocument.Paragraphs(sEnd).Range.End)
                    rg.Borders.Enable = False
                    rg.Shading.BackgroundPatternColor = wdColorAutomatic
                    rg.ConvertToTable Separator:=wdSeparateByParagraphs, NumColumns:=1
                    For Each aCell In rg.Cells
                        aCell.Range.Shading.BackgroundPatternColor = CellBGcolor
                    Next
                    'make sure pg and i are the same
                    i = i - 1
                    Set pg = ActiveDocument.Paragraphs(i)
                    sStart = 0
                End If
            End If
        End If
        Set pg = pg.Next
    Loop
    For Each thisTable In ActiveDocument.Tables
        MW_Convert_Table thisTable, TableTemplateParagraphFrame
        DoEvents
    Next thisTable
    
    'Find paragraphs with tabs, which probably look like tables
    MediaWikiConvert_TabTables
    For Each thisTable In ActiveDocument.Tables
        MW_Convert_Table thisTable, TableTemplateNoFrame
        'DoEvents
    Next thisTable
    
    'convert colored paragraphs to tables
    sStart = 0
    i = 0
    Set pg = ActiveDocument.Paragraphs.First
    i = 0
    Do While Not pg Is Nothing 'i < ActiveDocument.Paragraphs.Count
        i = i + 1
        If i Mod 100 = 0 Then MW_Statusbar True, "checking for colored paragraphs " & i & " of " & ActiveDocument.Paragraphs.Count
        'pg.Range.Select 'just for testing
        If pg.Shading.BackgroundPatternColor <> wdColorAutomatic Then
            If sStart = 0 Then If pg.Range.Information(wdWithInTable) = False Then _
                    sStart = i: CellBGcolor = pg.Shading.BackgroundPatternColor
        End If
        If sStart > 0 Then
            If pg.Range.Information(wdWithInTable) Or pg.Shading.BackgroundPatternColor <> CellBGcolor Then
                If sStart > 0 Then
                    sEnd = i - 1
                    Set rg = ActiveDocument.Range( _
                        Start:=ActiveDocument.Paragraphs(sStart).Range.Start, _
                        End:=ActiveDocument.Paragraphs(sEnd).Range.End)
                    rg.Select
                    rg.Borders.Enable = False
                    rg.Shading.BackgroundPatternColor = wdColorAutomatic
                    rg.ConvertToTable Separator:=wdSeparateByParagraphs, NumColumns:=1
                    For Each aCell In rg.Cells
                        aCell.Range.Shading.BackgroundPatternColor = CellBGcolor
                    Next
                    'make sure pg and i are the same
                    i = i - 1
                    Set pg = ActiveDocument.Paragraphs(i)
                    sStart = 0
                End If
            End If
        End If
        Set pg = pg.Next
    Loop
    For Each thisTable In ActiveDocument.Tables
        MW_Convert_Table thisTable, TableTemplateParagraphNoFrame
        DoEvents
    Next thisTable
    
    'merged cells
    MW_ReplaceString "||" & DeleteCellMarker, ""
    MW_ReplaceString "|" & DeleteCellMarker & "|", ""
    
    'nested tables
    MW_ReplaceString "}" & WordParagraph & "||", "}" & WordParagraph & "|"
    MW_ReplaceString "||", WordParagraph & "|"

Exit_MediaWikiConvert_Tables:
    MW_Statusbar False
    Exit Sub

Err_MediaWikiConvert_Tables:
    DisplayError "MediaWikiConvert_Tables"
    Resume Exit_MediaWikiConvert_Tables
End Sub

Private Sub MediaWikiConvert_TabTables()
' -------------------------------------------------------------------
' Function: Tabs are not supported by wiki, but in word commonly used for equal spacing.
'           So we make a table without frame.
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: June 11, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiConvert_TabTables

    Dim pg As Paragraph
    Dim TabInLineBefore&, LastTabAlign&, p&
    Dim IsIndention As Boolean
    
    If convertTableWithTabs Then
        
        'Delete all tabs at end of manual line break (or word messes up the table)
        MW_ReplaceString "^t" & WordNewLine, WordNewLine, True
        'Delete all tabs if only char in paragraph
        MW_ReplaceString WordParagraph & "^t" & WordParagraph, WordParagraph & WordParagraph, True
        
        'Paragraph must have a blank at the beginning then
        For Each pg In ActiveDocument.Paragraphs
            IsIndention = False
            'Do
                TabInLineBefore = TabInLineBefore - 1
                p = InStr(1, pg.Range.Text, vbTab, vbBinaryCompare)
                If p > 0 Then
                    'only indention?
                    pg.Range.Select
                    If TabInLineBefore <= 1 Then 'do we continue a table? 'IsIndention = False Or
                        pg.Range.Select
                        
                        'Line Breaks might move to different cells, this is a word bug/feature
                        MW_ReplaceString WordNewLine, "<br>", , wdFindStop
                        
                        LastTabAlign = wdAlignTabLeft
                        If pg.TabStops.Count > 0 Then LastTabAlign = pg.TabStops(pg.TabStops.Count).Alignment
                        
                        Selection.ConvertToTable Separator:=wdSeparateByTabs, AutoFitBehavior:=wdAutoFitFixed
                        If LastTabAlign = wdAlignTabRight Then
                            Selection.Cells(Selection.Cells.Count).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                        End If
                        TabInLineBefore = 2
                    End If
                End If
            'Loop Until p <= 1
        Next
        'replace remaining tabs
        MW_ReplaceString "^t", String(TabBlanksNo, " ")
        'MediaWikiConvert_Tables 'again
    End If
    
    'Delete all tabs at end of line
    MW_ReplaceString "^t" & WordParagraph, WordParagraph, True

Exit_MediaWikiConvert_TabTables:
    Exit Sub

Err_MediaWikiConvert_TabTables:
    DisplayError "MediaWikiConvert_TabTables"
    Resume Exit_MediaWikiConvert_TabTables
End Sub

Private Sub MediaWikiConvert_TextFormat()

    MW_FontFormat ("Hidden")
    MW_FontFormat ("Bold")
    MW_FontFormat ("Italic")
    MW_FontFormat ("Underline")
    MW_FontFormat ("StrikeThrough")
    MW_FontFormat ("Superscript")
    MW_FontFormat ("Subscript")
    MW_FontFormat ("FontSize")
    MediaWikiConvert_FontColors

End Sub

Public Sub MediaWikiImageUpload(Optional ExternalCall As Boolean = False)
' -------------------------------------------------------------------
' Function: The images of the document will be uploaded to your wiki
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: Nov 7, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiImageUpload
    
    Dim lRet&, i&, PauseUploadAfterXImages&
    
    PauseUploadAfterXImages = GetReg("PauseUploadAfterXImages")

    If GetReg("ImageUploadAuto") Or ExternalCall Then
    
        'Any images?
        On Error Resume Next
        i = UBound(ImageArr)
        If Err.Number > 0 Then
            Err.Clear
            On Error GoTo Err_MediaWikiImageUpload
            Exit Sub
        End If
        On Error GoTo Err_MediaWikiImageUpload
        
        For i = 1 To UBound(ImageArr)
            lRet = MW_ImageUpload_File(ImageArr(i))
            If lRet = -1 Then Exit For
            If i Mod PauseUploadAfterXImages = 0 Then
                Application.Activate
                lRet = MsgBox("Pause. Check Browser", vbOKCancel)
                If lRet <> vbOK Then Exit For
            End If
        Next i
    
        Application.Activate
    
    End If

Exit_MediaWikiImageUpload:
    Exit Sub

Err_MediaWikiImageUpload:
    DisplayError "MediaWikiImageUpload"
    Resume Exit_MediaWikiImageUpload
End Sub

Private Sub MediaWikiOpen()
'opens the page in your wiki

    If WikiOpenPage And GetReg("isCustomized") Then
        'Open the new document in browser
        IExplorer MW_SearchAddress & DocInfo.ArticleName
        'Wrong address: IExplorer MW_SearchAddress & DocInfo.ArticleName & "&action=edit"
        Sleep 1000
    End If
    
    'Finished
    If WikiOpenPage And GetReg("isCustomized") Then
        If AppActivatePlus(MW_CheckFileNameTitle(DocInfo.ArticleName) & " -", False, , "Microsoft Word") Then
            'we could implement some coding to open the article in edit mode
            'this would need another wiki link, since the alias function does not work
        Else
            'open the search
            AppActivatePlus GetReg("WikiSearchTitle"), False
        End If
    Else
        MsgBox Msg_Finished, vbInformation, ConverterPrgTitle
    End If

End Sub

Private Sub MediaWikiReplaceQuotes()
    ' Replace all smart quotes with their dumb equivalents

    Dim quotes As Boolean

    quotes = Options.AutoFormatAsYouTypeReplaceQuotes
    Options.AutoFormatAsYouTypeReplaceQuotes = False

    MW_ReplaceString ChrW(8220), """"
    MW_ReplaceString ChrW(8221), """"
    MW_ReplaceString "‘", "'"
    MW_ReplaceString "’", "'"

    Options.AutoFormatAsYouTypeReplaceQuotes = quotes

End Sub

Private Function MW_CheckFileName(FileName$) As String
' -------------------------------------------------------------------
' Function: Some wikis have a problem with special characters in the file name
'           Trying to catch the most common characters
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: FilePathName
'
' returns: converted FilePathName
'
' released: June 04, 2006
' -------------------------------------------------------------------
On Error Resume Next

    Dim FPath$, FName
    
    FPath = GetFilePath(FileName)
    FName = GetFileName(FileName)

    FName = Replace(FName, " ", "_")
    FName = Replace(FName, "ß", "ss")
    FName = Replace(FName, "ä", "ae", , , vbBinaryCompare)
    FName = Replace(FName, "Ä", "AE", , , vbBinaryCompare)
    FName = Replace(FName, "ö", "oe", , , vbBinaryCompare)
    FName = Replace(FName, "Ö", "OE", , , vbBinaryCompare)
    FName = Replace(FName, "ü", "ue", , , vbBinaryCompare)
    FName = Replace(FName, "Ü", "UE", , , vbBinaryCompare)
    
    'We do not allow some special characters in the filename
    FName = Replace(FName, ".", "_")
    FName = Replace(FName, ":", "_")
    FName = Replace(FName, "*", "_")
    FName = Replace(FName, "/", "_")
    FName = Replace(FName, "\", "_")
    FName = Replace(FName, "?", "_")
    FName = Replace(FName, """", "_")
    FName = Replace(FName, "<", "_")
    FName = Replace(FName, ">", "_")
    FName = Replace(FName, "|", "_")
    
    MW_CheckFileName = IIf(FPath = "", FName, FormatPath(FPath) & FName)

End Function

Private Function MW_CheckFileNameTitle(FileName$) As String
' -------------------------------------------------------------------
' Function: Trying to evalutate the wiki title, which in return hat blanks instead of _
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: FilePathName
'
' returns: converted FilePathName
'
' released: June 04, 2006
' -------------------------------------------------------------------
On Error Resume Next

    Dim FPath$, FName
    
    FPath = GetFilePath(FileName)
    FName = GetFileName(FileName)

    FName = Replace(FName, "_", " ")
    
    MW_CheckFileNameTitle = IIf(FPath = "", FName, FormatPath(FPath) & FName)

End Function

Private Sub MW_ClearFormatting(rg As Range, Optional ResetFont As Boolean = False)
    'Clear all formats of selection
    'needed, if different formats are combined
    
    If ResetFont Then rg.Font.Reset: Exit Sub

    'Dim fSize&
    With rg.Font
        'fSize = .Size
        '.Reset 'kills size too and color too
        '.Size = fSize
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .Strikethrough = False
        .Subscript = False
        .Superscript = False
        'maybe more?
    End With

End Sub

Private Sub MW_Convert_Table(thisTable As Table, Optional ByVal TableFormat$ = "", Optional ByVal TableIndex& = 0, Optional ByVal NestedTableIndex& = 0, Optional ByVal MaxWidth#)
' -------------------------------------------------------------------
' Function: converts one table into wiki syntax.
'           Will call itself recursivly if nested tables are found.
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input:    Table object
'           TableFormat for the whole table, but not width, e.g. {{prettytable}}
'           TableIndex will get the width from TableInfoArr
'           MaxWidth will override PageWidth, used in nested tables (cell width)
'
'
' returns: nothing
'
' released: Nov. 25, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_Convert_Table
   
    Const DeleteCellMarker$ = "$$DeleteCell$$"
   
    Dim aRow As Row
    Dim aCell As Cell
    Dim r&, c&, i&
    Dim RowSpan&, mergedRow As Boolean
    Dim arrCellInfo() As CellInfoType
    Dim arrRowCellsMissing&()  'Number of cells to add to row
    Dim thisTableFormat$, thisTableInsertAfter$
    Dim thisTableWidth#, pageWidth#, thisTableMaxWidth#
    
    If TableFormat = "" Then TableFormat = TableTemplate
        
    'Mediawiki-Tables can not contain line breaks, these must be eliminated
    If WordParagraph = "" Then
        thisTable.Range.Find.Execute FindText:="^p", ReplaceWith:="<br>", Replace:=wdReplaceAll
        thisTable.Range.Find.Execute FindText:="^l", ReplaceWith:="<br>", Replace:=wdReplaceAll
    Else
        thisTable.Range.Find.Execute FindText:=WordParagraph, ReplaceWith:="<br>", Replace:=wdReplaceAll
        thisTable.Range.Find.Execute FindText:=WordNewLine, ReplaceWith:="<br>", Replace:=wdReplaceAll
    End If
    DoEvents
    
    MW_Statusbar True, "converting table No. " & TableIndex
    
    With ActiveDocument.PageSetup
        thisTableMaxWidth = .pageWidth - .RightMargin - .LeftMargin
        '.LeftMargin = CentimetersToPoints(2)
        '.RightMargin = CentimetersToPoints(4)
        '.PageWidth = CentimetersToPoints(21)
    End With
    If MaxWidth > 0 Then thisTableMaxWidth = MaxWidth
    
    With thisTable
        
        thisTableFormat = TableFormat
        thisTableInsertAfter = ""
        
        'table width
        'set table to width=100% if AutoWidth in word
        Select Case .PreferredWidthType And &HF 'WordBug = .PreferredWidthType
            Case wdPreferredWidthAuto, wdPreferredWidthPoints
                'in some cases the table is not full width, we need to calculate the width of the tabel
                'we only check the width of the first row and assume it is the table width
                thisTableWidth = 0
                If TableIndex > 0 And NestedTableIndex >= 0 Then
                    thisTableWidth = TableInfoArr(TableIndex, NestedTableIndex).tableWidth
                    If TableInfoArr(TableIndex, NestedTableIndex).ParentCellWidth > 0 Then
                        thisTableMaxWidth = TableInfoArr(TableIndex, NestedTableIndex).ParentCellWidth
                        If NestedTableIndex > 0 Then thisTableMaxWidth = thisTableMaxWidth - 15 'some border around the cell
                    End If
                Else
                    'This does deliver the correct width only if the changes did not alter the original table width
                    For Each aCell In thisTable.Range.Cells
                        If aCell.RowIndex = 1 Then thisTableWidth = thisTableWidth + aCell.Width Else Exit For
                    Next
                    If NestedTableIndex < 0 Then thisTableWidth = thisTableMaxWidth
                End If
                If (thisTableWidth / thisTableMaxWidth) < 0.95 Then 'smaller than 95% of page width
                    thisTableFormat = thisTableFormat & " width=""" & Round(thisTableWidth / thisTableMaxWidth * 100) & "%"""
                Else
                    'full width
                    thisTableFormat = thisTableFormat & " width=""100%"""
                End If
            Case wdPreferredWidthPercent
                If .preferredWidth < 999999 Then
                    thisTableFormat = thisTableFormat & " width=""" & Round(.preferredWidth) & "%"""
                ElseIf TableInfoArr(TableIndex, NestedTableIndex).preferredWidth > 0 And TableInfoArr(TableIndex, NestedTableIndex).preferredWidth < 999999 Then
                    thisTableFormat = thisTableFormat & " width=""" & Round(TableInfoArr(TableIndex, NestedTableIndex).preferredWidth) & "%"""
                Else
                    'should not happen
                    thisTableFormat = thisTableFormat & " width=""100%"""
                End If
        End Select
        
        'alignment of table itself
        Select Case .Rows.Alignment
            Case wdAlignRowLeft
                If .Rows.WrapAroundText = True Then
                    thisTableInsertAfter = "<br clear=""all"">"                  'align = ""right""
                    thisTableFormat = thisTableFormat & " align=""left"""
                End If
            Case wdAlignRowCenter
                thisTableFormat = thisTableFormat & " align=""center"""
                If .Rows.WrapAroundText = False Then thisTableInsertAfter = "<br clear=""all"">"                  'align = ""right""
            Case wdAlignRowRight
                thisTableFormat = thisTableFormat & " align=""right"""
                If .Rows.WrapAroundText = False Then thisTableInsertAfter = "<br clear=""all"">"  'align = ""right""
        End Select
        
        'Nesting tables
        Dim X As Object
        
        For i = 1 To .Tables.Count - 100
            'recursion
            .Tables(1).Select
            Set X = Selection
            MW_Convert_Table .Tables(1)
        Next
        If .Tables.Count > 0 And 1 = 1 Then
            'we need the cell width of the cell which is hosting the table
            Dim j&
            For i = .Tables.Count To 1 Step -1
                For Each aCell In .Range.Cells
                    Do While aCell.Tables.Count > 0
                        'recursion
                        j = j + 1
                        MW_Convert_Table aCell.Tables(1), , TableIndex, IIf(NestedTableIndex > 0, -1, j), aCell.Width '- 14 'within a table we have a boarder
                    Loop
                Next
            Next
        End If
        
        'delete all tabs within the table
        'otherwise the table will get lost due to TabTables
        '#ToDo# one could make a nested table...
        thisTable.Select
        MW_ReplaceString "^t", "", , wdFindStop
        Selection.Collapse
        
        'add blank space in empty cells
        For i = 1 To thisTable.Range.Cells.Count
            If Trim$(thisTable.Range.Cells(i).Range.Text) = vbCr & Chr(7) Then
                thisTable.Range.Cells(i).Range.InsertBefore "&nbsp;"
            End If
        Next
        
        'Get info on merged cells
        If .Uniform = False Then
        
            'Find merged rows and split them
            Do
                mergedRow = False
                For i = 1 To thisTable.Range.Cells.Count
                    thisTable.Range.Cells(i).Select 'Word has a bug, this does not work with ranges
                    RowSpan = (Selection.Information(wdEndOfRangeRowNumber) - Selection.Information(wdStartOfRangeRowNumber)) + 1
                    If RowSpan > 1 Then
                        On Error Resume Next 'can happen at end of table
                        Selection.MoveLeft wdCell
                        Do
                            Err.Clear
                            Selection.Cells.Split NumRows:=RowSpan, NumColumns:=1, MergeBeforeSplit:=False
                            If Err.Number <> 0 Then RowSpan = RowSpan + 1
                        Loop Until Err.Number = 0 Or RowSpan = 100
                        If Err.Number <> 0 Then
                            Err.Clear
                            On Error GoTo Err_MW_Convert_Table
                            Exit For 'assume end of table
                        End If
                        On Error GoTo Err_MW_Convert_Table
                        mergedRow = True
                        Selection.InsertBefore "rowspan = """ & RowSpan & """|"
                        thisTable.Range.Cells(i).Select
                        Do While RowSpan > 1
                            Selection.MoveDown
                            Selection.InsertBefore DeleteCellMarker
                            RowSpan = RowSpan - 1
                        Loop
                        Exit For 'since the number of cells changed
                    End If
                Next
            Loop Until Not mergedRow
            'now we found all merged rows, but those in the last row and column! tx to Microsoft!
            'check last cell
            i = thisTable.Range.Cells.Count
            thisTable.Range.Cells(i).Select
            Selection.MoveRight
            If Not Selection.Information(wdAtEndOfRowMarker) Then
                'Now we know the last cell is merged
                RowSpan = (Selection.Information(wdEndOfRangeRowNumber) - Selection.Information(wdStartOfRangeRowNumber)) + 2
                On Error Resume Next
                Do
                    Selection.Cells.Split NumRows:=RowSpan, NumColumns:=1, MergeBeforeSplit:=False
                    If Err.Number > 0 Then
                        Err.Clear
                        RowSpan = RowSpan + 1
                        Selection.Cells.Split NumRows:=RowSpan, NumColumns:=1, MergeBeforeSplit:=False
                    End If
                Loop Until Err.Number = 0 Or RowSpan = 100
                Selection.Collapse
                Selection.InsertBefore "rowspan = """ & RowSpan & """|"
                Selection.MoveLeft wdCell
                Selection.MoveRight wdCell
                Do While RowSpan > 1
                    Selection.MoveDown
                    Selection.InsertBefore DeleteCellMarker
                    RowSpan = RowSpan - 1
                Loop
            End If
            
            'Find merged columns and split them
            'cell width
            'Word just makes the cells with a specific width
            'If we do not want that, we need to make guess, which columns are merged
            'If we use the width we have a good chance, unless the columns have different width in each row
            ReDim arrCellInfo(1 To .Rows.Count, 1 To .Columns.Count)
            For i = 1 To thisTable.Range.Cells.Count
                With thisTable.Range.Cells(i)
                    arrCellInfo(.RowIndex, .ColumnIndex).cWidth = .Width
                    arrCellInfo(.RowIndex, .ColumnIndex).cHeight = .Height
                    arrCellInfo(.RowIndex, .ColumnIndex).cIndex = i
                End With
            Next
            
            Dim SingleRowOk As Boolean
            'Do
            mergedRow = False
            For Each aRow In thisTable.Rows
                SingleRowOk = False
                If aRow.Cells.Count < thisTable.Columns.Count Then
                    'At least one cell is missing
                    For i = 1 To aRow.Cells.Count
                        'This is very simple. One could calculate the width and then determine if there are two merged cells in a row!
                        If aRow.Cells(i).Width > MW_FindNormalWidth(arrCellInfo, i) Then
                            mergedRow = True
                            aRow.Cells(i).Select
                            RowSpan = thisTable.Columns.Count - aRow.Cells.Count + 1
                            If 1 = 2 And RowSpan < thisTable.Columns.Count Then 'no need to split, if only cell in row
                                Selection.Cells.Split NumRows:=1, NumColumns:=RowSpan
                                Selection.InsertBefore "colspan = """ & RowSpan & """" & IIf(InStr(1, Selection.Text, "|") > 0, " ", "|")
                                Selection.Collapse
                                Do
                                    Selection.MoveRight wdCell
                                    Selection.InsertBefore DeleteCellMarker
                                    RowSpan = RowSpan - 1
                                Loop Until RowSpan = 1
                            Else
                                Selection.InsertBefore "colspan = """ & RowSpan & """" & IIf(InStr(1, Selection.Text, "|") > 0, " ", "|")
                            End If
                            SingleRowOk = True
                            Exit For
                            
                            'continue with next line, as it is not affected
                        End If
                    Next
                    'Check if we got enough cells
                    If aRow.Cells.Count < thisTable.Columns.Count And Not SingleRowOk Then
                        'we missed some due to simple width calculation
                        mergedRow = True
                        aRow.Cells(aRow.Cells.Count).Select
                        RowSpan = thisTable.Columns.Count - aRow.Cells.Count + 1
                        'Selection.Cells.Split NumRows:=1, NumColumns:=RowSpan
                        Selection.InsertBefore "colspan = """ & RowSpan & """" & IIf(InStr(1, Selection.Text, "|") > 0, " ", "|")
                        'Selection.Collapse
                        'Do
                        '    Selection.MoveRight wdCell
                        '    Selection.InsertBefore DeleteCellMarker
                        '    RowSpan = RowSpan - 1
                        'Loop Until RowSpan = 1
                        'continue with next line, as it is not affected
                    End If
                End If
            Next
            'Loop Until mergedRow = False
                   
            
        End If 'uniform
        
        'Now the table formats and stuff
        For Each aRow In thisTable.Rows
            If aRow.Index Mod 10 = 0 Then MW_Statusbar True, "Converting text formats of table row " & aRow.Index
            For Each aCell In aRow.Cells
                With aCell
                    
                'aCell.Select 'only to watch in debug.
                
                If InStr(1, aCell.Range.Text, DeleteCellMarker) = 0 Then
                
                    'in the case a - is the first character we have a problem
                    If aCell.Range.Characters(1) = "-" Then
                        aCell.Range.Characters(1).InsertAfter "</nowiki>"
                        aCell.Range.Characters(1).InsertBefore "<nowiki>"
                    End If
                
                    'Background colors
                    Dim CellBGcolor&
                    CellBGcolor = aCell.Range.Shading.BackgroundPatternColor
                    'Debug.Print CellBGcolor, RGB2HTML(CellBGcolor)
                    If CellBGcolor <> wdColorAutomatic And RGB2HTML(CellBGcolor) <> "#FFFFFF" Then 'assume white as normal color
                        Dim Brightness&
                        With aCell.Range
                            '.Select 'just for testing
                            If .Font.color = wdColorAutomatic Then
                                'If we have a dark background, we nicht white as font color
                                'Brightness = (Red + Green + Blue) \ 3
                                Brightness = ((CellBGcolor And vbRed) + ((CellBGcolor And vbGreen) \ &H100) + ((CellBGcolor And vbBlue) \ &H10000)) \ 3
                                
                                'Hintergrundfarbe anpassen:
                                If Brightness < 128 Then
                                    '.Font.color = vbWhite
                                    .InsertBefore "<font color=""" & RGB2HTML(.Font.color) & """>"
                                    'In tables we need to close within the cell
                                    'pg.Range.Characters.Count
                                    .Characters(.Characters.Count).InsertBefore "</font>"
                                End If
                            End If
                            .InsertBefore "bgcolor = """ & RGB2HTML(CellBGcolor) & """" & IIf(InStr(1, aCell.Range.Text, "|") > 0, " ", "|")
                        End With
                    End If
                    'If aCell.Range.Shading.BackgroundPatternColor = wdColorAutomatic Then 'assume white as normal color 'RGB2HTML(CellBGcolor) <> "#000000"
                    '    aCell.Range.InsertBefore "bgcolor = """ & RGB2HTML(CellBGcolor) & """" & IIf(InStr(1, aCell.Range.Text, "|") > 0, " ", "|")
                    'End If
                    
                    'Paragraph orientation: check first paragraph and accept center and right
                    Select Case aCell.Range.Paragraphs(1).Alignment
                        Case wdAlignParagraphCenter
                            aCell.Range.InsertBefore "align = ""center""" & IIf(InStr(1, aCell.Range.Text, "|") > 0, " ", "|")
                            'aCell.Range.InsertBefore "<center>"
                            'aCell.Range.InsertAfter "</center>"
                            
                        Case wdAlignParagraphRight
                            aCell.Range.InsertBefore "align = ""right""" & IIf(InStr(1, aCell.Range.Text, "|") > 0, " ", "|")
                    
                        'Wiki does not interpret this, maybe in the future
                        Case wdAlignParagraphJustify
                            If Cell_justify Then
                                aCell.Range.InsertBefore "align = ""justify""" & IIf(InStr(1, aCell.Range.Text, "|") > 0, " ", "|")
                            End If
                    
                    End Select
                End If
                
                'Divider
                aCell.Range.InsertBefore "|"
                    
                End With
            Next aCell
            If Not aRow.IsLast Then aRow.Range.InsertAfter vbCr + vbCr + "|-"
        Next aRow

        .Range.InsertBefore "{|" & thisTableFormat & vbCr
        .Range.InsertAfter vbCr & vbCr & "|}" & thisTableInsertAfter
        
'        'now the cells that need to be deleted must be cleaned of formating
'        Dim myRange As Range
'        If 1 = 2 Then 'not needed anymore
'        For i = 1 To thisTable.Range.Cells.Count
'            With thisTable.Range.Cells(i).Range
'                If i Mod 20 = 0 Then
'                    .Select 'give the user a hint where we are
'                    MW_Statusbar True, "working on cell " & i & " of " & thisTable.Range.Cells.Count
'                End If
'                If InStr(1, .Text, "{|") = 0 Or InStr(1, .Text, "{|") = 0 Then 'no nested tables
'                    r = InStr(1, .Text, "|" & DeleteCellMarker)
'                    If r > 1 Then
'                        '.Select'just for testing
'                        Set myRange = thisTable.Range.Cells(i).Range
'                        'Delete all formating before the marker
'                        myRange.Collapse
'                        myRange.MoveEnd , r - 1
'                        myRange.Delete
'                        '.Text = Mid$(.Text, r, Len(.Text) - r - 1)
'                        'we need the row end marker
'                    End If
'                End If
'            End With
'        Next i
'        End If
        
       'Change Headings, must begin at first character in row (who makes headings within a table, anyhow?)
        .Range.Select
        '.Range.Font.Name = "Arial"
        MW_ReplaceString "|==", "|" & WordParagraph & "==", , wdFindStop
        MW_ReplaceString "<br>==", WordParagraph & "==", , wdFindStop
        MW_ReplaceString "==<br>", "==" & WordParagraph, , wdFindStop
        'Lists to the first position
        MW_ReplaceString "|*", "|" & WordParagraph & "*", , wdFindStop
        MW_ReplaceString "<br>*", WordParagraph & "*", , wdFindStop
        MW_ReplaceString "|#", "|" & WordParagraph & "#", , wdFindStop
        MW_ReplaceString "<br>#", WordParagraph & "#", , wdFindStop
        
        .ConvertToText "|"
    
    End With

    Selection.InsertAfter vbCr
    'make sure the table begins on a new line, but not if nested
    'do not know what the problem was, but the following code destroyes the tables
'    If Selection.Information(wdWithInTable) = False Then
'        Selection.MoveLeft , , True
'        If Selection.Text <> vbCr Then Selection.InsertAfter vbCr
'    End If

Exit_MW_Convert_Table:
    Exit Sub

Err_MW_Convert_Table:
    DisplayError "MW_Convert_Table"
    Resume Exit_MW_Convert_Table
End Sub

Private Sub MW_ReplaceSpecialCharactersFirst()
    'replaces one specific Character in whole document
    'released Nov. 28, 2006
    
    Dim pg As Paragraph
    Dim c&
    
    If GetReg("AllowWiki") Then Exit Sub
    
        For Each pg In ActiveDocument.Paragraphs '#Slow#
            c = c + 1
            If c Mod 50 = 0 Then MW_Statusbar True, "Replacing wiki characters, paragraph " & c & " of " & ActiveDocument.Paragraphs.Count
            Select Case pg.Range.Characters.First
                Case "#", "*", ":", ";"
                    pg.Range.Text = NoWikiOn & pg.Range.Characters.First & NoWikiOff & Mid$(pg.Range.Text, 2)
            End Select
        Next

End Sub

Private Function MW_ReplaceCharacter(Char As String, Optional FirstOnly As Boolean = False)
    'replaces one specific Character in whole document
    'released Nov. 17, 2006
    
    Dim pg As Paragraph
    Dim c&
    'Dim T As Single
    
    If GetReg("AllowWiki") Then Exit Function
    
    If FirstOnly Then
        For Each pg In ActiveDocument.Paragraphs '#Slow#
            c = c + 1
            If c Mod 50 = 0 Then MW_Statusbar True, "Replacing Wiki Character """ & Char & """, Paragraph " & c & " of " & ActiveDocument.Paragraphs.Count
            If pg.Range.Characters.First = Char Then
                pg.Range.Text = NoWikiOn & Char & NoWikiOff & Mid$(pg.Range.Text, 2)
            End If
        Next
    Else
        'replace all
        MW_ReplaceString Char, NoWikiOn & Char & NoWikiOff
    End If

End Function

Private Function MW_FindNormalWidth(aCellInfo() As CellInfoType, colNo) As Single
'From all the width information, return the mostly used or the smallest

    Dim r&, i&, arrS&, c&, mc&
    Dim Found As Boolean
    Dim Smallest As Single, sng As Single, Biggest As Single
    Dim arrW() As Single, arrC() As Long
    
    arrS = UBound(aCellInfo, 1)
    'Smallest = aCellInfo(1, colNo).cWidth
    'For r = 1 To arrS
    '    If Smallest > aCellInfo(r, colNo).cWidth Then Smallest = aCellInfo(r, colNo).cWidth
    '    If Biggest < aCellInfo(r, colNo).cWidth Then Biggest = aCellInfo(r, colNo).cWidth
    'Next r
    
    'Find most common width
    ReDim arrW(1 To arrS)
    ReDim arrC(1 To arrS)
    c = 0
    Found = False
    'count each width
    For r = 1 To arrS
        sng = aCellInfo(r, colNo).cWidth
        For i = 1 To c
            If sng = arrW(i) Then
                arrC(i) = arrC(i) + 1
                Found = True
                Exit For
            End If
        Next i
        If Not Found Then
            c = c + 1
            arrW(c) = sng
            arrC(c) = 1
        End If
    Next r
    'find highest count
    r = 0
    For i = 1 To c
        If r < arrC(i) Then
            mc = i
            r = arrC(i)
        ElseIf r = arrC(i) Then
            'Same count, use smallest
            If arrW(mc) > arrW(i) Then mc = i
        End If
    Next i
    MW_FindNormalWidth = arrW(mc)

End Function

Private Sub MW_FontFormat(chFormat$)
' -------------------------------------------------------------------
' Function: Replaces Word Font style with wiki syntax
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: Format to be changed
'
' returns: nothing
'
' released: June 16, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_FontFormat
   
    Dim iBefore$, iAfter$
    Dim FoundSome As Boolean, Insertion As Boolean
    Dim fSize&
    Const MaxFontSize = 50
    Dim rg As Range
   
    MW_Statusbar True, "converting font " & chFormat
   
    fSize = 4
    If chFormat = "FontSize" And GetReg("convertFontSize") = False Then Exit Sub
    If chFormat <> "FontSize" Then fSize = MaxFontSize 'to end after one loop
    
    Do While fSize <= MaxFontSize
        fSize = fSize + 1
        'If fSize > DefaultFontSize - 2 And fSize < DefaultFontSize + 2 Then fSize = fSize + 4 'at least two points difference
        If fSize = 9 Then fSize = 13 'at least two points difference
        Selection.HomeKey wdStory
        With Selection.Find
    
            .ClearFormatting
            Select Case chFormat
                Case "FontSize"
                    .Font.Size = fSize
                    'If fSize < DefaultFontSize Then
                    '    iBefore = "<small>"
                    '    iAfter = "</small>"
                    'Else
                    '    iBefore = "<big>"
                    '    iAfter = "</big>"
                    'End If
                    Select Case fSize / DefaultFontSize * 100
                        'Case Is < 50 'Percentage
                        '    iBefore = "<font size = ""1"">"
                        Case Is < 100
                            iBefore = "<font size = ""1"">" 'MediaWiki does not seem to interpret size 2
                        Case Is > 300
                            iBefore = "<font size = ""7"">"
                        Case Is > 200
                            iBefore = "<font size = ""6"">"
                        Case Is > 150
                            iBefore = "<font size = ""5"">"
                        Case Is > 100
                            iBefore = "<font size = ""4"">"
                    End Select
                    iAfter = "</font>"
                Case "Hidden"
                    .Font.Hidden = True
                    iBefore = ""
                    iAfter = ""
                Case "Bold"
                    .Font.Bold = True
                    iBefore = "'''"
                    iAfter = iBefore
                Case "Italic"
                    .Font.Italic = True
                    iBefore = "''"
                    iAfter = iBefore
                Case "StrikeThrough"
                    .Font.Strikethrough = True
                    iBefore = "<s>"
                    iAfter = "</s>"
                Case "Subscript"
                    .Font.Subscript = True
                    iBefore = "<sub>"
                    iAfter = "</sub>"
                Case "Superscript"
                    .Font.Superscript = True
                    iBefore = "<sup>"
                    iAfter = "</sup>"
                Case "Underline"
                    .Font.Underline = True 'wdUnderlineSingle
                    iBefore = "<u>"
                    iAfter = "</u>"
                Case Else
                    Exit Sub
            End Select
            .Text = ""
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Forward = True
            .Wrap = wdFindContinue
        End With

        FoundSome = Selection.Find.Execute
        Do While FoundSome
            Insertion = False
            
            'Now check results, there seems to be a word bug
            Select Case chFormat
                Case "FontSize"
                    FoundSome = Selection.Font.Size = fSize
                Case "Hidden"
                    FoundSome = Selection.Font.Hidden
                Case "Bold"
                    FoundSome = Selection.Font.Bold
                Case "Italic"
                    FoundSome = Selection.Font.Italic
                Case "StrikeThrough"
                    FoundSome = Selection.Font.Strikethrough
                Case "Subscript"
                    FoundSome = Selection.Font.Subscript
                Case "Superscript"
                    FoundSome = Selection.Font.Superscript
                Case "Underline"
                    FoundSome = Selection.Font.Underline <> wdUnderlineNone
            End Select
            
            If FoundSome Then
                With Selection
    
                    If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                        ' Just process the chunk before any newline characters
                        ' We'll pick-up the rest with the next search
                        .Collapse
                        .MoveEndUntil vbCr
                    End If
                    '.Style = ActiveDocument.Styles(wdStyleNormal)
                    
                    'We have a problem if different formats like bold and underline do not start and end at the same character
                    'In this case we just do all letters that have the same format
                    Do While .Font.Bold = 9999999
                        Selection.MoveLeft , , True
                    Loop
                    Do While .Font.Italic = 9999999
                        Selection.MoveLeft , , True
                    Loop
                    Do While .Font.Strikethrough = 9999999
                        Selection.MoveLeft , , True
                    Loop
                    Do While .Font.Subscript = 9999999
                        Selection.MoveLeft , , True
                    Loop
                    Do While .Font.Superscript = 9999999
                        Selection.MoveLeft , , True
                    Loop
                    Do While .Font.Underline = 9999999
                        Selection.MoveLeft , , True
                    Loop
                    Do While .Font.Size = 9999999
                        Selection.MoveLeft , , True
                    Loop
                    Do While .Font.color = 9999999
                        Selection.MoveLeft , , True
                    Loop
'                    Do While .Font.Shading.BackgroundPatternColor = 9999999
'                        Selection.MoveLeft , , True
'                    Loop
'                    Do While .Range.HighlightColorIndex = 9999999
'                        Selection.MoveLeft , , True
'                    Loop
                    
                    ' Don't bother to markup newline characters (prevents a loop, as well)
                    Dim s As Selection 'just for testing
                    Set s = Selection
                    If Not (Len(.Range.Text) = 1 And Asc(.Range.Text) <= 13) Then
                        Insertion = True
                        
                        'we need the formatting to expand in front
                        Selection.CopyFormat
                        
                        Set rg = Selection.Range
                        'If rg.Characters.Count = 1 Then Stop
                        'so we insert a charcter, move it to the front
                        'rg.Characters(1).InsertAfter .Characters(1)
                        'rg.MoveStart 'jump one step to the right ;)
                        rg.InsertBefore iBefore
                        'rg.MoveStart , -1
                        'If rg.Characters.Count = 1 + Len(iBefore) Then
                        '    rg.MoveEnd
                        '    rg.Select
                        'End If
                        
                        'This might seem a little strange, but
                        '.insertAfter does not insert at the right position within table
                        'seems to be a word bug
                        .Collapse wdCollapseEnd
                        .InsertAfter iAfter
                        rg.MoveEnd , Len(iAfter)
                        rg.Select
                        
                        'give Format to inserted text
                        .Collapse
                        .MoveRight , Len(iBefore), True
                        .PasteFormat
                        
                        rg.Select
                        
                    End If
    
                    Select Case chFormat
                        Case "FontSize"
                            .Font.Size = DefaultFontSize
                        Case "Hidden"
                            If GetReg("deleteHiddenChars") Then
                                .Delete
                            Else
                                .Font.Hidden = False
                                .Font.ColorIndex = wdGray50 'mark as special, because hidden
                            End If
                        Case "Bold"
                            .Font.Bold = False
                        Case "Italic"
                            .Font.Italic = False
                        Case "StrikeThrough"
                            .Font.Strikethrough = False
                        Case "Subscript"
                            .Font.Subscript = False
                        Case "Superscript"
                            .Font.Superscript = False
                        Case "Underline"
                            .Font.Underline = wdUnderlineNone
                        'Case Else
                        '    Exit Sub
                    End Select
                    
                    .Collapse wdCollapseStart
                    'If Insertion Then
                    '    .Delete
                    'End If
    
                End With
                FoundSome = Selection.Find.Execute
            
            End If
        Loop
    Loop

Exit_MW_FontFormat:
    MW_Statusbar False
    Selection.Find.ClearFormatting
    Exit Sub

Err_MW_FontFormat:
    DisplayError "MW_FontFormat"
    Resume Exit_MW_FontFormat
End Sub

Public Function MW_FormatCategoryString(Category$, Optional Image As Boolean = False) As String
'Makes wiki syntax from category name like [[category:...]]
On Error Resume Next
    
    Dim catArr$(), i&
    Dim cip$
    
    cip = Trim$(GetReg("CategoryImagePreFix"))
    If cip <> "" Then cip = cip & " "

    If InStr(1, Category, ",") = 0 Then
        'Only one category
        If Trim$(Category$) = "" Then Exit Function 'returns ""
        MW_FormatCategoryString = "[[" & GetReg("WikiCategoryKeyWord") & ":" & IIf(Image, cip, "") & Trim$(Category$) & "]]"
    Else
        'several categories
        catArr() = Split(Category, ",")
        MW_FormatCategoryString = "[[" & GetReg("WikiCategoryKeyWord") & ":" & IIf(Image, cip, "") & Trim$(catArr(0)) & "]]"
        For i = 1 To UBound(catArr)
            MW_FormatCategoryString = MW_FormatCategoryString & vbCr & "[[" & GetReg("WikiCategoryKeyWord") & ":" & IIf(Image, cip, "") & Trim$(catArr(i)) & "]]"
        Next i
    End If

End Function

Public Function MW_GetEditorPath(Optional Simulate As Boolean = False) As String
' -------------------------------------------------------------------
' Function: Checks path of provided editor, looks for the MS Photo Editor, if empty
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: Simulate: check if the editor can be called
'
' returns: EditorPathInt
'
' released: June 04, 2006
' -------------------------------------------------------------------
On Error Resume Next

    Dim EditorPathInt$

    'Assume the position of Microsoft Photo Editor
    EditorPathInt = EditorPrgPath
    If EditorPathInt = "" Then EditorPathInt = FormatPath(GetSpecialfolder(CSIDL_PROGRAM_FILES_COMMON)) & "Microsoft Shared\PhotoEd\PhotoEd.exe"
    If Not FileExists(EditorPathInt) Then
        If GetReg("ImageExtractionPE") Then
            MsgBox EditorTitle & " not found. Check path:" & vbCr & vbCr & EditorPathInt & vbCr & vbCr & "Unable to convert images.", vbExclamation, ConverterPrgTitle
        End If
        EditorPathInt = ""
    End If
    MW_GetEditorPath = EditorPathInt
    
    If Simulate And EditorPathInt <> "" Then
        'Open Editor
        Shell EditorPathInt, vbMaximizedFocus
        DoEvents
        Sleep 1500
        If AppActivatePlus(EditorTitle, False) Then
            'SendKeys "%{F4}", True 'End Photo Editor
            MW_CloseProgramm EditorTitle
        Else
            MsgBox "Error: Could not activate your Photo Editor. You will not be able to convert pictures!", vbCritical
            Exit Function
        End If
        Application.Activate
    End If

End Function

Private Function MW_GetImageNameFromFile(ImageTxt$, ImagePath$) As String
' -------------------------------------------------------------------
' Function: Searches for an existing file with that name
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: ImageBaseName
'
' returns: nothing
'
' released: Nov. 11, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_GetImageNameFromFile

    Dim Name1$

    Name1 = Dir(ImagePath & ImageTxt & "*", vbDirectory)
    Do While Name1 <> ""
        ' check that it is not a dir!
        If (GetAttr(ImagePath & Name1) And vbDirectory) <> vbDirectory Then
            MW_GetImageNameFromFile = Name1
            Exit Function
        End If    ' um ein Verzeichnis handelt.
        Name1 = Dir    ' Nächsten Eintrag abrufen.
    Loop
    'if we are here, no appropriate file was found
    'take a guess
    MW_GetImageNameFromFile = ImageTxt & ".png"
    
Exit_MW_GetImageNameFromFile:
    Exit Function

Err_MW_GetImageNameFromFile:
    DisplayError "MW_GetImageNameFromFile"
    Resume Exit_MW_GetImageNameFromFile
End Function

Private Function MW_GetImagePath() As String
    
    MW_GetImagePath = GetReg("ImagePath")
    If MW_GetImagePath = "" Then MW_GetImagePath = Documents(DocInfo.DocName).Path
    MW_GetImagePath = FormatPath(MW_GetImagePath) & DocInfo.DocNameNoExt & "\"

End Function

Private Sub MW_GetScaleIS(MyInlineShape As InlineShape, ScaleW#, ScaleH#)
' -------------------------------------------------------------------
' Function: retrieves real scale values.
' Word Bug: Wrong scale values, if picture is cropped
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: InlineShape
'
' returns: real ScaleWidth and Height
'
' released: Nov 9, 2006
' -------------------------------------------------------------------
On Error Resume Next
    Dim Crop&
    
    With MyInlineShape
        'real scale values
        Crop = .PictureFormat.CropBottom + .PictureFormat.CropLeft + .PictureFormat.CropRight + .PictureFormat.CropTop
        Err.Clear
    
        On Error GoTo Err_MW_GetScaleIS
        'real scale values
        If Crop < 4 Then
            ScaleW = .ScaleWidth
            ScaleH = .ScaleHeight
        Else
            ScaleH = .Height / (.Height / .ScaleHeight * 100 - .PictureFormat.CropBottom - .PictureFormat.CropTop) * 100
            ScaleW = .Width / (.Width / .ScaleWidth * 100 - .PictureFormat.CropLeft - .PictureFormat.CropRight) * 100
            'round
            If ScaleH > 99 And ScaleH < 103 Then ScaleH = 100
            If ScaleW > 99 And ScaleW < 103 Then ScaleW = 100
        End If
    End With

Exit_MW_GetScaleIS:
    Exit Sub

Err_MW_GetScaleIS:
    DisplayError "MW_GetScaleIS"
    If DebugMode Then Stop: Resume Next
    Resume Exit_MW_GetScaleIS
End Sub

Private Function MW_GetUserLanguage() As String
' -------------------------------------------------------------------
' Function: returns the language code of the current user
'   only known languages
'
' Input: nothing
'
' returns: 3-letter language code of the current user
'
' released: June 18, 2006
' -------------------------------------------------------------------
On Error Resume Next

    'Windows Languages
    Const LANG_NEUTRAL = &H0      '   Neutral
    Const LANG_ARABIC = &H1       '   Arabic
    Const LANG_BULGARIAN = &H2    '   Bulgarian
    Const LANG_CATALAN = &H3      '   Catalan
    Const LANG_CHINESE = &H4      '   Chinese
    Const LANG_CZECH = &H5        '   Czech
    Const LANG_DANISH = &H6       '   Danish
    Const LANG_GERMAN = &H7       '   German
    Const LANG_GREEK = &H8        '   Greek
    Const LANG_ENGLISH = &H9      '   English
    Const LANG_SPANISH = &HA      '   Spanish
    Const LANG_FINNISH = &HB      '   Finnish
    Const LANG_FRENCH = &HC       '   French
    Const LANG_HEBREW = &HD       '   Hebrew
    Const LANG_HUNGARIAN = &HE    '   Hungarian
    Const LANG_ICELANDIC = &HF    '   Icelandic
    Const LANG_ITALIAN = &H10     '   Italian
    Const LANG_JAPANESE = &H11    '   Japanese
    Const LANG_KOREAN = &H12      '   Korean
    Const LANG_DUTCH = &H13       '   Dutch
    Const LANG_NORWEGIAN = &H14   '   Norwegian
    Const LANG_POLISH = &H15      '   Polish
    Const LANG_PORTUGUESE = &H16  '   Portuguese
    Const LANG_ROMANIAN = &H18    '   Romanian
    Const LANG_RUSSIAN = &H19     '   Russian
    Const LANG_CROATIAN = &H1A    '   Croatian
    Const LANG_SERBIAN = &H1A     '   Serbian
    Const LANG_SLOVAK = &H1B      '   Slovak
    Const LANG_ALBANIAN = &H1C    '   Albanian
    Const LANG_SWEDISH = &H1D     '   Swedish
    Const LANG_THAI = &H1E        '   Thai
    Const LANG_TURKISH = &H1F     '   Turkish
    Const LANG_URDU = &H20        '   Urdu
    Const LANG_INDONESIAN = &H21  '   Indonesian
    Const LANG_UKRAINIAN = &H22   '   Ukrainian
    Const LANG_BELARUSIAN = &H23  '   Belarusian
    Const LANG_SLOVENIAN = &H24   '   Slovenian
    Const LANG_ESTONIAN = &H25    '   Estonian
    Const LANG_LATVIAN = &H26     '   Latvian
    Const LANG_LITHUANIAN = &H27  '   Lithuanian
    Const LANG_FARSI = &H29       '   Farsi
    Const LANG_VIETNAMESE = &H2A  '   Vietnamese
    Const LANG_ARMENIAN = &H2B    '   Armenian
    Const LANG_AZERI = &H2C       '   Azeri
    Const LANG_BASQUE = &H2D      '   Basque
    Const LANG_MACEDONIAN = &H2F  '   Macedonian (FYROM)
    Const LANG_AFRIKAANS = &H36   '   Afrikaans
    Const LANG_GEORGIAN = &H37    '   Georgian
    Const LANG_FAEROESE = &H38    '   Faeroese
    Const LANG_HINDI = &H39       '   Hindi
    Const LANG_MALAY = &H3E       '   Malay
    Const LANG_KAZAK = &H3F       '   Kazak
    Const LANG_KYRGYZ = &H40      '   Kyrgyz
    Const LANG_SWAHILI = &H41     '   Swahili
    Const LANG_UZBEK = &H43       '   Uzbek
    Const LANG_TATAR = &H44       '   Tatar
    Const LANG_BENGALI = &H45     '   Not supported.
    Const LANG_PUNJABI = &H46     '   Punjabi
    Const LANG_GUJARATI = &H47    '   Gujarati
    Const LANG_ORIYA = &H48       '   Not supported.
    Const LANG_TAMIL = &H49       '   Tamil
    Const LANG_TELUGU = &H4A      '   Telugu
    Const LANG_KANNADA = &H4B     '   Kannada
    Const LANG_MALAYALAM = &H4C   '   Not supported.
    Const LANG_ASSAMESE = &H4D    '   Not supported.
    Const LANG_MARATHI = &H4E     '   Marathi
    Const LANG_SANSKRIT = &H4F    '   Sanskrit
    Const LANG_MONGOLIAN = &H50   '   Mongolian
    Const LANG_GALICIAN = &H56    '   Galician
    Const LANG_KONKANI = &H57     '   Konkani
    Const LANG_MANIPURI = &H58    '   Not supported.
    Const LANG_SINDHI = &H59      '   Not supported.
    Const LANG_SYRIAC = &H5A      '   Syriac
    Const LANG_KASHMIRI = &H60    '   Not supported.
    Const LANG_NEPALI = &H61      '   Not supported.
    Const LANG_DIVEHI = &H65      '   Divehi
    Const LANG_INVARIANT = &H7F   '   unknown
    
    Dim LngCode&
    LngCode = GetUserDefaultLangID
    LngCode = LngCode And &H3FF
    
    Select Case LngCode
        Case LANG_GERMAN
            MW_GetUserLanguage = "GER" 'NoTag
        Case LANG_ENGLISH
            MW_GetUserLanguage = "ENG" 'NoTag
        Case Else
            MsgBox "Your language is not supported. If you do not use english programs, you need to change the coding. Look into MW_LanguageTexts.", vbExclamation, ConverterPrgTitle
            MW_GetUserLanguage = "ENG" 'NoTag
    End Select

End Function

Private Function MW_ImageExportPowerpointPNG(ExportPath$, ExportPathExtra$, ImageName$, Optional SpecialSize& = 0) As Long
' -------------------------------------------------------------------
' Function: Exports a wmz or emz picture via powerpoint
' actually it could be used for any picture
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: ExportPath, ImageName without extension
'
' returns: Size of main picture, saves pictures in different sizes on disk
'
' released: Nov. 11, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_ImageExportPowerpointPNG

    Dim ppt As Object
    Dim PicWidth&, PicHeight&, UseSize&
    Dim T As Single
    Dim MyS As Object ' Powerpoint.shape
    
    T = Timer
    DocInfo.PowerpointStarted = True
    UseSize = SpecialSize
    If UseSize < 10 Then
        Select Case GetReg("ImageMaxWidth")
            Case Is < 1024
                UseSize = 800
            Case Is < 1280
                UseSize = 1024
            Case Else
                UseSize = 1280
        End Select
    End If
    
    'Select picture
    If ActiveDocument.InlineShapes.Count > 0 Then
        ActiveDocument.InlineShapes(1).Select
    ElseIf ActiveDocument.Shapes.Count > 0 Then
        ActiveDocument.Shapes(1).Select
    Else
        Exit Function
    End If
    
    If Not DirExists(ExportPath) Then MkDir ExportPath
    DoEvents
    Selection.Copy
    DoEvents
    
    'Open new Powerpoint presentation
    Set ppt = CreateObject("Powerpoint.Application")
    ppt.Visible = True
    ppt.Activate

    Dim MyPr As Object 'Presentation
    Set MyPr = ppt.Presentations.Add(True) 'Set to true, to see what Powerpoint ist doing
    'MyPr.Slides.Add 1, ppLayoutBlank
    MyPr.Slides.Add 1, 12
    DoEvents
    Sleep 100 'Sometime the image is not yet in the clipboard
    MyPr.Slides(1).Shapes.Paste
    
    'In some cases there seem to be more than one pictures, we take the last one
    
    Set MyS = MyPr.Slides(1).Shapes(MyPr.Slides(1).Shapes.Count)
    
    With MyS
        '.Left = 0#
        '.Top = 0#
        .LockAspectRatio = msoFalse
        PicWidth = .Width
        PicHeight = .Height
        '.ScaleHeight 1, msoTrue
        '.ScaleWidth 1, msoTrue
    End With
    
    With MyPr.PageSetup
        '.SlideSize = ppSlideSizeCustom
        '.SlideSize = 7
        .SlideWidth = PicWidth
        .SlideHeight = PicHeight
        '.FirstSlideNumber = 1
        '.SlideOrientation = msoOrientationHorizontal
        '.NotesOrientation = msoOrientationVertical
    End With

    'repeat, because resize of page dimension resizes image!
    With MyS
        .Left = 0#
        .Top = 0#
        .Width = PicWidth
        .Height = PicHeight
        '.ScaleHeight 1, msoTrue
        '.ScaleWidth 1, msoTrue
    End With

    'ExportPath = "E:\work\Wiki-Test\"
    
    'Save as PNG
'    With ppt.ActiveWindow.View.Slide
    With MyPr.Slides(1)
        'original size
        PicWidth = PointsToPixels(PicWidth)
        If DocInfo.ImageResizeOption = 3 Then UseSize = PicWidth
        'If it smaller than almost page width, than we do not scale up
        If PicWidth < 500 Then UseSize = PicWidth               'Use Original
        ' Set height proportional to slide height
        PicHeight = (PicWidth * MyPr.PageSetup.SlideHeight) / MyPr.PageSetup.SlideWidth
        If UseSize = PicWidth Then
            .Export ExportPath & ImageName & "_" & PicWidth & ".png", "PNG", PicWidth, PicHeight
        Else
            .Export ExportPathExtra & ImageName & "_" & PicWidth & ".png", "PNG", PicWidth, PicHeight
        End If
    
        PicWidth = 800
        PicHeight = (PicWidth * MyPr.PageSetup.SlideHeight) / MyPr.PageSetup.SlideWidth
        If UseSize = PicWidth Then
            .Export ExportPath & ImageName & "_" & PicWidth & ".png", "PNG", PicWidth, PicHeight
        Else
            .Export ExportPathExtra & ImageName & "_" & PicWidth & ".png", "PNG", PicWidth, PicHeight
        End If
    
        PicWidth = 1024
        PicHeight = (PicWidth * MyPr.PageSetup.SlideHeight) / MyPr.PageSetup.SlideWidth
        If UseSize = PicWidth Then
            .Export ExportPath & ImageName & "_" & PicWidth & ".png", "PNG", PicWidth, PicHeight
        Else
            .Export ExportPathExtra & ImageName & "_" & PicWidth & ".png", "PNG", PicWidth, PicHeight
        End If
    
        PicWidth = 1280
        PicHeight = (PicWidth * MyPr.PageSetup.SlideHeight) / MyPr.PageSetup.SlideWidth
        If UseSize = PicWidth Then
            .Export ExportPath & ImageName & "_" & PicWidth & ".png", "PNG", PicWidth, PicHeight
        Else
            .Export ExportPathExtra & ImageName & "_" & PicWidth & ".png", "PNG", PicWidth, PicHeight
        End If
    
        'Individual Size (additional file)
        If SpecialSize > 10 Then
            PicWidth = UseSize
            PicHeight = (PicWidth * MyPr.PageSetup.SlideHeight) / MyPr.PageSetup.SlideWidth
            If UseSize = PicWidth Then
                .Export ExportPath & ImageName & "_" & PicWidth & ".png", "PNG", PicWidth, PicHeight
            Else
                .Export ExportPathExtra & ImageName & "_" & PicWidth & ".png", "PNG", PicWidth, PicHeight
            End If
        End If
    
    End With
    
    'ImageInfo.PPTUsed = True
    MW_ImageExportPowerpointPNG = UseSize
    MyPr.Close
    Set ppt = Nothing
    
    'Debug.Print Format(Timer - T, "0.0000")

Exit_MW_ImageExportPowerpointPNG:
    Application.Activate 'back to word
    Exit Function

Err_MW_ImageExportPowerpointPNG:
    DisplayError "MW_ImageExportPowerpointPNG"
    If DebugMode Then Stop: Resume Next
    Resume Exit_MW_ImageExportPowerpointPNG
End Function

Private Sub MW_ImageExtract(MyS As Object, ImageInfo As ImageInfoType, Optional usePNG As Boolean = False, Optional SaveImageName As Boolean = True)
' -------------------------------------------------------------------
' Function: Extracts the selected object
' Some words to the extraction prozess
' All pictures are copied to a new document, which is then saved as webpage (html).
' There will be two pictures created, the original picture and another one in a common graphic format, just the size of display
' Depending on the settings and the format one of the two will be used.
' We do not use Compressed Windows Enhanced Metafile (*.emz) and Compressed Windows Metafile (*.wmz),
'  these formats write also a gif-file, which is loosing some colors, so there is a quality loss.
'  A solution to this problem would be, to paste in PowerPoint and save as png or jpg
'  ActiveWindow.Selection.SlideRange.Shapes("Picture 7").Select
'  ActivePresentation.SaveAs FileName:="D:\Eigene Dateien\Bild5.jpg", FileFormat:=ppSaveAsJPG, EmbedTrueTypeFonts:=msoFalse
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: MyS Shape or InlineShape
'   ImagePath and ImageTxt (=ImageName without extension)
'
' returns: ImageName incl. extension, adds an entry in imageArr
'
' released: Nov. 07, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_ImageExtract

    Dim p&
    Dim ImageExtractionInt As Boolean
    Dim DocN As Document, DocPNG As Document ', DocHtml As Document

    ImageInfo.PPTUsed = False
    ImageExtractionInt = GetReg("ImageExtraction")
    'ImageExtractionInt = True 'for test only
    If ImageExtractionInt Then
        DoEvents
        'Save as HTML, so we get the original picture out of the document
        'Make new document to identify the correct image
        
        'Make new document with hyperlink on original picture
        'Save document as HTML
'        On Error Resume Next
'        If Selection.InlineShapes.Count = 0 Then If Selection.ShapeRange.Count = 0 Then Err.Clear: Exit Sub      'nothing selected
'        Err.Clear
'        On Error GoTo Err_MW_ImageExtract
        MyS.Select
        Selection.Copy
        
        'Make two documents for differnt kinds of conversion
        'Insert as png for future reference
        If MW_WordVersion >= 2002 Then
            Set DocPNG = Documents.Add 'DocumentType:=wdNewBlankDocument
            Selection.PasteSpecial Link:=False, DataType:=14, Placement:=wdFloatOverText, DisplayAsIcon:=False
        End If
        
        'Work with normal copy first
        Set DocN = Documents.Add 'DocumentType:=wdNewBlankDocument
        On Error Resume Next
        Selection.PasteSpecial Link:=True, DataType:=wdPasteHyperlink
        If Err.Number > 0 Then
            Err.Clear
            On Error GoTo Err_MW_ImageExtract
            Selection.Paste
        End If
        On Error GoTo Err_MW_ImageExtract
        
        MW_ImageExtract2 MyS, DocN, ImageInfo
        'if emz or wmz, we can not convert and ImageName is blank, so we use png copy
        If ImageInfo.Name = "" Or usePNG Then MW_ImageExtract2 MyS, DocPNG, ImageInfo, True
        
    End If
    
    If (ImageExtractionInt Or GetReg("ImageReload")) And SaveImageName And ImageInfo.Name <> "" Then
        'Save Imagename in ImageArray for MW_ImageExtract
        On Error Resume Next
        p = UBound(ImageArr)
        If Err.Number <> 0 Then ReDim ImageArr(1 To 1): p = 1: Err.Clear
        On Error GoTo Err_MW_ImageExtract
        If ImageArr(p) <> "" Then
            p = p + 1
            ReDim Preserve ImageArr(1 To p)
        End If
        
        ImageArr(p) = DocInfo.ImagePath & ImageInfo.Name
    End If
    
Exit_MW_ImageExtract:
    If IsObjectValid(DocN) Then DocN.Close wdDoNotSaveChanges
    If IsObjectValid(DocPNG) Then DocPNG.Close wdDoNotSaveChanges
    Exit Sub

Err_MW_ImageExtract:
    DisplayError "MW_ImageExtract"
    If DebugMode Then Stop: Resume Next
    Resume Exit_MW_ImageExtract
End Sub

Private Sub MW_ImageExtract2(MyS As Object, DocH As Document, ImageInfo As ImageInfoType, Optional UseFullSize As Boolean = False)
On Error GoTo Err_MW_ImageExtract2
    
    Dim Name1$, ExportPath$, ExportPathExtra$
    Static ImageHtmlPath$
    Dim PicArr$(), c&
    
    ReDim PicArr(1 To 10)
    ImageInfo.Name = ""
    ImageInfo.NameDisplay = ""
            
    ExportPath = FormatPath(DocInfo.ImagePath)
    DoEvents
    MakeDir ExportPath
    ExportPathExtra = ExportPath & "Extra_Formats\"
    MakeDir ExportPathExtra
        
    'Now we need to find the correct picture
    'First picture is original
    'Last picture is displayed size, sometimes smaller
    
    If Not IsObjectValid(DocH) Then Exit Sub
    
    DoEvents
    If FileExists(DocInfo.ImagePath & "ImageConversion.htm") Then KillToBin DocInfo.ImagePath & "ImageConversion.htm"
    DoEvents
    DocH.SaveAs FileName:=DocInfo.ImagePath & "ImageConversion.htm", FileFormat:=wdFormatHTML, AddToRecentFiles:=False, SaveNativePictureFormat:=False
    If DirExists(ImageHtmlPath) = False Then
        'Find correct dirname, as it is language specific
        Name1 = Dir(DocInfo.ImagePath, vbDirectory)
        Do While Name1 <> ""
            If Left$(Name1, 15) = "ImageConversion" Then
                ' Mit bit-weisem Vergleich sicherstellen, daß Name1 ein Verzeichnis ist.
                If (GetAttr(DocInfo.ImagePath & Name1) And vbDirectory) = vbDirectory Then
                    ImageHtmlPath = DocInfo.ImagePath & Name1 & "\"
                    Exit Do
                End If    ' um ein Verzeichnis handelt.
            End If
            Name1 = Dir    ' Nächsten Eintrag abrufen.
        Loop
    End If
    
    'Find image files
    If DirExists(ImageHtmlPath) Then
        Name1 = Dir(ImageHtmlPath, vbDirectory)
        Do While Name1 <> ""
            If Name1 <> "." And Name1 <> ".." Then
                If Left$(Name1, 5) = "image" Then
                    c = c + 1
                    If c Mod 10 = 0 Then ReDim Preserve PicArr(1 To c + 10)
                    PicArr(c) = Name1
                Else
                    'Get rid of unused files
                    'killtobin ImageHtmlPath & Name1
                End If
            End If
            Name1 = Dir
        Loop
            
        Select Case c
            Case 1
                'only picture, we use it
                'should be usable like png, jpg, gif
                ImageInfo.Name = ImageInfo.NameNoExt & right$(PicArr(1), 4)
                FileCopy ImageHtmlPath & PicArr(1), ExportPath & ImageInfo.Name
            
            Case 2
                'Copy the display picture, which is the last one
                'document display size, usually jpg or gif
                ImageInfo.NameDisplay = ImageInfo.NameNoExt & "-display" & right$(PicArr(c), 4)
                FileCopy ImageHtmlPath & PicArr(c), ExportPathExtra & ImageInfo.NameDisplay
                
                'original picture and display picture, normal outcome
                'we always use a png and gif regardless of its size
                Select Case right$(PicArr(1), 3)
                    Case "png", "gif", "jpg"
                        'usable picture format
                        'we always use a png and gif regardless of its size
                        ImageInfo.Name = ImageInfo.NameNoExt & right$(PicArr(1), 4)
                        FileCopy ImageHtmlPath & PicArr(1), ExportPath & ImageInfo.Name
                    
                    Case "emz", "wmz"
                        'can't use that one, so we use the copy image002
                        FileCopy ImageHtmlPath & PicArr(1), ExportPathExtra & ImageInfo.NameNoExt & "-FullSize" & right$(PicArr(1), 4)
                        'Powerpoint Export here, only big pictures, small ones are good with word png
                        Dim usePPT As Boolean, pptSize&
                        usePPT = GetReg("UsePowerpoint")
                        If usePPT Then
                            If MyS.Type = 1 And MyS.Width < 100 Then usePPT = False ' "EmbeddedOLEObject"
                            If usePPT Then
                                'pptSize = GetReg("UsePPTSize")
                                pptSize = MW_ImageExportPowerpointPNG(ExportPath, ExportPathExtra, ImageInfo.NameNoExt) ', pptSize
                                ImageInfo.Name = ImageInfo.NameNoExt & "_" & pptSize & ".png"
                                If Not FileExists(ExportPath & ImageInfo.Name) Then ImageInfo.Name = "" Else ImageInfo.PPTUsed = True
                            End If
                        End If
                
                    Case Else
                        MsgBox "Unknown file format in MW_ImageExtract2. Please send file to author.", vbExclamation
                        If DebugMode Then Stop 'for unknown formats
                        
                End Select
                
            Case Is >= 3
                'we have a grouped picture, so we use the display picture
                'Copy the display picture, which is the last one
                'document display size, usually jpg or gif
                ImageInfo.NameDisplay = ImageInfo.NameNoExt & "-display" & right$(PicArr(c), 4)
                If right$(PicArr(c), 4) <> ".png" And GetReg("UsePowerpoint") Then
                    'we try to make a better quality picture
                    pptSize = MW_ImageExportPowerpointPNG(ExportPath, ExportPathExtra, ImageInfo.NameNoExt) ', pptSize
                    ImageInfo.Name = ImageInfo.NameNoExt & "_" & pptSize & ".png"
                    If FileExists(ExportPath & ImageInfo.Name) Then ImageInfo.PPTUsed = True
                End If
                'move the display picture to correct dir
                If FileExists(ExportPath & ImageInfo.Name) Then
                    FileCopy ImageHtmlPath & PicArr(c), ExportPathExtra & ImageInfo.NameDisplay
                Else
                    ImageInfo.Name = ImageInfo.NameDisplay
                    FileCopy ImageHtmlPath & PicArr(c), ExportPath & ImageInfo.NameDisplay
                End If
                
                'Copy all other pictures
                Do While c > 1
                    c = c - 1 'because we do not want the display picture
                    ImageInfo.NameDisplay = ImageInfo.NameNoExt & "_" & right$(PicArr(c), 7)
                    FileCopy ImageHtmlPath & PicArr(c), ExportPathExtra & ImageInfo.NameDisplay
                Loop
                
        End Select
        
        RemoveDir ImageHtmlPath ', False
    
    End If
        
    DocH.Close wdDoNotSaveChanges
    'On Error Resume Next
    If FileExists(DocInfo.ImagePath & "ImageConversion.htm") Then KillToBin DocInfo.ImagePath & "ImageConversion.htm"

Exit_MW_ImageExtract2:
    Exit Sub

Err_MW_ImageExtract2:
    DisplayError "MW_ImageExtract2"
    If DebugMode Then Stop: Resume Next
    Resume Exit_MW_ImageExtract2
End Sub

Private Function MW_ImagePathName(FPath$, PicFormat$, Optional KillFile As Boolean = False) As String
' -------------------------------------------------------------------
' Function: returns a new name according to the picture format
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: KillFile will delete fill if exists, Picformat can be any file extension
'
' returns: nothing
'
' released: June 04, 2006
' -------------------------------------------------------------------
On Error Resume Next
    Dim txt$
    txt = Left$(FPath, Len(FPath) - 3) & PicFormat 'assuming 3 letters
    If KillFile Then If FileExists(txt) Then KillToBin txt: DoEvents
    MW_ImagePathName = txt
End Function

Public Function MW_ImageUpload_File(ByVal FilePathName$, Optional Simulate As Boolean = False) As Long
' -------------------------------------------------------------------
' Function: Load one file into your wiki. Tab steps must be customized
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: File to be uploaded
'
' returns: 0 if no error, -1 if user cancels, errorcode
'
' released: Nov 8, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_ImageUpload_File

    Dim UploadFilePathName
    Dim i&, FLen&, c&, errC&
    Dim txt$
    Static UploadPrepared&
    Static LastCheck As Single
    Const ImageUploadAddress$ = "Special:Upload"
    
    If FileExists(FilePathName) Or Simulate Then
        
        UploadFilePathName = GetFilePath(FilePathName) & "Uploaded\"
        If Not DirExists(UploadFilePathName) Then MakeDir UploadFilePathName
        UploadFilePathName = UploadFilePathName & GetFileName(FilePathName)
        
        If Simulate Then
            FLen = -1
        Else
            FLen = FileLen(FilePathName)
        End If
    
        If FLen >= ImageMaxUploadSize Then
            
            Application.Activate
            i = MsgBox("The file " & GetFileName(FilePathName) & " is very big (" & Int(FLen / 1024) & "KB)! Do you really want to upload this file?", vbExclamation + vbYesNoCancel, ConverterPrgTitle)
            If i = vbCancel Then MW_ImageUpload_File = -1
            If i <> vbYes Then Exit Function
            
        End If
            
            'Prepare Upload Message
            If Timer - LastCheck > 60 Then UploadPrepared = False 'Reset Upload after a minute
            If UploadPrepared <> vbOK Then
                IExplorer WikiAddressRoot & ImageUploadAddress
                Sleep ImageUploadWaitTime
                DoEvents
                Application.Activate
                DoEvents
                UploadPrepared = MsgBox(Msg_Upload_Info, vbInformation + vbOKCancel, ConverterPrgTitle)
                DoEvents
                If UploadPrepared <> vbOK Then MW_ImageUpload_File = -1: Exit Function
            End If
            
            'Open the upload adress in browser
            IExplorer WikiAddressRoot & ImageUploadAddress
            DoEvents
            
            c = 0
            Do
                c = c + 1
                Sleep ImageUploadWaitTime
                If AppActivatePlus(GetReg("WikiUploadTitle"), False) Then
                    For i = 1 To GetReg("ImageUploadTabToFileName")
                        SendKeys "{Tab}", True: DoEvents: Sleep 100
                    Next
                    If DebugMode Then Debug.Print FilePathName
                    SendKeys FilePathName, True: DoEvents: Sleep 100
                    For i = 1 To ImageUploadTabToDescription
                        SendKeys "{Tab}", True: DoEvents: Sleep 100
                    Next
                    txt = Trim$(GetReg("ImageDescription"))
                    If txt <> "" Then
                        SendKeys txt & "{Enter}{Enter}", True: DoEvents: Sleep 100
                    End If
                    If GetReg("CategoryImagesUse") Then
                        txt = MW_FormatCategoryString(GetReg("CategoryImages"), True)
                        If txt <> "" Then
                            SendKeys txt, True: Sleep 100
                            If ImageIconOnlyType Then
                                If Left$(GetFileName(FilePathName), 4) = "Icon" Then SendKeys "{Enter}" & MW_FormatCategoryString("Icon", True), True: Sleep 100
                            End If
                        End If
                    End If
                    For i = 1 To ImageUploadTabToEnter
                        SendKeys "{Tab}", True:  DoEvents: Sleep 100
                    Next
                    If Not Simulate Then SendKeys "{Enter}", True: DoEvents
                    
                    'Wait some time
                    Sleep ImageUploadWaitTime
                    
                    'Check if successfull
                    'Close brower window, if successfull #ToDo#
                    
                    'Copy picture to upload directory
                    If Not Simulate Then
                        On Error Resume Next
                        FileCopy FilePathName, UploadFilePathName
                        If Err.Number > 0 Then
                            Err.Clear
                            Application.Activate
                            MsgBox "Can not copy " & FilePathName & " to upload directory. Please check if file is in use (browser and picture viewers)!", vbExclamation + vbOKOnly, ConverterPrgTitle
                            'Try again
                            FileCopy FilePathName, UploadFilePathName
                            If Err.Number > 0 Then
                                MsgBox "Still can not copy. Program will continue without copying file.", vbInformation, ConverterPrgTitle
                            End If
                            Err.Clear
                        End If
                        On Error GoTo Err_MW_ImageUpload_File
                    End If
                    
'                    'Move picture to upload directory
'                    On Error Resume Next
'                    If FileExists(UploadFilePathName) Then KillToBin UploadFilePathName
'                    Name FilePathName As UploadFilePathName
'                    If Err.Number > 0 Then
'                        Err.Clear
'                        Application.Activate
'                        MsgBox "Can not move " & FilePathName & " to upload directory. Please check if file is in use (browser and picture viewers)!", vbExclamation + vbOKOnly, ConverterPrgTitle
'                        'Try again
'                        Name FilePathName As UploadFilePathName
'                        If Err.Number > 0 Then
'                            MsgBox "Still can not move. Program will continue without movement of file.", vbInformation, ConverterPrgTitle
'                            FileCopy FilePathName, UploadFilePathName
'                        End If
'                        Err.Clear
'                    End If
                    
                    LastCheck = Timer
                    Exit Do
                ElseIf c > 3 Then
                    'ups, we did not find the browser window to upload!
                    Application.Activate
                    MsgBox "I could not identify the browser window. Maybe you have chosen the wrong language. Set WikiUploadTitle correctly." & vbCrLf & "Uploading will be disabled.", vbExclamation, ConverterPrgTitle
                    SetReg "ImageUploadAuto", False
                    MW_ImageUpload_File = -1
                    Exit Do
                End If
            Loop Until c > 4
        End If
    
Exit_MW_ImageUpload_File:
    Exit Function

Err_MW_ImageUpload_File:
    MW_ImageUpload_File = Err.Number
    DisplayError "MW_ImageUpload_File"
    Resume Exit_MW_ImageUpload_File
End Function

Public Function MW_Initialize(Optional Force As Boolean) As Boolean
' -------------------------------------------------------------------
' Function: Set Variables and do some checks
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: Force: Check again
'
' returns: true if ok
'
' released: June 13, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_Initialize

    Static RegSaveAsked As Boolean
    Dim Answer&, p&
    
    'since the variables keep their values until word is closed, some need to be updated in every run
    MW_LanguageTexts
    
    If GetReg("isCustomized") = False Then
        'customize first
        isInitialized = True
        frmW2MWP_Config.Show
        If Not isInitialized Then Exit Function
        isInitialized = False
    End If
               
    'ArticleName
    If Documents.Count = 0 Then Exit Function
    DocInfo.ArticleName = DocInfo.DocNameNoExt

    'internal variables
    convertImagesOnly = False
    
    If Force Then isInitialized = False
    If isInitialized Then MW_Initialize = True: Exit Function
    
'--- check only once ---
    
    'Check Word Version
    If MW_WordVersion <= 1997 Then
        'Word '97
        WordParagraph = "^a"
        WordNewLine = "^z"
        'not supported
        MsgBox "Sorry, you need Word 2000 or above to run this converter!", vbCritical
        Exit Function
    Else
        'Word 2000 and above
        WordParagraph = "^p"
        WordNewLine = "^l"
        WordForceBlank = "^s"
    End If
    WordManualPageBreak = "^m"
    
    'Check if the image path exists
    Dim ImagePathName$
    ImagePathName = IIf(GetReg("ImagePath") <> "", GetReg("ImagePath"), ActiveDocument.Path)
    ImagePathName = FormatPath(ImagePathName)
    If Not DirExistsCreate(ImagePathName) Then
        SetReg "ImageExtraction", False 'Create Dir, if it does not exist
        Answer = MsgBox("Your image path does not exist!" & vbCr & ImagePathName & vbCr & vbCr & "Create the path and configure Word2MediaWikiPlus!" & vbCr & vbCr, vbCritical, ConverterPrgTitle)
        If Answer <> vbYes Then Exit Function 'always exit
    End If

    'Assume the position of Microsoft Photo Editor
    EditorPath = MW_GetEditorPath
    'If EditorPath = "" Then Exit Function 'blank EditorPath disables image conversion

    isInitialized = True
    MW_Initialize = True

Exit_MW_Initialize:
    Exit Function

Err_MW_Initialize:
    DisplayError "MW_Initialize"
    Resume Exit_MW_Initialize
End Function

Private Sub MW_InsertPageHeaders()
' -------------------------------------------------------------------
' Function: inserts and copys page header and footer
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: June 11, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_InsertPageHeaders

    Dim pg As Paragraph
    Dim s&, H&, h2&, headerC&
    Dim TOCinserted As Boolean, NeedsPageTitle As Boolean

    ' insert Titlepage if there is text before first heading
    If GetReg("insertTitlePageIfNeeded") Then
        'only if there are headings
        For Each pg In ActiveDocument.Paragraphs
            Select Case pg.Style
                Case ActiveDocument.Styles(wdStyleHeading1), ActiveDocument.Styles(wdStyleHeading2), ActiveDocument.Styles(wdStyleHeading3)
                    'Heading is ok, TOC will be on top of page
                    Exit For
                Case Else
                    NeedsPageTitle = True
                    Exit For
            End Select
        Next
    End If
    
    Selection.HomeKey Unit:=wdStory
    
    'insert page headers
    If GetReg("convertPageHeaders") Then
        'What headers do we have?
        For s = 1 To ActiveDocument.Sections.Count
            H = 0
            Do
                If H = 0 Then
                    H = 2
                ElseIf H = 1 Then
                    H = 3
                ElseIf H = 2 Then
                    H = 1
                End If
                If ActiveDocument.Sections(s).Headers(H).Exists Then
                    If Len(ActiveDocument.Sections(s).Headers(H).Range.Text) > 1 Then
                        'convert shapes first
                        MediaWikiExtract_Images 1, s, H
                        ActiveDocument.Sections(s).Headers(H).Range.Copy
                        headerC = headerC + 1
                        With Selection
                            If TOCinserted = False Then
                                .InsertAfter "__TOC__" & vbCr & "----" & vbCr
                                TOCinserted = True
                            End If
                            .InsertAfter GetReg("txt_PageHeader") & " " & headerC & vbCr & "----" & vbCr
                            .Style = wdStyleNormal
                            '.Style = wdStyleHeading1
                            Selection.Collapse wdCollapseEnd
                            Selection.Paste
                            Selection.Collapse wdCollapseEnd
                            .InsertAfter "----" & vbCr
                            .Style = wdStyleNormal
                            '.ClearFormatting 'not in Word 2000
                            Selection.Collapse wdCollapseEnd
                        End With
                        ActiveDocument.Sections(s).Headers(H).Range.Delete
                    End If
                End If
            Loop Until H = 3
        Next s
    End If 'headers
    
    ' insert TOC
    
    If TOCinserted = False And NeedsPageTitle Then
        'only if there are headings
        With Selection
            .InsertAfter "__TOC__" & vbCr
            .InsertAfter "----" & vbCr
            .Style = wdStyleNormal
            '.ClearFormatting
            .Collapse wdCollapseEnd
        End With
    End If
    
    ' insert Titlepage if there is text before first heading
    If NeedsPageTitle Then
        With Selection
            .InsertAfter GetReg("txt_TitlePage") & vbCr & "----" & vbCr
            .Style = wdStyleNormal
            '.ClearFormatting
        End With
    End If
    
    'insert page footers
    If GetReg("convertPageFooters") Then
        'What headers do we have?
        Selection.EndKey Unit:=wdStory
        Selection.InsertAfter vbCr
        Selection.Style = wdStyleNormal
        'Selection.ClearFormatting
        Selection.Collapse wdCollapseEnd
        headerC = 0
        For s = 1 To ActiveDocument.Sections.Count
            H = 0
            Do
                If H = 0 Then
                    H = 2
                ElseIf H = 1 Then
                    H = 3
                ElseIf H = 2 Then
                    H = 1
                End If
                If ActiveDocument.Sections(s).Footers(H).Exists Then
                    If Len(ActiveDocument.Sections(s).Footers(H).Range.Text) > 1 Then
                        'convert shapes first
                        MediaWikiExtract_Images 1, s, H
                        ActiveDocument.Sections(s).Footers(H).Range.Copy
                        headerC = headerC + 1
                        With Selection
                            .InsertAfter "----" & vbCr & GetReg("txt_Pagefooter") & " " & headerC & vbCr & "----" & vbCr
                            .Style = wdStyleNormal
                            '.ClearFormatting
                            Selection.Collapse wdCollapseEnd
                            Selection.Paste
                            Selection.Collapse wdCollapseEnd
                        End With
                        ActiveDocument.Sections(s).Footers(H).Range.Delete
                    End If
                End If
            Loop Until H = 3
        Next s
    End If 'footers
    
Exit_MW_InsertPageHeaders:
    Exit Sub

Err_MW_InsertPageHeaders:
    DisplayError "MW_InsertPageHeaders"
    Resume Exit_MW_InsertPageHeaders
End Sub

Public Sub MW_LanguageTexts(Optional Force As Boolean = False)
' -------------------------------------------------------------------
' Function: Localization Texts for the macro
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing, sets a lot of modul variables
'
' released: Nov. 23, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_LanguageTexts

    Static LastLanguage$
    
    Dim i&

    'Fill language array with available languages
    ReDim languageArr(1 To 6, 1 To 6)
    i = 1
    languageArr(i, 1) = "GER"
    languageArr(i, 2) = "Deutsch"
    i = i + 1
    languageArr(i, 1) = "ENG"
    languageArr(i, 2) = "English"
    i = i + 1
    languageArr(i, 1) = "ESP"
    languageArr(i, 2) = "Español"
    i = i + 1
    languageArr(i, 1) = "FRA"
    languageArr(i, 2) = "Français"
    i = i + 1
    languageArr(i, 1) = "NDL"
    languageArr(i, 2) = "Nederlands"
    i = i + 1
    languageArr(i, 1) = "RUS"
    languageArr(i, 2) = "Russian"

    'Messages
    If LastLanguage = GetReg("Language") Then
        If Not Force Then Exit Sub
    ElseIf LastLanguage <> "" Then
        SetReg "CategoryImagePreFix", ""
    End If
    
    'delete language entries, so they can be set freshly
    'If Delete Then
        
'        SetReg "WikiSearchTitle", ""
'        SetReg "WikiCategoryKeyWord", ""
'        SetReg "WikiUploadTitle", ""
'        SetReg "EditorKeyLoadPic", ""
'        SetReg "EditorKeySavePic", ""
'        SetReg "EditorKeyPastePicAsNew", ""
'        SetReg "EditorPaletteKey", ""
'        SetReg "ClickChartText", ""
'        SetReg "UnableToConvertMarker", ""
'
'        SetReg "txt_TitlePage", ""
'        SetReg "txt_PageHeader", ""
'        SetReg "txt_PageFooter", ""
'        SetReg "txt_Footnote", ""
    
    'End If
    
    LastLanguage = GetReg("Language")
    Select Case LastLanguage
        'description look in ENG
        
        Case "ENG" 'english
        
            'SendKeys
            SetReg "EditorKeyLoadPic", "^o"              'Ctrl+o for open
            SetReg "EditorKeySavePic", "%fa"             'File save as
            SetReg "EditorKeyPastePicAsNew", "%en"       'paste as new
            SetReg "EditorPaletteKey", "p"               'change GIF Palette
        
            'just fill the registry with defaults, if no value is present
            'these must be present in setregValidate for the people that turn of writing back to registry (Demo document)
            SetReg "WikiSearchTitle", "Search -"         'Title of browser window after search in wiki
            SetReg "WikiCategoryKeyWord", "category"     'category key word according to your language, "category" works always
            If GetReg("CategoryImagePreFix") = "not set" Then SetReg "CategoryImagePreFix", "Images "      'Standard Prefix to the image category, add blank at last character to separate words
            SetReg "WikiUploadTitle", "Upload"           'Title of browser window when uploading to wiki
            SetReg "ClickChartText", "click me!"         'clickable charts will have additional text in wiki: "click me!"
            SetReg "UnableToConvertMarker", "## Error converting ##: " 'some links can not be converted, give hint
            
            SetReg "txt_TitlePage", "Title page"
            SetReg "txt_PageHeader", "Page header"
            SetReg "txt_PageFooter", "Page footer"
            SetReg "txt_Footnote", "Footnotes"
        
            'Messages; will not be stored in registry
            Msg_Upload_Info = "Now the image file upload will begin. Before you start you need to set your browser right:" & vbCr & vbCr & _
            "1. Close all sidebars like favorites." & vbCr & vbCr & _
            "2. Sign in into your wiki." & vbCr & vbCr & _
            "Do not click ok before you checked this."
            Msg_Finished = "Converting finished. Paste your clipboard contents into your wiki editor."
            Msg_NoDocumentLoaded = "No document was loaded."
            Msg_LoadDocument = "Please load the document to convert."
            Msg_CloseAll = "Please close all documents but the one you want to convert! The macro will stop now."
        
        Case "GER" 'german
            
            'SendKeys
            SetReg "EditorKeyLoadPic", "^o"
            SetReg "EditorKeySavePic", "%du"
            SetReg "EditorKeyPastePicAsNew", "%bn"
            SetReg "EditorPaletteKey", "p"
            
            SetReg "WikiSearchTitle", "Suchergebnisse -"
            SetReg "WikiCategoryKeyWord", "Kategorie"
            If GetReg("CategoryImagePreFix") = "not set" Then SetReg "CategoryImagePreFix", "Bilder "
            SetReg "WikiUploadTitle", "Hochladen"
            SetReg "ClickChartText", "Klick mich!"
            SetReg "UnableToConvertMarker", "## Fehler bei Konvertierung ##: "
        
            SetReg "txt_TitlePage", "Titelblatt"
            SetReg "txt_PageHeader", "Kopfzeile"
            SetReg "txt_PageFooter", "Fußzeile"
            SetReg "txt_Footnote", "Fußnoten"
    
            'Messages
            Msg_Upload_Info = "Jetzt werden die Bilder hochgeladen. Vorher muss der Browser korrekt eingestellt sein, damit es funktioniert:" & vbCr & vbCr & _
            "1. Schließen Sie alle Seitenleisten wie z.B. die Favoriten." & vbCr & vbCr & _
            "2. Melden Sie sich an Ihrem Wiki an." & vbCr & vbCr & _
            "Klicken Sie erst OK wenn Sie dies durchgeführt haben."
            Msg_Finished = "Konvertierung beendet. Fügen Sie die Daten aus der Zwischenablage in Ihr Wiki ein."
            Msg_NoDocumentLoaded = "Es wurde kein Dokument geladen."
            Msg_LoadDocument = "Bitte laden Sie das zu konvertierende Dokument."
            Msg_CloseAll = "Bitte schließen Sie alle Dokumente bis auf das, welches Sie konvertieren möchten! Das Makro wird jetzt beendet."
            
        Case "FRA" 'french
            
            'SendKeys
            SetReg "EditorKeyLoadPic", "^o"              'Ctrl+o pour ouvrir
            SetReg "EditorKeySavePic", "%fa"             'Sauvegarder sous
            SetReg "EditorKeyPastePicAsNew", "%en"       'coller en tant que nouveau
            SetReg "EditorPaletteKey", "p"               'changer la palette GIF
            
            'Juste mettre des valeurs pas défaut dans la base de registres, si aucune valeur n'est présente
            'celles-ci doivent être présentes dans setregValidate pour ceux qui empêchent l'écriture dans la base de registres (Document de démo)
            SetReg "WikiSearchTitle", "Rechercher -"                       'Titre de la fenêtre du navigateur après une recherche dans le wiki
            SetReg "WikiCategoryKeyWord", "catégorie"                   'mot-clé de catégorie d'après votre langue, "category" fonctionne toujours
            If GetReg("CategoryImagePreFix") = "not set" Then SetReg "CategoryImagePreFix", "Images "                    'Préfixe standard de la catégorie des images, ajoutez une espace pour séparer les mots
            SetReg "WikiUploadTitle", "Télécharger"                         'Titre de la fenêtre du navigateur pendant le téléchargement vers le wiki
            SetReg "ClickChartText", "cliquez moi !"                       'Les graphes cliquables dans le wiki auront un nouveau texte : "cliquez moi !"
            SetReg "UnableToConvertMarker", "## Erreur de conversion de ##: " 'certains liens ne peuvent pas être convertis, donnez un indice
            
            SetReg "txt_TitlePage", "Page de titre"
            SetReg "txt_PageHeader", "En-tête de page"
            SetReg "txt_PageFooter", "Pied de page"
            SetReg "txt_Footnote", "Notes de bas de page"
            
            'Messages; ne seront pas sauvegardés dans la base de registres
            Msg_Upload_Info = "Maintenant le fichier image va être téléchargé. Avant de commencer, vous devez organiser correctement votre navigateur :" & vbCr & vbCr & _
            "1. Fermez toutes les fenêtres latérales comme celle des favoris." & vbCr & vbCr & _
            "2. Identifiez vous dans votre wiki." & vbCr & vbCr & _
            "Ne cliquez pas sur OK avant d'avoir vérifé tout cela."
            Msg_Finished = "Conversion terminée. Collez le contenu du bloc-note dans l'éditeur du wiki."
            Msg_NoDocumentLoaded = "Aucun document n'a été chargé."
            Msg_LoadDocument = "Veuillez charger le document à convertir."
            Msg_CloseAll = "Veuillez fermer tous les documents sauf celui que vous souhaitez convertir ! La macro va maintenant s'arrêter."
    
        Case "ESP" 'spanish
    
            'SendKeys
            SetReg "EditorKeyLoadPic", "^a"              'Ctrl+a para abrir
            SetReg "EditorKeySavePic", "%fa"             'Guardar Como
            SetReg "EditorKeyPastePicAsNew", "%en"       'Pegar como neuvo
            SetReg "EditorPaletteKey", "p"               'cambiar palette GIF
                 
            'terraplén justo el registro con defectos, si no hay valor presente
            'éstos deben estar presentes en setregValidate para la gente que vuelta de escribir de nuevo al registro (el documento de la versión parcial de programa)
            SetReg "WikiSearchTitle", "Buscar -"            'Título de la ventana de browser después de la búsqueda en wiki
            SetReg "WikiCategoryKeyWord", "Categoría"       'La palabra clave de la categoría según tu lengua, “categoría” trabaja siempre
            If GetReg("CategoryImagePreFix") = "not set" Then SetReg "CategoryImagePreFix", "Imagen "         'El prefijo estándar a la categoría de la imagen, agrega el espacio en blanco en el carácter pasado a las palabras separadas
            SetReg "WikiUploadTitle", "Subir un Archivo"    'Título de la ventana de browser al uploading al wiki
            SetReg "ClickChartText", "¡Chascarme!"          'las cartas clickable tendrán texto adicional en wiki: ¡“chascarme!"
            SetReg "UnableToConvertMarker", "## Error de convertir ##: " 'Mensajes; ningún registro del en del almacenado del será
                      
            SetReg "txt_TitlePage", "Portada"
            SetReg "txt_PageHeader", "Ecabezado"
            SetReg "txt_PageFooter", "Pie de página"
            SetReg "txt_Footnote", "Nota al pie"
                  
            'Mensajes; no será almacenado en registro
            Msg_Upload_Info = "Ahora el upload del archivo de la imagen comenzará. Antes de que te comiences necesidad de fijar la tu derecha del browser:" & vbCr & vbCr & _
            "1. Cerrar todo barras laterales como marcadores." & vbCr & vbCr & _
            "2. Registrarse a su wiki." & vbCr & vbCr & _
            "No tecleo acepter antes de comprobar esto."
            Msg_Finished = "Convertier Listo. Pegar su contentido portapapeles en su editor de wiki."
            Msg_NoDocumentLoaded = "No se cargó ningún documento."
            Msg_LoadDocument = "Cargar por favor el documento para convertir."
            Msg_CloseAll = "¡Cerrar por favor todos los documentos pero el que deseas convertir! La macro ahora parará."
    
    
        Case "NDL" 'dutch, Nederlands
            
            'SendKeys
            SetReg "EditorKeyLoadPic", "^o"              'Ctrl+o om te openen
            SetReg "EditorKeySavePic", "%fa"             'Bestand opslaan als
            SetReg "EditorKeyPastePicAsNew", "%en"       'plakken als nieuw
            SetReg "EditorPaletteKey", "p"               'verander GIF Palet
                 
            'vul de registry in met de defaults, als er geen waarde is
            'deze moeten aanwezig zijn in de setregValidate voor mensen die het schrijven naar de registry uit hebben staan(Demo document)
            SetReg "WikiSearchTitle", "Zoeken -"                       'Titel of het browser venster na het zoeken in wiki
            SetReg "WikiCategoryKeyWord", "categorie"                   'categorie sleutel woord naar gelang de taal, "category" werkt altijd
            If GetReg("CategoryImagePreFix") = "not set" Then SetReg "CategoryImagePreFix", "Afbeelding "                    'Standaard Prefix van de afbeelding categorie, voeg een spatie toe aan het laatste karakter om de worden te scheiden
            SetReg "WikiUploadTitle", "Upload"                         'Titel van het browser venster tijdens het uploaden naar wiki
            SetReg "ClickChartText", "klik mij!"                       'Grafieken die aangeklikt kunnen worden in wiki hebben een extra tekst: "Klik mij!"
            SetReg "UnableToConvertMarker", "## converteer Error ##: " 'sommige links kunnen niet geconverteerd  worden, geef een hint
                      
            SetReg "txt_TitlePage", "Hoofdpagina"
            SetReg "txt_PageHeader", "Koptekst"
            SetReg "txt_PageFooter", "Voettekst"
            SetReg "txt_Footnote", "Voetnoten"
                  
            'Messages; zullen niet worden opgeslagen in de registry
            Msg_Upload_Info = "Nu zal de afbeelding upload beginnen. Voordat je begint moet je je internet browser goed zetten:" & vbCr & vbCr & _
            "1. SLuit alle zijbalken zoals bladwijzers of favorieten." & vbCr & vbCr & _
            "2. Meld je aan bij je wiki." & vbCr & vbCr & _
            "Klik niet op ok voordat je dit gecontroleerd hebt."
            Msg_Finished = "Converteren klaar. Plak de klembord inhoud in je wiki editor."
            Msg_NoDocumentLoaded = "Het document is niet geladen."
            Msg_LoadDocument = "AUB laad de documenten om te converteren."
            Msg_CloseAll = "AUB sluit alle documenten behalve degene die je wilt converteren! De macro zal nu stoppen."
        
        Case "RUS" 'russian
            'unicode problem in german word editor
            'setreg does not work either...
            'no russian for now
            'maybe this works though
            
            'SendKeys
            SetReg "EditorKeyLoadPic", "^o"              'Ctrl+o for open
            SetReg "EditorKeySavePic", "%fa"             'File save as
            SetReg "EditorKeyPastePicAsNew", "%en"       'paste as new
            SetReg "EditorPaletteKey", "p"               'change GIF Palette
                 
            'ïðîñòî îñòàâòå çíà÷åíèÿ ðåãèñòðà ïî óìîë÷àíèþ, åñëè çíà÷åíèÿ íå îïðåäåëåíû
            SetReg "WikiSearchTitle", "Ðåçóëüòàòû ïîèñêà -"                       'Title of browser window after search in wiki
            SetReg "WikiCategoryKeyWord", "Êàòåãîðèè -"                   'category key word according to your language, "category" works always
            If GetReg("CategoryImagePreFix") = "not set" Then SetReg "CategoryImagePreFix", "Images "                    'Standard Prefix to the image category, add blank at last character to separate words
            SetReg "WikiUploadTitle", "Çàãðóçèòü ôàéë -"                         'Title of browser window when uploading to wiki
            SetReg "ClickChartText", "click me!"                       'clickable charts will have additional text in wiki: "click me!"
            SetReg "UnableToConvertMarker", "## Error converting ##: " 'some links can not be converted, give hint
                      
            SetReg "txt_TitlePage", "Çàãëàâíàÿ ñòðàíèöà"
            SetReg "txt_PageHeader", "Çàãîëîâîê ñòðàíèöû"
            SetReg "txt_PageFooter", "Page footer"
            SetReg "txt_Footnote", "Ïðèìå÷àíèÿ"
                  
            'Ñîîáùåíèÿ íå ñîõðàíÿåìûå â ðååñòðå
            Msg_Upload_Info = "Ñåé÷àñ íà÷íåòñÿ çàãðóçêà ñîîáùåíèé. Ïåðåä íà÷àëîì ïðàâèëüíî íàñòðîéòå îêíî ïðîâîäíèêà Èíòåðíåò:" & vbCr & vbCr & _
            "1. Çàêðîéòå âñå ïàíåëè âðîäå Çàêëàäêè." & vbCr & vbCr & _
            "2. Âîéäèòå ïîä ñâîèì ëîãèíîì â Wiki." & vbCr & vbCr & _
            "Íå íàæèìàéòå ÎÊ ïîêà íå ñäåëàåòå ýòî!"
            Msg_Finished = "Êîíâåðòàöèÿ çàâåðøåíà.Âñòàâòå òåêñò èç áóôåðà îáìåíà â ðåäàêòîð WiKi"
            Msg_NoDocumentLoaded = "Äîêóìåíòû íå çàãðóæåíû."
            Msg_LoadDocument = "Ïîæàëóéñòà çàãðóçèòå äîêóìåíò â êîíâåðòåð."
            Msg_CloseAll = "Ïîæàëóéñòà çàêðîéòå âñå äîêóìåíòû êðîìå òîãî êîòîðûé íóæíî ñêîíâåðòèðîâàòü! Ìàêðîñ îñòàíîâëåí."
    
    End Select

Exit_MW_LanguageTexts:
    Exit Sub

Err_MW_LanguageTexts:
    DisplayError "MW_LanguageTexts"
    If DebugMode Then Stop: Resume Next
    Resume Exit_MW_LanguageTexts
End Sub

Public Sub MW_PhotoEditor_Convert(Action$, ByVal FilePathName$, Optional Simulate As Boolean = False, Optional PxWidth& = 0, Optional PxHeight& = 0)
' -------------------------------------------------------------------
' Function: converts an image from clipboard or file to desired image format file
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: Action: Paste or Load, FileName, ImageSave...
'
' returns: nothing
'
' released: June 12, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_PhotoEditor_Convert
        
    Dim txt$

    'Convert images
    If FileExists(FilePathName) Or Simulate Or Action = "Paste" Or Action = "PastePixel" Then
        
        'Make sure Photo Editor is activated

           If Not AppActivatePlus(EditorTitle, False) Then         'We need to open it
            'Open Photo Editor
            If EditorPath = "" Then MW_GetEditorPath 'happens only in Debug
            Shell EditorPath, vbMaximizedFocus
            DoEvents
            Sleep 1500
            'Try again
            If Not AppActivatePlus(EditorTitle, False) Then
                MsgBox "Error: Could not activate your Photo Editor", vbCritical
                Exit Sub
            End If
            DoEvents
            Sleep 500
        End If
        
        'SendKeys
        Select Case Action
            Case "Paste"
                'Paste Picture
                SendKeys GetReg("EditorKeyPastePicAsNew"), True
                DoEvents
                Sleep 1500 'maybe greater if loading takes more time
            Case "PastePixel"
                'Paste Picture
                SendKeys "^n", True
                SendKeys "{TAB}{TAB}{TAB}{TAB}p", True
                SendKeys "+{TAB}", True
                SendKeys "+{TAB}", True
                SendKeys PxWidth & "{TAB}", True
                SendKeys PxHeight & "{Enter}", True
                DoEvents
                SendKeys "^v" 'paste
                DoEvents
                Sleep 1500 'maybe greater if loading takes more time
            Case "Load"
                'Load Picture
                SendKeys GetReg("EditorKeyLoadPic"), True
                SendKeys FilePathName, True
                SendKeys "{Enter}", True
                Sleep 2000 'maybe greater if loading takes more time
            Case Else
                MsgBox "Program error in MW_PhotoEditor_Convert", vbCritical
                Exit Sub
        End Select
        
        'Save Picture in different format
        If ImageSaveJPG Then
            'Save JPG
            SendKeys GetReg("EditorKeySavePic"), True
            SendKeys "+{END}{DEL}", True
            SendKeys MW_ImagePathName(FilePathName, "jpg", True), True
            SendKeys "{TAB}j", True 'filetype
            SendKeys "{Enter}", True
            Sleep 1000
        End If
        
        If ImageSavePNG Then
            'Save PNG
            SendKeys GetReg("EditorKeySavePic"), True
            SendKeys "+{END}{DEL}", True
            SendKeys MW_ImagePathName(FilePathName, "png", True), True
            SendKeys "{TAB}j", True 'filetype jpg, so it will not become pix
            SendKeys "p", True 'filetype
            SendKeys "{Enter}", True
            Sleep 1000
        End If
        
        If ImageSaveBMP Then
            'Save BMP
            SendKeys GetReg("EditorKeySavePic"), True
            SendKeys "+{END}{DEL}", True
            SendKeys MW_ImagePathName(FilePathName, "bmp", True), True
            SendKeys "+{END}{DEL}", True
            SendKeys "{TAB}w", True 'filetype
            SendKeys "{Enter}", True
            Sleep 1000
        End If
        
        If ImageSaveGIF Then
            'Save GIF
            'Will change picture, so must be in the end
            'change colors
            SendKeys "%{Enter}", True
            SendKeys GetReg("EditorPaletteKey") & "{Enter}", True
            
            SendKeys GetReg("EditorKeySavePic"), True
            SendKeys MW_ImagePathName(FilePathName, "gif", True), True
            SendKeys "{TAB}g", True 'filetype
            SendKeys "{Enter}", True
            Sleep 1000
        End If
        
        'Close picture
        SendKeys "^{F4}", True
        DoEvents
        Sleep 200
        
    End If
        
Exit_MW_PhotoEditor_Convert:
    Exit Sub

Err_MW_PhotoEditor_Convert:
    DisplayError "MW_PhotoEditor_Convert"
    Resume Exit_MW_PhotoEditor_Convert
End Sub

Sub MW_PowerpointQuit()

    If Not DocInfo.PowerpointStarted Then Exit Sub

    Dim ppt As Object
    
    Set ppt = CreateObject("Powerpoint.Application")
    ppt.Visible = True
    
    If ppt.Presentations.Count = 0 Then ppt.Quit
    
    Set ppt = Nothing
    
    DocInfo.PowerpointStarted = False

End Sub

Public Sub MW_ReplaceString(findStr As String, replacementStr As String, Optional ByVal Repeat As Boolean = False, Optional Wrap& = wdFindContinue, Optional MaxRepeat& = 100)
' -------------------------------------------------------------------
' Function: replaces text in the whole document (replace all)
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: Wrap = wdFindStop replaces only all findings in selection
'        MaxRepeat: If a search must be repeated it may go into an endless loop, here it is stopped
'
' returns: nothing
'
' released: Nov 6, 2006
' -------------------------------------------------------------------
On Error Resume Next
    'ActiveDocument.Range.Text = Replace(ActiveDocument.Range.Text, findStr, replacementStr, , , vbTextCompare)
    'does not work, because ^p etc work diffent
    
    Dim c&
    Dim rg As Range
    
    MW_Statusbar True, "replacing " & findStr & " with " & replacementStr
    
    If Wrap = wdFindContinue Then
        Set rg = ActiveDocument.Range
    Else
        Set rg = Selection.Range
    End If
    
    Do
        With rg.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = findStr
            .Replacement.Text = replacementStr
            .Forward = True
            .Wrap = Wrap 'wdFindContinue means whole document
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            .Execute Replace:=wdReplaceAll 'In Word 2003 the active document will change!!! Bug.
            
        End With
        
        If Repeat Then 'might lead to an endless loop, lets see
            rg.Find.ClearFormatting
            rg.Find.Replacement.ClearFormatting
            With rg.Find
                .Text = findStr
                .Replacement.Text = replacementStr
                .Forward = True
                .Wrap = Wrap 'wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Repeat = rg.Find.Execute
            c = c + 1 'TimeOut
        End If
        DoEvents
    Loop Until Not Repeat Or c = MaxRepeat
'    If c = MaxRepeat And DebugMode Then
'        MsgBox "Replacement of " & findStr & " could not be completed. Mosty harmless, makro will continue.", vbInformation
'    End If
        
End Sub

Private Sub MW_ScaleMax(MyIS As InlineShape)
' -------------------------------------------------------------------
' Function: try to max picture to 100%, but limit size to 800x600 or word will not copy correctly
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: InlineShape
'
' returns: nothing
'
' released: Nov. 09, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_ScaleMax

    Dim pWidth&, pHeight&, Factor#, r#, ScaleW#, ScaleH#
    Dim CropBottom#, CropLeft#, CropTop#, CropRight#
    Dim cLR#, cTB#
    Dim MaxWidth&, MaxHeight&
    
    MaxWidth = PixelsToPoints(800)
    MaxHeight = PixelsToPoints(600) 'Limitations of MS Word or MS Photo Editor
    
    'Word Bug: wrong values for ScaleHeight if cropped
    'So we set cropped = 0, but use cropped values to calculate size
    
    MW_GetScaleIS MyIS, ScaleW, ScaleH
    
    If ScaleH = 0 Then ScaleH = 100
    If ScaleW = 0 Then ScaleW = 100
    
    With MyIS
        
        'Check if > 100% 'why?
        'If .Height / ScaleH * 100 < MaxHeight And .Width / ScaleW * 100 < MaxWidth Then Exit Sub
        
        pWidth = .Width
        pHeight = .Height
        
        'we lower MaxSizes to limit either direction to 100%
        If .Height / ScaleH * 100 < MaxHeight Then MaxHeight = .Height / ScaleH * 100
        If .Width / ScaleW * 100 < MaxWidth Then MaxWidth = .Width / ScaleW * 100
        
        'calc Max Size
        If pWidth * MaxHeight / pHeight < MaxWidth Then
            .Height = MaxHeight
            .Width = pWidth * MaxHeight / pHeight
        Else
            .Width = MaxWidth
            .Height = pHeight * MaxWidth / pWidth
        End If
        
    End With
        
Exit_MW_ScaleMax:
    Exit Sub

Err_MW_ScaleMax:
    DisplayError "MW_ScaleMax"
    If DebugMode Then Stop: Resume Next
    Resume Exit_MW_ScaleMax
End Sub

Private Function MW_ScaleMaxOK(MyIS As InlineShape, Optional MaxWidth& = 800) As Boolean
' -------------------------------------------------------------------
' Function: checks, if a picture exceeds MS-Word size limits, if it would be expanded to 100%
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: InlineShape
'
' returns: true or false
'
' released: Nov 7, 2006
' -------------------------------------------------------------------

    Dim pWidth&, pHeight&
    Const MaxHeight& = 600 'Limitations of MS Word
    'If MaxWidth = 0 Then MaxWidth = 800
    
    'calculate points at 100%
    pWidth = PointsToPixels(MyIS.Width) * (100 / MyIS.ScaleWidth)
    pHeight = PointsToPixels(MyIS.Height) * (100 / MyIS.ScaleHeight)
    
    If pHeight > MaxHeight Or pWidth > MaxWidth Then
        MW_ScaleMaxOK = False
    Else
        MW_ScaleMaxOK = True
    End If

End Function

Public Function MW_SearchAddress(Optional ByVal AddressRoot$ = "") As String
'returns the search path for MediaWiki

    Const WikiSearch$ = "Special:Search?search=" 'all languages
    Dim p&
    
    If AddressRoot = "" Then AddressRoot = WikiAddressRoot
    
    If right$(AddressRoot, 1) = "/" Then
        MW_SearchAddress = AddressRoot & WikiSearch
    Else
        p = InStrRev(AddressRoot, "?")
        If p > 0 Then
            MW_SearchAddress = Left$(AddressRoot, p) & "search="
        End If
    End If

    If MW_SearchAddress = "" Then MW_SearchAddress = AddressRoot & WikiSearch

End Function

Public Sub MW_SetWikiAddressRoot(Optional url$ = "")

    If url = "" Then
        WikiAddressRoot = GetReg("WikiAddressRoot" & GetReg("WikiSystem"))
        If WikiAddressRoot = "" Then
            'MsgBox "Could not retrieve your wiki URL", vbExclamation, ConverterPrgTitle
            WikiAddressRoot = GetReg("WikiAddressRootTest")
        End If
    Else
        WikiAddressRoot = url
    End If

End Sub

Private Sub MW_ChangeView(MyView&)
'Changes the view of the document (header, footer, page view)

    Select Case MyView
        Case 0 ' Normal view
            If ActiveWindow.View.SplitSpecial = wdPaneNone Then
                ActiveWindow.ActivePane.View.Type = wdNormalView
            Else
                ActiveWindow.View.Type = wdPrintView
            End If
            'ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
            Exit Sub
        
        Case 1, 2 ' Header
            If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
                ActiveWindow.Panes(2).Close
            End If
            If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow.ActivePane.View.Type = wdOutlineView Then
                ActiveWindow.ActivePane.View.Type = wdPrintView
            End If
        
        Case 4 ' Print / Layout view
            If ActiveWindow.View.SplitSpecial = wdPaneNone Then
                ActiveWindow.ActivePane.View.Type = wdPrintView
            Else
                ActiveWindow.View.Type = wdPrintView
            End If
            ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
            Exit Sub
    
    End Select

    Select Case MyView
        Case 1, 2 ' Header
            ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        
        Case 2 ' Footer
            ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
        
    End Select

End Sub

'Public Function MW_SnagIt_Check_Installed(CreateRef As Boolean) As Boolean
'' -------------------------------------------------------------------
'' Function: Checks if SnagIt is referenced, tries to reference
'' copyright by Gunter Schmidt www.beadsoft.de
'' Released under GPL
''
'' Input: nothing
''
'' returns: true, if SnagIt is referenced
''
'' released: June 04, 2006
'' -------------------------------------------------------------------
'On Error GoTo Err_MW_SnagIt_Check_Installed
'
'    Dim myRef As Object
'
'    For Each myRef In NormalTemplate.VBProject.References
'        'Debug.Print myRef.Name
'        If myRef.Name = "SNAGITLib" Then
'            'Debug.Print "SNAGITLib", myRef.GUID, myRef.major, myRef.minor; ""
'            'If SnagIt is installed, but has a differend GUID, then reference manualle and use this GUID to install
'            MW_SnagIt_Check_Installed = True
'            Exit Function
'        End If
'    Next
'
'    'Not found, lets try to install
'    If CreateRef Then
'        On Error Resume Next
'        NormalTemplate.VBProject.References.AddFromGuid "{49A23F8B-91B7-49AB-8DC3-8E4F56FCB17A}", 1, 0
'        DoEvents
'        'check if it worked
'        If Err.Number = 0 Then
'            MW_SnagIt_Check_Installed = MW_SnagIt_Check_Installed(False)
'        Else
'            Err.Clear
'            MW_SnagIt_Check_Installed = False
'        End If
'    End If
'
'Exit_MW_SnagIt_Check_Installed:
'    Exit Function
'
'Err_MW_SnagIt_Check_Installed:
'    DisplayError "MW_SnagIt_Check_Installed"
'    Resume Exit_MW_SnagIt_Check_Installed
'End Function
'
'
'Private Sub MW_SnagIt_Clipboard_to_File(ByVal FilePathName$)
'' -------------------------------------------------------------------
'' Function: Uses SnagIt to write the file from clipboard
''           some screenshots loose their color
''
'' copyright by Gunter Schmidt www.beadsoft.de
'' Released under GPL
''
'' Input: Filename
''
'' returns: nothing
''
'' released: June 04, 2006
'' -------------------------------------------------------------------
'On Error GoTo Err_SnagIt_Clipboard_to_file
'
'    Dim ImageCapture As SNAGITLib.ImageCapture
'    'You will get a compiler error if SnagIt is not installed
'    'If you have SnagIt installed, run MW_SnagIt_Check_Installed
'    Dim T As Single, p&
'
'    MW_Statusbar True, "SnagIt: converting picture " & GetFileName(FilePathName)
'
'    Set ImageCapture = CreateObject("SnagIt.ImageCapture")
'
'    With ImageCapture
'        ImageCapture.Input = siiClipboard
'        ImageCapture.Output = sioFile
'
'        .OutputImageFile.Directory = FormatPath(GetFilePath(FilePathName))
'        .OutputImageFile.FileNamingMethod = sofnmFixed
'        FilePathName = GetFileName(FilePathName)
'        p = InStrRev(FilePathName, ".")
'        If p > 0 Then FilePathName = Left$(FilePathName, p - 1)
'        .OutputImageFile.FileName = FilePathName 'SnagIt adds format extension
'
'        'save to different formats
'
'        If ImageSaveBMP Then
'            T = Timer
'            '.OutputImageFile.FileName = GetFileName(MW_ImagePathName(FilePathName, "bmp"))
'            .OutputImageFile.FileType = siftBMP
'            .OutputImageFile.FileSubType = sifstBMP_Uncompressed
'            .Capture
'            Do
'                Sleep 100
'            Loop Until .IsCaptureDone Or Timer - T > 8 'give x seconds max to save the file
'        End If
'
'        If ImageSavePNG Then
'            T = Timer
'            '.OutputImageFile.FileName = GetFileName(MW_ImagePathName(FilePathName, "png"))
'            .OutputImageFile.FileType = siftPNG
'            .Capture
'            Do
'                Sleep 100
'            Loop Until .IsCaptureDone Or Timer - T > 8 'give x seconds max to save the file
'        End If
'
'        If ImageSaveGIF Then
'            T = Timer
'            '.OutputImageFile.FileName = GetFileName(MW_ImagePathName(FilePathName, "gif"))
'            .OutputImageFile.FileType = siftGIF
'            .OutputImageFile.FileSubType = sifstGIF_NonInterlaced
'            .Capture
'            Do
'                Sleep 100
'            Loop Until .IsCaptureDone Or Timer - T > 8 'give x seconds max to save the file
'        End If
'
'        If ImageSaveJPG Then
'            T = Timer
'            '.OutputImageFile.FileName = GetFileName(MW_ImagePathName(FilePathName, "jpg"))
'            .OutputImageFile.FileType = siftJPEG
'            .OutputImageFile.Quality = 80
'            .Capture
'            Do
'                Sleep 100
'            Loop Until .IsCaptureDone Or Timer - T > 8 'give x seconds max to save the file
'        End If
'
'    End With
'
'    Debug.Print "SnagIt " & Timer - T
'
'Exit_SnagIt_Clipboard_to_file:
'    Exit Sub
'
'Err_SnagIt_Clipboard_to_file:
'    DisplayError "SnagIt_Clipboard_to_file"
'    Resume Exit_SnagIt_Clipboard_to_file
'End Sub

Private Sub MW_Statusbar(ShowGlass As Boolean, Optional Message$ = "", Optional ScreenUpdate As Boolean = True)
'gives some information to the user in the status bar

    Application.ScreenUpdating = ScreenUpdate

    If ShowGlass Then
        If Message <> "" Then Application.StatusBar = Message
        DoEvents
        system.Cursor = wdCursorWait
    Else
        system.Cursor = wdCursorNormal
    End If

End Sub

Private Sub MW_SurroundHeader(rg As Range, Level&)
' -------------------------------------------------------------------
' Function: Formats Header Wiki Markup
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input:
'
' returns: nothing
'
' released: Nov. 06, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_SurroundHeader

    Dim hLevel$, txt$
   
    hLevel = HeaderFirstLevel & String(Level - 1, "=")

    With rg
   
    'If .Tables.Count = 0 Then 'no headings in tables

        If InStr(1, rg.Text, vbCr) > 0 Then
            .Collapse
            .MoveEndUntil vbCr
        End If
        If Len(.Text) > 0 Then
            .InsertAfter hLevel
            .InsertBefore hLevel
            .Font.Reset
            HeaderCount = HeaderCount + 1
        End If
        .Style = wdStyleNormal
        
       
        'delete pagebreak
        If InStr(1, .Text, Chr(12)) > 0 Then 'faster
            .MoveStartUntil Chr(12)
            .Collapse
            .Delete
        End If
       
        'remove manual numbering of heading
        .StartOf wdParagraph
        .Collapse
        .MoveStartWhile "="
        .MoveEnd
        Do While (Asc(.Text) >= 49 And Asc(.Text) <= 57) Or .Text = "." Or .Text = " " Or .Text = vbTab 'remove 1 to 9, .

            .Move
            .Delete , -1 'only this gets blanks
            .MoveEnd
        Loop
       
        'remove remaining tabs from heading
        .MoveEndUntil vbCr
        Do While InStr(1, .Text, vbTab) > 0 'faster
            .MoveStartUntil vbTab
            .Collapse
            .Delete
            .MoveEndUntil vbCr
        Loop
       
        'check if there is still some text
        txt = Replace(.Text, "=", "")
        If txt = "" Then .InsertBefore "Header"
       
    'Else
        '.Style = wdStyleNormal
    'End If
       
    End With

Exit_MW_SurroundHeader:
    Exit Sub

Err_MW_SurroundHeader:
    DisplayError "MW_SurroundHeader"
    Resume Exit_MW_SurroundHeader
End Sub

Private Sub MW_TableInfo()
' -------------------------------------------------------------------
' Function: store some information about tables, they might get lost during conversion
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: nothing
'
' returns: nothing
'
' released: June 04, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MW_TableInfo

    Dim thisTable As Table
    Dim aCell As Cell
    Dim i&, j&, k&, cTab&, cTabN&
    Dim pageWidth#
    
    cTab = ActiveDocument.Tables.Count
    If cTab = 0 Then Exit Sub
    
    With ActiveDocument.PageSetup
        pageWidth = .pageWidth - .RightMargin - .LeftMargin
        '.LeftMargin = CentimetersToPoints(2)
        '.RightMargin = CentimetersToPoints(4)
        '.PageWidth = CentimetersToPoints(21)
    End With
    
    ReDim TableInfoArr(1 To cTab, 0)
    'First dimension: Index of tables of the document
    'Second dimension: Index of tables of the top level tables
    'we do not care about a nested table in a nested table as it is only the width affected, which will be ridiculous small anyhow
    
    For i = 1 To cTab
        TableInfoArr(i, 0).tableWidth = 0
        'calculate table width
        For Each aCell In ActiveDocument.Tables(i).Range.Cells
            If aCell.RowIndex = 1 Then TableInfoArr(i, 0).tableWidth = TableInfoArr(i, 0).tableWidth + aCell.Width Else Exit For
        Next
        TableInfoArr(i, 0).preferredWidth = ActiveDocument.Tables(i).preferredWidth
        If TableInfoArr(i, 0).preferredWidth > 999999 Then
            'I do not understand MS Word logic here
            TableInfoArr(i, 0).preferredWidth = Round(TableInfoArr(i, 0).tableWidth / pageWidth * 100)
        End If
        'Nested Table?
        cTabN = ActiveDocument.Tables(i).Tables.Count
        If cTabN > 0 Then
            'calculate table width of nested table
            ReDim Preserve TableInfoArr(1 To cTab, 0 To cTabN)
            For j = 1 To cTabN
                TableInfoArr(i, j).preferredWidth = ActiveDocument.Tables(i).Tables(j).preferredWidth
                For Each aCell In ActiveDocument.Tables(i).Tables(j).Range.Cells
                    If aCell.RowIndex = 1 Then TableInfoArr(i, j).tableWidth = TableInfoArr(i, j).tableWidth + aCell.Width Else Exit For
                Next
            Next
            'we need the cell width of the hosting cell, I did not find a simpler way to retrieve it
            j = 0
            For Each aCell In ActiveDocument.Tables(i).Range.Cells
                If aCell.Tables.Count > 0 Then
                    j = j + 1 'TableNo
                    TableInfoArr(i, j).ParentCellWidth = aCell.Width
                End If
            Next
        End If
    Next
    
Exit_MW_TableInfo:
    Exit Sub

Err_MW_TableInfo:
    DisplayError "MW_TableInfo"
    Resume Exit_MW_TableInfo
End Sub

Private Function MW_WordVersion() As Long
'retrieves the Version and returns the value before the point 8.3 --> 8
'For easier handling, now the year is retrieved
'Comment by E.Lorenz on 01.03.2018
'Actually, we only want to consider Office 365 (Word 2016) here because it's
'the relevant version for IS4IT GmbH
    
    MW_WordVersion = Int(Val(Application.Version))
    Select Case MW_WordVersion
        Case Is <= 8
            MW_WordVersion = 1997
        Case 9
            MW_WordVersion = 2000
        Case 10
            MW_WordVersion = 2002
        Case 11
            MW_WordVersion = 2003
        Case 16
            MW_WordVersion = 2016   ' added this Case to take office 2016 into consideration, else and error message will bother
        Case Else
            MsgBox "Word Version not recognized. Macros may not function correctly."
            MW_WordVersion = 2003
    End Select

End Function

Public Function GetRegValidate(ByVal RegKey$, ByVal KeyValue As Variant) As Variant
' -------------------------------------------------------------------
' Function: checks the values of registry entries, checks variable type
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input: RegKey and its value
'
' returns: nothing
'
' released: Nov 11, 2006
' -------------------------------------------------------------------
On Error GoTo Err_GetRegValidate

    Dim RegFormat As RegFormatEnum

    'Default values if empty
    'In our case most values are in ...R const, so we do not need defaults
    If KeyValue = "" Then
        Select Case RegKey
            'boolean
            Case "CategoryArticleUse", "CategoryImagesUse", "LikeArticleCategory", "ListNumbersManual", "OptionSmartParaSelection" 'NoTag
                KeyValue = True
            'This is optional, all boolean values will be false by default
                    
            'normal settings
            Case "CategoryArticle":         KeyValue = CategoryArticleR
            Case "CategoryImages":          KeyValue = CategoryImagesR
            Case "convertPageFooters":      KeyValue = convertPageFootersR
            Case "convertPageHeaders":      KeyValue = convertPageHeadersR
            Case "convertFontSize":         KeyValue = True
            Case "deleteHiddenChars":       KeyValue = deleteHiddenCharsR
            Case "ImageExtraction":        KeyValue = ImageExtractionR
            Case "ImageConverter":          KeyValue = ImageConverterR
            Case "ImageConvertCheckFileExists": KeyValue = ImageConvertCheckFileExistsR
            Case "ImageMaxPixelSize": KeyValue = ImageMaxPixelSizeR
            Case "ImageNamePreFix":         KeyValue = ImageNamePreFixR
            Case "ImagePath"
                KeyValue = ImagePathR
                If KeyValue = "" Then KeyValue = FormatPath(GetSpecialfolder(CSIDL_MYPICTURES)) & "wiki"
            Case "ImageUploadAuto":         KeyValue = ImageUploadAutoR
            Case "ImageUploadTabToFileName": KeyValue = ImageUploadTabToFileNameR
            Case "insertTitlePageIfNeeded": KeyValue = insertTitlePageIfNeededR
            Case "isCustomized":            KeyValue = isCustomizedR
            Case "ImageMaxWidth":           KeyValue = 1024&
            Case "PauseUploadAfterXImages": KeyValue = 50&
            Case "UsePPTSize":              KeyValue = 1024&
            Case "WikiAddressRootTest":     KeyValue = WikiAddressRootTestR
            Case "WikiAddressRootProd":     KeyValue = WikiAddressRootProdR
                    
                    
            'Language specific settings
            Case "Language":                KeyValue = MW_GetUserLanguage
            Case "WikiSearchTitle":         KeyValue = "Search -"         'Title of browser window after search in wiki
            Case "WikiCategoryKeyWord":     KeyValue = "category"         'category key word according to your language: keyvalue = "category" works always
            Case "CategoryImagePreFix":     KeyValue = "not set"          'Standard Prefix to the image category: keyvalue = add blank at last character to separate words
            Case "WikiUploadTitle":         KeyValue = "Upload"           'Title of browser window when uploading to wiki
            Case "EditorKeyLoadPic":        KeyValue = "^o"               'Ctrl+o for open
            Case "EditorKeySavePic":        KeyValue = "%fa"              'File save as
            Case "EditorKeyPastePicAsNew":  KeyValue = "%en"              'paste as new
            Case "EditorPaletteKey":        KeyValue = "p"                'change GIF Palette
            Case "ClickChartText":          KeyValue = "click me!"        'clickable charts will have additional text in wiki: "click me!"
            Case "UnableToConvertMarker":   KeyValue = "## Error converting ##: " 'some links can not be converted: keyvalue = give hint
            
            Case "txt_TitlePage":           KeyValue = "Title page"
            Case "txt_PageHeader":          KeyValue = "Page header"
            Case "txt_PageFooter":          KeyValue = "Page footer"
            Case "txt_Footnote":            KeyValue = "Footnotes"
            
        End Select
        
        'write back to registry
        If isInitialized Or RegKey <> "isCustomized" Then
            If KeyValue <> "" Then SetReg RegKey, KeyValue
        End If
    End If
    
    'Format of the value (type)
    Select Case RegKey
        'boolean
        Case "AllowWiki", "CategoryArticleUse", "CategoryImagesUse", "convertPageHeaders", "convertPageFooters", "convertFontSize", "deleteHiddenChars", "ImageExtraction", _
            "ImageConvertCheckFileExists", "ImageExtractionPE", "ImagePastePixel", "ImagePixelSize", "ImageMaxPixel", "ImageReload", "ImageUploadAuto", "insertTitlePageIfNeeded", _
            "isCustomized", "LikeArticleCategory", "LikeArticleName", "ListNumbersManual", "OptionSmartParaSelection", "UsePowerpoint", "Z_finished" 'NoTag
            
            RegFormat = regBoolean
        
        'regLong
        Case "ImageUploadTabToFileName", "ImageMaxPixelSize", "PauseUploadAfterXImages", "UsePPTSize", "ImageMaxWidth" 'NoTag
            
            RegFormat = regLong
            
        'regDate
        Case "GFrom" 'NoTag
            
            RegFormat = regDate
    
    End Select
    
    On Error Resume Next
    Select Case RegFormat
        Case regBoolean
            If UCase(KeyValue) = "TRUE" Then KeyValue = True 'NoTag
            KeyValue = KeyValue = True
        
        Case regDate
            If IsDate(KeyValue) Then KeyValue = CDate(KeyValue) Else KeyValue = #12:00:00 AM#
            
        Case regDouble
            KeyValue = CDbl(KeyValue)
            If Err.Number <> 0 Then Err.Clear: KeyValue = 0
        
        Case regLong
            KeyValue = CLng(KeyValue)
            If Err.Number <> 0 Then Err.Clear: KeyValue = 0
    End Select
    On Error GoTo Err_GetRegValidate

    'Value Validations
    Select Case RegKey
        Case "Language" 'NoTag
            'localization: supported languages
            KeyValue = UCase(KeyValue)
            Dim i&
            For i = 1 To UBound(languageArr, 1)
                If languageArr(i, 1) = KeyValue Then
                    i = -1
                    Exit For
                End If
            Next i
            If i > 0 Then KeyValue = MW_GetUserLanguage 'unidentified language
            
        Case "ImagePath"
            'Create Dir if not exists
            DirExistsCreate KeyValue, False
            
        Case "PauseUploadAfterXImages"
            If KeyValue <= 0 Then KeyValue = 50
    End Select
    
    
    GetRegValidate = KeyValue

Exit_GetRegValidate:
    Exit Function

Err_GetRegValidate:
    DisplayError "GetRegValidate" 'NoTag
    Resume Exit_GetRegValidate
End Function

Private Function RemoveDir(ByVal DirName$, Optional ToBin As Boolean = True, Optional StopError As Boolean = False, Optional DirCount& = 0) As Boolean
'removes all files and the directory
'Nov 20, 2006
On Error GoTo Err_RemoveDir

    Const MaxDirDel = 3 'Safety feature, will not remove more then 3 directories

    Dim Name1$
    
    If DirName = "" Then Exit Function 'Prevent deletion of full disc!
    If DirCount >= MaxDirDel Then Exit Function
    DirCount = DirCount + 1
    
    DirName = FormatPath(DirName)
    If DirExists(DirName) Then
        
        'Find files
        Name1 = Dir(DirName, vbDirectory)
        Do While Name1 <> ""
            If Name1 <> "." And Name1 <> ".." Then
                If (GetAttr(DirName & Name1) And vbDirectory) = vbDirectory Then
                    RemoveDir DirName & Name1, ToBin, StopError, DirCount
                    Name1 = Dir(DirName, vbDirectory)
                Else
                    If ToBin Then
                        KillToBin FormatPath(DirName) & Name1
                    Else
                        Kill FormatPath(DirName) & Name1
                    End If
                End If
            End If
            Name1 = Dir
        Loop
    
        'Removed directory itself
        '#ToDo# Delete to recycle bin
        RmDir DirName
        DoEvents
    
    End If
    
    RemoveDir = True

Exit_RemoveDir:
    DoEvents
    Exit Function

Err_RemoveDir:
    If StopError Then DisplayError "RemoveDir" Else Err.Clear
    'If DebugMode Then Stop: Resume Next
    Resume Exit_RemoveDir
End Function

Private Sub MediaWikiExtract_ImagesHtml2002(MyS As Shape, ImageInfo As ImageInfoType)
' -------------------------------------------------------------------
' Function: Extracts Canvas Images, these do not exist in Word 2000
' to prevent compiler error, we have a separate sub
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input:
'
' returns: nothing
'
' released: Nov. 12, 2006
' -------------------------------------------------------------------
On Error GoTo Err_MediaWikiExtract_ImagesHtml2002

    'extract canvas items
    Dim cTxt$
    Dim ci&
    Dim cItem As Shape
    
    For ci = 1 To MyS.CanvasItems.Count
        Set cItem = MyS.CanvasItems(1) 'as we delete all items
        With cItem
        
        'in this case we copy the text and make a table
        If .TextFrame.HasText Then
            '#ToDo# convert to wiki markup
            cTxt = .TextFrame.TextRange.Text
            cTxt = Replace(cTxt, vbCr, "<br>" & vbCr)
            'for now we just add a wiki table
            '#ToDo# Line type and frame size
            cTxt = vbCr & "{|" & TableTemplateParagraphFrame & vbCr & "|" & cTxt & "|}" & vbCr & vbCr
            Selection.MoveLeft
            Selection.InsertAfter cTxt
        End If
        
        Select Case .Type
            Case msoPicture, msoAutoShape
                .Delete
            
            Case msoTextBox
                ImageInfo.hasFrame = False
                .Delete
            
            Case Else
                If DebugMode Then Stop 'just for curiosity
                .Delete
                
        End Select
        End With
    Next

Exit_MediaWikiExtract_ImagesHtml2002:
    Exit Sub

Err_MediaWikiExtract_ImagesHtml2002:
    DisplayError "MediaWikiExtract_ImagesHtml2002"
    If DebugMode Then Stop: Resume Next
    Resume Exit_MediaWikiExtract_ImagesHtml2002
End Sub

Private Sub MakeDir(ByVal DirName$)
'creates Dir, sometimes windows is slow

    Dim c&

    If right$(DirName, 1) = "\" Then DirName = Left$(DirName, Len(DirName) - 1)
    
    On Error Resume Next
    c = 3
    Do While c > 0
        If Not DirExists(DirName) Then
            MkDir DirName
            If Err.Number > 0 Then
                DoEvents
                Sleep 100
                Err.Clear
                c = c - 1
                If c = 0 Then Stop 'can not create dir!
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    DoEvents

End Sub

Private Sub TestSendMessage()

    Const WM_SYSCOMMAND = &H112
    Const WM_MENUSELECT = &H11F

    Dim hWnd&, lRet&
    
    hWnd = GetWindowTitleHandle(":Grafik hoch", False, , True)
    If hWnd = 0 Then Exit Sub
    lRet = SendMessage(hWnd, WM_CLOSE, 0, 0)

End Sub

Sub TestImageInfo()
'gather some information on image position for smart grouping

    Dim arrIm() As ImageInfoType
    
    Dim DocA As Document
    Dim DocStory As Object
    Dim MyIS As InlineShape, MyS As Shape
    Dim ic&, lRet&
    
    ReDim arrIm(1 To 100)
    
    Set DocA = ActiveDocument
    
    For Each DocStory In DocA.StoryRanges
        For Each MyIS In DocStory.InlineShapes
            'If DocStory.StoryType <> wdMainTextStory Then MyIS.Select
            ic = ic + 1
            If ic Mod 100 = 0 Then ReDim Preserve arrIm(1 To ic + 100)
            With arrIm(ic)
                .ImageNo = ic
                .StoryType = DocStory.StoryType
                .TextStart = MyIS.Range.Start
                'If .TextStart = 0 Then MyIS.Select: Stop
                .IsInlineShape = True
                .Type = MyIS.Type
            End With
        Next
        On Error Resume Next
        lRet = DocStory.ShapeRange.Count
        If lRet > 0 Then
            For Each MyS In DocStory.ShapeRange
                'If DocStory.StoryType <> 1 Then MyS.Select: Stop
                ic = ic + 1
                If ic Mod 100 = 0 Then ReDim Preserve arrIm(1 To ic + 100)
                With arrIm(ic)
                    .ImageNo = ic
                    .StoryType = DocStory.StoryType
                    .TextStart = MyS.Anchor.Start
                    .Left = MyS.Left
                    .Top = MyS.Top
                    .IsInlineShape = False
                    .Type = MyS.Type
                End With
            Next
        Else
            Err.Clear
        End If
    Next
    
    If ic > 0 Then ReDim Preserve arrIm(1 To ic) Else Erase arrIm

End Sub

Sub TestUnicode()
    
    Dim txt$

    txt = ChrW(1044) & ChrW(1086) & ChrW(1082) & ChrW(1091) & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1090) & ChrW(1099) & ChrW(32) & ChrW(1085) & ChrW(1077) & ChrW(32) & ChrW(1079) & ChrW(1072) & ChrW(1075) & ChrW(1088) & ChrW(1091) & ChrW(1078) & ChrW(1077) & ChrW(1085) & ChrW(1099) & ChrW(46)
    
    'MsgBox txt

    Dim Text As String
    Dim Überschrift As String
    
    Text = "Hallo!" & vbNewLine & ChrW$(&H3B1&) & ChrW$(&H3B2&) & ChrW$(&H3B3&) 'einige griechische Zeichen
    Überschrift = "Ein Beispiel"
    
    Call MessageBoxW(0, StrPtr(txt), StrPtr(Überschrift), MB_ICONINFORMATION Or MB_TASKMODAL)
    Call MessageBoxA(0, txt, Überschrift, MB_ICONINFORMATION Or MB_TASKMODAL)

End Sub

Sub TestReadUnicode()

    Dim txt$
    Dim i&
    Dim s As Selection
    
    Set s = Selection
    
    For i = 1 To s.Characters.Count
        
        txt = txt & "ChrW(" & AscW(s.Characters(i)) & ") & "
    
    Next i

    Debug.Print txt

End Sub

Sub TestCopyDoc()

    Dim DocOrg As Document
    Dim DocWork As Document
    
    Set DocOrg = ActiveDocument
    
    Documents.Add
    Set DocWork = ActiveDocument
    
    'DocWork.Range.FormattedText = DocOrg.Range.FormattedText
    DocWork.Sections(1).Range.FormattedText = DocOrg.Sections(1).Range.FormattedText
    

End Sub

Private Sub MW_SetOptions_2003(SetOption As Boolean)

    'Store user options
    If SetOption Then
        If GetReg("Z_finished") Then
            SetReg "OptionSmartParaSelection", Application.Options.SmartParaSelection
        End If
        Options.SmartParaSelection = False
    Else
        Application.Options.SmartParaSelection = GetReg("OptionSmartParaSelection")
    End If

End Sub


