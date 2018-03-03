VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmW2MWP_Config 
   Caption         =   "Word2MediaWikiPlus"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   OleObjectBlob   =   "frmW2MWP_Config.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmW2MWP_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheckProd_Click()
    IExplorer MW_SearchAddress(Me.txtURLProd) & "WikiTest Word2MediaWikiPlus"
End Sub

Private Sub cmdCheckTest_Click()
    IExplorer MW_SearchAddress(Me.txtURLTest) & "WikiTest Word2MediaWikiPlus"
End Sub

Private Sub cmdOK_Click()

    If GetReg("Language") <> Me.cboLanguage Then
        SetReg "Language", Me.cboLanguage
        MW_LanguageTexts
    End If

    SetReg "WikiAddressRootTest", Me.txtURLTest
    SetReg "WikiAddressRootProd", Me.txtURLProd
    SetReg "ImageUploadTabToFileName", Val(Me.txtTabToFileName)
    SetReg "ImageExtractionPE", Me.OptionPhotoEditor
    MW_SetWikiAddressRoot
    
    '### Needs to be included in the config dialog, as well as the format choice
    'SnagIt available?
    'ImageConverterReg = "BMPinternally"
    'ImageConverterReg = "MSPhotoEditor"
    'ImageConverterReg = "SnagIt"
    'If GetReg("ImageConverter") = "SnagIt" Then If Not MW_SnagIt_Check_Installed(True) Then SetReg "ImageConverter", "MSPhotoEditor"
    
    If DirExistsCreate(Me.txtImagePath) Then
        SetReg "ImagePath", Me.txtImagePath
        SetReg "isCustomized", True
        isInitialized = True
        Unload Me
    Else
        MsgBox "Please check image path", vbOKOnly + vbExclamation, ConverterPrgTitle
    End If
End Sub

Private Sub cmdSimulateUpload_Click()
    SetReg "ImageUploadTabToFileName", Val(Me.txtTabToFileName)
    MW_SetWikiAddressRoot Me.txtURLTest
    
    If Me.OptionPhotoEditor Then
        'Check Photo Editor
        MW_GetEditorPath True
        If EditorPath <> "" Then MW_ImageUpload_File "Test upload.png", True
    Else
        MW_ImageUpload_File "Test upload.png", True
    End If
    Application.Activate
End Sub

Private Sub lblHelp_Click()
    IExplorer "http://meta.wikimedia.org/wiki/Word2MediaWikiPlus_Documentation#Configuration"
End Sub

Private Sub txtTabToFileName_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(Me.txtTabToFileName) Then Me.txtTabToFileName = 2
End Sub

Private Sub txtURLProd_AfterUpdate()
    Me.txtURLProd = BaseUrl(Me.txtURLProd)
End Sub

Private Sub txtURLTest_AfterUpdate()
    Me.txtURLTest = BaseUrl(Me.txtURLTest)
End Sub

Private Sub UserForm_Initialize()
' -------------------------------------------------------------------
' Function:
'
' copyright by Gunter Schmidt www.beadsoft.de
' Released under GPL
'
' Input:
'
' returns: nothing
'
' released: June 04, 2006
' changed:
' -------------------------------------------------------------------
On Error GoTo Err_UserForm_Initialize
    
    Dim i&
    
    Me.Caption = ConverterPrgTitle & " " & WMPVersion
    
    If Not isInitialized Then MW_Initialize
    isInitialized = False
    
    EditorPath = MW_GetEditorPath
    If EditorPath = "" Then
        Me.lblMSPE_Path = "not available"
        SetReg "ImageExtractionPE", False
    Else
        Me.lblMSPE_Path = EditorPath
    End If
    If GetReg("ImageExtractionPE") Then Me.OptionPhotoEditor = True Else Me.OptionHtml = True
    If Me.lblMSPE_Path = "not available" Then
        'Me.OptionHtml.Enabled = False
        Me.OptionPhotoEditor.Enabled = False
    End If

    'fill language combo
    With Me.cboLanguage
        .ColumnCount = 2
        .BoundColumn = 1
        .TextColumn = 2
        .MatchEntry = fmMatchEntryFirstLetter
        .Clear
        .List() = languageArr()
        .ColumnWidths = "25;30"
        'position
        For i = 0 To .ListCount - 1
            If .List(i, 0) = GetReg("Language") Then .ListIndex = i: Exit For
        Next i
    End With
    MW_LanguageTexts True
    
    Me.txtURLTest = GetReg("WikiAddressRootTest")
    Me.txtURLProd = GetReg("WikiAddressRootProd")
    Me.txtImagePath = GetReg("ImagePath")
    Me.txtTabToFileName = GetReg("ImageUploadTabToFileName")
    
Exit_UserForm_Initialize:
    Exit Sub

Err_UserForm_Initialize:
    DisplayError "UserForm_Initialize"
    Resume Exit_UserForm_Initialize
End Sub

Private Function BaseUrl$(url$)
'The wiki base address is the part of the url before the page name
'As some wikis use index.php?title= and some don't, we must use diffent URL formats
'We also have a language problem, as the Main_page is called different in other languages
'e.g. http://beadsoft.net/wiki/index.php?title=Hauptseite

    Dim p&

    BaseUrl = url

    'To remain compatible with V0.6b, we assume an URL ending with / as valid address
    If right$(url, 1) = "/" Then BaseUrl = url: Exit Function
    
    p = InStrRev(url, "=")
    If p > 0 Then
        BaseUrl = Left$(url, p)
        'Exit Function
    Else
        BaseUrl = url & "/"
    End If

    'MsgBox "Your URL was not recognized as valid URL!", vbExclamation

End Function
