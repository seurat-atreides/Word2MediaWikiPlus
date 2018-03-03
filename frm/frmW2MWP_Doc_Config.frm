VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmW2MWP_Doc_Config 
   Caption         =   "Word2MediaWiki Plus"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   OleObjectBlob   =   "frmW2MWP_Doc_Config.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmW2MWP_Doc_Config"
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

Private Sub boxImageExtraction_Change()
    If Me.boxImageExtraction Then
        Me.boxImageUpload.Visible = True
        'Me.boxImageReload.Visible = True
        'Me.cmdImagesOnly.Enabled = True
        If GetReg("ImageExtractionPE") = False Then
            Me.boxOverwrite = False
            Me.boxOverwrite.Enabled = False
            Me.boxPastePixel = False
            Me.boxPastePixel.Enabled = False
            Me.boxMaxPixel = False
            Me.boxMaxPixel.Enabled = False
            Me.boxPixelSize = False
            Me.boxPixelSize.Enabled = False
            Me.boxImageReload = False
            Me.boxImageReload.Enabled = False
        Else
            Me.boxOverwrite.Visible = True
            Me.boxPastePixel.Visible = True
            Me.boxMaxPixel.Visible = True
            Me.boxPixelSize.Visible = True
        End If
    Else
        Me.boxImageUpload.Visible = False
        Me.boxOverwrite.Visible = False
        Me.boxPastePixel.Visible = False
        Me.boxMaxPixel.Visible = False
        Me.boxPixelSize.Visible = False
        'Me.boxImageReload.Visible = False
        'Me.cmdImagesOnly.Enabled = False
    End If
    boxImageUpload_Change
End Sub

Private Sub boxImageUpload_Change()
    If Me.boxImageUpload And Me.boxImageExtraction Then 'And Me.boxOverwrite
        Me.boxImageReload.Visible = True
    Else
        Me.boxImageReload.Visible = False
    End If
End Sub

Private Sub boxLikeArticleCategory_Click()
    If Me.boxLikeArticleCategory Then Me.txtImageCategory = Me.txtCategory
End Sub

Private Sub boxLikeArticleName_Click()
    If Me.boxLikeArticleName Then Me.txtCategory = Me.txtArticleName
End Sub

Private Sub boxOverwrite_Click()
    boxImageUpload_Change
End Sub

Private Sub cmdCancel_Click()
    'isInitialized = False
    '### should have other exit code
    Unload Me
End Sub

Private Sub cmdImagesOnly_Click()
    If Me.boxImageExtraction = True Then
        convertImagesOnly = True
        cmdOK_Click
    Else
        Me.boxImageExtraction = True
        MsgBox "Turned Image conversion on, now choose image handling settings and try again.", vbInformation, ConverterPrgTitle
    End If
End Sub

Private Sub cmdOK_Click()
    
    Me.txtCategory = Trim$(Me.txtCategory)
    Me.txtImageCategory = Trim$(Me.txtImageCategory)
    
    'Save settings to registry
    SetReg "CategoryArticle", Trim$(Me.txtCategory)
    SetReg "CategoryImages", Trim$(Me.txtImageCategory)
    'CategoryStandardImagesInt = MW_FormatCategoryString(CategoryStandardImagesInt, True)
    SetReg "CategoryImagePreFix", Trim$(Me.txtImageCategoryPrefix)
    SetReg "CategoryArticleUse", Me.boxArticleCategory
    SetReg "CategoryImagesUse", Me.boxImageCategory
    SetReg "ImageDescription", Me.txtImageDescription

    Me.txtArticleName = Trim$(Me.txtArticleName)
    DocInfo.ArticleName = IIf(Me.txtArticleName = "", "No Name", Me.txtArticleName)
    SetReg "LastUserArticleName", DocInfo.ArticleName
    
    SetReg "ImageExtraction", Me.boxImageExtraction
    SetReg "ImageUploadAuto", Me.boxImageUpload
    SetReg "ImageConvertCheckFileExists", Not Me.boxOverwrite
    SetReg "AllowWiki", Me.boxAllowWiki
    SetReg "convertFontSize", Me.boxFontSize
    SetReg "convertPageHeaders", Me.boxHeader
    SetReg "convertPageFooters", Me.boxFooter
    SetReg "ImagePixelSize", Me.boxPixelSize
    SetReg "ImageMaxPixel", Me.boxMaxPixel
    SetReg "ImagePastePixel", Not Me.boxPastePixel
    SetReg "ImageReload", Me.boxImageReload
    SetReg "ListNumbersManual", Me.boxListNumberManual
    SetReg "LikeArticleName", Me.boxLikeArticleName
    SetReg "LikeArticleCategory", Me.boxLikeArticleCategory
    
    If Me.optImageSize1 Then
        SetReg "ImageResizeOption", 1
        DocInfo.ImageResizeOption = 1
    ElseIf Me.optImageSize2 Then
        SetReg "ImageResizeOption", 2
        DocInfo.ImageResizeOption = 2
    Else
        SetReg "ImageResizeOption", 3
        DocInfo.ImageResizeOption = 3
    End If
    
    If optSystemTest Then SetReg "WikiSystem", "Test" Else SetReg "WikiSystem", "Prod" 'It is important that these are spelled correctly (case)
    MW_SetWikiAddressRoot
    
    isInitialized = True
        
    Unload Me

End Sub

Private Sub lblHelp_Click()
    IExplorer "http://meta.wikimedia.org/wiki/Word2MediaWikiPlus_Documentation#Document_conversion"
End Sub

Private Sub txtArticleName_Change()
    If Me.boxLikeArticleName Then Me.txtCategory = Me.txtArticleName
End Sub

Private Sub txtCategory_Change()
    If Me.boxLikeArticleCategory Then Me.txtImageCategory.Text = Me.txtCategory.Text
End Sub

Private Sub txtImageDescription_AfterUpdate()
    If InStr(1, Me.txtImageDescription, vbCr) > 0 Then
        Me.txtImageDescription = Replace(Me.txtImageDescription, vbCr, "<br>")
    End If
End Sub

Private Sub UserForm_Initialize()
    
    Me.Caption = ConverterPrgTitle & " " & WMPVersion
    
    If Not isInitialized Then MW_Initialize
    isInitialized = False 'to remember if ok was pressed
    
    If GetReg("LastArticleName") <> DocInfo.ArticleName Or GetReg("LastUserArticleName") = "" Then
        SetReg "LastArticleName", DocInfo.ArticleName
        SetReg "LastUserArticleName", DocInfo.ArticleName
    End If
    
    'Load registry settings
    Me.txtArticleName = GetReg("LastUserArticleName")
    Me.txtCategory = GetReg("CategoryArticle")
    Me.txtImageCategory = GetReg("CategoryImages")
    Me.boxArticleCategory = GetReg("CategoryArticleUse")
    Me.boxImageCategory = GetReg("CategoryImagesUse")
    Me.txtImageDescription = GetReg("ImageDescription")
    Me.txtImageCategoryPrefix = GetReg("CategoryImagePreFix")
    
    Me.boxLikeArticleName = GetReg("LikeArticleName")
    Me.boxLikeArticleCategory = GetReg("LikeArticleCategory")
    Me.txtImageCategoryPrefix = GetReg("CategoryImagePreFix")
    Me.boxImageExtraction = GetReg("ImageExtraction")
    Me.boxImageUpload = GetReg("ImageUploadAuto")
    Me.boxImageReload = GetReg("ImageReload")
    Me.boxOverwrite = Not GetReg("ImageConvertCheckFileExists")
    Me.boxAllowWiki = GetReg("AllowWiki")
    Me.boxFontSize = GetReg("convertFontSize")
    Me.boxHeader = GetReg("convertPageHeaders")
    Me.boxFooter = GetReg("convertPageFooters")
    Me.boxPixelSize = GetReg("ImagePixelSize")
    Me.boxMaxPixel = GetReg("ImageMaxPixel")
    Me.boxPastePixel = Not GetReg("ImagePastePixel")
    Me.boxListNumberManual = GetReg("ListNumbersManual")
    boxImageExtraction_Change
    Select Case GetReg("WikiSystem")
        Case "Test"
            Me.optSystemTest = True
        Case "Prod"
            Me.optSystemProd = True
        Case Else
            Me.optSystemTest = True
    End Select

    Select Case GetReg("ImageResizeOption")
        Case 1
            Me.optImageSize1 = True
        Case 3
            Me.optImageSize3 = True
        Case Else
            Me.optImageSize2 = True
    End Select
    

    If EditorPath = "" And GetReg("ImageExtractionPE") Then
        Me.boxImageExtraction = False
        Me.boxImageExtraction.Enabled = False
        SetReg "ImageExtraction", False
        boxImageExtraction_Change
        Me.cmdImagesOnly.Enabled = False
    End If

End Sub
