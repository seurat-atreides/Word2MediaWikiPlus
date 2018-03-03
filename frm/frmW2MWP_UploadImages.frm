VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmW2MWP_UploadImages 
   Caption         =   "UserForm1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   OleObjectBlob   =   "frmW2MWP_UploadImages.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmW2MWP_UploadImages"
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

Public Function GetFiles(Optional ByVal sTitle As String = "Select files to upload") As String
'sTitle: Optional Title of Dialog
'New Version by Manfred Gerwing, August 2006
On Error GoTo ProcError
  
    Dim sFilenames$, TempDir$
    
    TempDir = CurDir
    
    sFilenames = modW2MWP_FileDialog.ahtCommonFileOpenSave( _
        &H4 Or &H800 Or &H40000 Or &H200 Or &H80000, _
        GetReg("ImagePath"), _
        "Images (*.png;*.jpg;*.gif)|*.png;*.jpg;*.gif", 1, , , sTitle, 0, True)

ProcExit:
On Error Resume Next
    GetFiles = sFilenames
    ChDir TempDir
    Exit Function

ProcError:
    If Err.Number = &H7FF3 Then Resume Next 'Cancel selected - Ignore
    MsgBox Err.Description & "(" & Err.Number & ")", vbExclamation, "Open error"
    sFilenames = ""
    Resume ProcExit
End Function

Private Sub cmdSelect_Click()
  
    Dim strFileNames() As String, AddFiles$
    Dim i As Integer
    
    strFileNames = Split(GetFiles, Chr(0))
    
    Me.List1.Clear
    ReDim ImageArr(0)
    
    If UBound(strFileNames) = 0 Then
        Label1.Caption = "Files selected from: " & GetFilePath(strFileNames(0))
        Me.List1.AddItem GetFileName(strFileNames(0))
        ReDim ImageArr(1 To 1)
        ImageArr(1) = strFileNames(0)
    ElseIf UBound(strFileNames) > 0 Then
        'Path is stored in index 0 of array files in index 1...
        Label1.Caption = "Files selected from: " & strFileNames(0)
        ReDim ImageArr(1 To UBound(strFileNames))
        For i = 1 To UBound(strFileNames)
            Me.List1.AddItem GetFileName(strFileNames(i))
            ImageArr(i) = FormatPath(strFileNames(0)) & strFileNames(i)
        Next
    Else
        Label1.Caption = "No files selected"
    End If
    
    If UBound(ImageArr) > 0 Then Me.cmdUpload.Enabled = True
    
End Sub

Private Sub cmdUpload_Click()
    MW_SetWikiAddressRoot
    If UBound(ImageArr) > 0 Then
        MediaWikiImageUpload True
        Unload Me
    End If
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.Caption = ConverterPrgTitle & ": Image Upload to Wiki"
    ReDim ImageArr(0)
    MW_LanguageTexts
    If Not isInitialized Then MW_Initialize
    Me.cmdUpload.Enabled = False
    
End Sub
