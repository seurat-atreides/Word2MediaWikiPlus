VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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

Private Sub cmdCopyModuls_Click()

    'copies the Word2MediaWikiPlus Moduls to Normal.dotm
    If CopyModulesToNormal Then
        'CreateSymbol True       ' E.Lorenz. don't want to create symbols because it's deprecated
        MsgBox "The macro have been installed.", vbInformation
    Else
        MsgBox "Macro could not be installed.", vbExclamation
    End If

End Sub

Private Sub CreateSymbol(Optional CheckExit As Boolean = False)
'not in use anymore, Nov. 14, 2006
'some problem with the display of the statusbar
'seem to be a word bug, if not using the organizer (Word 2002)

    Dim MyC As CommandBar
    Dim Answer&
    
    For Each MyC In Application.CommandBars
         
        If MyC.Name = "Word2MediaWikiPlus" And InStr(1, MyC.Context, "Normal.dotm", vbTextCompare) > 0 Then
            If CheckExit Then Exit Sub 'leave if exists
            Answer = MsgBox("Overwrite existing symbol bar?", vbYesNo)
            If Answer = vbYes Then Application.OrganizerDelete Source:= _
            "Normal.dotm", Name:="Word2MediaWikiPlus", Object _
            :=wdOrganizerObjectCommandBars
        End If
    Next MyC
    
    'copy Word2MediaWikiPlus
    Application.OrganizerCopy Source:= _
        ActiveDocument.Path & "\" & ActiveDocument.Name _
        , Destination:=Templates("Normal.dotm").Path & "\" & "Normal.dotm", Name _
        :="Word2MediaWikiPlus", Object:=wdOrganizerObjectCommandBars
    Application.CommandBars("Word2MediaWikiPlus").Visible = True

End Sub

Private Sub CreateSymbol2()

    Const Word2WikiBar = "Word2MediaWikiPlus"

    Dim MyControl As Object
    Dim myDoc As Document
    Dim c&
    Dim oExists As Boolean
    
    CustomizationContext = NormalTemplate
    'Create a new CommandBar
    For Each MyControl In Application.CommandBars
        If MyControl.Name = Word2WikiBar Then oExists = True: Exit For
    Next
    If Not oExists Then
        Application.CommandBars.Add Word2WikiBar, msoBarTop
    End If
    
    With Application.CommandBars(Word2WikiBar)
        
        'Convert document
        'Find existing Symbol, so we do not create a second one
        Set MyControl = .FindControl(msoControlButton, Tag:="W2MW Convert")
        If MyControl Is Nothing Then
            'make new button
            Set MyControl = .Controls.Add(Type:=msoControlButton)
        End If
        MyControl.OnAction = "Word2MediaWikiPlus"
        MyControl.Caption = "Convert to Wiki"
        MyControl.TooltipText = "Word2MediaWikiPlus - Convert to MediaWiki syntax"
        MyControl.DescriptionText = "Converts Word Documents in MediaWiki format."
        MyControl.Tag = "W2MW Convert"
        If MyControl.FaceId < 2 Then MyControl.FaceId = 2144
        MyControl.Style = msoButtonIconAndCaption
        MyControl.BeginGroup = True
        
        'Upload Images
        Set MyControl = .FindControl(msoControlButton, Tag:="W2MW Upload")
        If MyControl Is Nothing Then
            'make new button
            Set MyControl = .Controls.Add(Type:=msoControlButton)
        End If
        MyControl.OnAction = "Word2MediaWikiPlus_Upload"
        MyControl.Caption = "Upload Images"
        MyControl.TooltipText = "Upload Images to wiki"
        MyControl.DescriptionText = ""
        MyControl.Tag = "W2MW Upload"
        If MyControl.FaceId < 2 Then MyControl.FaceId = 38
        MyControl.Style = msoButtonIconAndCaption
        MyControl.BeginGroup = True
    
        'Config
        Set MyControl = .FindControl(msoControlButton, Tag:="W2MW Config")
        If MyControl Is Nothing Then
            'make new button
            Set MyControl = .Controls.Add(Type:=msoControlButton)
        End If
        MyControl.OnAction = "Word2MediaWikiPlus_Config" '"normal.modWord2MediaWikiPlus.Word2MediaWikiPlus"
        MyControl.Caption = "Configure Word2MediaWikiPlus"
        MyControl.TooltipText = "Word2MediaWikiPlus - Configuration"
        MyControl.DescriptionText = ""
        MyControl.Tag = "W2MW Config"
        If MyControl.FaceId < 2 Then MyControl.FaceId = 277
        MyControl.Style = msoButtonIconAndCaption
        MyControl.BeginGroup = True
        
    End With

    ActiveDocument.CommandBars(Word2WikiBar).Visible = True
    DoEvents

    'Now check if there are more then one CommandBar
    c = 0
    For Each MyControl In Application.CommandBars
        If MyControl.Name = Word2WikiBar Then
            If c > 0 Then MyControl.Delete Else c = c + 1
        End If
    Next


    'MsgBox "Symbols were created.", vbInformation

End Sub

Private Function CopyModulesToNormal() As Boolean
On Error GoTo cn_error

    Dim vbCom As Object, vbncom As Object
    Dim ObjExists As Boolean
    Dim i&, deleteAll As Boolean
    Dim ImportName$
    
    For Each vbCom In ActiveDocument.VBProject.VBComponents
        If vbCom.Name <> "ThisDocument" And vbCom.Name <> "NewMacros" Then
            'Do we have it already?
            ObjExists = False
            For Each vbncom In NormalTemplate.VBProject.VBComponents
                If vbncom.Name = vbCom.Name Then ObjExists = True: Exit For
            Next
            If ObjExists Then
                If Not deleteAll Then
                    i = MsgBox("Do you want to update your existing code?", vbQuestion + vbYesNo)
                    If i = vbYes Then deleteAll = True Else Exit Function
                End If
                NormalTemplate.VBProject.VBComponents.Remove vbncom
                ObjExists = False
            End If
            If Not ObjExists Then
                'export it from macro document
                ImportName = FormatPath(ActiveDocument.Path) & vbCom.Name
                vbCom.Export ImportName
                'import it
                If FileExists(ImportName) Then
                    NormalTemplate.VBProject.VBComponents.Import ImportName
                    Kill ImportName
                    If vbCom.Type = 3 Then
                        If FileExists(ImportName & ".frx") Then Kill ImportName & ".frx"
                    End If
                Else
                    MsgBox "File not found: " & ImportName, vbCritical
                    Exit Function
                End If
                
            End If
            Debug.Print "imported: " & vbCom.Name
        End If
    Next
    
    'remove old coding
    If deleteAll Then
        For Each vbncom In NormalTemplate.VBProject.VBComponents
            Select Case vbncom.Name
                Case "Word2MediaWikiPlus", "Word2MediaWikiPlusGlobal"
                    NormalTemplate.VBProject.VBComponents.Remove vbncom
            End Select
        Next
    End If
    
    'MsgBox "The macro is now installed.", vbInformation
    
    CopyModulesToNormal = True
    Exit Function

cn_error:
    MsgBox "Error copying modules." & vbCr & vbCr & Err.Description, vbCritical
End Function

Private Sub cmdSymbols_Click()
    CreateSymbol
End Sub

Private Sub cmdUninstall_Click()
    Uninstall_Word2MediaWikiPlus
End Sub
