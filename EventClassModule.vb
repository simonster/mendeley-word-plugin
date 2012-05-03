' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2009-2012 Mendeley Ltd.
'
' Licensed under the Educational Community License, Version 1.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
' http://www.opensource.org/licenses/ecl1.php
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
' ***** END LICENSE BLOCK *****

' (No code from Zotero in this source file)

' author: steve.ridout@mendeley.com

Option Explicit

Public WithEvents App As Word.Application
Public WithEvents ComboBox As CommandBarComboBox

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    ' Don't do anything if we're in the middle of an operation
    If uiDisabled Or Not initialised Or unitTest Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler

    If USE_RIBBON Then
        Call recoverRibbonUi
    End If
    
    Dim selectedFields() As Object
    Call getFieldsAtSelection(2, selectedFields)

    Dim currentField As Field
    If selectedFields(1) Is Nothing Then
        Set currentField = Nothing
    Else
        Set currentField = selectedFields(1)
    End If
    
    Dim undoEditButton As CommandBarButton
    Set undoEditButton = getUndoEditButton()
    
    Dim insertCitationButton As CommandBarButton
    Set insertCitationButton = getInsertCitationButton()
    
    If Not previouslySelectedField Is Nothing Then
        If currentField Is Nothing Then
            Call checkForCitationEdit
        Else
            On Error GoTo ResetPreviouslySelected
            If Not previouslySelectedField.index = currentField.index Then Call checkForCitationEdit
            If DEBUG_MODE Then
                On Error GoTo 0
            Else
                On Error GoTo ErrorHandler
            End If
            If False Then
ResetPreviouslySelected:
                Set previouslySelectedField = Nothing
            End If
        End If
    End If

    Dim prevRibbonCitationButtonState As String
    prevRibbonCitationButtonState = ribbonCitationButtonState
    
    If Not selectedFields(2) Is Nothing Then
        ribbonCitationButtonState = RIBBON_MERGE_CITATIONS
    ElseIf Not currentField Is Nothing Then
        Dim position As Long
        Dim markName As String
        markName = getMarkName(currentField)
        position = InStr(markName, CSL_CITATION)
        
        If position > 0 Then
        
            insertCitationButton.Caption = TOOLBAR_EDIT_CITATION
            insertCitationButton.TooltipText = TOOLTIP_EDIT_CITATION
            ribbonCitationButtonState = RIBBON_EDIT_CITATION
                    
            If mendeleyApiClient().hasManualEdit(markName) Then
                ribbonCitationButtonState = RIBBON_UNDO_CITATION
            End If
        
        ElseIf isMendeleyBibliographyField(currentField.code.Text) Or _
            startsWith(currentField.code.Text, MENDELEY_CITATION) Then
            insertCitationButton.Caption = TOOLBAR_EDIT_CITATION
            insertCitationButton.TooltipText = TOOLTIP_EDIT_CITATION
            ribbonCitationButtonState = RIBBON_EDIT_CITATION
        ElseIf startsWith(currentField.code.Text, MENDELEY_EDITED_CITATION) Then
            ribbonCitationButtonState = RIBBON_UNDO_CITATION
        Else
            insertCitationButton.Caption = TOOLBAR_INSERT_CITATION
            insertCitationButton.TooltipText = TOOLTIP_INSERT_CITATION
            ribbonCitationButtonState = RIBBON_INSERT_CITATION
        End If
    Else
        insertCitationButton.Caption = TOOLBAR_INSERT_CITATION
        insertCitationButton.TooltipText = TOOLTIP_INSERT_CITATION
        ribbonCitationButtonState = RIBBON_INSERT_CITATION
    End If
    
    If undoEditButton.Visible <> (ribbonCitationButtonState = RIBBON_UNDO_CITATION) Then
        undoEditButton.Visible = (ribbonCitationButtonState = RIBBON_UNDO_CITATION)
    End If
    
    If USE_RIBBON And (prevRibbonCitationButtonState <> ribbonCitationButtonState) Then
        ribbonUi.Invalidate
    End If

    If currentField Is Nothing Then
        previouslySelectedFieldResultText = ""
        Set previouslySelectedField = Nothing
    Else
        If Not previouslySelectedField Is Nothing Then
            If Not (currentField.index = previouslySelectedField.index) Then
                Call updatePrevious
            End If
        Else
            Call updatePrevious
        End If
    End If
    
    GoTo EndOfSub
ErrorHandler:
    ' Can't call reportError here since unfortunately we expect to get errors while the user
    ' is using the Word spell-checker
EndOfSub:

    ' stop word trying to save changes to this template
    ThisDocument.Saved = True
End Sub

Sub updatePrevious()
    Dim updatePrevious As Boolean
                
    Dim currentField As Field
    Set currentField = getFieldAtSelection()
    
    If previouslySelectedField Is Nothing Then
        updatePrevious = True
    ElseIf previouslySelectedField.index = currentField.index Then updatePrevious = True
    End If
    
    If updatePrevious Then
        previouslySelectedFieldResultText = currentField.result.Text
        Set previouslySelectedField = currentField
    End If
End Sub

Private Sub ComboBox_Change(ByVal ctrl As CommandBarComboBox)
    If uiDisabled Then
        Exit Sub
    End If
    If Not DEBUG_MODE Then
        On Error GoTo EndOfSub
    End If
    uiDisabled = True

    Call updateStyleFromComboBox
    Call refreshDocument
    
EndOfSub:
    uiDisabled = False
End Sub

Private Sub App_DocumentOpen(ByVal Doc As Document)
    If uiDisabled Then
        Exit Sub
    End If
    If Not DEBUG_MODE Then
        On Error GoTo EndOfSub
    End If
    uiDisabled = True

    Call refreshDocument(True)
    
EndOfSub:
    uiDisabled = False
End Sub




