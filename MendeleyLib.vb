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

' author: steve.ridout@mendeley.com

Option Explicit

#If VBA7 Then
   Private Declare PtrSafe Function AllowSetForegroundWindow Lib "User32" (ByVal processId As Long) As Boolean
   Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
   Private Declare PtrSafe Function GetActiveWindow Lib "User32" () As LongPtr
   Private Declare PtrSafe Function SetForegroundWindow Lib "User32" (ByVal hwnd As LongPtr) As Boolean
#Else
   Private Declare Function AllowSetForegroundWindow Lib "User32" (ByVal processId As Long) As Boolean
   Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
   Private Declare Function GetActiveWindow Lib "User32" () As Long
   Private Declare Function SetForegroundWindow Lib "User32" (ByVal hwnd As Long) As Boolean
#End If

Function buildingPlugin() As Boolean
    buildingPlugin = Left(ActiveDocument.Name, Len(TEMPLATE_NAME_DURING_BUILD)) = TEMPLATE_NAME_DURING_BUILD
End Function

Function isMendeleyInstalled() As Boolean
    Dim myWS As Object
    Dim executablePath As String
    On Error GoTo ErrorHandler
    Set myWS = CreateObject("WScript.Shell")
    executablePath = myWS.RegRead("HKEY_CURRENT_USER\Software\Mendeley Ltd.\Mendeley Desktop\ExecutablePath")
    
    isMendeleyInstalled = (Dir(executablePath) <> "")
    Exit Function
  
ErrorHandler:
    isMendeleyInstalled = False
End Function

' Gives Mendeley Desktop permission to set the foreground window.
' See AllowSetForegroundWindow()
Function allowMendeleyToSetForeground() As Boolean
    Dim processId As Long
    processId = mendeleyApiClient().mendeleyProcessId()
    allowMendeleyToSetForeground = AllowSetForegroundWindow(processId)
End Function
' Attempt to bring Word to the foreground
' This requires that either Word or Mendeley Desktop currently
' has the focus
Sub bringWordToForeground()
    ' If Mendeley has focus, ask it to allow Word to steal focus
    Dim processId As Long
    processId = GetCurrentProcessId()
    Call mendeleyApiClient().allowSetFocus(processId)
   
#If VBA7 Then
    Dim windowId As LongPtr
#Else
    Dim windowId As Long
#End If
    windowId = GetActiveWindow()
    Call SetForegroundWindow(windowId)
End Sub

Sub reportError()
    Dim errorDescription
    Dim errorSource

    errorDescription = Err.Description
    errorSource = Err.source

    Dim vbCrLf
    vbCrLf = Chr(13)

    If isMendeleyInstalled() = False Then
        Exit Sub
    End If

    MsgBox errorDescription + " in " + errorSource, Title:="Mendeley Word Plugin Problem"
End Sub

Sub sendWordProcessorVersion()
    Call mendeleyApiClient().setWordProcessor("WinWord", Application.Version)
End Sub


' Refresh the citations in this document and update the
' citation selector combo-box
'
' @param openingDocument Set to true if the refresh is being
' called whilst opening a new document or false if refreshing
' an existing already-open document
'
Function refreshDocument(Optional openingDocument As Boolean = False) As Boolean

    Dim currentDocumentPath As String
    currentDocumentPath = activeDocumentPath()

    refreshDocument = False
    Call ActiveDocument.Activate
    
    ZoteroUseBookmarks = False
    
    If openingDocument = True Then
        If Not unitTest Then
            Dim ComboBox2 As CommandBarComboBox
            Set ComboBox2 = getCitationStyleComboBox()
            ComboBox2.Text = getStyleNameFromId(getCitationStyleId())
        End If
        ThisDocument.Saved = True
        Exit Function
    End If
    
    If launchMendeleyIfNecessary() <> CONNECTION_CONNECTED Then
        Exit Function
    End If
    
    If Not isDocumentLinkedToCurrentUser Then
        Exit Function
    End If
    
    Dim documentState As DocumentStateType
    documentState = startUpdatingDocument(ActiveDocument)
    ' Update document
    Call beginUndoTransaction("Format Mendeley Citations and Bibliography")
    
    Call sendWordProcessorVersion
    
    Call setCitationStyle(getCitationStyleId())
    If Not unitTest Then
        Call updateCitationStylesComboBox
    End If
    
    If USE_RIBBON Then
        Call recoverRibbonUi
        ribbonUi.Invalidate
    End If

    ' Subscribe to events (e.g. WindowSelectionChange) doing on refreshDocument as it
    ' doesn't work in initialise() when addExternalFunctions() is also called
    If Not openingDocument Then
        Set eventClassModuleInstance.App = Word.Application
    End If

    Dim citationNumberCount As Long
    citationNumberCount = 0
    
    Dim bibliography As String

    Call mendeleyApiClient().resetCitations
    
    Dim marks
    marks = fnGetMarks(ZoteroUseBookmarks)
    
    Dim markName As String
    
    Dim thisField As Field

    Dim mark

    Dim citationNumber As Long
    citationNumber = 0

    
    For Each mark In marks
        Set thisField = mark
        
        markName = getMarkName(thisField)
        
        If startsWith(markName, "ref Mendeley") Then
            markName = Right(markName, Len(markName) - 4)
            thisField.code.Text = markName
        End If
        
        If isMendeleyCitationField(markName) Then
            citationNumber = citationNumber + 1
            
            ' Just send an empty string if the displayed text is a temporary placeholder
            Dim displayedText As String
            displayedText = getMarkText(thisField)
            'displayedText = getMarkTextWithFormattingTags(thisField)
            If displayedText = INSERT_CITATION_TEXT Or displayedText = MERGING_TEXT Then
                displayedText = ""
            End If
            mendeleyApiClient().addCitation markName, displayedText
            
            thisField.Locked = True
        End If
    Next
    
    Dim oldCitationStyle As String
    oldCitationStyle = getCitationStyleId()
    
    ' Now that we've compiled the list of uuids, give them to Mendeley Desktop
    ' and tell it to format the citations and bibliography
    If Not mendeleyApiClient().formatCitationsAndBibliography() Then
        Call bringWordToForeground
        GoTo ExitFunction
    End If
    
    citationNumber = 0
    
    marks = fnGetMarks(ZoteroUseBookmarks)
    For Each mark In marks 'ActiveDocument.Fields
        If currentDocumentPath <> activeDocumentPath() Then
            GoTo ExitFunction
        End If

        Set thisField = mark
        Dim fieldText As String
        fieldText = ""
        markName = getMarkName(thisField)

        If IsObjectValid(thisField) = False Then
            GoTo NextIterationLoop
        End If
        
        If (isMendeleyCitationField(markName)) Then

            fieldText = mendeleyApiClient().getFormattedCitation(citationNumber)
            
            Dim previousFormattedCitation As String
            previousFormattedCitation = mendeleyApiClient().getPreviouslyFormattedCitation(citationNumber)
            
            Dim jsonData As String
            jsonData = mendeleyApiClient().getFieldCode(citationNumber)

            If currentDocumentPath <> activeDocumentPath() Then
                GoTo ExitFunction
            End If
            Set thisField = fnRenameMark(thisField, jsonData)
            
            If fieldText <> getMarkText(thisField) Then
                If currentDocumentPath <> activeDocumentPath() Then
                    GoTo ExitFunction
                End If
                
                ' if Mendeley sends us an empty field, leave it alone since we want to
                ' preserve the user's formatting options
                If fieldText <> "" Then
                    Call applyFormatting(fieldText, thisField)
                End If
            End If
            
            citationNumber = citationNumber + 1
        ElseIf isMendeleyBibliographyField(markName) Then
            If Not InStr(markName, CSL_BIBLIOGRAPHY) > 0 Then
                    Call fnRenameMark(mark, markName & " " & CSL_BIBLIOGRAPHY)
            End If
        
            If bibliography = "" Then
                bibliography = bibliography + mendeleyApiClient().getFormattedBibliography()
            End If
                Dim range As range
                Set range = thisField.result
                
                ' get font used at start of bibliography
                Dim endOfRange As Long
                endOfRange = range.End
                range.Collapse (wdCollapseStart)
                
                Dim currentFontName As String
                Dim currentSize As Long
                currentFontName = range.Font.Name
                currentSize = range.Font.Size

                ' restore range end
                range.End = endOfRange

                range.InsertFile (bibliography)
                ' apply font to whole range
                Set range = thisField.result
                range.Font.Name = currentFontName
                range.Font.Size = currentSize

                ' Delete first two characters (spaces and newline) (inserted for a workaround in OO.org)
                range.End = range.Start + 2
                range.Text = ""
            
            'fieldText = bibliography
            'Call applyFormatting(fieldText, thisField)
        End If
        
        If Not (fieldText = "") Then
            ' Put text in field
                If thisField.ShowCodes Then
                    thisField.ShowCodes = False
                End If
        End If
        
        thisField.Locked = True
NextIterationLoop:
    Next
    
    If Not unitTest Then
        Dim newCitationStyle As String
        newCitationStyle = mendeleyApiClient().getCitationStyleId()
        
        If (newCitationStyle <> oldCitationStyle) Then
            ' set new citation style
            Call setCitationStyle(newCitationStyle)
            
            ' update citation styles list
            Call updateCitationStylesComboBox
        End If
        
            Set previouslySelectedField = getFieldAtSelection()
        If Not IsNull(previouslySelectedField) And Not IsEmpty(previouslySelectedField) Then
            previouslySelectedFieldResultText = getMarkText(previouslySelectedField)
        Else
            previouslySelectedFieldResultText = ""
        End If
    End If

    refreshDocument = True

ExitFunction:
    Call endUndoTransaction
    Call finishUpdatingDocument(documentState)
End Function

Sub setCitationStyle(style As String)
    Dim currentStyle As String
    currentStyle = getCitationStyleId(returnDefaultIfNotSet:=False)
    If style <> currentStyle Then
        Call subSetProperty(MENDELEY_CITATION_STYLE, style)
    End If
    
    Call mendeleyApiClient().setCitationStyle(style)
End Sub

' Returns the citation style currently saved in the document
' or DEFAULT_CITATION_STYLE otherwise.  If 'returnDefaultIfNotSet' is False,
' an empty string is returned instead of DEFAULT_CITATION_STYLE if
' no style is saved in the document
Function getCitationStyleId(Optional returnDefaultIfNotSet As Boolean = True) As String
    getCitationStyleId = fnGetProperty(MENDELEY_CITATION_STYLE)
    If getCitationStyleId = "" And returnDefaultIfNotSet Then
        getCitationStyleId = DEFAULT_CITATION_STYLE
    End If
End Function

Function getStyleNameFromId(styleId As String) As String
    ' For compatibility with old system where the name was used as the identifier
    getStyleNameFromId = styleId
    
    If Not (startsWith(styleId, "http://") Or startsWith(styleId, "https://")) Then
        Exit Function
    End If
    
    getStyleNameFromId = StyleListModel.nameForStyle(styleId)
End Function

' returns the user account which this document is currently linked to
Function mendeleyUserAccount() As String
    On Error GoTo CatchError
    
    mendeleyUserAccount = ActiveDocument.CustomDocumentProperties(MENDELEY_USER_ACCOUNT).value
    Exit Function
    
CatchError:
    mendeleyUserAccount = ""
End Function

Sub setMendeleyUserAccount(value As String)
    On Error GoTo CatchError
    
    Dim test As String
    
    ' if MENDELEY_DOCUMENT property not set this will throw an exception
    test = ActiveDocument.CustomDocumentProperties(MENDELEY_USER_ACCOUNT).value
    
    ActiveDocument.CustomDocumentProperties(MENDELEY_USER_ACCOUNT).value = value

    Exit Sub
CatchError:
    ActiveDocument.CustomDocumentProperties.Add Name:=MENDELEY_USER_ACCOUNT, _
        LinkToContent:=False, value:=value, Type:=msoPropertyTypeString
End Sub

Function getFieldAtSelection() As Field
    Dim fields() As Object
    Call getFieldsAtSelection(1, fields)
        
    If Not IsEmpty(fields) Then
        If Not fields(LBound(fields)) Is Nothing Then
            Set getFieldAtSelection = fields(LBound(fields))
        End If
    End If
End Function

Sub getFieldsAtSelection(maximumNumberToFind As Long, ByRef result() As Object)
    ' Check whether the selection is within a current field:
    '   Should be an easier way but couldn't find one.
    '   Selection.Fields only contains fields for which the start or
    '   end of the field appears within the selection
    
    Dim currentRange As range
    Set currentRange = Nothing
    
    ' Search the cursor range in the footnotes
    Dim thisFootnote As Footnote
    For Each thisFootnote In ActiveDocument.Footnotes
        If Not currentRange Is Nothing Then
            Exit For
        End If
        
        If Selection.InRange(thisFootnote.range()) = True Then
            Set currentRange = thisFootnote.range()
        End If
    Next

    ' Search the cursor range in the shapes
    Dim thisShape As Shape
    For Each thisShape In ActiveDocument.Shapes
        If Not currentRange Is Nothing Then
            Exit For
        End If
        
        If thisShape.Type = msoTextBox Then
            If thisShape.TextFrame.HasText Then
                If Selection.InRange(thisShape.TextFrame.TextRange) = True Then
                    Set currentRange = thisShape.TextFrame.TextRange
               End If
            End If
        End If
    Next

    If currentRange Is Nothing Then
        Set currentRange = ActiveDocument.range()
    End If

    ' currentRange contains the range where the cursor is right now

    Dim currentRangeStartOriginal As Long
    Dim currentRangeEndOriginal As Long

    currentRangeStartOriginal = currentRange.Start
    currentRangeEndOriginal = currentRange.End

    Dim charsToCheck As Long
    charsToCheck = 5000
    
    ' Expand selection 1 character to the right so we detect the field
    ' even if we're only at the very start of it
    Dim currentSelection As range
    Set currentSelection = Selection.range
    
    Dim shiftedSelection As range
    Set shiftedSelection = Selection.range
    Call shiftedSelection.MoveEnd(wdCharacter, 1)
    Call shiftedSelection.Collapse(wdCollapseEnd)

    ' To not check for all fields of all document (speed), it
    ' checks only the fields of the zone where we are now (-100 +100)

    If currentSelection.Start > charsToCheck Then
        currentRange.Start = currentSelection.Start - charsToCheck
    Else
        currentRange.Start = 1
    End If

    If currentSelection.End + charsToCheck < currentRange.End Then
        currentRange.End = currentSelection.End + charsToCheck
    Else
        currentRange.End = currentRange.End
    End If

    Dim currentRangeFieldsCount As Long
    currentRangeFieldsCount = currentRange.fields.count

    Dim currentRangeFields As fields
    Set currentRangeFields = currentRange.fields

    Dim i As Long
    i = 1
    Dim thisField As Field
   
    ReDim result(1 To maximumNumberToFind)
    Dim numFound As Long
    numFound = 0

    While i <= currentRangeFieldsCount
        Set thisField = currentRangeFields(i)
        If thisField.Type = wdFieldQuote Or thisField.Type = wdFieldRef Or thisField.Type = wdFieldAddin Then
            If currentSelection.InRange(thisField.result) Or _
               shiftedSelection.InRange(thisField.result) Or _
               (thisField.result.Start > currentSelection.Start And thisField.result.Start < shiftedSelection.Start) Then
                numFound = numFound + 1
                Set result(numFound) = currentRange.fields(i)
                If numFound >= maximumNumberToFind Then
                    GoTo EndOfSub
                End If
            End If
        End If
        i = i + 1
    Wend

EndOfSub:
    
    currentRange.Start = currentRangeStartOriginal
    currentRange.End = currentRangeEndOriginal
End Sub

' Returns connection status
Function launchMendeleyIfNecessary() As Long
    ' Only need to launch Mendeley if the document contains the
    ' MENDELEY property
    If Not isMendeleyDocument Then
        launchMendeleyIfNecessary = CONNECTION_NOT_A_MENDELEY_DOCUMENT
        Exit Function
    End If

    If mendeleyApiClient().launchMendeley() Then
        launchMendeleyIfNecessary = CONNECTION_CONNECTED
    Else
        launchMendeleyIfNecessary = CONNECTION_MENDELEY_DESKTOP_NOT_FOUND
    End If
End Function

Function isMendeleyRunning() As Boolean
    isMendeleyRunning = mendeleyApiClient().isMendeleyRunning()
End Function

Function isMendeleyDocument() As Boolean
    If fnGetProperty(MENDELEY_DOCUMENT) = "True" Then
        isMendeleyDocument = True
    Else
        isMendeleyDocument = False
    End If
End Function

Sub setMendeleyDocument(value As Boolean)
    If value Then
        Call subSetProperty(MENDELEY_DOCUMENT, "True")
    Else
        Call subSetProperty(MENDELEY_DOCUMENT, "False")
    End If
End Sub

Sub createCitationStyleComboBox()
    Dim mendeleyControl As CommandBarControl
    
    For Each mendeleyControl In CommandBars(TOOLBAR).Controls
        If mendeleyControl.Parameter = TOOLBAR_CITATION_STYLE Then
            ' clear existing styles in combo box
            Call mendeleyControl.Clear
            Exit Sub
        End If
    Next
    
    ' Create combo box
    Call CommandBars(TOOLBAR).Controls.Add( _
        Type:=msoControlComboBox, Parameter:=TOOLBAR_CITATION_STYLE, Before:=5)
End Sub

Sub setUpContextMenus()
    ' add to "Text" command bar, which is the context menu which shows
    ' outside Mendeley fields

    Dim textCommandBar As CommandBar
    Set textCommandBar = CommandBars("Text")
    Dim menuItem As CommandBarControl
    
    Dim foundMendeleyItem As Boolean
    foundMendeleyItem = False
    
    Const ENABLE_MENU_ICONS = False
    
    ' Get the icon from the Insert/Edit citation button
    Dim mendeleyControl As CommandBarControl
    
    If ENABLE_MENU_ICONS Then
        For Each mendeleyControl In CommandBars(TOOLBAR).Controls
            If mendeleyControl.Caption = TOOLBAR_INSERT_CITATION Or mendeleyControl.Caption = TOOLBAR_EDIT_CITATION Then
                ' clear existing styles in combo box
                Call mendeleyControl.CopyFace
            End If
        Next
    End If
    
    Dim insertCitationButton As CommandBarControl
    Dim beforeButtonIndex As Long
    beforeButtonIndex = 0
    For Each menuItem In textCommandBar.Controls
        If menuItem.id = 22 Then ' id 22 has caption "&Paste" in English
            beforeButtonIndex = menuItem.index + 1
        End If
        ' Delete old menu item
        If menuItem.Caption = "Insert Mendeley Citation" Then
            foundMendeleyItem = True
            Set insertCitationButton = menuItem
            insertCitationButton.Delete
        End If
    Next
    
    Set insertCitationButton = textCommandBar.Controls.Add(msoControlButton, Before:=beforeButtonIndex)
    insertCitationButton.Caption = "Insert Mendeley Citation"
    insertCitationButton.OnAction = "insertCitation"
    insertCitationButton.BeginGroup = True
    If ENABLE_MENU_ICONS Then
        Call insertCitationButton.PasteFace
    End If
    
    ' add to "Fields" context menu which shows when inside a field
    ' (slight problem is that it is also used for non-Mendeley fields)
    Dim editCitationButton As CommandBarControl
    beforeButtonIndex = 0
    For Each menuItem In CommandBars("Fields").Controls
        If menuItem.id = 22 Then ' id 22 has caption "&Paste" in English
            beforeButtonIndex = menuItem.index + 1
        End If
        ' Delete old menu item
        If menuItem.Caption = "Edit Mendeley Citation" Then
            foundMendeleyItem = True
            Set editCitationButton = menuItem
            editCitationButton.Delete
        End If
    Next
    Set editCitationButton = CommandBars("Fields").Controls.Add(msoControlButton, Before:=beforeButtonIndex)
    editCitationButton.Caption = "Edit Mendeley Citation"
    editCitationButton.OnAction = "insertCitation"
    editCitationButton.BeginGroup = True
    If ENABLE_MENU_ICONS Then
        Call editCitationButton.PasteFace
    End If
End Sub

Sub createCitationStyleComboBoxAndQuit()
    Call createCitationStyleComboBox
    ThisDocument.Saved = False
    ' Save to file and quit
    Call ThisDocument.Save
    ThisDocument.Close
    
    Call Application.Quit
End Sub

Function getCitationStyleComboBox() As CommandBarComboBox
    Dim mendeleyControl As CommandBarControl
    
    ' @todo: cache comboBox again, but need a map of documents -> comboBoxes
    '       instead of just the one
    
    For Each mendeleyControl In CommandBars(TOOLBAR).Controls
        If mendeleyControl.Parameter = TOOLBAR_CITATION_STYLE Then
            Set getCitationStyleComboBox = mendeleyControl
            Exit Function
        End If
    Next
    ' if here, combo box hasn't been created yet
    ' (this was happening very rarely, looks safe to create now
    ' and subsequent calls will have the combobox)
    
    ' (maybe caused by a commented race condition in initialise)
    'Call createCitationStyleComboBox
    If Not buildingPlugin() Then
        MsgBox "Combo box not found"
    End If
End Function

Function getUndoEditButton() As CommandBarButton
    Dim mendeleyControl As CommandBarControl
    
    For Each mendeleyControl In CommandBars(TOOLBAR).Controls
        If mendeleyControl.Caption = TOOLBAR_UNDO_EDIT Then
            Set getUndoEditButton = mendeleyControl
            Exit Function
        End If
    Next
    ' if here, button hasn't been created yet
    MsgBox "Undo edit button not found"
End Function

Function getInsertCitationButton() As CommandBarButton
    Dim mendeleyControl As CommandBarControl
    
    For Each mendeleyControl In CommandBars(TOOLBAR).Controls
        If mendeleyControl.Caption = TOOLBAR_INSERT_CITATION Or mendeleyControl.Caption = TOOLBAR_EDIT_CITATION Then
            Set getInsertCitationButton = mendeleyControl
            Exit Function
        End If
    Next
    ' if here, button hasn't been created yet
    MsgBox "Insert citation button not found"
End Function

' Returns true if the comboBox contains item in its list
Function comboBoxContains(ComboBox As CommandBarComboBox, item As String) As Boolean
    Dim index As Long
    
    For index = 1 To ComboBox.ListCount
        If item = ComboBox.List(index) Then
            comboBoxContains = True
            Exit Function
        End If
    Next
    
    comboBoxContains = False
End Function

' Update the citation style for the current document
' from the current selection in the style combobox
Sub updateStyleFromComboBox()
    Dim citationStyleComboBox As CommandBarComboBox
    Set citationStyleComboBox = getCitationStyleComboBox()
    
    Dim newStyleId As String
    newStyleId = StyleListModel.styleIdFromIndex(citationStyleComboBox.ListIndex)
    If newStyleId <> "" Then
        Call setCitationStyle(newStyleId)
    End If
End Sub

Sub PullStyles()
    Dim ComboBox As CommandBarComboBox
    Set ComboBox = getCitationStyleComboBox()
    
    If Not isMendeleyRunning() Then
        Dim temp
        temp = mendeleyApiClient().launchMendeley()
    End If

    If ComboBox.ListCount = 1 Then        ' Select me to fetch styles
        Call updateCitationStylesComboBox
    End If
    
    If ComboBox.ListCount > 1 And ComboBox.ListIndex = ComboBox.ListCount Then   ' More Styles...
        Call OpenCitationsFromMendeley
    End If
End Sub

' Refresh the contents of the citation style combobox.  If 'fetchStyles'
' is false or Mendeley Desktop is not running, the combobox will display
' a single item indicating that the citation style list will be updated when
' "Insert Citation" is clicked
Sub updateCitationStylesComboBox(Optional fetchStyles As Boolean = True)
    Dim mendeleyControl As CommandBarControl
        
    Dim ComboBox As CommandBarComboBox
    Set ComboBox = getCitationStyleComboBox()
    
    ComboBox.TooltipText = TOOLTIP_CITATION_STYLE
    ComboBox.OnAction = "PullStyles"

    Call StyleListModel.refreshStyles(fetchStyles)
    
    If StyleListModel.count() = 0 Then
        ComboBox.Clear
        ComboBox.AddItem SELECT_ME_FETCH_STYLES, 1
        ComboBox.Width = 180
        ComboBox.ListIndex = 1
        Exit Sub
    End If

    ComboBox.Clear
    Dim citationStyle As Variant
    Dim index As Long
        
    For index = 1 To StyleListModel.count()
        ComboBox.AddItem StyleListModel.styleNameFromIndex(index)
    Next
    
    Call ComboBox.AddItem("More Styles...")
    
    ' Set default style if none set
    If getCitationStyleId() = "" Then
        Dim defaultStyle As String
        defaultStyle = DEFAULT_CITATION_STYLE_NAME
        If comboBoxContains(ComboBox, defaultStyle) Then
            setCitationStyle (defaultStyle)
        Else
            setCitationStyle (StyleListModel.styleNameFromIndex(1))
        End If
    End If
    
    ' display current style in combo box
    ' check first whether combo box has a different text - we do this because
    ' although getCitationStyle gets the citation style for the current document,
    ' ComboBox.Text = ?? will set it for ALL the currently open documents which isn't desired
    Dim styleName As String
    styleName = StyleListModel.nameForStyle(getCitationStyleId())
    If ComboBox.Text <> styleName Then
        ComboBox.Text = styleName
    End If
    
    ' set current style for Ribbon
    Dim currentStyleId As String
    currentStyleId = getCitationStyleId()
    ribbonSelectedStyleIndex = StyleListModel.indexFromStyleId(currentStyleId) - 1
    
    ' subscribe to events
    Set eventClassModuleInstance.ComboBox = ComboBox
    
    ' adjust width of combo box as default is a bit small
    ComboBox.Width = 160
    
    ' stop word trying to save changes to this template
    ThisDocument.Saved = True
End Sub

Sub addTooltipsToButtons()
    If CommandBars(TOOLBAR).Visible Then
        CommandBars(TOOLBAR).Controls(TOOLBAR_INSERT_CITATION).TooltipText = TOOLTIP_INSERT_CITATION
        CommandBars(TOOLBAR).Controls(TOOLBAR_UNDO_EDIT).TooltipText = TOOLTIP_UNDO_EDIT
        CommandBars(TOOLBAR).Controls(TOOLBAR_MERGE_CITATIONS).TooltipText = TOOLTIP_MERGE_CITATIONS
        CommandBars(TOOLBAR).Controls(TOOLBAR_INSERT_BIBLIOGRAPHY).TooltipText = TOOLTIP_INSERT_BIBLIOGRAPHY
        CommandBars(TOOLBAR).Controls(TOOLBAR_REFRESH).TooltipText = TOOLTIP_REFRESH
        
        ' The "Export..." button doesn't have a Parameter member so find by Caption
        Dim mendeleyControl As CommandBarControl
        Dim exportMenu As CommandBarPopup
        
        For Each mendeleyControl In CommandBars(TOOLBAR).Controls
            If mendeleyControl.Caption = "Export..." Then
                Set exportMenu = mendeleyControl
            End If
        Next
        
        If Not exportMenu.Caption = TOOLBAR_EXPORT Then
            MsgBox "export popup not found"
        Else
            exportMenu.TooltipText = TOOLTIP_EXPORT
            exportMenu.Controls("Compatible with OpenOffice").TooltipText = TOOLTIP_EXPORT_OPENOFFICE
        End If
    End If
    
    ' stop word trying to save changes to this template
    ThisDocument.Saved = True
End Sub

Function isDocumentLinkedToCurrentUser() As Boolean
    Dim currentMendeleyUser As String
    Dim thisDocumentUser As String
    
    currentMendeleyUser = mendeleyApiClient().getUserAccount()
    thisDocumentUser = fnGetProperty(MENDELEY_USER_ACCOUNT)
    
    ' remove server protocol from account string
    thisDocumentUser = Replace(thisDocumentUser, "http://", "")
    thisDocumentUser = Replace(thisDocumentUser, "https://", "")
    
    If currentMendeleyUser = thisDocumentUser Then
        isDocumentLinkedToCurrentUser = True
    Else
        Dim result ' As VbMsgBoxResult
        
        Dim vbCrLf
        vbCrLf = Chr(13)
        
        If thisDocumentUser = "" Then
            ' if no user currently linked then set without asking user
            result = MSGBOX_RESULT_YES
        Else
            ' ask user if they want to link the document to their account
            result = MsgBox("This document has been edited by another Mendeley user: " + thisDocumentUser + vbCrLf + vbCrLf + _
                "Do you wish to enable the Mendeley plugin to edit the citations and bibliography yourself?" + vbCrLf + vbCrLf, _
                MSGBOX_BUTTONS_YES_NO, "Enable Mendeley plugin for this document?")
        End If

        If result = MSGBOX_RESULT_YES Then
            Call subSetProperty(MENDELEY_USER_ACCOUNT, currentMendeleyUser)
            isDocumentLinkedToCurrentUser = True
        Else
            isDocumentLinkedToCurrentUser = False
        End If
    End If
End Function


' Returns true if mainString starts with the subString
Function startsWith(mainString As String, subString As String) As Boolean
    startsWith = Left(mainString, Len(subString)) = subString
End Function

Sub applyFormatting(markup As String, mark As Field)
    ' parse range and apply following formatting:
    ' <i></i> italics
    ' <b></b> bold
    ' <u></u> underline
    ' <sup></sup> superscript
    ' <sub></sub> subscript

    ' add extra space at start because the Range.Delete function will
    ' delete the whole field if we attempt to delete the first character
    ' (it gets deleted later)

    Call subSetMarkText(mark, markup)

    Dim range As range
    Set range = mark.result

    Dim subRange As range
    Set subRange = range.Duplicate
    
    ' remove currently existing formatting
    range.bold = False
    range.italic = False
    range.underline = False
    range.Font.superscript = False
    range.Font.subscript = False
    
    Dim startPosition As Long
    Dim endPosition As Long
    
    startPosition = range.Start
    endPosition = range.End
    
    Call applyStyleToTagPairs("i", subRange, startPosition, endPosition)
    Call applyStyleToTagPairs("b", subRange, startPosition, endPosition)
    Call applyStyleToTagPairs("u", subRange, startPosition, endPosition)
    Call applyStyleToTagPairs("sup", subRange, startPosition, endPosition)
    Call applyStyleToTagPairs("sub", subRange, startPosition, endPosition)
    
    If InStr(getMarkText(mark), "second-field-align") > 0 Or InStr(getMarkText(mark), "hanging-indent") Then
        ' hanging indent
        range.ParagraphFormat.FirstLineIndent = 0
        range.ParagraphFormat.LeftIndent = 0
        range.ParagraphFormat.RightIndent = 0
    End If
    
    Call applyStyleToTagPairs("second-field-align", subRange, startPosition, endPosition)
    Call applyStyleToTagPairs("hanging-indent", subRange, startPosition, endPosition)
    
    ' convert <unicode>$CHARCODE</unicode> tags to corresponding characters
    ' this is still used by the applyFormatting() test but is not used
    ' by the plugin itself
    Call applyStyleToTagPairs("unicode", subRange, startPosition, endPosition)

    ' Add paragraph breaks in place of <p>
    Call applyStyleToIndividualTags("p", subRange, startPosition, endPosition)
End Sub

Sub applyStyleToTagPairs(tag As String, wholeRange As range, _
    startPosition As Long, endPosition As Long)

    Dim startTag As String
    Dim endTag As String
    
    Dim thisRange As range
    Set thisRange = wholeRange.Duplicate
    
    startTag = "<" + tag + ">"
    endTag = "</" + tag + ">"
    
    ' Maximum number of characters used in the first field
    ' Used for setting the second-field-align tab stopa
    Dim maxFirstFieldLength As Long
    maxFirstFieldLength = 0

    Do While Not (rangeString(thisRange) = "") And Not (InStr(rangeString(thisRange), startTag) = 0)
        ' find and remove start tag
        thisRange.Start = startPosition
        thisRange.End = endPosition
        
        Dim startTagPosition As Long

        startTagPosition = InStr(rangeString(thisRange), startTag) - 1
        thisRange.Start = startPosition + startTagPosition
        ' Conditional needed to deal with strange VBA behaviour
        If startTagPosition = 0 Then
            ' deleting this way when at the start avoids deleting the whole field
            thisRange.End = thisRange.Start + 2 + Len(tag)
            thisRange.Text = ""
            'thisRange.Clear is not deleting
        Else
            ' deleting this way when not at the start avoids adding an extra space
            ' in place of the deleted range (VBA is crazy)
            thisRange.End = thisRange.Start
            Call thisRange.Delete(Word.WdUnits.wdCharacter, 2 + Len(tag))
        End If
        
        ' find and remove end tag
        
        thisRange.Start = startPosition
        thisRange.End = endPosition

        Dim endTagPosition As Long
        
        endTagPosition = InStr(rangeString(thisRange), endTag) - 1
        thisRange.Start = startPosition + endTagPosition
        thisRange.End = thisRange.Start + 3 + Len(tag)
        thisRange.Text = ""
        
        thisRange.Start = startPosition + startTagPosition
        thisRange.End = startPosition + endTagPosition
        
        ' apply style
        Select Case tag
            Case "b"
                thisRange.bold = 1
            Case "i"
                thisRange.italic = 1
            Case "u"
                thisRange.underline = 1
            Case "sup"
                thisRange.Font.superscript = 1
            Case "sub"
                thisRange.Font.subscript = 1
            Case "second-field-align"
                ' Remove spaces at the end of the range
                ' (@todo remove this if fixed in Mendeley Desktop)
                Do While Right(thisRange.Text, 1) = " "
                    thisRange.Start = startPosition + endTagPosition - 1
                    thisRange.End = startPosition + endTagPosition
                    
                    Call thisRange.Delete
                    endTagPosition = endTagPosition - 1
                Loop
                
                If endTagPosition - startTagPosition > maxFirstFieldLength Then
                    maxFirstFieldLength = endTagPosition - startTagPosition
                End If
                
                ' remove subsequent spaces after the range
                thisRange.Start = startPosition + endTagPosition
                thisRange.End = startPosition + endTagPosition + 1
                Do While thisRange.Text = " "
                    Call thisRange.Delete(Word.WdUnits.wdCharacter, 2)
                    
                    thisRange.Start = startPosition + endTagPosition
                    thisRange.End = startPosition + endTagPosition + 1
                Loop
                thisRange.End = startPosition + endTagPosition
                
                ' insert tab
                thisRange.InsertAfter (vbTab)
                
            Case "hanging-indent"
                ' set indent to 1 tab stop for whole range
                Call setHangingIndent(wholeRange)
                
            Case "unicode"
                Dim characterCode As Long
                characterCode = thisRange.Text
                
                ' Note: an old version of the code used thisRange.Delete(wdCharacter, length)
                ' to avoid a supposed bug where a space was removed after the range,
                ' but this introduced bug #17191
                ' The following line doesn't have any unwanted side-effects I'm aware of
                thisRange.Text = ""
                Call thisRange.InsertAfter(ChrW(characterCode))
        End Select

        thisRange.Start = startPosition
        thisRange.End = endPosition
    Loop
    
    If tag = "second-field-align" And maxFirstFieldLength > 0 Then
        ' Create tab stop - the position is calculated approximately
        ' and works reasonably well for numbered citations
        Call setHangingIndent(wholeRange, 4 + 6 * maxFirstFieldLength)
    End If
    
End Sub

Sub setHangingIndent(range As range, Optional length As Long)
    If Not (length = 0) Then
        Call range.ParagraphFormat.TabStops.ClearAll
        range.ParagraphFormat.TabStops.Add (length)
    End If
    range.ParagraphFormat.FirstLineIndent = 0
    range.ParagraphFormat.LeftIndent = 0
    range.ParagraphFormat.TabHangingIndent (1)
End Sub

Function rangeString(range As range) As String
        rangeString = range.Text
End Function

Sub applyStyleToIndividualTags(tag As String, wholeRange As range, _
    startPosition As Long, endPosition As Long)

    Dim startTag As String
    Dim endTag As String
    
    startTag = "<" + tag + ">"
    endTag = "</" + tag + ">"
    
    Dim thisRange As range
    Set thisRange = wholeRange.Duplicate
    thisRange.Start = startPosition
    thisRange.End = endPosition

    Do While Not (rangeString(thisRange) = "") And Not (InStr(rangeString(thisRange), startTag) = 0)
        ' find and remove start tag
        Dim startTagPosition As Long
        
            startTagPosition = InStr(rangeString(thisRange), startTag) - 1
            thisRange.Start = startPosition + startTagPosition
            thisRange.End = thisRange.Start
            Call thisRange.Delete(Word.WdUnits.wdCharacter, 2 + Len(tag))
            
            thisRange.Start = startPosition + startTagPosition
            thisRange.End = startPosition + startTagPosition

        ' apply formatting
        Select Case tag
            Case "p"
                Call thisRange.InsertParagraph
        End Select
        
            thisRange.Start = startPosition
            thisRange.End = endPosition
    Loop
End Sub

Sub checkForCitationEdit()
    ' We can't update the document until we've updated the previously
    ' selected field, so if the user cancels an edit, needUpdate is set
    ' to true and we UpdateDocument() at the end
    Dim needUpdate As Boolean
    needUpdate = False
    
    ' TODO: deal with edits of bibliographies
    Dim objectExists As Boolean
    objectExists = IsObjectValid(previouslySelectedField)
    If Not (previouslySelectedField Is Nothing) And Not IsMissing(previouslySelectedField) And Not IsEmpty(previouslySelectedField) And objectExists Then
        If Not (previouslySelectedField.result.Text = previouslySelectedFieldResultText) Then
            
            If Not isMendeleyRunning() Then
                ' Don't need do anything - this edit will get detected next time the user refreshes
                Exit Sub
            End If
            
            Dim markName As String
            markName = getMarkName(previouslySelectedField)
            
            Dim displayedText As String
            displayedText = getMarkText(previouslySelectedField)
            displayedText = Replace(displayedText, Chr(13), "<p>")
            
            Dim newMarkName As String
            newMarkName = mendeleyApiClient().checkManualFormatAndGetFieldCode(markName, displayedText)
            
            If markName <> newMarkName Then
                Call fnRenameMark(previouslySelectedField, newMarkName)
                ' Disabled until we can send rich text formatting tags from the displayed
                ' citations to Mendeley Desktop
                'Call subSetMarkText(previouslySelectedField, displayedText)
                needUpdate = True
            End If
                        
        End If
    End If
    
    If needUpdate Then
        Call refreshDocument
    End If
End Sub

Sub beginUndoTransaction(action As String)
    ' Application.UndoRecord exists in Word 2010 and later
#If VBA7 Then
    Application.UndoRecord.StartCustomRecord action
#End If
End Sub

Sub endUndoTransaction()
#If VBA7 Then
    Application.UndoRecord.EndCustomRecord
#End If
End Sub
' ---------------------------
'    Utility Functions
' ---------------------------

' Returns the index of item in the array, or -1 if not found
' (doesn't permit arrays with -ve lower bound)
Function indexOf(container() As String, item As String) As Long
    Dim index As Long
    
    If LBound(container) < 0 Then
        MsgBox "indexOf doesn't permit lower bounds < 0"
        Exit Function
    End If
    
    For index = LBound(container) To UBound(container)
        If container(index) = item Then
            indexOf = index
            Exit Function
        End If
    Next
    
    ' not found
    indexOf = -1
End Function