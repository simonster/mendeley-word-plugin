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

' The API connection point
Private theApiClient As Object

' These need to be Global as they are accessed from EventClassModule
Global previouslySelectedField As Field
Global previouslySelectedFieldResultText As String

Global Const DEBUG_MODE = False

Global Const TEMPLATE_NAME_DURING_BUILD = "MendeleyPlugin"

Global Const MENDELEY_DOCUMENT = "Mendeley Document"
Global Const MENDELEY_USER_ACCOUNT = "Mendeley User Name"
Global Const MENDELEY_CITATION = "Mendeley Citation"
Global Const MENDELEY_CITATION_EDITOR = "Mendeley Citation Editor"
Global Const MENDELEY_CITATION_MAC = " PRINTDATE Mendeley Citation"
Global Const MENDELEY_EDITED_CITATION = "Mendeley Edited Citation"
Global Const MENDELEY_BIBLIOGRAPHY = "Mendeley Bibliography"
Global Const MENDELEY_BIBLIOGRAPHY_MAC = " PRINTDATE Mendeley Bibliography"
Global Const MENDELEY_CITATION_STYLE = "Mendeley Citation Style"
Global Const DEFAULT_CITATION_STYLE = "http://www.zotero.org/styles/apa"
Global Const DEFAULT_CITATION_STYLE_NAME = "American Psychological Association"
Global Const SELECT_ME_FETCH_STYLES = "Select me to fetch the styles"

Global Const CSL_CITATION = "CSL_CITATION "
Global Const CSL_BIBLIOGRAPHY = "CSL_BIBLIOGRAPHY"
Global Const CSL_BIBLIOGRAPHY_OLD = "CSL_BIBLIOGRAPHY "
Global Const INSERT_CITATION_TEXT = "{Formatting Citation}"
Global Const CITATION_EDIT_TEXT = ""
Global Const BIBLIOGRAPHY_TEXT = "{Bibliography}"
Global Const MERGING_TEXT = "{Merging Citations}"
Global Const TOOLBAR = "Mendeley Toolbar"
Global Const TOOLBAR_CITATION_STYLE = "Citation Style"
Global Const TOOLBAR_INSERT_CITATION = "Insert Citation"
Global Const TOOLBAR_EDIT_CITATION = "Edit Citation"
Global Const TOOLBAR_MERGE_CITATIONS = "Merge Citations"
Global Const TOOLBAR_INSERT_BIBLIOGRAPHY = "Insert Bibliography"
Global Const TOOLBAR_REFRESH = "Refresh"
Global Const TOOLBAR_EXPORT = "Export..."
Global Const TOOLBAR_UNDO_EDIT = "Undo Edit"
Global Const MERGE_CITATIONS_NOT_ENOUGH_CITATIONS = "Please select at least two citations to merge."
Global Const CITATIONS_NOT_ADJACENT = "Citations must be adjacent to merge."
Global Const CITATION_ADJECENT_LIMIT = 4
Global Const MACRO_ALREADY_RUNNING = "Waiting For Response From Mendeley Desktop"
Global Const RECENT_STYLE_NAME = "Mendeley Recent Style Name"
Global Const RECENT_STYLE_ID = "Mendeley Recent Style Id"

Global Const TOOLTIP_INSERT_CITATION = "Insert a new citation (Alt-M)"
Global Const TOOLTIP_EDIT_CITATION = "Edit the selected citation (Alt-M)"
Global Const TOOLTIP_UNDO_EDIT = "Undo custom edit of the selected citation"
Global Const TOOLTIP_MERGE_CITATIONS = "Merge the selected citations into one"
Global Const TOOLTIP_INSERT_BIBLIOGRAPHY = "Insert a bibliography"
Global Const TOOLTIP_REFRESH = "Refresh citations and bibliographies"
Global Const TOOLTIP_CITATION_STYLE = "Select a citation style"
Global Const TOOLTIP_EXPORT_OPENOFFICE = "Export a copy of the document compatible with OpenOffice"
Global Const TOOLTIP_EXPORT = "Export the document with different options"

Global Const CONNECTION_CONNECTED = 0
Global Const CONNECTION_VERSION_MISMATCH = 1
Global Const CONNECTION_MENDELEY_DESKTOP_NOT_FOUND = 2
Global Const CONNECTION_NOT_A_MENDELEY_DOCUMENT = 3

' The following dictate the maximum length of the strings to store in each of these data types
Global Const MAX_PROPERTY_LENGTH = 255
Global Const BOOKMARK_ID_STRING_LENGTH = 10
Global Const ZOTERO_BOOKMARK_REFERENCE_PROPERTY = "Mendeley_Bookmark"

'The following constants describe a location in a document
Global Const ZOTERO_ERROR = 0 'Frame, comments, header, footer
Global Const ZOTERO_MAIN = 1 'Main document including things like tables (wdMainTextStory)
Global Const ZOTERO_FOOTNOTE = 2 'Footnote (wdFootnotesStory)
Global Const ZOTERO_ENDNOTE = 3 'Endnote (wdEndnotesStory)
Global Const ZOTERO_TABLE = 4 ' Inside a Table

Global Const MSGBOX_RESULT_YES = 6
Global Const MSGBOX_BUTTONS_YES_NO = 4

Global eventClassModuleInstance As New EventClassModule
Global StyleListModel As New StyleListModel

Global initialised As Boolean

' to prevent user from performing actions while we're still in the middle of another one
Global uiDisabled As Boolean
Global awaitingResponseFromMD As Boolean
Global ZoteroUseBookmarks As Boolean

Global seedGenerated As Boolean

Global unitTest As Boolean

' initialise on word startup and on new / open document
Public Sub AutoExec()
    If Not USE_RIBBON Then
        Call initialise
    End If
End Sub
Public Sub AutoNew()
    Call initialise
End Sub
Public Sub AutoOpen()
    If buildingPlugin() Then
        Exit Sub
    End If
    
    Call initialise
End Sub
' Returns the COM server used for communication with Mendeley Desktop.
' The COM server is started on-demand the first time that this function is called.
' Except when loading a document that already has Mendeley citations, this should
' not be called until the user interacts with the 'Mendeley Cite-O-Matic' toolbar/ribbon.
Public Function mendeleyApiClient() As Object
    If theApiClient Is Nothing Then
        Dim PLUGIN_CLASS_NAME As String
        PLUGIN_CLASS_NAME = "MendeleyWordPlugin.PluginComServer.1"
        
        On Error GoTo StartApiClient
        Set theApiClient = GetObject(, PLUGIN_CLASS_NAME)
        
StartApiClient:
        If theApiClient Is Nothing Then
            Set theApiClient = CreateObject(PLUGIN_CLASS_NAME)
        End If
    End If
    Set mendeleyApiClient = theApiClient
End Function
Public Sub initialise()
    uiDisabled = True
    ThisDocument.Saved = True
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    
    ' Subscribe to events
    Set eventClassModuleInstance.App = Word.Application
    
    ' Update the citation styles box.  We avoid fetching the list
    ' of citation styles here to avoid blocking Word startup whilst
    ' loading the plugin's COM server and trying to contact Mendeley Desktop
    
    ' Possible race condition here
    If initialised Then
        Call updateCitationStylesComboBox(fetchStyles:=False)
        ThisDocument.Saved = True
        GoTo EndOfSub
    End If
    initialised = True

    ' set up combo box
    Call updateCitationStylesComboBox(fetchStyles:=False)
    
    ' Set up context menus
    ' DISABLED until we find a way to keep the items only for one instance
    ' or to remove on uninstall of Mendeley Desktop (#15163)
    'Call setUpContextMenus
    
    Call addTooltipsToButtons

    ' hide undo button
    getUndoEditButton().Visible = False

    ' stop word trying to save changes to this template
    ThisDocument.Saved = True
    
    GoTo EndOfSub
    
ErrorHandler:
    Call reportError
    
EndOfSub:
    ' enable the mendeley plugin ui
    uiDisabled = False
End Sub


' ----- Top level functions - those directly triggered by user actions -----

Sub privateInsertCitation(Optional hintText As String, Optional documentUuid As String)
    Dim currentMark As Field
    
    Dim bringToForeground As Boolean
    bringToForeground = False
    
    Dim citeField As Field

    Set currentMark = getFieldAtSelection()

    Dim markName As String
    Dim undoActionText As String
    undoActionText = "Insert Mendeley Citation"
    
    If Not (currentMark Is Nothing) And Not IsEmpty(currentMark) Then
        Dim fieldType As String
    
        markName = getMarkName(currentMark)
        
        If isMendeleyCitationField(markName) Then
            undoActionText = "Edit Mendeley Citation"
        ElseIf isMendeleyBibliographyField(markName) Then
            MsgBox "Bibliographies are generated automatically and cannot be manually edited"
            GoTo EndOfSub
        Else
            MsgBox "This is not an editable citation."
            GoTo EndOfSub
        End If
    End If
    
    Call setMendeleyDocument(True)
      
    Dim connectionStatus As Long
    connectionStatus = launchMendeleyIfNecessary()
    If connectionStatus = CONNECTION_MENDELEY_DESKTOP_NOT_FOUND Then
        MsgBox "Please run Mendeley Desktop before using the plugin"
        GoTo EndOfSub
    End If
    
    Call sendWordProcessorVersion
    
    Dim useCitationEditor As Boolean
    useCitationEditor = True
    
    ZoteroUseBookmarks = False

    Dim selectedRange As range
    Set selectedRange = fnSelection()
    If (selectedRange Is Nothing) Then Return
    
    Dim citationText As String
    If useCitationEditor Then
        citationText = CITATION_EDIT_TEXT
        Set citeField = getFieldAtSelection()
    Else
        citationText = MENDELEY_CITATION
    End If
    
    If (citeField Is Nothing) Or IsEmpty(citeField) Then
        If selectedRange.Characters.count > 15 Then
            Dim result
            result = MsgBox("Are you sure you want to replace the following with a citation:" & _
                vbCrLf & vbCrLf & selectedRange.Text, MSGBOX_BUTTONS_YES_NO, "Insert Citation?")
            
            If result <> MSGBOX_RESULT_YES Then
                GoTo EndOfSub
            End If
        End If
    End If
        
    If connectionStatus = CONNECTION_CONNECTED Then
        If Not isDocumentLinkedToCurrentUser Then
            GoTo EndOfSub
        End If
    
        Dim buttonText As String
        buttonText = "Send Citation\nto Word;Cancel\nCitation"
        
        awaitingResponseFromMD = True
        
        ' TODO Pass hint
        Call allowMendeleyToSetForeground
        
        Dim fieldCode As String
        If documentUuid <> "" Then
            fieldCode = mendeleyApiClient().getFieldCodeFromUuid(documentUuid)
        Else
            fieldCode = mendeleyApiClient().getFieldCodeFromCitationEditor(markName)
        End If
        
        awaitingResponseFromMD = False
        
        bringToForeground = True
        Call bringWordToForeground
     
        ' check for null result:
        If (Len(fieldCode) = 0) Or ((Len(fieldCode) = 1) And (fieldCode = "")) Then
            ' MsgBox "No Citation Received from Mendeley"
            GoTo EndOfSub
        End If
        
        ' check if another instance of Word is awaiting a response from Mendeley
        If fieldCode = "<CURRENTLY-PROCESSING-REQUEST>" Then
            MsgBox "You can only make one citation at a time, " + _
                "please choose the documents for your initial citation in Mendeley Desktop first"
            GoTo EndOfSub
        End If
        
        Call beginUndoTransaction(undoActionText)
        
        If (currentMark Is Nothing) Or IsEmpty(currentMark) Then
            Set citeField = fnAddMark(selectedRange, citationText)
        Else
            Set citeField = currentMark
        End If
           
        'citeField.result.Text = INSERT_CITATION_TEXT
        citeField.code.Text = fieldCode
        Call refreshDocument(False)
        
        Call endUndoTransaction
    End If
    
    GoTo EndOfSub
    
ErrorHandler:
    Call reportError
    
EndOfSub:
    If Not (citeField Is Nothing) And IsObjectValid(citeField) Then
        If getMarkText(citeField) = INSERT_CITATION_TEXT Or _
            getMarkText(citeField) = CITATION_EDIT_TEXT Then
                citeField.Delete
        End If
    End If
    
    If bringToForeground Then
        Call bringWordToForeground
    End If
End Sub

Sub insertBibliography()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    ZoteroUseBookmarks = False
    
    If Not getFieldAtSelection() Is Nothing Then
        MsgBox "A bibliography cannot be inserted within another citation or bibliography."
        GoTo EndOfSub
    End If
    
    Call setMendeleyDocument(True)
    
    If Not (launchMendeleyIfNecessary() = CONNECTION_CONNECTED) Then
        GoTo EndOfSub
    End If
    
    If Not isDocumentLinkedToCurrentUser Then
        GoTo EndOfSub
    End If
    
    Call beginUndoTransaction("Insert Mendeley Bibliography")
    
    Dim thisField 'As Field
    Set thisField = fnAddMark(fnSelection(), "ADDIN " & MENDELEY_BIBLIOGRAPHY & " " & CSL_BIBLIOGRAPHY_OLD)
    Call refreshDocument
    
    Call endUndoTransaction
    
    GoTo EndOfSub
    
ErrorHandler:
    Call reportError
    
EndOfSub:
    uiDisabled = False
End Sub

Sub undoEdit()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Const NOT_IN_EDITABLE_CITATION_TITLE = "Undo Citation Edit"
    Const NOT_IN_EDITABLE_CITATION_TEXT = "Place cursor within an edited citation and press this button to undo the edit"
    
    Dim currentMark As Field
    
    If Not IsEmpty(getFieldAtSelection()) Then
        Set currentMark = getFieldAtSelection()
    End If
    
    If currentMark Is Nothing Or IsEmpty(currentMark) Then
        MsgBox NOT_IN_EDITABLE_CITATION_TEXT, 1, NOT_IN_EDITABLE_CITATION_TITLE
        GoTo EndOfSub
    End If
    
    If Not isDocumentLinkedToCurrentUser Then
        GoTo EndOfSub
    End If
    
    Dim markName As String
    markName = getMarkName(currentMark)
        
    Dim newMarkName As String
    newMarkName = mendeleyApiClient().undoManualFormat(markName)
    
    Call fnRenameMark(currentMark, newMarkName)
    Call subSetMarkText(currentMark, INSERT_CITATION_TEXT)
    
    getUndoEditButton().Visible = False
    
    If USE_RIBBON Then
        Call recoverRibbonUi
        ribbonCitationButtonState = RIBBON_INSERT_CITATION
        Call ribbonUi.Invalidate
    End If
    
    Call refreshDocument
    GoTo EndOfSub
    
ErrorHandler:
   Call reportError
    
EndOfSub:
   uiDisabled = False
End Sub

Sub refresh()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True

    Call beginUndoTransaction("Refresh Mendeley Citations and Bibliography")
    Call refreshDocument
    Call endUndoTransaction
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub chooseCitationStyle()
    If isUiDisabled() Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Dim chosenStyle As String
    
    Call setMendeleyDocument(True)
        
    Call beginUndoTransaction("Change Citation Style")
    
    If launchMendeleyIfNecessary() = CONNECTION_CONNECTED Then
        chosenStyle = mendeleyApiClient().getCitationStyleFromDialog(getCitationStyleId())
        Call setCitationStyle(chosenStyle)
        Call refreshDocument
    End If
    Call refreshDocument
    
    Call endUndoTransaction
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

' Called from "More styles..." in the combobox
Sub OpenCitationsFromMendeley()
    If isUiDisabled() Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Dim chosenStyle As String
    
    Call setMendeleyDocument(True)
        
    If launchMendeleyIfNecessary() = CONNECTION_CONNECTED Then
        Dim ComboBox As CommandBarComboBox
        chosenStyle = mendeleyApiClient().getCitationStyleFromDialog(getCitationStyleId())
        
        Call updateCitationStylesComboBox
        
        Set ComboBox = getCitationStyleComboBox()
        ComboBox.Text = getStyleNameFromId(chosenStyle)
        ThisDocument.Saved = True
        
        Call setCitationStyle(chosenStyle)
        Call refreshDocument
    End If
    
    Call refreshDocument
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub afterSave()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Call refreshDocument
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub afterOpen()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Call refreshDocument
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub insertCitationButton()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Dim tipText As String
    tipText = "Tip: You can press Alt-M instead of clicking Insert Citation."
    
    Dim insertCitationButton As CommandBarButton
    Set insertCitationButton = getInsertCitationButton()
    
    If Not (insertCitationButton.Caption = TOOLBAR_INSERT_CITATION) Then
        tipText = "Tip: You can press ALT-M instead of clicking Edit Citation."
    End If
    
    Call privateInsertCitation(tipText)
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub insertCitation()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Call privateInsertCitation("")
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub exportAsBookmarks()
    Call exportCompatibleOpenOffice
End Sub

Sub exportCompatibleOpenOffice()
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    Dim saveAsDialog As FileDialog
    Set saveAsDialog = Application.FileDialog( _
        FileDialogType:=msoFileDialogSaveAs)
    
    ' -1 means the user wishes to save
    If saveAsDialog.Show = -1 Then
        Dim selectedItem As Variant
        For Each selectedItem In saveAsDialog.SelectedItems
        Dim selectedItemString As String
        selectedItemString = selectedItem
            Call privateExportCompatibleOpenOffice(selectedItemString)
            Exit For
        Next selectedItem
    End If
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
EndOfSub:
    uiDisabled = False
End Sub

Sub privateExportCompatibleOpenOffice(exportFilename As String)
    ZoteroUseBookmarks = True
    
    Dim documentState As DocumentStateType
    documentState = startUpdatingDocument(ActiveDocument)
    Dim marks

    marks = fnGetMarks(ZoteroUseBookmarks)
    
    Call ActiveDocument.SaveAs(filename:=exportFilename)
    Call finishUpdatingDocument(documentState)
End Sub

Sub exportWithoutMendeleyFields()
    ' Opens a Save As file dialog and saves the document
    ' converting the fields into plain text
    If isUiDisabled Then Exit Sub
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True
    
    If ActiveDocument.Path = "" Then
        MsgBox "Please save this document using the ""File->Save As..."" before exporting."
        GoTo EndOfSub
    End If
        
    Dim saveAsDialog As FileDialog
    Set saveAsDialog = Application.FileDialog( _
        FileDialogType:=msoFileDialogSaveAs)
    
    ' -1 means the user wishes to save
    If saveAsDialog.Show = -1 Then
        Dim Doc As Document
        
        Call ActiveDocument.SaveAs(ActiveDocument.FullName)

        'the next line copies the active document
        Call Application.Documents.Add(ActiveDocument.FullName)
        
        Set Doc = ActiveDocument

        Call removeMendeleyFields(Doc)

        Dim selectedItem
        For Each selectedItem In saveAsDialog.SelectedItems
             'the next line saves the copy to your location and name
            Call ActiveDocument.SaveAs(filename:=selectedItem)
        Next selectedItem
        
        ' next line closes the copy leaving you with the original document
        Call ActiveDocument.Close
    End If
    Set previouslySelectedField = Nothing
    
    GoTo EndOfSub
ErrorHandler:
    Call reportError
    
EndOfSub:
        uiDisabled = False
End Sub

Sub removeMendeleyFields(Doc As Document)
    Dim documentState As DocumentStateType
    documentState = startUpdatingDocument(Doc)
    ' Delete the custom properties clearing them, since deleting
    ' doesn't get saved
    
    ' Write directly to the document to avoid calling to
    ' Mendeley Desktop during this part as it would happen
    ' if setCitationStyle function is used.
    Call subSetProperty(MENDELEY_CITATION_STYLE, "")
    Call subSetProperty(MENDELEY_USER_ACCOUNT, "")
    Call subSetProperty(MENDELEY_DOCUMENT, "")
        
    ' Fields in document
    Dim fld As Field
    For Each fld In Doc.fields
        Call ConvertMendeleyFieldToText(fld)
    Next

    ' Fields in footnotes
    Dim i As Long
    i = 1
    While i <= Doc.Footnotes.count
        For Each fld In Doc.Footnotes(i).range.fields
            Call ConvertMendeleyFieldToText(fld)
        Next
        i = i + 1
    Wend

    ' Fields in TextBoxes
    i = 1
    While i <= Doc.Shapes.count
        If Doc.Shapes(i).Type <> msoPicture Then
            For Each fld In Doc.Shapes(i).TextFrame.TextRange.fields
                    Call ConvertMendeleyFieldToText(fld)
            Next
        End If
        i = i + 1
    Wend
    Call finishUpdatingDocument(documentState)
End Sub

' ----- end of top level functions -----
Function isUiDisabled() As Boolean
        If awaitingResponseFromMD Then
            Call MsgBox("Please finish selecting a citation from Mendeley Desktop first.", _
                vbOKOnly, MACRO_ALREADY_RUNNING)
        End If
        
        isUiDisabled = uiDisabled
End Function

Function mergeCitations()
    If isUiDisabled Then Exit Function
    If Not DEBUG_MODE Then
        On Error GoTo ErrorHandler
    End If
    uiDisabled = True

    ' Gathers the selected UUIDS to uuids and deletes the fields
    Dim selectedFields As fields
    Set selectedFields = Selection.range.fields
    
    If selectedFields.count < 2 Then
        MsgBox MERGE_CITATIONS_NOT_ENOUGH_CITATIONS
        GoTo EndOfSub
    End If
    
    If Not (launchMendeleyIfNecessary() = CONNECTION_CONNECTED) Then
        GoTo EndOfSub
    End If
    
    Dim markName As String
    Dim markUuids As String

    Dim mark As Field

    Dim previousStart As Long
    Dim previousEnd As Long
    previousStart = -1

    Dim mergeFieldCount As Long
    Dim mergeFieldCodes() As String

    For Each mark In selectedFields
        If previousStart > 0 Then
            Dim inbetween As range
            If mark.result.Start > previousStart Then
                Set inbetween = ActiveDocument.range(previousEnd, mark.result.Start)
            Else
                Set inbetween = ActiveDocument.range(mark.result.End, previousStart)
            End If
        
            If Len(inbetween.Text) > CITATION_ADJECENT_LIMIT Then
                MsgBox CITATIONS_NOT_ADJACENT
                GoTo EndOfSub
            End If
        End If
        
        previousStart = mark.result.Start
        previousEnd = mark.result.End
    
        markName = getMarkName(mark)
        If isMendeleyCitationField(markName) = False Then
            GoTo SkipField
        End If
        
        mergeFieldCount = mergeFieldCount + 1
        ReDim Preserve mergeFieldCodes(0 To mergeFieldCount)
        mergeFieldCodes(UBound(mergeFieldCodes)) = markName
SkipField:
    Next
    
    ' Creates a new field with the previously selected UUIDS
    Dim selectedRange As range
    
    Dim newFieldCodeText As String
    newFieldCodeText = mendeleyApiClient().mergeFields(mergeFieldCodes)
    
    Call beginUndoTransaction("Merge Mendeley Citations")
    
    Dim citeField As Field
    Set selectedRange = fnSelection()
    Set citeField = fnAddMark(selectedRange, newFieldCodeText)
    citeField.result.Text = MERGING_TEXT
    Call refreshDocument
    
    Call endUndoTransaction
    
    GoTo EndOfSub
    
ErrorHandler:
    Call reportError
    
EndOfSub:
    uiDisabled = False
End Function