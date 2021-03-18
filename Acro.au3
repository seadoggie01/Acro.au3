#include-once

; #INDEX# =======================================================================================================================
; Title .........: Adobe Acrobat Automation
; AutoIt Version : 3.3.14.5
; UDF Version ...: 1.0.2.0
; Language ......: English
; Description ...: A collection of functions for accessing and manipulating PDFs though Adobe Acrobat Pro
; Author(s) .....: Seadoggie01
; Modified.......: 20210318 (YYYYMMDD)
; Contributors ..:
; Resources .....: Acrobat and PDF Library API Reference - (More info than you could ever want on Acrobat)
;                : 		https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/acrobat_pdfl_api_reference.pdf
;                : Inter-Application Communication and API Reference - (A manageable but annoyingly laid out API refernce.)
;                : 		https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/iac_api_reference.pdf
;                : Developer Guide - (Specifically see page 30 for an overview of how the main objects work)
;                : 		https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/iac_developer_guide.pdf
;                : JavaScript API Reference - (This is where stuff gets interesting)
;                : 		https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/js_api_reference.pdf
; ===============================================================================================================================

#Region #GLOBAL VARIABLES# ======================================================================================================

; PDSaveFlags - Flags for the PDDocSave saveFlags parameter. All undefined flags should be set to zero.
; https://help.adobe.com/pdfl_sdk/15/PDFL_SDK_HTMLHelp/PDFL_SDK_HTMLHelp/API_References/PDFL_API_Reference/PD_Layer/General.html#PDSaveFlags1439
Global Enum $PDSaveIncremental = 		0x00, _ ; Save only those portions of the document that have changed. This is provided only as the opposite of PDSaveFull, since there is no bit value of 0.
			$PDSaveFull = 				0x01, _ ; Save the entire document. Plug-ins that set PDSaveFull are also encouraged to set PDSaveCollectGarbage.
			$PDSaveCopy = 				0x02, _ ; Save a copy of the document (the PDDoc continues to use the old file). This flag is ignored if PDSaveFull is off.
			$PDSaveLinearized = 		0x04, _ ; Write the PDF file in a format that is optimized for page-served remote (network) access (linearized). This flag is ignored if PDSaveFull is off. Prior to Acrobat 7, linearizing a file caused Cos objects to be invalidated, which required that some plug-ins use notifications to release and re-acquire objects. Acrobat 7 onwards, Cos objects are no longer invalidated after a linearized save.
			$PDSaveWithPSHeader = 		0x08, _ ; (Obsolete. In effect, it is always off). Write a PostScript header as part of the saved file.
			$PDSaveBinaryOK = 			0x10, _ ; (Obsolete. In effect, it is always on). It is okay to store binary data in the PDF file.
			$PDSaveCollectGarbage = 	0x20, _ ; Remove unreferenced objects, often reducing file size. Plug-ins are encouraged to use this flag. This flag is ignored if PDSaveFull is off.
			$PDSaveForceIncremental = 	0x40, _ ; Perform an incremental save even if the save is to a different file or the document's version number has changed.
			$PDSaveKeepModDate = 		0x80, _ ; Do not update ModDate in InfoDict.
			$PDSaveLeaveOpen = 			0x100   ; Leave the file open after the save (do not Close the file)

; PDPageMode Enumeration
; https://help.adobe.com/pdfl_sdk/15/PDFL_SDK_HTMLHelp/PDFL_SDK_HTMLHelp/API_References/PDFL_API_Reference/PD_Layer/PDPage.html#PDPageMode1940
Global Enum $PDDontCare = 0, _	; Leaves the view mode as is.
			$PDUseNone, _		; Displays the document, but displays neither thumbnails nor bookmarks.
			$PDUseThumbs, _		; Displays the document plus thumbnails.
			$PDUseBookmarks, _	; Displays the document plus bookmarks.
			$PDFullScreen, _	; Displays the document in full-screen viewing mode. This is equivalent to AVAppBeginFullScreen().
			$PDContents, _ 		; [Undefined]
			$PDUseOC, _			; [Undefined]
			$PDUseAttachments	; Displays the document plus attachments.

; PDInsertFlags
; https://help.adobe.com/pdfl_sdk/15/PDFL_SDK_HTMLHelp/PDFL_SDK_HTMLHelp/API_References/PDFL_API_Reference/PD_Layer/PDDoc.html#PDInsertFlags2079
Global Enum $PDInsertBookmarks = 	0x0001, _ ; Insert bookmarks as well as pages.
			$PDInsertAll = 			0x1000, _ ; Insert all Catalog and Info dictionary values as well as pages.
			$PDInsertThreads = 		0x0002    ; Insert articles as well.

#EndRegion GLOBAL VARIABLES =====================================================================================================

; #CURRENT# =====================================================================================================================
; _Acro_AppCreate
; _Acro_AppExit
; _Acro_AppShow
; _Acro_DocAppend
; _Acro_DocBookmarkAdd
; _Acro_DocBookmarkAddEx
; _Acro_DocBookmarkGet
; _Acro_DocBookmarkProperties
; _Acro_DocBookmarkRemove
; _Acro_DocClose
; _Acro_DocDisplay
; _Acro_DocJSObject
; _Acro_DocNew
; _Acro_DocOpen
; _Acro_DocSave
; _Acro_PageCount
; _Acro_PageDelete
; _Acro_PageInsert
; _Acro_PageGet
; _Acro_PageMove
; _Acro_PageRotate
; _Acro_PageSize
; _Acro_PageGetText
; _Acro_PageViewMode
; ===============================================================================================================================

; #INTERNAL_USE_ONLY#============================================================================================================
; __Acro_Rect
; __Acro_FileGetFolder
; __Acro_ObjCheck
; __Acro_ErrHnd
; __Acro_SecPolicyGet -- Beta, unused
; __Acro_PrivContext -- Beta, unused
; __Acro_SecHandler -- Beta, unused
; ===============================================================================================================================

#ToDo: (Feature) Add ability to password protect document, but it doesn't seem possible without elevated permission or elevated JavaScript file.

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_AppCreate
; Description ...: Creates an Acrobat Application
; Syntax ........: _Acro_AppCreate([$bVisible = True])
; Parameters ....: $bVisible            - [optional] a boolean value. Default is True.
; Return values .: Success - an Acrobat Application object
;                  Failure - Sets @error to 1
;                  |1 - Returns False, Unable to create applicaiton
;                  |2 - Returns application, unable to set visibility, sets @extended to _Acro_AppShow's error
; Author ........: Seadoggie01
; Modified ......: August 19, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_AppCreate($bVisible = True)

	Local $iExtended = 0
	Local $oAcroApp = ObjGet("", "AcroExch.App")
	If Not @error Then
		$iExtended = 1
	Else
		$oAcroApp = ObjCreate("AcroExch.App")
		If @error Then Return SetError(1, 0, False)
	EndIf

	If Not IsKeyword($bVisible) Then _Acro_AppShow($oAcroApp, $bVisible)
	If @error Then Return SetError(2, @error, False)

	Return SetExtended($iExtended, $oAcroApp)

EndFunc   ;==>_Acro_AppCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_AppExit
; Description ...: Closes an Acrobat Application
; Syntax ........: _Acro_AppExit($oApp)
; Parameters ....: $oApp                - an object.
; Return values .: Success - The page mode of the PDDocument
;                  Failure - False and sets @error:
;                  |1 - $oApp isn't an Acrobat Application
;                  |2 - Unable to close documents, sets @extended if COM error
;                  |3 - Unable to hide application, sets @extended if COM error
;                  |4 - Unable to completely exit, sets @extended if COM error
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......:
; Related .......: This will fail if the application is locked. This function will mostly likely return @error = 4 with no COM
;                  + error due to Acrobat taking its time closing. This seems to be normal behavior, but the documentation says
;                  + otherwise.
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_AppExit($oApp)

	If Not __Acro_ObjCheck($oApp, "App") Then Return SetError(1, 0, False)

	Local $oError = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oError

	Local $vRet = $oApp.CloseAllDocs
	If @error Then Return SetError(2, @error, False)
	If Not ($vRet = -1) Then Return SetError(2, 0, False)

	; If the application is showing, it can't be closed
	$vRet = $oApp.Hide()
	If @error Then Return SetError(3, @error, False)
	If Not ($vRet = -1) Then Return SetError(3, 0, False)

	$vRet = $oApp.Exit
	If @error Then Return SetError(4, @error, False)
	If Not ($vRet = -1) Then Return SetError(4, 0, False)
	Return True

EndFunc   ;==>_Acro_AppExit

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_AppShow
; Description ...: Shows and hides the Acrobat Application.
; Syntax ........: _Acro_AppShow($oApp[, $bShow = False])
; Parameters ....: $oApp                - an object.
;                  $bShow               - [optional] a boolean value. Default is True.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - $oApp isn't an Acrobat Application
;                  |2 - Unable to toggle visibility, sets @extended if COM error
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......: When the application is visible, the user can control it and it will stay open if all documents are closed.
;                : When the application is hidden, the application will close after all objects are closed.
;                : This means that if you hide an application with open documents, it will continue running even after the
;                : script is closed. You should explicitly call _Acro_DocClose for each document opened to avoid this.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_AppShow($oApp, $bShow = Default)

	If Not __Acro_ObjCheck($oApp, "App") Then Return SetError(1, 0, False)
	If IsKeyword($bShow) Then $bShow = True

	Local $oError = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oError

	Local $vRet
	If $bShow Then
		$vRet = $oApp.Show()
	Else
		$vRet = $oApp.Hide()
	EndIf
	If @error Then Return SetError(2, @error, False)
	If Not ($vRet = -1) Then Return SetError(2, 0, False)
	Return True

EndFunc   ;==>_Acro_AppShow

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocAppend
; Description ...: Appends an entire document to the end of another
; Syntax ........: _Acro_DocAppend($oPdDoc, $vAppend)
; Parameters ....: $oPdDoc              - a PD Document to append to.
;                  $vAppend             - a PD Document to append or the full path to a PDF to open and append.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PD Document
;                  |2 - $sFullPath doesn't exist
;                  |3 - Can't open document, @extended set to _Acro_DocOpen's error
;                  |4 - Can't get page count of original, @extended set to _Acro_PageCount's error
;                  |5 - Can't get page count of appending document, @extended set to _Acro_PageCount's error
;                  |6 - Can't insert pages, @extended set to _Acro_PageInsert's error
;                  |7 - Can't close append document, @extended set to _Acro_DocClose's error
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocAppend($oPdDoc, $vAppend)

	Local $oAppendPdDoc, $bCloseAppend

	If IsString($vAppend) Then
		If Not FileExists($vAppend) Then Return SetError(2, 0, False)

		$oAppendPdDoc = _Acro_DocOpen($vAppend)
		If @error Then Return SetError(3, @error, False)

		; Close the document we opened
		$bCloseAppend = True
	Else

		If Not __Acro_ObjCheck($oPdDoc, "PDDoc") Then Return SetError(1, 0, False)
		$oAppendPdDoc = $vAppend
	EndIf

	Local $iPdPages = _Acro_PageCount($oPdDoc)
	If @error Then Return SetError(4, @error, False)

	Local $iAppendPages = _Acro_PageCount($oAppendPdDoc)
	If @error Then Return SetError(5, @error, False)

	_Acro_PageInsert($oPdDoc, $iPdPages - 1, $oAppendPdDoc, 0, $iAppendPages - 1, True)
	If @error Then Return SetError(6, @error, False)

	If $bCloseAppend Then _Acro_DocClose($oAppendPdDoc)
	If @error Then Return SetError(7, @error, False)

	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocBookmarkAdd
; Description ...: Adds a bookmark to a PdDoc that sets the current view to a specified page.
; Syntax ........: _Acro_DocBookmarkAdd($oPdDoc, $sName, $iPage[, $iIndex = 0[, $oParent = Default]])
; Parameters ....: $oPdDoc              - an PDDoc object.
;                  $sName               - the text to display.
;                  $iPage               - the 0-based destination page.
;                  $iIndex              - [optional] the 0-based index of bookmarks to place the bookmark at. Default is 0.
;                  $oParent             - [optional] the parent bookmark. Default creates a top-level bookmark.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc
;                  |2 - _Acro_DocJSObject error, sets @extended to the @error.
;                  |3 - Error accessing root bookmark, sets @extended to the COM Error
;                  |4 - Error creating bookmark, sets @extended to the COM Error
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......: Setting a destination page is achieved by executing JavaScript. If the pages are re-arranged later, the
;                  + JavaScript will NOT change and the destination will be incorrect. Adding a bookmark by page without JavaScript
;                  + is not supported by Acrobat through the Interapplication API or JavaScript.
; Related .......: _Acro_DocBookmarkGet
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocBookmarkAdd($oPdDoc, $sName, $iPage, $iIndex = 0, $oParent = Default)

	Local $vRet = _Acro_DocBookmarkAddEx($oPdDoc, $sName, "this.pageNum = " & $iPage, $iIndex, $oParent)
	Return SetError(@error, @extended, $vRet)

EndFunc   ;==>_Acro_DocBookmarkAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocBookmarkAddEx
; Description ...: Adds a bookmark to a PdDoc that executes JavaScript when clicked.
; Syntax ........: _Acro_DocBookmarkAddEx($oPdDoc, $sName, $sExecutableJavaScript[, $iIndex = 0[, $oParent = Default]])
; Parameters ....: $oPdDoc              - a PDDocument object.
;                  $sName               - the text shown in the bookmark.
;                  $sExecutableJavaScript- a string containing the JavaScript to be executed each time the bookmark is clicked.
;                  $iIndex              - [optional] the index to insert the new bookmark under $oParent. Default is 0.
;                  $oParent             - [optional] a bookmark object. Default is Default.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc
;                  |2 - _Acro_DocJSObject error, sets @extended to the @error.
;                  |3 - Error accessing root bookmark, sets @extended to the COM Error
;                  |4 - Error creating bookmark, sets @extended to the COM Error
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......: $sExecutableJavaScript is not checked for errors. Note that bookmarks are not in a priveledged context.
; Related .......: _Acro_DocBookmarkAdd
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocBookmarkAddEx($oPdDoc, $sName, $sExecutableJavaScript, $iIndex = 0, $oParent = Default)

	If Not __Acro_ObjCheck($oPdDoc, "PDDoc") Then Return SetError(1, 0, False)
	Local $oJS = _Acro_DocJSObject($oPdDoc)
	If @error Then Return SetError(2, @error, False)

	Local $oError = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oError

	If IsKeyword($oParent) Then $oParent = $oJS.bookmarkRoot
	If @error Then Return SetError(3, @error, False)
	$oParent.createChild($sName, $sExecutableJavaScript, $iIndex)
	If @error Then Return SetError(4, @error, False)
	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocBookmarkGet
; Description ...: Gets a bookmark or an array of bookmarks from a document
; Syntax ........: _Acro_DocBookmarkGet($oPdDoc[, $vBookmark = Default[, $oParent = Default]]])
; Parameters ....: $oPdDoc              - an object.
;                  $vBookmark           - [optional] A number gets the child bookmark at the index of $oParent.
;                                               A string searches $oParent for a bookmark with that name.
;                                               -1 returns a 0-based 1D array of all of $oParent's children.
;                                               Default is the root bookmark ($oParent is ignored).
;                  $oParent             - [optional] a Bookmark object returned from this function. Default is the root bookmark.
; Return values .: Success - the target bookmark object(s)
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc
;                  |2 - _Acro_DocJSObject returned an error: @extended set to it's @error.
;                  |3 - $oParent doesn't contain a (direct) child bookmark with $vBookmark's name
;                  |4 - $vBookmark contains an index greater than the number of children in $oParent
;                  |5 - Can't access root bookmark, sets @extended to COM Error
;                  |6 - $oParent doesn't have any child bookmarks, sets @extended to COM Error
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......: Get the root bookmark:			_Acro_DocBookmarkGet($oPdDoc)
;				   Get a top-level named bookmark: 	_Acro_DocBookmarkGet($oPdDoc, "Name")
;                  Get the child of a bookmark: 	_Acro_DocBookmarkGet($oPdDoc, "Name", $oBookmark)
;                  Get all children of a bookmark: 	_Acro_DocBookmarkGet($oPdDoc, -1, $oBookmark) (returns an array)
; Related .......: _Acro_DocBookmarkAdd
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocBookmarkGet($oPdDoc, $vBookmark = Default, $oParent = Default)

	If Not __Acro_ObjCheck($oPdDoc, "PDDoc") Then Return SetError(1, 0, False)

	Local $oJS = _Acro_DocJSObject($oPdDoc)
	If @error Then Return SetError(2, @error, False)

	Local $oError = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oError

	If IsKeyword($oParent) Then $oParent = $oJS.bookmarkRoot
	If @error Then Return SetError(5, @error, False)

	Local $aoChildren

	If IsKeyword($vBookmark) Then
		$vBookmark = $oJS.bookmarkRoot
		If @error Then Return SetError(5, @error, False)
	ElseIf IsString($vBookmark) Then

		$aoChildren = $oParent.children
		If @error Then Return SetError(6, @error, False)
		Local $bFound = False
		For $i=0 To UBound($aoChildren) - 1
			If $aoChildren[$i].Name = $vBookmark Then
				$vBookmark = $aoChildren[$i]
				$bFound = True
				ExitLoop
			EndIf
		Next

		If Not $bFound Then Return SetError(3, 0, False)
	ElseIf -1 = $vBookmark Then
		$aoChildren = $oParent.children
		If @error Then Return SetError(6, @error, False)
		Return $aoChildren
	ElseIf IsNumber($vBookmark) Then
		$aoChildren = $oParent.children
		If @error Then Return SetError(6, @error, False)
		If $vBookmark >= UBound($aoChildren) Then Return SetError(4, 0, False)
		$vBookmark = $aoChildren[$vBookmark]
	EndIf

	Return $vBookmark

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocBookmarkProperties
; Description ...: Sets the various properties of a bookmark and returns an array of properties
; Syntax ........: _Acro_DocBookmarkProperties($oBookmark[, $sName = Default[, $aColor = Default[, $bOpen = Default[,
;                  $iStyle = Default[, $sExecutableJavaScript = Default]]]]])
; Parameters ....: $oBookmark           - an object.
;                  $sName               - [optional] a string value. Default is Default.
;                  $aColor              - [optional] an array, see remarks. Default is Default.
;                  $bOpen               - [optional] True - expended, False - collapsed. Default is Default.
;                  $iStyle              - [optional] an integer, see remarks. Default is Default.
;                  $sExecutableJavaScript- [optional] a string value. Default is Default.
; Return values .: Success - a 0-based 1D array of properties, see remarks for details
;                  Failure - False and sets @error:
;                  |1 - $oBookmark isn't an object
;                  |2 - Error setting name, sets @extended to COM error
;                  |3 - Error setting color, sets @extended to COM error
;                  |4 - Error setting open, sets @extended to COM error
;                  |5 - Error setting style, sets @extended to COM error
;                  |6 - Error setting action (JavaScript), sets @extended to COM error
; Author ........: Seadoggie01
; Modified ......: November 23, 2020
; Remarks .......: Use like: _Acro_DocBookmarkProperties($oBookmark) to get an array of values
;
;                  Default means values won't be changed.
;                  $aColor is ["T"] for transparent, ["G", x] for grayscale,
;                  +["RGB", x, x, x] for a RGB value, or ["CMYK", x, x, x, x] for a CMYK value.
;                  +N.B. The JavaScript color object isn't available through the COM.
;                  $iStyle: 0 is normal, 1 is italic, 2 is bold, and 3 is bold-italic
;                  $sExecutableJavaScript is NOT checked for valid JavaScript
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocBookmarkProperties($oBookmark, $sName = Default, $aColor = Default, $bOpen = Default, $iStyle = Default, $sExecutableJavaScript = Default)

	If Not IsObj($oBookmark) Then Return SetError(1, 0, False)
	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr
	If Not IsKeyword($sName) Then $oBookmark.Name = $sName
	If @error Then Return SetError(2, @error, False)
	If Not IsKeyword($aColor) Then $oBookmark.Color = $aColor
	If @error Then Return SetError(3, @error, False)
	If Not IsKeyword($bOpen) Then $oBookmark.Open = $bOpen
	If @error Then Return SetError(4, @error, False)
	If Not IsKeyword($iStyle) Then $oBookmark.Style = $iStyle
	If @error Then Return SetError(5, @error, False)
	If Not IsKeyword($sExecutableJavaScript) Then $oBookmark.setAction($sExecutableJavaScript)
	If @error Then Return SetError(6, @error, False)

	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocBookmarkRemove
; Description ...: Removes a bookmark from a document.
; Syntax ........: _Acro_DocBookmarkRemove($oPdDoc, $vBookmark[, $oParent = Default])
; Parameters ....: $oPdDoc              - an object.
;                  $vBookmark           - a bookmark object or a bookmark name.
;                  $oParent             - [optional] a bookmark object. Default is the root bookmark.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc object
;                  |2 - $oParent isn't an object
;                  |3 - Unable to remove bookmark, sets @extended to COM Error
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocBookmarkRemove($oPdDoc, $vBookmark, $oParent = Default)

	If Not __Acro_ObjCheck($oPdDoc, "PDDoc") Then Return SetError(1, 0, False)
	If IsKeyword($oParent) Then $oParent = _Acro_DocBookmarkGet($oPdDoc)
	If Not IsObj($oParent) Then Return SetError(2, 0, False)

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd())
	#forceref $oErr
	$vBookmark.Remove()
	If @error Then Return SetError(3, @error, False)

	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocCreate
; Description ...: Creates an empty document
; Syntax ........: _Acro_DocCreate([$oPdDoc = Default])
; Parameters ....: $oPdDoc              - [optional] a PD Document object. Default creates one.
; Return values .: Success - the PD Document object
;                  Failure - False and sets @error:
;                  |1 - Error getting new document, sets @extended to _Acro_DocNew's error
;                  |2 - Error creating document, sets @extended to COM error
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocCreate($oPdDoc = Default)

	If IsKeyword($oPdDoc) Then
		$oPdDoc = _Acro_DocNew()
		If @error Then Return SetError(1, @error, False)
	EndIf

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd())
	#forceref $oErr

	$oPdDoc.Create()
	If @error Then Return SetError(2, @error, False)

	Return $oPdDoc

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocClose
; Description ...: Closes an AV or PD Document
; Syntax ........: _Acro_DocClose($oDoc)
; Parameters ....: $oDoc                - a document object.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - $oDoc isn't an AV/PD Doc object
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......: Every document opened should have a corresponding _Acro_DocClose call. Acrobat doesn't manage this.
; Related .......: _Acro_DocOpen
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocClose($oDoc)

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr
	If __Acro_ObjCheck($oDoc, "PDDoc") Then
		Return $oDoc.Close() = -1
	ElseIf __Acro_ObjCheck($oDoc, "AVDoc") Then
		; Always returns -1, no reason to check
		$oDoc.Close(1)
		Return True
	Else
		Return SetError(1, 0, False)
	EndIf

EndFunc   ;==>_Acro_DocClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocDisplay
; Description ...: Opens the PDDocument in a Window with the requested title or a default one.
; Syntax ........: _Acro_DocDisplay($oPdDoc[, $sTitle = Default])
; Parameters ....: $oPdDoc              - an object.
;                  $sTitle              - [optional] a string value. Default is determined by Acrobat.
;                  $iPage               - [optional] an integer. Default is 0 (the first page).
; Return values .: Success - An AVDoc object
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc object
;                  |2 - AVDoc not created, sets @extended to COM error
;                  |3 - Unable to get Page View, sets @extended to COM error
;                  |4 - Unable to go to page, sets @extneded to COM error
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......: If a document is displayed and the app is then hidden, the document will continue to exist in the hidden app
;                : Call _Acro_DocClose for each document opened to release it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocDisplay($oPdDoc, $sTitle = Default, $iPage = Default)

	If IsKeyword($sTitle) Then $sTitle = ""

	If Not __Acro_ObjCheck($oPdDoc, "PDDoc") Then Return SetError(1, 0, False)
	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr
	Local $oAvDoc = $oPdDoc.OpenAVDoc($sTitle)
	If @error Then Return SetError(2, @error, False)
	If Not IsKeyword($iPage) Then
		Local $oPageView = $oAvDoc.GetAVPageView
		If @error Then Return SetError(3, @error, False)
		Local $vRet = $oPageView.GoTo($iPage)
		If @error Then Return SetError(4, @error, False)
		If Not ($vRet = -1) Then Return SetError(4, 0, False)
	EndIf

	Return $oAvDoc

EndFunc   ;==>_Acro_DocDisplay

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocJSObject
; Description ...: Exposes the JavaScript Object of a PdDoc
; Syntax ........: _Acro_DocJSObject($oPdDoc)
; Parameters ....: $oPdDoc              - an object.
; Return values .: Success - a JavaScript Object
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc
;                  |2 - Unable to get JavaScript object, sets @extended if COM Error
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocJSObject($oPdDoc)

	If Not (ObjName($oPdDoc) = "CAcroPDDoc") Then Return SetError(1, 0, False)

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr
	Local $oJS = $oPdDoc.GetJSObject
	If @error Then Return SetError(2, @error, False)
	If Not IsObj($oJS) Then Return SetError(2, 0, False)

	Return $oJS

EndFunc   ;==>_Acro_DocJSObject

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocNew
; Description ...: Creates an empty reference to a new PD or AV document object
; Syntax ........: _Acro_DocNew([$bAVDoc = False])
; Parameters ....: $bAVDoc              - [optional] a boolean value. Default is False.
; Return values .: Success - a PD or AV document object
;                  Failure - False and sets @error to ObjCreate's error
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......: Use _Acro_DocCreate or _Acro_DocOpen to start using the object
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocNew($bAVDoc = Default)

	Local $sClassName = "AcroExch.PdDoc"

	If $bAVDoc Then $sClassName = "AcroExch.AvDoc"

	Local $oDoc = ObjCreate($sClassName)
	If @error Then Return SetError(1, @error, False)

	Return $oDoc

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocOpen
; Description ...: Creates a PDDocument from a file.
; Syntax ........: _Acro_DocOpen($sFullPath[, $oDoc = Default])
; Parameters ....: $sFullPath           - the path to the file to open.
;                  $oDoc                - [optional] an existing AV or PD document object. Default uses a new PDDocument.
; Return values .: Success - a new PDDocument, or $oDoc if passed
;                  Failure - False and sets @error:
;                  |1 - $sFullPath doesn't exist
;                  |2 - Document not opened. Sets @error if COM Error.
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......:
; Related .......: _Acro_DocClose
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocOpen($sFullPath, $oDoc = Default)

	If Not FileExists($sFullPath) Then Return SetError(1, 0, False)

	If IsKeyword($oDoc) Or Not IsObj($oDoc) Then
		$oDoc = ObjCreate("AcroExch.PDDoc")
		If @error Then Return SetError(1, @error, False)
	EndIf

	Local $vRet, $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd())
	#forceref $oErr

	If StringInStr(ObjName($oDoc), "av") Then
		$vRet = $oDoc.Open($sFullPath, "")
	Else
		$vRet = $oDoc.Open($sFullPath)
	EndIf

	If @error Then Return SetError(2, @error, False)
	If Not ($vRet = -1) Then Return SetError(2, 0, False)

	Return $oDoc

EndFunc   ;==>_Acro_DocOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_DocSave
; Description ...: Saves a PDDocument and optionally creates a copy.
; Syntax ........: _Acro_DocSave($oPdDoc[, $sFullPath = Default[, $iSaveMode = Default]])
; Parameters ....: $oPdDoc              - a PD or AV Document object.
;                  $sFullPath           - [optional] the full path to save as. Default doesn't change the save location.
;                  $iSaveMode           - [optional] a PDSaveOptions value. Default is $PDSaveFull + $PDSaveLinearized.
;                  $bCreatePath         - [optional] a boolean value. Default is False.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc object
;                  |2 - Document not saved
;                  |3 - Document failed to save, sets @extended to COM Error
;                  |4 - Error getting PDDoc from AVDoc
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......: The default save options may fail if you've edited the file since opening it. Use $PDSaveFull if you aren't
;                  + sure. Specifying default save options and a new path after editing may cause the new path to be ignored!
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_DocSave($oDoc, $sFullPath = Default, $iSaveMode = Default, $bCreatePath = Default)

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr

	If Not __Acro_ObjCheck($oDoc, "PDDoc") Then
		If Not __Acro_ObjCheck($oDoc, "AVDoc") Then Return SetError(1, 0, False)
		$oDoc = $oDoc.GetPDDoc
		If @error Then Return SetError(3, @error, False)
	EndIf

	If IsKeyword($sFullPath) Then $sFullPath = ""
	If IsKeyword($iSaveMode) Then $iSaveMode = BitAND($PDSaveFull, $PDSaveLinearized)
	If IsKeyword($bCreatePath) Then $bCreatePath = False

	If $bCreatePath Then DirCreate(__Acro_FileGetFolder($sFullPath))

	Local $vRet = $oDoc.Save($iSaveMode, $sFullPath)
	If @error Then Return SetError(3, @error, False)
	If $vRet = -1 Then Return True
	; Document failed to save
	Return SetError(2, 0, False)

EndFunc   ;==>_Acro_DocSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_PageCount
; Description ...: Gets the number of pages in a PDDocument
; Syntax ........: _Acro_PageCount($oPdDoc)
; Parameters ....: $oPdDoc              - an object.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc object
;                  |2 - Can't get number of pages, sets @extended if COM Error
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_PageCount($oPdDoc)

	If Not __Acro_ObjCheck($oPdDoc, "PDDoc") Then Return SetError(1, 0, False)
	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr
	Local $iPages = $oPdDoc.GetNumPages()
	If @error Then Return SetError(2, @error, False)
	If $iPages = -1 Then Return SetError(2, 0, False)
	Return $iPages

EndFunc   ;==>_Acro_PageCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_PageDelete
; Description ...: Deletes a list of pages from a PDF document
; Syntax ........: _Acro_PageDelete($oPdDoc, $sPages)
; Parameters ....: $oPdDoc              - an object.
;                  $sPages              - a string value.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc object
;                  |2 - Error deleting pages. @extended set to 0-based index of failed page/s to delete.
;                  |3 - Error deleting pages, @extended set to COM Error
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......: Pass the pages in ascending order to avoid deleting the wrong pages: _Acro_PageDelete($oPdDoc, "0,3,5-7")
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_PageDelete($oPdDoc, $sPages)

	If Not __Acro_ObjCheck($oPdDoc, "PDDoc") Then Return SetError(1, 0, False)
	Local $aPages = StringSplit($sPages, ",", 2)
	Local $iPos, $vRet, $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr
	For $i = UBound($aPages) - 1 To 0 Step -1
		$iPos = StringInStr($aPages[$i], "-")
		If $iPos Then
			$vRet = $oPdDoc.DeletePages(StringLeft($aPages[$i], $iPos - 1), StringRight($aPages[$i], StringLen($aPages[$i]) - $iPos - 1))
		Else
			; Delete a single page
			$vRet = $oPdDoc.DeletePages($aPages[$i], $aPages[$i])
		EndIf
		If @error Then Return SetError(3, Hex(@error), False)
		If Not ($vRet = -1) Then Return SetError(2, UBound($aPages) - $i, False)
	Next
	Return True

EndFunc   ;==>_Acro_PageDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_PageInsert
; Description ...: Inserts pages into PDF from another PDF
; Syntax ........: _Acro_PageInsert($oDestPdDoc, $iInsertPage, $oSourcePdDoc[, $iStart = 1[, $iCount = 1[, $bBookmark = True]]])
; Parameters ....: $oDestPdDoc          - a PD Doc to insert pages into.
;                  $iInsertPage         - the page to insert after (0-based). Use -1 to insert before first page. Default is 0.
;                  $oSourcePdDoc        - a PD Doc to copy pages from.
;                  $iStart              - [optional] the first page (0-based) to add from $oSourcePdDoc. Default is 0.
;                  $iCount              - [optional] number of pages to insert from $oSourcePdDoc. Default is 1.
;                  $bBookmark           - [optional] Not working! boolean to copy existing bookmarks. Default is True.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - $oDestPdDoc isn't a PDDoc
;                  |2 - $oSourcePdDoc isn't a PDDoc
;                  |3 - Failed to insert pages to $iInsertPage, sets @extended if COM Error
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......: Despite the documentation, there doesn't appear to be a way to insert pages without adding bookmarks. Per the
;                  + docs, sending 0 shouldn't copy bookmarks, and any positive number will. Bookmarks are always copied.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_PageInsert($oDestPdDoc, $iInsertPage, $oSourcePdDoc, $iStart = Default, $iCount = Default, $bBookmark = Default)

	If Not __Acro_ObjCheck($oDestPdDoc, "PDDoc") Then Return SetError(1, 0, False)
	If Not __Acro_ObjCheck($oSourcePdDoc, "PDDoc") Then Return SetError(2, 0, False)

	If IsKeyword($iInsertPage) Then $iInsertPage = 0
	If IsKeyword($iStart) Then $iStart = 0
	If IsKeyword($iCount) Then $iCount = 1
	If IsKeyword($bBookmark) Then $bBookmark = 0
	; I attempted to use a value I pulled from the SDK for a while, but I can't get this working

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr
	Local $vRet = $oDestPdDoc.InsertPages($iInsertPage, $oSourcePdDoc, $iStart, $iCount, 	$bBookmark)
	If @error Then Return SetError(3, @error, False)
	If Not ($vRet = -1) Then Return SetError(3, 0, False) ; ==> See Remarks about why this is commented
	Return True

EndFunc   ;==>_Acro_PageInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_PageGet
; Description ...: Get a PDPage object from a document
; Syntax ........: _Acro_PageGet($oPdDoc, $iPage)
; Parameters ....: $oPdDoc              - an object.
;                  $iPage               - an integer value.
; Return values .: Success - a PDPage object
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc
;                  |2 - Unable to get page
;                  |3 - Unable to get page and sets COM Error
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_PageGet($oPdDoc, $iPage)

	If Not __Acro_ObjCheck($oPdDoc, "PDDoc") Then Return SetError(1, 0, False)
	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr
	Local $oPdPage = $oPdDoc.AcquirePage($iPage)
	If @error Then Return SetError(3, @error, False)
	If Not __Acro_ObjCheck($oPdPage, "PDPage") Then Return SetError(2, 0, False)
	Return $oPdPage

EndFunc   ;==>_Acro_PageGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_PageMove
; Description ...: Moves a page inside of a PDDocument
; Syntax ........: _Acro_PageMove($oPdDoc, $iPage, $iNewPage)
; Parameters ....: $oPdDoc              - an object.
;                  $iPage               - an integer value.
;                  $iNewPage            - an integer value.
; Return values .: Success - True
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc
;                  |2 - Unable to move page
;                  |3 - Unable to move page and sets @extended to COM Error
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_PageMove($oPdDoc, $iPage, $iNewPage)

	If Not __Acro_ObjCheck($oPdDoc, "PDDoc") Then Return SetError(1, 0, False)
	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr
	Local $vRet = $oPdDoc.MovePage($iNewPage, $iPage)
	If @error Then Return SetError(3, @error, False)
	If Not ($vRet = -1) Then Return SetError(2, 0, False)
	Return True

EndFunc   ;==>_Acro_PageMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_PageRotate
; Description ...: Gets or sets the rotation of a page
; Syntax ........: _Acro_PageRotate($oPdDoc, $vPage[, $iDegrees = Default])
; Parameters ....: $oPdDoc              - a PDDoc object.
;                  $vPage               - a PDPage object or the page number.
;                  $iDegrees            - [optional] an integer value. Default is Default.
; Return values .: Success - the rotation of $vPage
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc
;                  |2 - Unable to get page
;                  |3 - $iDegrees is invalid
;                  |4 - Unable to rotate page
;                  |5 - Error getting page rotation, sets @extended to COM Error
;                  |6 - Error setting page rotation, sets @extended to COM Error
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_PageRotate($oPdDoc, $vPage, $iDegrees = Default)

	Local $oPage
	; If the page object was passed
	If __Acro_ObjCheck($vPage, "PDPage") Then
		$oPage = $vPage
	Else
		If Not __Acro_ObjCheck($vPage, "PDDoc") Then Return SetError(1, 0, False)
		$oPage = _Acro_PageGet($oPdDoc, $vPage)
		If @error Then Return SetError(2, @error, False)
	EndIf

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr

	If Not IsKeyword($iDegrees) Then
		Switch Number($iDegrees)
			Case 0, 90, 180, 270
				; Do nothing, good value
			Case Else
				Return SetError(3, 0, False)
		EndSwitch
		If $oPage.SetRotate($iDegrees) = -1 Then Return SetError(4, 0, False)
		If @error Then Return SetError(6, @error, False)
	EndIf

	$iDegrees = $oPage.GetRotate()
	If @error Then Return SetError(5, @error, False)
	Return $iDegrees

EndFunc   ;==>_Acro_PageRotate

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_PageSize
; Description ...: Gets the height and width of a page
; Syntax ........: _Acro_PageSize($oPdDoc, $iPage)
; Parameters ....: $oPdDoc              - an object.
;                  $iPage               - an integer value.
; Return values .: Success - a 0-based 1D array of [width, height]
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc
;                  |2 - Unable to get page. @extended = _Acro_PageGet's error
;                  |3 - Unable to get size of page
;                  |4 - Unable to obtain height and width, @extended set to COM Error
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_PageSize($oPdDoc, $iPage)

	If Not (ObjName($oPdDoc) = "CAcroPDDoc") Then Return SetError(1, 0, False)
	Local $oPdPage = _Acro_PageGet($oPdDoc, $iPage)
	If @error Then Return SetError(2, @error, False)
	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr
	Local $oPoint = $oPdPage.GetSize
	If Not (ObjName($oPoint) = "CAcroPoint") Then Return SetError(3, 0, False)
	Local $aPoint = [$oPoint.X, $oPoint.Y]
	If @error Then Return SetError(4, @error, False)
	Return $aPoint

EndFunc   ;==>_Acro_PageSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_PageGetText
; Description ...: Gets the text of the indicated page
; Syntax ........: _Acro_PageGetText($oPdDoc, $iPage)
; Parameters ....: $oPdDoc              - an object.
;                  $iPage               - an integer value.
; Return values .: Success - The text of the current page
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc
;                  |2 - Unable to get page size. @extended set to _Acro_PageSize's error.
;                  |3 - Unable to select text in PDDocument. There is likely no selectable text on the page. @extended set to COM Error
;                  |4 - Can't create Rectangle for text selection
;                  |5 - Error getting text from selection, @extended set to COM Error
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_PageGetText($oPdDoc, $iPage)

	If Not (ObjName($oPdDoc) = "CAcroPDDoc") Then Return SetError(1, 0, False)

	Local $aPoint = _Acro_PageSize($oPdDoc, $iPage)
	If @error Then Return SetError(2, @error, False)

	Local $oRect = __Acro_Rect(0, $aPoint[1], 0, $aPoint[0])
	If @error Then Return SetError(4, @error, False)

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr

	Local $oTextSelection = $oPdDoc.CreateTextSelect($iPage, $oRect)
	If @error Or Not IsObj($oTextSelection) Then Return SetError(3, @error, False)

	Local $sText = ""
	; For each text piece
	For $i = 0 To $oTextSelection.GetNumText - 1
		$sText &= $oTextSelection.GetText($i)
		If @error Then Return SetError(5, @error, False)
	Next
	Return $sText

EndFunc   ;==>_Acro_PageGetText

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_PageViewMode
; Description ...: Gets or sets the PageMode of the current document.
; Syntax ........: _Acro_PageViewMode($oPdDoc[, $iPageMode = Default])
; Parameters ....: $oPdDoc              - an object.
;                  $iPageMode           - [optional] the new Page Mode. Default returns current value.
; Return values .: Success - The page mode of the PDDocument
;                  Failure - False and sets @error:
;                  |1 - $oPdDoc isn't a PDDoc
;                  |2 - Error setting page mode, @extended set to COM Error
;                  |3 - Error getting page mode, @extended set to COM Error
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_PageViewMode($oPdDoc, $iPageMode = Default)

	If Not (ObjName($oPdDoc) = "CAcroPDDoc") Then Return SetError(1, 0, False)

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr

	If Not IsKeyword($iPageMode) Then $oPdDoc.SetPageMode($iPageMode)
	If @error Then Return SetError(2, @error, False)

	$iPageMode = $oPdDoc.GetPageMode()
	If @error Then Return SetError(3, @error, False)

	Return $iPageMode

EndFunc   ;==>_Acro_PageViewMode

Func _Acro_PagesExtract($oPdDoc, $sSavePath, $iStart = 0, $iEnd = $iStart)

	Local $oJsDoc = _Acro_DocJSObject($oPdDoc)
	$oJsDoc.extractPages($iStart, $iEnd, $sSavePath)
	If @error Then Return SetError(1, @error, False)
	Return True

EndFunc

Func _Acro_PagesExtractDoc($oPdDoc, $iStart = 0, $iEnd = $iStart)

	Local $oJsDoc = _Acro_DocJSObject($oPdDoc)
	Local $oNewDoc = $oJsDoc.extractPages($iStart, $iEnd)
	If @error Then Return SetError(1, @error, False)
	Return $oNewDoc

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_WrapperDocsAppend
; Description ...: Appends a 0-based 1D array of PDF filenames
; Syntax ........: _Acro_WrapperDocsAppend($aDocs[, $sFinalPath = Default])
; Parameters ....: $aDocs               - a 0-based 1D array of PDF filenames.
;                  $sFinalPath          - where to save the combined PDF.
; Return values .: Success - True
;                  Failure - False and sets @error
;                  |10x - _Acro_DocOpen, x is error, sets @extended
;                  |20x - _Acro_DocAppend, x is error, sets @extended
;                  |30x - _Acro_DocSave, x is error, sets @extended
;                  |40x - _Acro_DocClose, x is error, sets @extended
; Author ........: Seadoggie01
; Modified ......: January 14, 2021
; Remarks .......: Not well tested. Leaves the first document open when @error is 200+. Use at your own risk.
;                  +I think this also needs to have the save properties changed. I constantly seem to change this function so
;                  +that it "works".
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_WrapperDocsAppend($aDocs, $sFinalPath)

	#ToDo: Check that all documents exist and are PDFs
	#ToDo: Use a FileCopy instead if Ubound($aDocs) = 1

	Local $oPdDoc = _Acro_DocOpen($aDocs[0])
	If @error Then Return SetError(100 + @error, @extended, False)

	For $i=1 To UBound($aDocs) - 1
		_Acro_DocAppend($oPdDoc, $aDocs[$i])
		If @error Then Return SetError(200 + @error, @extended, False)
	Next

	_Acro_DocSave($oPdDoc, $sFinalPath, $PDSaveFull)
	If @error Then Return SetError(300 + @error, @extended, False)

	_Acro_DocClose($oPdDoc)
	If @error Then Return SetError(400 + @error, @extended, False)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Acro_WrapperDocText
; Description ...: Gets the full text of a document
; Syntax ........: _Acro_WrapperDocText($sFile[, $vOptions = 0])
; Parameters ....: $sFile               - the full path to the file.
;                  $vOptions            - [optional] 0 for a string, 1 for a 0-based 1D array of page's text. Default is 0.
; Return values .: Success - the text of $sFile, a string or a 0-based 1D array
;                  Failure - False and sets @error and @extended (where x is the function's error value):
;                  |10x - _Acro_DocOpen
;                  |20x - _Acro_PageCount
;                  |30x - _Acro_PageGetText
;                  |40x - _Acro_DocClose
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Acro_WrapperDocText($sFile, $vOptions = 0)

	Local $oPdDoc = _Acro_DocOpen($sFile)
	If @error Then Return SetError("10" & @error, @extended, False)

	Local $iPages = _Acro_PageCount($oPdDoc)
	If @error Then Return SetError("20" & @error, @extended, False)

	Local $sText, $sFinalText, $aText[$iPages]

	For $i=0 To $iPages - 1

		$sText = _Acro_PageGetText($oPdDoc, $i)
		If @error Then Return SetError("30" & @error, @extended, False)

		If $vOptions = 0 Then $sFinalText &= $sText
		If $vOptions = 1 Then $aText[$i] = $sText
	Next

	_Acro_DocClose($oPdDoc)
	If @error Then Return SetError("40" & @error, @extended, False)

	If $vOptions = 0 Then Return $sFinalText
	If $vOptions = 1 Then Return $aText

EndFunc

#Region ### Internal Functions ###

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __Acro_Rect
; Description ...: Creates a Rectangle from the points passed.
; Syntax ........: __Acro_Rect($iBottom, $iTop, $iLeft, $iRight)
; Parameters ....: $iBottom             - an integer value.
;                  $iTop                - an integer value.
;                  $iLeft               - an integer value.
;                  $iRight              - an integer value.
; Return values .: AcroExch.Rect
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Acro_Rect($iBottom, $iTop, $iLeft, $iRight)

	Local $oRect = ObjCreate("AcroExch.Rect")
	If Not __Acro_ObjCheck($oRect, "Rect") Then Return SetError(1, 0, False)
	$oRect.Bottom = $iBottom
	$oRect.Top = $iTop
	$oRect.Left = $iLeft
	$oRect.Right = $iRight
	Return $oRect

EndFunc   ;==>__Acro_Rect

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __Acro_FileGetFolder
; Description ...: $sPath's parent folder
; Syntax ........: __Acro_FileGetFolder($sPath)
; Parameters ....: $sPath               - a string value.
; Return values .: Returns everything left of the last backslash to get the folder
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Acro_FileGetFolder($sPath)

	Return StringLeft($sPath, StringInStr($sPath, "\", 0, -1))

EndFunc   ;==>__Acro_FileGetFolder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __Acro_ObjCheck
; Description ...: Checks that an Object returns the expected name from ObjName
; Syntax ........: __Acro_ObjCheck($oObject, $sType)
; Parameters ....: $oObject             - an object.
;                  $sType               - a string value.
; Return values .: Success - True
;                  Failure - False (Object isn't of expected type). May set @error:
;                  | 1 - The object type isn't supported.
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......: This is because I'm not sure that ObjName will return the correct values on everyone's system.
;                : By centralizing the checks I can change the process better.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Acro_ObjCheck($oObject, $sType)

	; Remember to leave a pipe after the final object name
	Local $sSupportedObjects = "PDDoc|App|AVDoc|AVPageView|PDPage|Rect|"
	; Check if the type is supported
	If StringInStr($sSupportedObjects, $sType & "|") Then
		; All Acrobat objects return their names pre-pended with CAcro (that I've found)
		Return ObjName($oObject) = "CAcro" & $sType
	Else
		ConsoleWrite("- __Acro_ObjCheck with undocumented option: " & $sType & @CRLF & "-	ObjName: " & ObjName($oObject) & @CRLF)
		Return SetError(1, 0, False)
	EndIf

EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __Acro_ErrHnd
; Description ...: An error handler
; Syntax ........: __Acro_ErrHnd()
; Parameters ....: None
; Return values .: None - sets @error in case of COM Errors
; Author ........: Seadoggie01
; Modified ......: October 5, 2020
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __Acro_ErrHnd()

EndFunc

; N.B. Security Policy Object's properties are Read Only
Func __Acro_SecPolicyGet($oJavaScript, $sPolicyName)

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr

	Local $oJsSecurity = $oJavaScript.security
	If @error Then Return SetError(1, @error, False)

	Local $aoPolicies = $oJsSecurity.getSecurityPolicies()
	If @error Then Return SetError(2, @error, False)

	Local $oPolicy
	For $i=0 To UBound($aoPolicies) - 1
		If $aoPolicies[$i].Name = $sPolicyName Then
			$oPolicy = $aoPolicies[$i]
		ElseIf @error Then
			Return SetError(3, @error, False)
		EndIf
	Next

	Return $oPolicy

EndFunc

; This method doesn't throw an exception or error but appears to do nothing
Func __Acro_PrivContext($oJavaScript, $bPriv = True)

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr

	Local $oJavaScriptApp = $oJavaScript.App
	If @error Then Return SetError(2, @error, False)

	If $bPriv Then
		$oJavaScriptApp.beginPriv()
	Else
		$oJavaScriptApp.endPriv()
	EndIf
	If @error Then Return SetError(1, @error, False)
	Return True

EndFunc

; N.B. Security Handler Object's properties are Read Only
;      Methods (None successfully used yet)
;		- login
;		- logout
;		- newDirectory
;		- newUser
;		- setPasswordTimeout
Func __Acro_SecHandler($oJavaScript)

	Local $oErr = ObjEvent("AutoIt.Error", __Acro_ErrHnd)
	#forceref $oErr

	Local $oRDN = $oJavaScript.RDN
	If @error Then Return SetError(1, @error, False)

	$oRDN.cn = "username"
	If @error Then Return SetError(11, @error, False)

	Local $oJsSecurity = $oJavaScript.security
	If @error Then Return SetError(2, @error, False)

	Local $oHandler = $oJsSecurity.getHandler($oJsSecurity.PPKLiteHandler, False)
	If @error Then Return SetError(3, @error, False)

	$oHandler.newUser("password", Default, $oRDN)
	If @error Then Return SetError(4, @error, False)

	Return $oHandler.Name

EndFunc

#EndRegion ### Internal Functions ###
