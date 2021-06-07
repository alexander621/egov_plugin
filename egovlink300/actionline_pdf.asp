<%
response.buffer = True 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: actionline_pdf.asp
' AUTHOR:   David Boyer
' CREATED:  02/22/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' DESCRIPTION:  CREATES PDF FILE CONTAINING ACTION LINE REQUEST INFORMATION FOR THE USER TO PRINT.
'
' MODIFICATION HISTORY
' 1.0  02/22/08  David Boyer - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

'INITIALIZE AND DECLARE VARIABLES
 Dim sPrimaryTemplate,sSecondaryTemplate
 Dim oPDF,sSystem,sDB,sRequestTitle
 sPrimaryTemplate = server.mappath("actionline_pdf_page.pdf")
 'sSecondaryTemplate = server.mappath("templates/wo_last_page.pdf")
 sSystem = request("sys")
 Select Case sSystem 
	  Case "DEV"
	 	  sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
  	Case "QA"
	 	  sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink400_QA_Test; UID=egovsa; PWD=egov_4303;"
  	Case "LIVE"
		   sDB = "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=sa; PWD=;"
  	Case Else
	 	  sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
 End Select 

'CREATE PDF OBJECT
 Set oPDF  = Server.CreateObject("APToolkit.Object")
 oDocument = oPDF.OpenOutputFile("MEMORY") 'CREATE THE OUTPUT IN MEMORY

'BUILD PDF DOCUMENT
 Dim iFont, iLineSpacing, iSpacer, iFontFace, sStaringXPosition, sStartingYPosition, iCurrentYPosition, sSubmitName, datSubmitDate 
 Dim iPageWidth,iPageHeight
 iFont                   = 10
 iLineSpacing            = 13
' iFontFace               = "arial" 'NEED MONOSPACE FONT TO MAINTAIN SPACING IN TEXT FILE
 iFontFace               = "Times" 'NEED MONOSPACE FONT TO MAINTAIN SPACING IN TEXT FILE
 sStartingXPosition      = 25      'DEFAULT TEXT X POSITION
 sStartingYPosition      = 765     'DEFAULT TEXT Y POSITION
' iPageWidth              = 500
 iPageWidth              = 582
 iPageHeight             = 792
 iPageNum                = 1
 oPDF.OutputPageWidth    = iPageWidth  '8.5 inches
 oPDF.OutputPageHeight   = iPageHeight '11 inches
 lcl_text_alignment      = ""  'Default is "" which means LEFT
 session("bold_start")   = ""
 session("italic_start") = ""

'ADD TEXT TO PAGE
 r = oPDF.AddLogo(sPrimaryTemplate,0)
 oPDF.PrintLogo

'---------------------------------------------------------------------------
'Get pdf data (form letter)
 Dim lcl_letter_id, lcl_request_id, sStatus, sSubStatus, lcl_org_id, lcl_user_id, lcl_add_text
 lcl_letter_id  = request("iletterid")
 lcl_request_id = request("action_autoid")
 sStatus        = request("status")
 sSubStatus     = request("substatus")
 lcl_org_id     = request("orgid")
 lcl_user_id    = request("userid")
 lcl_add_text   = request("add_text")

'Get the title
 sSQL = "SELECT flid, orgid, sequence, fltitle, flbody, blnallmergefields "
 sSQL = sSQL & " FROM FormLetters "
 sSQL = sSQL & " WHERE FLid = " & CLng(lcl_letter_id)

 set oFormLetter = Server.CreateObject("ADODB.Recordset")
 oFormLetter.Open sSQL, sDB, 3, 1

 if NOT oFormLetter.EOF then
    lcl_fl_title = oFormLetter("FLtitle")

  		lcl_body = oFormLetter("FLbody")
'  		lcl_body = replace(lcl_body,"<p>","[p]")
'	  	lcl_body = replace(lcl_body,"</p>","[/p]")
'  		lcl_body = replace(lcl_body,"<P>","[P]")
'		  lcl_body = replace(lcl_body,"</P>","[/P]")

'   	lcl_body = replace(lcl_body,"</br>","[/br]")
' 		 lcl_body = replace(lcl_body,"<br>","[br]")
'	 	 lcl_body = replace(lcl_body,"</BR>","[/BR]")
'		  lcl_body = replace(lcl_body,"<BR>","[BR]")
    lcl_body = replace(lcl_body,chr(10),"[line_break]")
'    lcl_body = replace(lcl_body,vbcrlf,"")
    lcl_body = replace(lcl_body,chr(13),"")

    oFormLetter.Close
 else
    lcl_fl_title = ""
    lcl_body     = ""
 end if

 Set oFormLetter = Nothing
'---------------------------------------------------------------------------

'Get the contact user id for this request
 sSQL = "SELECT userid "
 sSQL = sSQL & " FROM egov_actionline_requests "
 sSQL = sSQL & " WHERE action_autoid = " & lcl_request_id

 set oRst = Server.CreateObject("ADODB.Recordset")
 oRst.Open sSQL, sDB, 3, 1

 if oRst.eof then
	   iUserid = 124
 else
	   iUserid = oRst("userid")
   	oRst.Close
 end if

 set oRst = Nothing
'---------------------------------------------------------------------------

'Build message for Activity Log
 sSQL = "SELECT * "
 sSQL = sSQL & " From egov_users "
 sSQL = sSQL & " WHERE userid = " & lcl_user_id
 Set oUser = Server.CreateObject("ADODB.Recordset")
 oUser.Open sSQL, sDB, 3, 1

 If NOT oUser.EOF Then
   	if IsNull(lcl_add_text) or lcl_add_text = "" then
		     lcl_add_text = ""
   	else
     		lcl_add_text = " - " & replace(lcl_add_text,"<br>"," ")
     		lcl_add_text = replace(lcl_add_text,"%20"," ")
   	end if 

   	oUser.Close

 End If

 Set oUser = Nothing
'---------------------------------------------------------------------------

'Record in the Activity Log that the PDF was printed.
 AddCommentTaskComment lcl_fl_title & " PDF Printed " & lcl_add_text,"",sStatus,lcl_request_id,lcl_user_id,lcl_org_id,sSubStatus
'---------------------------------------------------------------------------

'HEADER INFORMATION 
'	oPDF.GreyBar sStartingXPosition - 5, sStartingYPosition - 1, 550, iFont + 5, 0
'	oPDF.SetTextColor 255,255,255,255 ' CHANGE TEXT COLOR TO WHITE
	oPDF.SetFont "Times-Bold", iFont + 2
	oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK

'	oPDF.PrintText sStartingXPosition,sStartingYPosition + 2, lcl_fl_title
	oPDF.SetFont iFontFace, iFont
'	oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK
	iCurrentYPosition = sStartingYPosition

'-----------------------------------------------------------------------
 lcl_body = fnFillMergeFields(lcl_body,iUserid,lcl_add_text,lcl_request_id)

 if lcl_body <> "" AND NOT isnull(lcl_body) AND instr(lcl_body,"[line_break]") > 0 THEN

   	arrComment = split(lcl_body,"[line_break]")
   	iNumLines  = UBOUND(arrComment)

    if iNumLines < 1 then
       iNumLines = 1
    end if

    for iCommentLines = 0 to iNumLines
  		   	fnNewLineCheck trim(arrComment(iCommentLines)),lcl_fl_title

        if trim(arrComment(iCommentLines)) = "" OR isnull(trim(arrComment(iCommentLines))) then
          	iCurrentYPosition = iCurrentYPosition - iLineSpacing
        end if
    Next

   	iCurrentYPosition = iCurrentYPosition - iLineSpacing
 else
    fnNewLineCheck trim(lcl_body),lcl_fl_title
 end if

'Display data on PDF page
'  if lcl_body <> "" then
'   		iCurrentYPosition = iCurrentYPosition - iLineSpacing
'   		oPDF.PrintText sStartingXPosition,iCurrentYPosition, lcl_body
'  end if

'		iCurrentYPosition = iCurrentYPosition - (iLineSpacing * 2)		

'---------------------------------------------------------------------------

'ADD LAST PAGE
'oPDF.NewPage
'oPDF.ClearLogosandImages ' CLEAR FRONT UNDERLAY
'oPDF.AddLogo sSecondaryTemplate,0 ' LOAD SECOND PAGE UNDERLAY INTO MEMORY
'oPDF.PrintLogo

'HEADER INFORMATION FOR LAST PAGE
'oPDF.GreyBar sStartingXPosition - 5, sStartingYPosition - 1, 550, iFont + 5, 0
'oPDF.SetTextColor 255,255,255,255 ' CHANGE TEXT COLOR TO WHITE
'oPDF.SetFont "arial", iFont + 2
'oPDF.PrintText sStartingXPosition,sStartingYPosition + 2,  sRequestTitle

'BODY FOR LAST PAGE
'oPDF.SetFont iFontFace, iFont
'oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK
'iCurrentYPosition = sStartingYPosition - iLineSpacing
'oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Tracking Number: " & request("iRequestID")  & replace(FormatDateTime(cdate(datSubmitDate),4),":","")
'iCurrentYPosition = iCurrentYPosition - iLineSpacing
'oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Date Time Received: " & datSubmitDate
'iCurrentYPosition = iCurrentYPosition - iLineSpacing
'oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Created By: " & sSubmitName
'iCurrentYPosition = iCurrentYPosition - (iLineSpacing * 2)	

oPDF.CloseOutputFile
oDocument = oPDF.binaryImage 

' STREAM PDF TO BROWSER
response.expires = 0
response.Clear
response.ContentType = "application/pdf"
response.AddHeader "Content-Type", "application/pdf"
response.AddHeader "Content-Disposition", "inline;filename=FORMS.PDF"
response.BinaryWrite oDocument  

' DESTROY OBJECTS
Set oPDF      = Nothing
Set oDocument = Nothing
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function FormatPhone( Number )
'--------------------------------------------------------------------------------------------------
Function FormatPhone( Number )
  If Len(Number) = 10 Then
     FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
  Else
     FormatPhone = Number
  End If
End Function

'--------------------------------------------------------------------------------------------------
'SUB SUBNEWPAGECHECK(ICURRENTYPOSITION,SHEADING)
'--------------------------------------------------------------------------------------------------
Sub subNewPageCheck(iCurrentYPosition, sHeading)

	' CHECK TO SEE IF WE NEED TO START A NEW PAGE
	If Clng(iCurrentYPosition) <= 36 Then
		oPDF.NewPage
		iCurrentYPosition = sStartingYPosition
  iPageNum = iPageNum + 1
		oPDF.AddLogo sPrimaryTemplate,0
		oPDF.PrintLogo
'		oPDF.SetFont "arial", iFont 
'		oPDF.SetFont "Times", iFont 

		' HEADER INFORMATION 
'		oPDF.GreyBar sStartingXPosition - 5, sStartingYPosition - 1, 550, iFont + 5, 0
'		oPDF.SetTextColor 255,255,255,255 ' CHANGE TEXT COLOR TO WHITE
'	oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK	
'	oPDF.SetFont "arial", iFont + 2
'	oPDF.SetFont "Times", iFont + 2
'	oPDF.PrintText sStartingXPosition,sStartingYPosition + 2,  sRequestTitle
'	oPDF.SetFont iFontFace, iFont
'	iCurrentYPosition = sStartingYPosition - iLineSpacing
'	oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Tracking Number: " & request("iRequestID") & replace(FormatDateTime(cdate(datSubmitDate),4),":","")
'	iCurrentYPosition = iCurrentYPosition - iLineSpacing
'	oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Date Time Received: " & datSubmitDate
'	iCurrentYPosition = iCurrentYPosition - iLineSpacing
'	oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Created By: " & sSubmitName
'		iCurrentYPosition = iCurrentYPosition - (iLineSpacing * 2)

		' CURRENT HEADER INFORMATION
'		oPDF.GreyBar sStartingXPosition - 5, iCurrentYPosition - 1, 550, iFont + 5, 0
'		oPDF.SetTextColor 255,255,255,255 ' CHANGE TEXT COLOR TO WHITE
	'	oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK
'		oPDF.SetFont "arial", iFont + 2
		oPDF.SetFont "Times", iFont + 2
		oPDF.PrintText sStartingXPosition,iCurrentYPosition + 2, sHeading & " (continued)"
'		oPDF.SetFont iFontFace, iFont
'		oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK
		iCurrentYPosition = iCurrentYPosition - iLineSpacing
	End If

End Sub

'--------------------------------------------------------------------------------------------------
' FUNCTION fnNewLineCheck(STEXT,SHEADING)
'--------------------------------------------------------------------------------------------------
Function fnNewLineCheck(sText,sHeading)
 'If the output contains a link in it is counted within the iLineWidth BEFORE the link is formatted for PDF
 'If the HREF value is longer than the page then the PDF output attempts to wrap it to the next line and 
 '   the link does not display propery in the PDF
 'Therefore check to see if the sText being passed is a LINK and if so check the length of ONLY the display text
 '  (i.e. <a href="url">display text</a>)
  if instr(UCASE(sText),"<A ") = 0 then
     lcl_display_line = sText
     lcl_display_line = REPLACE(lcl_display_line, "<B>","")
     lcl_display_line = REPLACE(lcl_display_line,"</B>","")
     lcl_display_line = REPLACE(lcl_display_line, "<I>","")
     lcl_display_line = REPLACE(lcl_display_line,"</I>","")

     lcl_display_line = REPLACE(lcl_display_line, "<b>","")
     lcl_display_line = REPLACE(lcl_display_line,"</b>","")
     lcl_display_line = REPLACE(lcl_display_line, "<i>","")
     lcl_display_line = REPLACE(lcl_display_line,"</i>","")

   	 'iLineWidth = oPDF.GetTextWidth(sText)
     iLineWidth = oPDF.GetTextWidth(lcl_display_line)

  else
    'Find the position of the END of the anchor open tag <a href="url"> <=== (looking for the FIRST greater than character)
     lcl_opentag_end_pos    = instr(sText,">")

    'Find the position of the START of the anchor close tag </A>
     lcl_closetag_start_pos = instr(UCASE(sText),"</A>")

    'Pull the text between these points and evaluate the length
     lcl_display_text = mid(sText,lcl_opentag_end_pos,lcl_closetag_start_pos)

   	 iLineWidth = oPDF.GetTextWidth(lcl_display_text)
  end if

 'Check for the setting of text alignment.  This will determine where the "starting X position is"
  if instr(UCASE(sText),"<P ALIGN='CENTER'>") = 1 then
     session("text_align_center") = "Y"
     lcl_value = ""

  elseif instr(UCASE(sText),"<P ALIGN='RIGHT'>") = 1 then
     session("text_align_right") = "Y"
     lcl_value = ""

  else
     lcl_value = sText
  end if

  if session("text_align_center") = "Y" OR session("text_align_right") = "Y" then
     if UCASE(sText) <> "</P>" then
        if instr(UCASE(sText),"<IMG ") = 0 AND instr(UCASE(sText),"<A ") = 0 then
           if session("text_align_center") = "Y" then
              lcl_starting_x_position = (iPageWidth/2) - (oPDF.GetTextWidth(lcl_display_line)/2)
           elseif session("text_align_right") = "Y" then
              lcl_starting_x_position = iPageWidth - oPDF.GetTextWidth(lcl_display_line)
           end if

           if lcl_starting_x_position < 25 then
              lcl_starting_x_position = 25
           end if
        end if
     else
        if session("text_align_center") = "Y" then
           session("text_align_center") = "N"
        elseif session("text_align_right") = "Y" then
           session("text_align_right") = "N"
        end if

        lcl_value = ""
        lcl_starting_x_position = sStartingXPosition
     end if
  else
     lcl_starting_x_position = sStartingXPosition
  end if

'----------------------------------------
 	If iLineWidth > iPageWidth Then
'----------------------------------------

    'CREATE WRAP LINE
    	sWrapLine = ""
    	arrWords  = split(replace(sText,"["," "),chr(32))
     'arrWords  = split(sText,"]")  'ORIGINAL CODE

   		For iWordCount = 0 To UBOUND(arrWords)
         If oPDF.GetTextWidth(sWrapLine) < iPageWidth-25 Then
     	   			sWrapLine = sWrapLine & " " & arrWords(iWordCount) 
      			Else
        				Exit For
   			   End If
   		Next

   	'REST OF LINE
   		iRestWordCount = 0
     sRestofLine    = ""
   		For iRestWordCount = iWordCount To UBOUND(arrWords)
         sRestofLine = sRestofLine & " " & arrWords(iRestWordCount) 
   		Next

   	'DISPLAY WRAPPED LINE
   		Call subNewPageCheck(iCurrentYPosition, sHeading)
     '		oPDF.PrintText sStartingXPosition,iCurrentYPosition, "2. " & trim(sWrapLine)
     '		iCurrentYPosition = iCurrentYPosition - iLineSpacing
     checkCustomFormatting trim(sWrapLine),lcl_starting_x_position

   	'CONTINUE WITH REST OF LINE WRAPPING AS NEEDED
   		Call subNewPageCheck(iCurrentYPosition, sHeading)
   		fnNewLineCheck sRestofLine, sHeading

'----------------------------------------
  Else
'----------------------------------------

  		'WRITE LINE AS IS - NOT WRAPPING REQUIRED
   		Call subNewPageCheck(iCurrentYPosition, sHeading)

     checkCustomFormatting lcl_value,lcl_starting_x_position

'----------------------------------------
  End If
'----------------------------------------

End Function

'------------------------------------------------------------------------------------------------------------
Function AddCommentTaskComment(sInternalMsg,sExternalMsg,sStatus,iFormID,iUserID,iOrgID,sSubStatus)
  if sSubStatus = "" then
     lcl_sub_status = 0
  else
     lcl_sub_status = sSubStatus
  end if

  sSQL = "INSERT egov_action_responses (action_status,action_internalcomment,action_externalcomment,action_userid,action_orgid,action_autoid,action_sub_status_id) "
  sSQL = sSQL & " VALUES ("
  sSQL = sSQL & "'" & sStatus              & "', "
  sSQL = sSQL & "'" & DBsafe(sInternalMsg) & "', "
  sSQL = sSQL & "'" & DBsafe(sExternalMsg) & "', "
  sSQL = sSQL & "'" & iUserID              & "', "
  sSQL = sSQL & "'" & iOrgID               & "', "
  sSQL = sSQL & "'" & iFormID              & "', "
  sSQL = sSQL       & lcl_sub_status       & ")"
  Set oComment = Server.CreateObject("ADODB.Recordset")
  oComment.Open sSQL, sDB , 3, 1
  Set oComment = Nothing

 'UPDATE THE TASK SUB-STATUS
  sSqls = "UPDATE egov_actionline_requests "
  sSqls = sSqls & " SET sub_status_id = "   & lcl_sub_status
  sSqls = sSqls & " WHERE action_autoid = " & iFormID

  Set oUpdate2 = Server.CreateObject("ADODB.Recordset")
  oUpdate2.Open sSqls, sDB, 3, 1
  Set oUpdate2 = Nothing

End Function

'------------------------------------------------------------------------------------------------------------
function dbsafe(p_value)
  if p_value = "" or isnull(p_value) then
     lcl_return = ""
  end if

  lcl_return = REPLACE(p_value,"'","''")

  dbsafe = lcl_return

end function

'------------------------------------------------------------------------------------------------------------
Function fnFillMergeFields(sValue,iuserid,sAddText,itrackid)

	'Fill General Purpose Fields
	 sValue = fnFillGeneralFields( sValue )

	'FILL STANDARD FIELDS
 	sValue = fnFillStandardFields(sValue,iuserid)

	'FILL ISSUE LOCATION FIELDS
	 sValue = fnFillIssueLocationFields(sValue,itrackid)

	'FILL DYNAMIC FIELDS
 	sValue = fnFillDynamicFields(sValue,itrackid)

	'FILL ADDITIONAL COMMENT FIELD
	 sValue = fnFillAdditionalCommentField(sValue,sAddText)

	'FILL CODE SECTIONS FIELDS
  sValue = fnFillCodeSectionsField(sValue,itrackid)

 'FILL TRACKING NUMBER FIELD
  sValue = fnFillTrackingNumberField(sValue,itrackid)

	'CLEAR ANY REMAINING UNFILLED MERGE FIELDS
' 	sValue = fnClearMergeFields(sValue)

	'RETURN VALUE
 	fnFillMergeFields = sValue

End Function

'------------------------------------------------------------------------------------------------------------
Function fnFillGeneralFields( ByVal sValue )
	Dim sNewValue, sDate

	sDate = MonthName(Month(Date())) & " " & Day(Date()) & ", " & Year(Date())
	
	' Today's Date
	sNewValue = Replace(sValue,"[*TodaysDate*]", sDate )
	
	fnFillGeneralFields = sNewValue 
	
End Function

'------------------------------------------------------------------------------------------------------------
' FUNCTION FNFILLADDITIONALCOMMENTFIELD(SVALUE)
'------------------------------------------------------------------------------------------------------------
Function fnFillAdditionalCommentField(sValue,sAddText)
	
	sValue = replace(sValue,"[*Admin_Additional_Comments*]",sAddText)
	
	' RETURN VALUE
	fnFillAdditionalCommentField = sValue 
	
End Function

'------------------------------------------------------------------------------------------------------------
Function fnFillStandardFields(sValue,iUserId)

'SET DEFAULT RETURN VALUE
	sReturnValue = sValue

'CONNECT TO DATABASE RETRIEVE VALUES FOR FORM FIELDS ASSOCIATED WITH THIS REQUEST
	sSQL ="SELECT * FROM egov_users WHERE userid='" & iUserId & "'"

	Set oFields = Server.CreateObject("ADODB.Recordset")
	oFields.Open sSQL, sDB, 3, 1

	If NOT oFields.EOF Then
  	'LOOP THRU FORM FIELDS
  		For Each Field In oFields.Fields

    			'REPLACE FORM FIELD PLACE HOLDER WITH ACTUAL DATA
     			sFieldPlaceHolder = "[*" & Field.Name & "*]"
     			if oFields(Field.Name) <> "" AND NOT IsNull(oFields(Field.Name)) then
       				sReturnValue = replace(sReturnValue,sFieldPlaceHolder,oFields(Field.Name))
     			end if
  		Next

	End If

'RETURN VALUE
	fnFillStandardFields = sReturnValue

End Function

'------------------------------------------------------------------------------------------------------------
Function fnFillIssueLocationFields(sValue,irequestid)

'SET DEFAULT RETURN VALUE
	sReturnValue = sValue

'CONNECT TO DATABASE RETRIEVE VALUES FOR ISSUE LOCATION FIELDS ASSOCIATED WITH THIS REQUEST
'	sSQL = "SELECT streetnumber, streetaddress, city, state, zip, comments "
' sSQL = sSQL & " FROM egov_action_response_issue_location "
' sSQL = sSQL & " WHERE actionrequestresponseid = '" & irequestid & "'"
	sSQL ="SELECT streetnumber, "
 sSQL = sSQL & " dbo.fn_buildAddress('', "
 sSQL = sSQL &                      "ISNULL(streetprefix, ''), "
 sSQL = sSQL &                      "ISNULL(streetaddress, ''), "
 sSQL = sSQL &                      "ISNULL(streetsuffix, ''), "
 sSQL = sSQL &                      "ISNULL(streetdirection, '')) AS streetaddress, "
 sSQL = sSQl & " city, state, zip, comments "
 sSQL = sSQL & " FROM egov_action_response_issue_location "
 sSQL = sSQL & " WHERE actionrequestresponseid = '" & irequestid & "'"

	Set oFields = Server.CreateObject("ADODB.Recordset")
	oFields.Open sSQL, sDB, 3, 1

	If NOT oFields.EOF Then
 		'LOOP THRU FORM FIELDS
   	For Each Field In oFields.Fields

			    'REPLACE FORM FIELD PLACE HOLDER WITH ACTUAL DATA
     			sFieldPlaceHolder = "[*" & Field.Name & "*]"
     			if oFields(Field.Name) <> "" AND NOT IsNull(oFields(Field.Name)) then
       				sReturnValue = replace(sReturnValue,sFieldPlaceHolder,oFields(Field.Name))
      		end if
   	Next
	End If

'RETURN VALUE
	fnFillIssueLocationFields = sReturnValue

End Function

'------------------------------------------------------------------------------------------------------------
Function fnFillDynamicFields( sValue, irequestid )

	'SET DEFAULT RETURN VALUE
	 sReturnValue = sValue

	'Retrieve values for form fields associated with this request
	 sSQL = "SELECT submitted_request_field_id, submitted_request_field_type_id "
  sSQL = sSQL & " FROM egov_submitted_request_fields "
  sSQL = sSQL & " WHERE submitted_request_id = '" & irequestid & "' "
  sSQL = sSQL & " ORDER BY submitted_request_field_sequence "

	 set oDynamicFields = Server.CreateObject("ADODB.Recordset")
	 oDynamicFields.Open sSQL, sDB, 3, 1

  if not oDynamicFields.eof then
     lcl_value = ""
     while not oDynamicFields.eof

       'Cycle through and build answer list of all of the answers that have been selected/entered for each field on the form.
        sSQLr = "SELECT submitted_request_field_id, submitted_request_field_response, submitted_request_form_field_name "
        sSQLr = sSQLr & " FROM egov_submitted_request_field_responses "
        sSQLr = sSQLr & " WHERE submitted_request_field_id = " & oDynamicFields("submitted_request_field_id")
        sSQLr = sSQLr & " AND submitted_request_form_field_name IS NOT NULL "
        sSQLr = sSQLr & " AND submitted_request_form_field_name <> '' "
        sSQLr = sSQLr & " ORDER BY submitted_request_field_id "

      	 set rsr = Server.CreateObject("ADODB.Recordset")
      	 rsr.Open sSQLr, sDB, 3, 1

        if not rsr.eof then
          'If the field type is a radio (2), select (4), or checkbox (6) then we need to cycle through the answers to before converting
          'the place holder.

           if clng(oDynamicFields("submitted_request_field_type_id")) = 2 _
           OR clng(oDynamicFields("submitted_request_field_type_id")) = 4 _
           OR clng(oDynamicFields("submitted_request_field_type_id")) = 6 then

              lcl_value = ""
              while not rsr.eof
                'An answer of "default_novalue" is the hidden field that has been set up to fix the bug
                'with radio, select, and checkbox fields and the user not selecting an option in these lists.
                 if rsr("submitted_request_field_response") <> "default_novalue" then
                    if lcl_value <> "" then
                       lcl_value = lcl_value & ", " & cleanup_display_value(rsr("submitted_request_field_response"))
                    else
                       lcl_value = cleanup_display_value(rsr("submitted_request_field_response"))
                    end if
                 else
                    lcl_value = lcl_value
                 end if

                'Get the form name.  It should be the same for all of the answers.
                 lcl_submitted_request_form_field_name = rsr("submitted_request_form_field_name")

                 rsr.movenext
              wend

              if lcl_submitted_request_form_field_name <> "" then
              			sFieldPlaceHolder = "[*" & lcl_submitted_request_form_field_name & "*]"
                	sReturnValue      = replace(sReturnValue,sFieldPlaceHolder,lcl_value)
              else
              			sFieldPlaceHolder = "[*" & lcl_submitted_request_form_field_name & "*]"
                	sReturnValue      = replace(sReturnValue,sFieldPlaceHolder,"")
              end if
          '----------------------------------------------------------------------------------
           else  'Any other type of field can simply convert the place holder with the answer
          '----------------------------------------------------------------------------------
              lcl_value = cleanup_display_value(rsr("submitted_request_field_response"))

             'Get the form name.  It should be the same for all of the answers.
              lcl_submitted_request_form_field_name = rsr("submitted_request_form_field_name")

              if lcl_submitted_request_form_field_name <> "" then
              			sFieldPlaceHolder = "[*" & lcl_submitted_request_form_field_name & "*]"
                	sReturnValue      = replace(sReturnValue,sFieldPlaceHolder,lcl_value)
              else
              			sFieldPlaceHolder = "[*" & lcl_submitted_request_form_field_name & "*]"
                	sReturnValue      = replace(sReturnValue,sFieldPlaceHolder,"")
              end if
           end if
        end if

        oDynamicFields.movenext
     wend
  end if

 'The form/pdf may have fields in it that are not related to this action_autoid.
 'If any exist in the form/pdf then replace the field with NULL so that the [*this_value*] doesn't show.
  sSQLrf = "SELECT DISTINCT pdfformname "
  sSQLrf = sSQLrf & " FROM egov_action_form_questions "
  sSQLrf = sSQLrf & " WHERE pdfformname IS NOT NULL "
  sSQLrf = sSQLrf & " AND pdfformname <> '' "
  sSQLrf = sSQLrf & " ORDER BY pdfformname "

	 set rsrf = Server.CreateObject("ADODB.Recordset")
	 rsrf.Open sSQLrf, sDB, 3, 1

  if not rsrf.eof then
     while not rsrf.eof
     			sFieldPlaceHolder = "[*" & rsrf("pdfformname") & "*]"
       	sReturnValue      = replace(sReturnValue,sFieldPlaceHolder,"")

        rsrf.movenext
     wend
  end if

 'RETURN VALUE
 	fnFillDynamicFields = sReturnValue

End Function

'------------------------------------------------------------------------------------------------------------
Function fnFillCodeSectionsField(sValue,itrackid)

'SET DEFAULT RETURN VALUE
	sReturnValue = sValue

 'sSql = "SELECT code_sections FROM egov_actionline_requests WHERE action_autoid = " & itrackid

 sSql = "SELECT scs.submitted_action_code_id, cs.code_name, cs.description "
	sSql = sSql & " FROM egov_submitted_request_code_sections scs, egov_actionline_code_sections cs "
	sSql = sSql & " WHERE scs.submitted_action_code_id = cs.action_code_id "
	sSql = sSql & " AND scs.submitted_request_id = " & itrackid
	sSql = sSql & " ORDER BY cs.code_name, cs.description "

	Set oGetCodes = Server.CreateObject("ADODB.Recordset")
	oGetCodes.Open sSql, sDB, 3, 1

	If NOT oGetCodes.EOF Then

 		'LOOP THRU CODE SECTIONS
    lcl_display_codes = ""
	  	while not oGetCodes.eof
		    'REPLACE FORM FIELD PLACE HOLDER WITH ACTUAL DATA
       if lcl_display_codes = "" then
          lcl_display_codes = "<b>" & oGetCodes("code_name") & "</b>[line_break]- " & oGetCodes("description") & "[line_break][line_break]"
       else
          lcl_display_codes = lcl_display_codes & "<b>" & oGetCodes("code_name") & "</b>[line_break]- " & oGetCodes("description") & "[line_break][line_break]"
       end if

       oGetCodes.movenext
		  wend
 else
	   lcl_display_codes = ""
	end if

'RETURN VALUE
 sFieldPlaceHolder = "[*Code_Sections*]"
 sReturnValue      = replace(sReturnValue,sFieldPlaceHolder,lcl_display_codes)

 fnFillCodeSectionsField = sReturnValue

end function

'------------------------------------------------------------------------------------------------------------
function fnFillTrackingNumberField(sValue,itrackid)

'SET DEFAULT RETURN VALUE
	sReturnValue = sValue

 sSQL = "SELECT [tracking number] "
	sSQL = sSQL & " FROM egov_rpt_actionline "
	sSQL = sSQL & " WHERE action_autoid = " & itrackid

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSQL, sDB, 3, 1

	if NOT rs.eof then
    lcl_tracking_number = rs("tracking number")
 else
	   lcl_tracking_number = ""
	end if

'RETURN VALUE
 sFieldPlaceHolder = "[*Tracking Number*]"
 sReturnValue      = replace(sReturnValue,sFieldPlaceHolder,lcl_tracking_number)

 fnFillTrackingNumberField = sReturnValue

end function

'------------------------------------------------------------------------------------------------------------
function cleanup_display_value(p_value)
  if p_value <> "" then
     lcl_value = p_value
'     lcl_value = REPLACE(lcl_value,vbcrlf,"")
     lcl_value = REPLACE(lcl_value,chr(10),"[line_break]")
     lcl_value = REPLACE(lcl_value,chr(13),"")

     cleanup_display_value = lcl_value
  end if

end function

'------------------------------------------------------------------------------------------------------------
Function fnClearMergeFields(sString)

'DEFAULT RETURN VALUE
	sReturnValue = sString
	sPattern     = "([\[\*].*[\*\]])"

'CREATE REGULAR EXPRESSION TO FIND MATCHING TAG SYNTAX OF [* any merge field name *]
	Dim oMergeFields
	Set oMergeFields        = New RegExp
	oMergeFields.Pattern    = sPattern
	oMergeFields.Global     = True
	oMergeFields.IgnoreCase = True
		
'REPLACE MATCHING TAGS
	sReturnValue = oMergeFields.Replace(sString, " ")

'DESTROY OBJECT
	Set oMergeFields = Nothing

'RETURN VALUE
	fnClearMergeFields =  sReturnValue

End Function

'------------------------------------------------------------------------------
sub checkCustomFormatting(p_value,p_starting_x_position)
  lcl_starting_x_position = p_starting_x_position

     'Cycle the the value to determine if there is any custom formatting
      if len(trim(p_value)) > 0 then
         lcl_length = len(trim(p_value))

        'Font Formatting
         lcl_bold_start        = "N"
         lcl_italic_start      = "N"

        'Image building
         lcl_img_start         = "N"
         lcl_img_str           = ""
         lcl_img_src_start     = "N"
         lcl_img_src_str       = ""
         lcl_img_width_start   = "N"
         lcl_img_height_start  = "N"
         lcl_img_width_str     = ""
         lcl_img_height_str    = ""

        'Link
         lcl_anchor_start      = "N"

         lcl_formatted_value   = ""

         for i = 0 to lcl_length
             if i > 0 then
                lcl_formatted_value = lcl_formatted_value & mid(p_value,i,1)

               'Font formatting
                lcl_bold_start        = lcl_bold_start
                lcl_italic_start      = lcl_italic_start

               'Image building
                lcl_img_start         = lcl_img_start
                lcl_img_str           = lcl_img_str
                lcl_img_src_start     = lcl_img_src_start
                lcl_img_src_str       = lcl_img_src_str
                lcl_img_width_start   = lcl_img_width_start
                lcl_img_height_start  = lcl_img_height_start
                lcl_img_width_str     = lcl_img_width_str
                lcl_img_height_str    = lcl_img_height_str

               'Link
                lcl_anchor_start      = lcl_anchor_start
                lcl_anchor_str        = lcl_anchor_str
                lcl_anchor_href_start = lcl_anchor_href_start
                lcl_href_str          = lcl_href_str
                lcl_anchor_text_start = lcl_anchor_text_start
                lcl_link_text         = lcl_link_text

               '-- Check for OPENING TAGS ------------------------------------------------

               '-- BOLD <B> --------------------------------------------------------------
                if UCASE(right(lcl_formatted_value,3)) = "<B>" then  'Check for BOLD start
               '--------------------------------------------------------------------------
                   lcl_bold_start        = "Y"
                   session("bold_start") = "Y"
                   lcl_formatted_value = replace(lcl_formatted_value,"<b>","")
                   lcl_formatted_value = replace(lcl_formatted_value,"<B>","")

                   if lcl_formatted_value <> "" then
                   		 oPDF.PrintText lcl_starting_x_position, iCurrentYPosition, lcl_formatted_value
                      lcl_starting_x_position = lcl_starting_x_position + oPDF.GetTextWidth(lcl_formatted_value)
                      lcl_formatted_value     = ""
                   end if

                   if lcl_italic_start = "Y" then
                     	oPDF.SetFont "Times-BoldItalic", iFont
                   else
                     	oPDF.SetFont "Times-Bold", iFont
                   end if
               '-- ITALICS <I> -----------------------------------------------------------
                elseif UCASE(right(lcl_formatted_value,3)) = "<I>" then  'Check for ITALICS start
               '--------------------------------------------------------------------------
                   lcl_italic_start        = "Y"
                   session("italic_start") = "Y"
                   lcl_formatted_value = replace(lcl_formatted_value,"<i>","")
                   lcl_formatted_value = replace(lcl_formatted_value,"<I>","")

                   if lcl_formatted_value <> "" then
                   		 oPDF.PrintText lcl_starting_x_position, iCurrentYPosition, lcl_formatted_value
                      lcl_starting_x_position = lcl_starting_x_position + oPDF.GetTextWidth(lcl_formatted_value)
                      lcl_formatted_value     = ""
                   end if

                   if lcl_bold_start = "Y" then
                     	oPDF.SetFont "Times-BoldItalic", iFont
                   else
                     	oPDF.SetFont "Times-Italic", iFont
                   end if

               '-- Check for END TAGS ----------------------------------------------------

               '-- Check for BOLD end ----------------------------------------------------
                elseif UCASE(right(lcl_formatted_value,4)) = "</B>" then
               '--------------------------------------------------------------------------
                   lcl_bold_start        = "N"
                   session("bold_start") = "N"
                   lcl_formatted_value = replace(lcl_formatted_value,"</b>","")
                   lcl_formatted_value = replace(lcl_formatted_value,"</B>","")

                   if lcl_formatted_value <> "" then
                   		 oPDF.PrintText lcl_starting_x_position, iCurrentYPosition, lcl_formatted_value
                      lcl_starting_x_position = lcl_starting_x_position + oPDF.GetTextWidth(lcl_formatted_value)
                      lcl_formatted_value     = ""
                   end if

                   if lcl_italic_start = "Y" then
                     	oPDF.SetFont "Times-Italic", iFont
                   else
                      oPDF.SetFont iFontFace, iFont
                   end if

               '-- Check for ITALICS end -------------------------------------------------
                elseif UCASE(right(lcl_formatted_value,4)) = "</I>" then  
               '--------------------------------------------------------------------------
                   lcl_italic_start        = "N"
                   session("italic_start") = "N"
                   lcl_formatted_value = replace(lcl_formatted_value,"</i>","")
                   lcl_formatted_value = replace(lcl_formatted_value,"</I>","")

                   if lcl_formatted_value <> "" then
                   		 oPDF.PrintText lcl_starting_x_position, iCurrentYPosition, lcl_formatted_value
                      lcl_starting_x_position = lcl_starting_x_position + oPDF.GetTextWidth(lcl_formatted_value)
                      lcl_formatted_value     = ""
                   end if

                   if lcl_bold_start = "Y" then
                     	oPDF.SetFont "Times-Bold", iFont
                   else
                      oPDF.SetFont iFontFace, iFont
                   end if

               '-- Check for opening tag for LINKS ---------------------------------------
                elseif UCASE(right(lcl_formatted_value,3)) = "<A " then  
               '--------------------------------------------------------------------------
                   lcl_anchor_start = "Y"

                  'Begin tracking the entire link to format out of the body text later
                   lcl_anchor_str = right(lcl_formatted_value,3)

               '-- Check for the LINK HREF -----------------------------------------------
                elseif UCASE(right(lcl_formatted_value,6)) = "HREF='" then  
               '--------------------------------------------------------------------------
                   lcl_anchor_href_start = "Y"

                   lcl_anchor_str = lcl_anchor_str & mid(p_value,i,1)

               '-- Build the LINK HREF ---------------------------------------------------
                elseif lcl_anchor_href_start = "Y" then  
               '--------------------------------------------------------------------------
                   lcl_anchor_str = lcl_anchor_str & mid(p_value,i,1)

                   if right(lcl_formatted_value,1) <> "'" then
                      lcl_href_str = lcl_href_str & mid(p_value,i,1)
                   else
                      lcl_anchor_href_start = "N"
                   end if

               '-- Look for the end of the LINK opening tag ------------------------------
                elseif lcl_anchor_start = "Y" then  
               '--------------------------------------------------------------------------
                   lcl_anchor_str = lcl_anchor_str & mid(p_value,i,1)

                   if right(lcl_formatted_value,1) = ">" then
                      if lcl_anchor_text_start = "Y" then
                         lcl_link_text = lcl_link_text & mid(p_value,i,1)
                      else
                         lcl_anchor_text_start = "Y"
                      end if
                  
                   elseif lcl_anchor_text_start = "Y" then  '-- Build the LINK TEXT to be displayed
                     'Track the text to be displayed for the link
                      lcl_link_text = lcl_link_text & mid(p_value,i,1)
                   end if

                  'Check for the closing LINK tag (</A>)
                   if UCASE(right(lcl_formatted_value,4)) = "</A>" then
                      lcl_anchor_start      = "N"
                      lcl_anchor_text_start = "N"

                      lcl_link_text = replace(lcl_link_text,"</a>","")
                      lcl_link_text = replace(lcl_link_text,"</A>","")

                      lcl_formatted_value = replace(lcl_formatted_value,lcl_anchor_str,"")

                     'Display the link
                      oPDF.PrintText lcl_starting_x_position, iCurrentYPosition, lcl_link_text
                      lcl_link_width = oPDF.GetTextWidth(lcl_link_text)
                      oPDF.AddHyperlink iPageNum, lcl_starting_x_position, iCurrentYPosition, lcl_link_width + lcl_starting_x_position, iCurrentYPosition + iLineSpacing, lcl_href_str, 6
                     	iCurrentYPosition = iCurrentYPosition - iLineSpacing

                      lcl_anchor_str      = ""
                      lcl_href_str        = ""
                      lcl_link_text       = ""
                   end if
               '-- Image Start -------------------------------------------------------
                elseif UCASE(right(lcl_formatted_value,5)) = "<IMG " then
               '----------------------------------------------------------------------
                   lcl_img_start = "Y"

                  'Begin tracking the entire IMG str to format out of the body text later
                   lcl_img_str = right(lcl_formatted_value,5)

               '-- Image SRC Start ---------------------------------------------------
                elseif UCASE(right(lcl_formatted_value,5)) = "SRC='" OR _
                       UCASE(right(lcl_formatted_value,5)) = "SRC=""" then
               '----------------------------------------------------------------------
                   lcl_img_src_start = "Y"

                   lcl_img_str = lcl_img_str & mid(p_value,i,1)

               '-- Build the IMG SRC -------------------------------------------------
                elseif lcl_img_src_start = "Y" then
               '----------------------------------------------------------------------
                   lcl_img_str = lcl_img_str & mid(p_value,i,1)

                   if right(lcl_formatted_value,1) <> "'" OR _
                      right(lcl_formatted_value,1) <> """" then
                      lcl_img_src_str = lcl_img_src_str & mid(p_value,i,1)
                   else
                      lcl_img_src_start = "N"
                   end if

               '-- Image WIDTH Start -------------------------------------------------
                elseif UCASE(right(lcl_formatted_value,8)) = " WIDTH='" OR _
                       UCASE(right(lcl_formatted_value,8)) = " WIDTH=""" then
               '----------------------------------------------------------------------
                   lcl_img_width_start = "Y"

                   lcl_img_str = lcl_img_str & mid(p_value,i,1)

               '-- Image HEIGHT Start ------------------------------------------------
                elseif UCASE(right(lcl_formatted_value,9)) = " HEIGHT='" OR _
                       UCASE(right(lcl_formatted_value,9)) = " HEIGHT=""" then
               '----------------------------------------------------------------------
                   lcl_img_height_start = "Y"

                   lcl_img_str = lcl_img_str & mid(p_value,i,1)

               '-- Build the WIDTH ---------------------------------------------------
                elseif lcl_img_width_start = "Y" then
               '----------------------------------------------------------------------
                   lcl_img_str = lcl_img_str & mid(p_value,i,1)

                   if right(lcl_formatted_value,1) <> "'" OR _
                      right(lcl_formatted_value,1) <> """" then
                      lcl_img_width_str = lcl_img_width_str & mid(p_value,i,1)
                   else
                      lcl_img_width_start = "N"
                   end if

               '-- Build the HEIGHT --------------------------------------------------
                elseif lcl_img_height_start = "Y" then
               '----------------------------------------------------------------------
                   lcl_img_str = lcl_img_str & mid(p_value,i,1)

                   if right(lcl_formatted_value,1) <> "'" OR _
                      right(lcl_formatted_value,1) <> """" then
                      lcl_img_height_str = lcl_img_height_str & mid(p_value,i,1)
                   else
                      lcl_img_height_start = "N"
                   end if

               '-- Look for the end of the IMG tag -----------------------------------
                elseif lcl_img_start = "Y" then
               '----------------------------------------------------------------------
                   lcl_img_str = lcl_img_str & mid(p_value,i,1)

                  'Check for the closing IMG tag (>)
                   if UCASE(right(lcl_formatted_value,1)) = ">" then
                      lcl_img_start = "N"

                      lcl_formatted_value = replace(lcl_formatted_value,lcl_img_str,"")

                     'Convert the Width and Height to numbers
                      if lcl_img_width_str <> "" then
                         lcl_width = CLng(lcl_img_width_str)
                      else
                         lcl_width = 0
                      end if

                      if lcl_img_height_str <> "" then
                         lcl_height = CLng(lcl_img_height_str)
                      else
                         lcl_height = 0
                      end if

                     'If a width has been entered then we can check for an alignment.
                      if lcl_width > 0 then
                         if session("text_align_center") = "Y" then
                            lcl_starting_x_position = (iPageWidth/2) - (lcl_width/2)
                         elseif session("text_align_right") = "Y" then
                            lcl_starting_x_position = iPageWidth - lcl_width
                         end if
                      else
                         lcl_starting_x_position = sStartingXPosition
                      end if

                     'Display the IMG
                      oDocument = oPDF.PrintImage(server.mappath("images/" & lcl_img_src_str), lcl_starting_x_position, iCurrentYPosition - iLineSpacing, lcl_width, lcl_height, False)
                     	iCurrentYPosition = iCurrentYPosition - iLineSpacing

                      lcl_formatted_value = replace(lcl_formatted_value,lcl_img_str,"")

                      lcl_img_str        = ""
                      lcl_img_src_str    = ""
                      lcl_img_width_str  = ""
                      lcl_img_height_str = ""
                   end if

               '-- Check for CENTER alignment --------------------------------------------
'                elseif UCASE(right(lcl_formatted_value,18)) = "<P ALIGN='CENTER'>" then  
               '--------------------------------------------------------------------------

               '--------------------------------------------------------------------------
                else
               '--------------------------------------------------------------------------
                   lcl_starting_x_position = lcl_starting_x_position
               '--------------------------------------------------------------------------
                end if
               '--------------------------------------------------------------------------

             else
                lcl_formatted_value = ""
             end if
         next

         if lcl_formatted_value <> "" AND NOT isnull(lcl_formatted_value) then
'            oPDF.SetFont iFontFace, iFont
            if session("bold_start") = "Y" AND session("italic_start") = "Y" then
              	oPDF.SetFont "Times-BoldItalic", iFont
            elseif session("bold_start") = "Y" AND session("italic_start") <> "Y" then
              	oPDF.SetFont "Times-Bold", iFont
            elseif session("bold_start") <> "Y" AND session("italic_start") = "Y" then
              	oPDF.SetFont "Times-Italic", iFont
            else
               oPDF.SetFont iFontFace, iFont
            end if
            		 oPDF.PrintText lcl_starting_x_position, iCurrentYPosition, lcl_formatted_value
            end if

          		iCurrentYPosition = iCurrentYPosition - iLineSpacing

      end if
end sub

'--------------------------------------------------------------------------------------------------
function buildStreetAddress(sStreetNumber, sPrefix, sStreetName, sSuffix, sDirection)
  lcl_street_name = ""

  if trim(sStreetNumber) <> "" then
     lcl_street_name = trim(sStreetNumber)
  end if

  if trim(sPrefix) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sPrefix)
     else
        lcl_street_name = trim(sPrefix)
     end if
  end if

  if trim(sStreetName) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sStreetName)
     else
        lcl_street_name = trim(sStreetName)
     end if
  end if

  if trim(sSuffix) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sSuffix)
     else
        lcl_street_name = trim(sSuffix)
     end if
  end if

  if trim(sDirection) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sDirection)
     else
        lcl_street_name = trim(sDirection)
     end if
  end if

  buildStreetAddress = trim(lcl_street_name)

end function
%>
