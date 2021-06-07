<%
response.buffer = True 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: DISPLAY_WAIVER.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/6/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' DESCRIPTION:  CREATES PDF FILE CONTAINING WAIVER INFORMATION AND RESERVATION DETAIL INFORMATION
' FOR THE CONSUMER TO PRINT AND SIGN.
'
' MODIFICATION HISTORY
' 1.0  02/06/06 John Stullenberger - Inital Version
' 2.0  03/02/06	John Stullenberger - Modified to add external PDFS
' 3.0  11/21/07 David Boyer - Combined to single line: 1. First/Last Name 2. City/State/Zip
'                           - Also removed "Form Information" and Internal questions/answers
'                             for those that had no answers.
' 4.0  03/22/08 David Boyer - Modified IP locations due to move to Time Warner.
' 4.1  08/17/09 David Boyer - Added check to show/hide Request Activity Log
'
' LINK: HTTPS://SECURE.ECLINK.COM/EGOVLINK/DISPLAY_WAIVER.ASP?MASK=X9
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

'Initialize and Declare Variables
 Dim sPrimaryTemplate,sSecondaryTemplate
 Dim oPDF,sSystem,sDB,sRequestTitle
 sPrimaryTemplate   = server.mappath("templates/wo_primary_page.pdf")
 sSecondaryTemplate = server.mappath("templates/wo_last_page.pdf")
 sSystem            = request("sys")

 if request("hideActLog") <> "" then
    sHideActLog = ucase(request("hideActLog"))
 else
    sHideActLog = "N"
 end if

 select Case sSystem 
	  case "DEV"
      'sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
		    sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink300; UID=egovsa; PWD=Egov_6814;"
 	 case "QA"
      'sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink400_QA_Test; UID=egovsa; PWD=egov_4303;"
		    sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink400_QA_Test; UID=egovsa; PWD=egov_4303;"
 	 case "LIVE"
	 	   'sDB = "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=sa; PWD=;"
	 	   sDB = "Driver={SQL Server}; Server=CO-SQL-03; Database=egovlink300; UID=egovsa; PWD=Egov_6814;"
 	 case Else
      'sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
		    'sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
		    sDB = "Driver={SQL Server}; Server=L3SQL2; Database=egovlink300; UID=egovsa; PWD=Egov_6814;"
 end select 

'Create PDF Object
 set oPDF = Server.CreateObject("APToolkit.Object")
 oDocument = oPDF.OpenOutputFile("MEMORY") 'CREATE THE OUTPUT INMEMORY

'Build PDF Document
 Dim iFont,iLineSpacing,iSpacer,iFontFace,sStaringXPosition,sStartingYPosition,iCurrentYPosition, sSubmitName,datSubmitDate 
 Dim iPageWidth,iPageHeight
 iFont                 = 10
 iLineSpacing          = 13
 iFontFace             = "arial" ' NEED MONOSPACE FONT TO MAINTAIN SPACING IN TEXT FILE
 sStartingXPosition    = 25 ' DEFAULT TEXT X POSITION
 sStartingYPosition    = 725 ' DEFAULT TEXT Y POSITION
 iPageWidth            = 500
 iPageHeight           = 792
 oPDF.OutputPageWidth  = iPageWidth ' 8.5 inches
 oPDF.OutputPageHeight = iPageHeight ' 11 inches

'Add Text to Page
 r = oPDF.AddLogo(sPrimaryTemplate,0)
 oPDF.PrintLogo
 Call GetRequestInformation(request("iRequestID"),oPDF)

'---------------------------------------------------------------------------
	sSQL = "SELECT *, (FirstName + ' ' + LastName) as EmployeeSubmitName "
 sSQL = sSQL & " FROM egov_actionline_requests "
	sSQL = sSQL &      " LEFT OUTER JOIN users ON egov_actionline_requests.employeesubmitid = users.userid "
	sSQL = sSQL &      " LEFT OUTER JOIN egov_users ON egov_actionline_requests.userid = egov_users.userid "
	sSQL = sSQL &      " LEFT OUTER JOIN egov_action_request_forms AS F ON egov_actionline_requests.category_id = F.action_form_id "
	sSQL = sSQL & " WHERE action_autoid='" & request("iRequestID") & "'"

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSQL, sDB, 3, 1
	
'Build Work Order
	if not rs.eof then
		
		 'Initial Values
  		datSubmitDate = rs("submit_date")
  		isubmitid     = rs("employeesubmitid")
		  sRequestTitle = Ucase(trim(rs("category_title")))

  	'Get Employee or Citizen that submitted the request
  		if isubmitid < 0 OR IsNull(isubmitid) OR isubmitid = "" then
 	  		'Use CITIZEN name as submitter
  			  sSubmitName = rs("userfname") & " " & rs("userlname") & " (Citizen)"
  		else
		  	 'Use EMPLOYEE name as submitter
			    sSubmitName = rs("EmployeeSubmitName") & " (Admin Employee)"
  		end if
 else
  		datSubmitDate = ""
  		sSubmitName   = ""
		  sRequestTitle = ""
 end if

 rs.close
 set rs = nothing

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

'Stream PDF to browser
 response.expires = 0
 response.Clear
 response.ContentType = "application/pdf"
 response.AddHeader "Content-Type", "application/pdf"
 response.AddHeader "Content-Disposition", "inline;filename=FORMS.PDF"
 response.BinaryWrite oDocument  

'Destroy Objects
 set oPDF      = nothing
 set oDocument = nothing

'------------------------------------------------------------------------------
sub GetRequestInformation(iRequestID,oPDF)
	
	sSQL = "SELECT *, (FirstName + ' ' + LastName) as EmployeeSubmitName, F.DeptID "
 sSQL = sSQL & " FROM egov_actionline_requests "
	sSQL = sSQL & " LEFT OUTER JOIN users ON egov_actionline_requests.employeesubmitid = users.userid "
	sSQL = sSQL & " LEFT OUTER JOIN egov_users ON egov_actionline_requests.userid = egov_users.userid "
	sSQL = sSQL & " LEFT OUTER JOIN egov_action_request_forms AS F ON egov_actionline_requests.category_id = F.action_form_id "
	sSQL = sSQL & " WHERE action_autoid='" & iRequestID & "'"

	set oRequest = Server.CreateObject("ADODB.Recordset")
	oRequest.Open sSQL, sDB, 3, 1
	
	' BUILD WORK ORDER
	If Not oRequest.EOF Then
		
		' INITIAL VALUES
		sTitle           = oRequest("category_title")
		sStatus          = oRequest("status")
		datSubmitDate    = oRequest("submit_date")
		sComment         = oRequest("comment")
		sTheUserid       = oRequest("userid")
		iemployeeid      = oRequest("assignedemployeeid")
		isubmitid        = oRequest("employeesubmitid")
		icontactmethodid = oRequest("contactmethodid")
		sRequestTitle    = Ucase(trim(oRequest("category_title")))

		'GET EMPLOYEE OR CITIZEN THAT SUBMITTED THE REQUEST
		If isubmitid < 0 OR IsNull(isubmitid) OR isubmitid = "" Then
 			'USER CITIZEN NAME AS SUBMITTER
			  sSubmitName = oRequest("userfname") & " " & oRequest("userlname") & " (Citizen)"
		Else
			 'USE EMPLOYEE NAME AS SUBMITTER
			  sSubmitName =  oRequest("EmployeeSubmitName") & " (Admin Employee)"
		End If

		' HEADER INFORMATION 
		oPDF.GreyBar sStartingXPosition - 5, sStartingYPosition - 1, 550, iFont + 5, 0
		oPDF.SetTextColor 255,255,255,255 ' CHANGE TEXT COLOR TO WHITE
		oPDF.SetFont "arial", iFont + 2
		oPDF.PrintText sStartingXPosition,sStartingYPosition + 2,  sRequestTitle
		oPDF.SetFont iFontFace, iFont
		oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK
		iCurrentYPosition = sStartingYPosition - iLineSpacing
		oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Tracking Number: " & iRequestID  & replace(FormatDateTime(cdate(datSubmitDate),4),":","")
		iCurrentYPosition = iCurrentYPosition - iLineSpacing
		oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Date Time Received: " & datSubmitDate
		iCurrentYPosition = iCurrentYPosition - iLineSpacing
		oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Created By: " & sSubmitName
		iCurrentYPosition = iCurrentYPosition - (iLineSpacing * 1.5)		

		' CONTACT INFORMATION
		oPDF.GreyBar sStartingXPosition - 5, iCurrentYPosition - 1, 550, iFont + 5, 0
		oPDF.SetTextColor 255,255,255,255 ' CHANGE TEXT COLOR TO WHITE
		oPDF.SetFont "arial", iFont + 2
		oPDF.PrintText sStartingXPosition,iCurrentYPosition + 2,  "CONTACT INFORMATION"
		oPDF.SetFont iFontFace, iFont
		oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK
'		iCurrentYPosition = iCurrentYPosition - iLineSpacing

	 if oRequest("userfname") <> "" OR oRequest("userlname") <> "" then
   		iCurrentYPosition = iCurrentYPosition - iLineSpacing
   		oPDF.PrintText sStartingXPosition,iCurrentYPosition, oRequest("userfname") & " " & oRequest("userlname")
  end if

	 if trim(oRequest("userbusinessname")) <> "" then
   		iCurrentYPosition = iCurrentYPosition - iLineSpacing
   		oPDF.PrintText sStartingXPosition,iCurrentYPosition, trim(oRequest("userbusinessname"))
  end if

  if trim(oRequest("useremail")) <> "" then
   		iCurrentYPosition = iCurrentYPosition - iLineSpacing
   		oPDF.PrintText sStartingXPosition,iCurrentYPosition, trim(oRequest("useremail"))
  end if

  if trim(oRequest("userhomephone")) <> "" then
   		iCurrentYPosition = iCurrentYPosition - iLineSpacing
   		oPDF.PrintText sStartingXPosition,iCurrentYPosition, FormatPhone( oRequest("userhomephone"))
  end if

	 if trim(oRequest("userfax")) <> "" then
	 	  iCurrentYPosition = iCurrentYPosition - iLineSpacing
   		oPDF.PrintText sStartingXPosition,iCurrentYPosition, FormatPhone( oRequest("userfax"))
  end if

	 if trim(oRequest("useraddress")) <> "" then
   		iCurrentYPosition = iCurrentYPosition - iLineSpacing
   	 oPDF.PrintText sStartingXPosition,iCurrentYPosition, trim(oRequest("useraddress"))
  end if

 'User City/State/Zip
  lcl_contact_csz   = ""
  lcl_contact_city  = trim(oRequest("usercity"))
  lcl_contact_state = trim(oRequest("userstate"))
  lcl_contact_zip   = trim(oRequest("userzip"))

 'Build the city/state/zip display variable
 'Evaluate the City
  if lcl_contact_city <> "" then
     lcl_contact_csz = lcl_contact_city
  end if

 'Evaluate the State
  if lcl_contact_state <> "" then
     if lcl_contact_csz <> "" then
        lcl_contact_csz = lcl_contact_csz & "/" & lcl_contact_state
     else
        lcl_contact_csz = lcl_contact_state
     end if
  end if

 'Evaluate the Zip
  if lcl_contact_zip <> "" then
     if lcl_contact_csz <> "" then
        lcl_contact_csz = lcl_contact_csz & "/" & lcl_contact_zip
     else
        lcl_contact_csz = lcl_contact_zip
     end if
  end if

 'If the city/state/zip display variable is not null then display it.
  if lcl_contact_csz <> "" then
   		iCurrentYPosition = iCurrentYPosition - iLineSpacing
'   		oPDF.PrintText sStartingXPosition,iCurrentYPosition, trim(oRequest("usercity")) & "/" &  trim(oRequest("userstate")) & "/" &  trim(oRequest("userzip"))
   		oPDF.PrintText sStartingXPosition,iCurrentYPosition, lcl_contact_csz
  end if

      'iCurrentYPosition = iCurrentYPosition - iLineSpacing
      'oPDF.PrintText sStartingXPosition,iCurrentYPosition, "Country: "        & trim(oRequest("usercountry"))
		iCurrentYPosition = iCurrentYPosition - (iLineSpacing * 1.5)		

	'ISSUE LOCATION
 'Check to see if the form has the feature turned on.  If not then do not display this section.
  If oRequest("action_form_display_issue") = True Then
   		Call SubDrawIssueLocationInformation(iRequestID)
  end if

	'FORM INFORMATION
		oPDF.GreyBar sStartingXPosition - 5, iCurrentYPosition - 1, 550, iFont + 5, 0
		oPDF.SetTextColor 255,255,255,255 ' CHANGE TEXT COLOR TO WHITE
		oPDF.SetFont "arial", iFont + 2
		oPDF.PrintText sStartingXPosition,iCurrentYPosition + 2,  "REQUEST DETAILS"
		oPDF.SetFont iFontFace, iFont
		oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK
		iCurrentYPosition = iCurrentYPosition - iLineSpacing
'		sComment   = replace(sComment,"<p>","",1,-1,1)
'		sComment   = replace(sComment,"</p>","",1,-1,1)
'		sComment   = replace(sComment,"</br>",vbcrlf,1,-1,1)
'		sComment   = replace(sComment,"<br>",vbcrlf,1,-1,1)
'		sComment   = replace(sComment,"</b>","",1,-1,1)
'		sComment   = replace(sComment,"<b>","",1,-1,1)
'  sComment   = replace(UCASE(sComment),"DEFAULT_NOVALUE","",1,-1,1)

		sComment   = replace(sComment,"<p>","")
		sComment   = replace(sComment,"</p>","")
		sComment   = replace(sComment,"<P>","")
		sComment   = replace(sComment,"</P>","")

		sComment   = replace(sComment,"</br>","")
		sComment   = replace(sComment,"<br>","")
		sComment   = replace(sComment,"</BR>","")
		sComment   = replace(sComment,"<BR>","")

		sComment   = replace(sComment,"</b>","")
		sComment   = replace(sComment,"<b>","")
		sComment   = replace(sComment,"</B>","")
		sComment   = replace(sComment,"<B>","")

		sComment   = replace(sComment,"<u>","")
		sComment   = replace(sComment,"</u>","")
		sComment   = replace(sComment,"<U>","")
		sComment   = replace(sComment,"</U>","")
		sComment   = replace(sComment,"&quot;","""")

		sComment   = replace(sComment,chr(13),"")

		'sComment   = replace(UCASE(sComment),"DEFAULT_NOVALUE","")
		sComment   = replace(sComment,"default_novalue","")
		sComment   = replace(sComment,"DEFAULT_NOVALUE","")

		arrComment = split(sComment,chr(10))
		iNumLines  = UBOUND(arrComment)
		Dim blankLineCount, initialLine
		blankLineCount = 0
		initialLine = True 
		For iCommentLines = 0 to iNumLines
   			'fnNewLineCheck Trim(arrComment(iCommentLines)),"REQUEST DETAILS"
			If initialLine Then 
				If Trim(arrComment(iCommentLines)) = "" Then
					okToPrint = False
				Else
					okToPrint = True
				End If 
				initialLine = False 
			Else
				If Trim(arrComment(iCommentLines)) <> "" Then 
					okToPrint = True 
					blankLineCount = 0
				Else
					blankLineCount = blankLineCount + 1
					If blankLineCount < 2 Then
						okToPrint = True
					Else
						okToPrint = False 
					End If 
				End If
			End If 

			If okToPrint Then 
				fnNewLineCheck Trim(arrComment(iCommentLines)),"REQUEST DETAILS"
			End If 
		Next

 'BEGIN: Internal Fields Only -------------------------------------------------

	'BEGIN: Request Activity Log -------------------------------------------------
  if sHideActLog <> "Y" then
   		oPDF.GreyBar sStartingXPosition - 5, iCurrentYPosition - 1, 550, iFont + 5, 0
   		oPDF.SetTextColor 255,255,255,255  'Change text color to WHITE
   		oPDF.SetFont "arial", iFont + 2
   		oPDF.PrintText sStartingXPosition,iCurrentYPosition + 2,  "REQUEST ACTIVITY"
   		oPDF.SetFont iFontFace, iFont
   		oPDF.SetTextColor 0,0,0,0  'Change text color to BLACK
   		iCurrentYPosition = iCurrentYPosition - iLineSpacing

   		Call Display_Request_Details(iRequestID)
   		iCurrentYPosition = iCurrentYPosition - (iLineSpacing * 1.5)
  end if
 'END: Request Activity Log ---------------------------------------------------

	End If

End Sub

'------------------------------------------------------------------------------
Function FormatPhone( Number )
  If Len(Number) = 10 Then
     FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
  Else
     FormatPhone = Number
  End If
End Function

'------------------------------------------------------------------------------
Sub Display_Request_Details(iID)

'	sSQL = "SELECT egov_action_responses.*, s.status_name AS SUB_STATUS_NAME "
	sSQL = "SELECT r.*, eu.*, u.*, status_name AS SUB_STATUS_NAME "
 sSQL = sSQL & " FROM egov_action_responses r "
 sSQL = sSQL & " LEFT OUTER JOIN egov_users eu ON r.action_userid = eu.userid "
 sSQL = sSQL & " LEFT OUTER JOIN users u ON r.action_userid = u.userid "
 sSQL = sSQL & " LEFT OUTER JOIN egov_actionline_requests_statuses s ON s.action_status_id = r.action_sub_status_id "
 sSQL = sSQL & " WHERE r.action_autoid = " & iID
 sSQL = sSQL & " ORDER BY r.action_editdate DESC"

	Set oCommentList = Server.CreateObject("ADODB.Recordset")
	oCommentList.Open sSQL, sDB , 3, 1

	If NOT oCommentList.EOF Then
		Do While NOT oCommentList.EOF 
     if oCommentList("SUB_STATUS_NAME") <> "" then
        lcl_sub_status =  " (" & oCommentList("SUB_STATUS_NAME") & ")"
     else
        lcl_sub_status = ""
     end if

			' DISPLAY ADMIN COMMENT CREATOR NAME
			fnNewLineCheck oCommentList("action_editdate") & " -- " & oCommentList("firstname") & " " & oCommentList("lastname") & " - " & UCASE(oCommentList("action_status")) & lcl_sub_status,"REQUEST ACTIVITY"

			' DISPLAY EXTERNAL COMMENTS
			If oCommentList("action_externalcomment") <> "" Then
  				sNote = ""
		  		sNote = replace( oCommentList("action_externalcomment"),"",vbcrlf,1,-1,1)
				  sNote = replace(sNote,vbcrlf," ")
  				sNote = fnCleanData(sNote)
      sNote = replace(sNote,"default_novalue","")
				  fnNewLineCheck "-----Note to Citizen: " & sNote,"REQUEST ACTIVITY"  
			End If

			' DISPLAY CITIZEN COMMENT CREATOR NAME
			If oCommentList("action_citizen") <> "" Then
  				fnNewLineCheck oCommentList("action_editdate") & " -- " & oCommentList("userfname")  & " " & oCommentList("userlname") & " : " & oCommentList("action_citizen"),"REQUEST ACTIVITY"
			End If

			' DISPLAY INTERNAL NOTE
			If oCommentList("action_internalcomment") <> "" Then
  				sNote = ""
		  		sNote = replace(oCommentList("action_internalcomment"),"",vbcrlf,1,-1,1)
 				 sNote = replace(sNote,vbcrlf," ")
	  			sNote = fnCleanData(sNote)
      sNote = replace(sNote,"default_novalue","")
  				fnNewLineCheck "-----Internal Note: " & sNote ,"REQUEST ACTIVITY"
			End If
			
			oCommentList.MoveNext

		Loop

			' DISPLAY SUBMIT DATE TIME AND USER
			fnNewLineCheck datSubmitDate & " -- " & sSubmitName & " - " & UCASE("SUBMITTED") & lcl_sub_status,"REQUEST ACTIVITY"

	Else
		
		' NO ACTIVITY FOR THIS REQUEST
		fnNewLineCheck  datSubmitDate & " -- " & "No activity Reported.","REQUEST ACTIVITY"


		' DISPLAY SUBMIT DATE TIME AND USER
		fnNewLineCheck datSubmitDate & " -- " & sSubmitName & " - " & UCASE("SUBMITTED") & lcl_sub_status,"REQUEST ACTIVITY" 
		
	End If

 displayFooter()

End Sub

'--------------------------------------------------------------------------------------------------
 sub displayFooter()

   iCurrentYPosition = iCurrentYPosition - (iLineSpacing * 2)
   lcl_footer_line = "--------------------------------"
   lcl_footer_line = lcl_footer_line & "                                                                  "
   lcl_footer_line = lcl_footer_line & "---------------------------------------------------"
   lcl_footer_line = lcl_footer_line & "                                             "
   lcl_footer_line = lcl_footer_line & "Date"
   lcl_footer_line = lcl_footer_line & "                                                                                             "
   lcl_footer_line = lcl_footer_line & "Authorized Signature"
   fnNewLineCheck lcl_footer_line,""

   fnNewLineCheck "Comments: ",""


 end sub

'--------------------------------------------------------------------------------------------------
'SUB SUBNEWPAGECHECK(ICURRENTYPOSITION,SHEADING)
'--------------------------------------------------------------------------------------------------
Sub subNewPageCheck(iCurrentYPosition,sHeading)

	' CHECK TO SEE IF WE NEED TO START A NEW PAGE
	If Clng(iCurrentYPosition) <= 36 Then
		oPDF.NewPage
		iCurrentYPosition = sStartingYPosition
		oPDF.AddLogo sPrimaryTemplate,0
		oPDF.PrintLogo
		oPDF.SetFont "arial", iFont 

		' HEADER INFORMATION 
		oPDF.GreyBar sStartingXPosition - 5, sStartingYPosition - 1, 550, iFont + 5, 0
		oPDF.SetTextColor 255,255,255,255 ' CHANGE TEXT COLOR TO WHITE
		oPDF.SetFont "arial", iFont + 2
		oPDF.PrintText sStartingXPosition,sStartingYPosition + 2,  sRequestTitle
		oPDF.SetFont iFontFace, iFont
		oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK
		iCurrentYPosition = sStartingYPosition - iLineSpacing
		oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Tracking Number: " & request("iRequestID") & replace(FormatDateTime(cdate(datSubmitDate),4),":","")
		iCurrentYPosition = iCurrentYPosition - iLineSpacing
		oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Date Time Received: " & datSubmitDate
		iCurrentYPosition = iCurrentYPosition - iLineSpacing
		oPDF.PrintText sStartingXPosition,iCurrentYPosition,  "Created By: " & sSubmitName
		iCurrentYPosition = iCurrentYPosition - (iLineSpacing * 2)	

		' CURRENT HEADER INFORMATION
		oPDF.GreyBar sStartingXPosition - 5, iCurrentYPosition - 1, 550, iFont + 5, 0
		oPDF.SetTextColor 255,255,255,255 ' CHANGE TEXT COLOR TO WHITE
		oPDF.SetFont "arial", iFont + 2
		oPDF.PrintText sStartingXPosition,iCurrentYPosition + 2,  sHeading & " (continued)"
		oPDF.SetFont iFontFace, iFont
		oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK
		iCurrentYPosition = iCurrentYPosition - iLineSpacing
	End If

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION fnNewLineCheck(STEXT,SHEADING)
'--------------------------------------------------------------------------------------------------
Function fnNewLineCheck(sText,sHeading)

	 iLineWidth = oPDF.GetTextWidth(sText)

 	If iLineWidth > iPageWidth Then
	  	'CREATE WRAP LINE
   		sWrapLine = ""
   		arrWords  = split(replace(sText,"["," "),chr(32))
   		For iWordCount = 0 To UBOUND(arrWords)
			      If oPDF.GetTextWidth(sWrapLine) < iPageWidth Then
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
   		Call subNewPageCheck(iCurrentYPosition,sHeading)
   		oPDF.PrintText sStartingXPosition,iCurrentYPosition, trim(sWrapLine)
   		iCurrentYPosition = iCurrentYPosition - iLineSpacing

  		'CONTINUE WITH REST OF LINE WRAPPING AS NEEDED
   		Call subNewPageCheck(iCurrentYPosition,sHeading)
   		fnNewLineCheck sRestofLine,sHeading
  Else
  		'WRITE LINE AS IS - NOT WRAPPING REQUIRED
   		Call subNewPageCheck(iCurrentYPosition,sHeading)
'   		oPDF.PrintText sStartingXPosition,iCurrentYPosition, trim(replace(sText,",","[,]"))

'       		oPDF.PrintText sStartingXPosition,iCurrentYPosition, "[" & sText & "] - [[" & lcl_label & "]] - [[[" & lcl_value & "]]]"
'         iCurrentYPosition = iCurrentYPosition - iLineSpacing

      lcl_br_position = instr(sText,"[")
      lcl_label = replace(left(sText,lcl_br_position),"[","")
      lcl_value = mid(sText,lcl_br_position+1)

'       		oPDF.PrintText sStartingXPosition,iCurrentYPosition, "[[" & lcl_label & " " & lcl_value & "("& sText &")]]"
'         iCurrentYPosition = iCurrentYPosition - (iLineSpacing * 2)	

     'Cycle through the answers if a comma exists and trim off the paragraph symbol that appears when the PDF is printed.
      lcl_display_list = replace(lcl_value,chr(10),"")

'      if instr(lcl_value,",") > 0 then
'         lcl_main_list = lcl_value & ","
'      else
'         lcl_main_list = lcl_value
'      end if

'      lcl_display_list = ""
'      if instr(lcl_main_list,",") > 0 then
'         do while right(lcl_main_list,1) = ","
'            lcl_comma_position = instr(lcl_main_list,",")
'            lcl_new_value = left(lcl_main_list,lcl_comma_position-2)
'            lcl_main_list = mid(lcl_main_list,lcl_comma_position+1)

'            if lcl_display_list = "" then
'               lcl_display_list = trim(lcl_new_value)
'            else
'               lcl_display_list = lcl_display_list & ", " & trim(lcl_new_value)
'            end if
'         loop
'      else
'         lcl_display_list = lcl_main_list
'   			End If

'      if lcl_display_list <> "" then
'      if len(sText) > 1 then
      if len(lcl_value) > 1 then
         if instr(lcl_label,":") = 0 AND instr(lcl_label,".") = 0 AND instr(lcl_label,"?") = 0 AND lcl_label <> "" then
            lcl_label = lcl_label & ":"
         end if

         if len(lcl_label) > 1 then
            lcl_label = lcl_label & " "
         end if

      		 oPDF.PrintText sStartingXPosition,iCurrentYPosition, lcl_label & lcl_display_list
         iCurrentYPosition = iCurrentYPosition - (iLineSpacing * 1.5)	
      end if
  End If

End Function

'--------------------------------------------------------------------------------------------------
' SUB SUBDRAWISSUELOCATIONINFORMATION(IREQUESTID)
'--------------------------------------------------------------------------------------------------
Sub SubDrawIssueLocationInformation(iRequestID)
  lcl_issue_location = ""
  lcl_issue_csz      = ""

 	sSQL = "SELECT * FROM egov_action_response_issue_location WHERE actionrequestresponseid='" & iRequestID & "'"

 	Set oIssueLocation = Server.CreateObject("ADODB.Recordset")
 	oIssueLocation.Open sSQL, sDB , 3, 1
	
 	If NOT oIssueLocation.EOF Then

	   	sHasValue = trim(oIssueLocation("streetnumber") & oIssueLocation("streetprefix") & oIssueLocation("streetaddress") & oIssueLocation("streetsuffix") & oIssueLocation("streetdirection") & oIssueLocation("city") &  oIssueLocation("state") & oIssueLocation("zip") & oIssueLocation("comments"))

   		If sHasValue <>  "" Then
   			 'DISPLAY HEADER
     			oPDF.GreyBar sStartingXPosition - 5, iCurrentYPosition - 1, 550, iFont + 5, 0
			     oPDF.SetTextColor 255,255,255,255 ' CHANGE TEXT COLOR TO WHITE
     			oPDF.SetFont "arial", iFont + 2
     			oPDF.PrintText sStartingXPosition,iCurrentYPosition + 2, "ISSUE LOCATION"
     			oPDF.SetFont iFontFace, iFont
     			oPDF.SetTextColor 0,0,0,0 ' CHANGE TEXT COLOR TO BLACK
     			iCurrentYPosition = iCurrentYPosition - iLineSpacing
			
       'Build the street number and address
        lcl_issue_location = buildStreetAddress(oIssueLocation("streetnumber"), oIssueLocation("streetprefix"), oIssueLocation("streetaddress"), oIssueLocation("streetsuffix"), oIssueLocation("streetdirection"))

'        if oIssueLocation("streetnumber") <> "" then
'           lcl_issue_location = oIssueLocation("streetnumber")
'        end if

'        if oIssueLocation("streetaddress") <> "" then
'           if lcl_issue_location <> "" then
'              lcl_issue_location = lcl_issue_location & " " & oIssueLocation("streetaddress")
'           else
'              lcl_issue_location = oIssueLocation("streetaddress")
'           end if
'        end if

       'Build the city, state, and zip
        if oIssueLocation("city") <> "" then
           lcl_issue_csz = oIssueLocation("city")
        end if

        if oIssueLocation("state") <> "" then
           if lcl_issue_csz <> "" then
              lcl_issue_csz = lcl_issue_csz & " / " & oIssueLocation("state")
           else
              lcl_issue_csz = oIssueLocation("state")
           end if
        end if

        if oIssueLocation("zip") <> "" then
           if lcl_issue_csz <> "" then
              lcl_issue_csz = lcl_issue_csz & " / " & oIssueLocation("zip")
           else
              lcl_issue_csz = oIssueLocation("zip")
           end if
        end if

     		'WRITE DETAILS
        if lcl_issue_location <> "" then
   		     	fnNewLineCheck	lcl_issue_location, "ISSUE LOCATION"
        end if

        if oIssueLocation("streetunit") <> "" then
           fnNewLineCheck "Unit: " & oIssueLocation("streetunit"), "ISSUE LOCATION"
        end if

        if lcl_issue_csz <> "" then
   		     	fnNewLineCheck lcl_issue_csz, "ISSUE LOCATION"
        end if

        if oIssueLocation("comments") <> "" then
   		     	fnNewLineCheck "Comments: " & oIssueLocation("comments") , "ISSUE LOCATION"
        end if

     		'ADD SEPARATION
'     			iCurrentYPosition = iCurrentYPosition - (iLineSpacing * 2)
     			iCurrentYPosition = iCurrentYPosition - iLineSpacing
   		End If
  End If
End Sub

'--------------------------------------------------------------------------------------------------
' FUNCTION FNCLEANDATA(STEXT)
'--------------------------------------------------------------------------------------------------
Function fnCleanData(sText)
 	sReturnValue = sText

 	For i=1 to 31
	     sReturnValue = replace(sReturnValue,chr(i)," ")
  Next

	'REMOVE HTML BREAKS
 	sReturnValue = replace(sReturnValue,"<BR>",", ")
  sReturnValue = replace(sReturnValue,"&quot;","""")

	 fnCleanData = sReturnValue
End Function

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

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"

 	set oDTB = Server.CreateObject("ADODB.Recordset")
	 oDTB.Open sSQL, sDB, 3, 1

end sub

%>
