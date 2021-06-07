<!DOCTYPE HTML>
<!-- #include file="includes/common.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: action_request_lookup.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Action Line Search Results.
'
' MODIFICATION HISTORY
' 1.0 ??/??/??  ??? - Initial Version
' 2.0 01/22/08  David Boyer - Added "isFeatureOffline" check to screen.
' 2.1 01/09/09		David Boyer - Added "View PDF" button
' 2.2 02/17/09  David Boyer - Added "Edit Display" for all "Action Line" display texts
' 2.3 06/17/09  David Boyer - Added "e=Y" to (action_respond.asp) urls in emails.
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'To help prevent hacks
 If Not IsNumeric(request("REQUEST_ID")) Or request("REQUEST_ID") = "3273201600" Or request("frmsubjecttext") <> "" Then 
    response.redirect "action.asp"
 End If 

'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 Dim sError, sActionDefaultEmail

 lcl_hidden = "hidden"  'Show/Hides all hidden fields.  HIDDEN = Hide, TEXT = Show

 datOrgDateTime = ConvertDateTimetoTimeZone(iorgid)

'If users supplied comments then update them
 if Request.ServerVariables("REQUEST_METHOD") = "POST" AND request("sMsg") <> "" then
   	sCitizenMsg = request("sMsg")
   	iFormID     = CLng(request("iFormID"))
   	iUserID     = CLng(request("iUSerID"))
   	sStatus     = request("sStatus")
   	iOrgID      = iorgid 

	If iFormID <> 327320 Then 
	   	AddCommentTaskComment sStatus,sCitizenMsg,iFormID,iUserID,iOrgID, datOrgDateTime

  		'Email the comment to those who get the message. - Steve Loar - 4/10/2006
	   	EmailComment sCitizenMsg, request("iCategoryId"), iFormID, request("REQUEST_ID"), iUserID, datOrgDateTime, sOrgName
	End If 
 end if

'Check for org features
 lcl_orghasfeature_requestmergeforms                    = orghasfeature(iorgid, "requestmergeforms")
 lcl_orghasfeature_action_line_substatus                = orghasfeature(iorgid, "action_line_substatus")
 lcl_orghasfeature_hide_email_actionline                = orghasfeature(iorgid, "hide email actionline")
 lcl_orghasfeature_actionline_display_duedate           = orghasfeature(iorgid, "actionline_display_duedate")
 lcl_orghasfeature_issue_location                       = orghasfeature(iorgid, "issue location")
 lcl_orghasfeature_actionline_formcreator_mobileoptions = orghasfeature(iorgid, "actionline_formcreator_mobileoptions")

'Set the "Action Line Request" label
 lcl_actionlinelabel = "Action Line Request"

 if OrgHasDisplay(iOrgID,"actionlinelabel_publicrequestlookup") then
    lcl_actionlinelabel = GetOrgDisplayWithId(iOrgID,getDisplayID("actionlinelabel_publicrequestlookup"),False)
 end if
 
  if OrgHasDisplay(iOrgID,"actionline_default_citizen_message") then
    lcl_defaultComment = GetOrgDisplayWithId(iOrgID,getDisplayID("actionline_default_citizen_message"),False)
  else
    lcl_defaultComment = "Your request was reviewed and/or status was updated."
 end if

%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

	<title>E-Gov Services <%=sOrgName%></title>

	<link rel="stylesheet" href="css/styles.css" />
	<link rel="stylesheet" href="global.css" />
	<link rel="stylesheet" href="css/style_<%=iorgid%>.css" />

<style>

.accountmenu
{
   width: 38px;
   height: 25px;
}

.mobileImg
{
   width: 100px;
   height: 100px;
   border: 1pt solid #000000;
   border-radius: 6px;
   margin: 0px 10px 10px 10px;
}

.topbanner td
{
   height: inherit !important;
}

.fieldset
{
   border: 1pt solid #808080;
   border-radius: 6px;
   background-color: #e0e0e0;
   margin: 10px;
   padding: 10px;
}

.fieldset legend
{
    background-color: #FFFFFF;
    border: 1pt solid #808080;
    border-radius: 6px;
    color: #800000;
    font-size: 1.125em;
    padding: 4px 8px;
}

#sMsg
{
   margin-bottom: 10px;
}

.mobilePicImg
{
   display: inline;
   float: left;
}

.mobilePicButtons
{
   display: inline;
}
/*--------------------------------------------------------------------------------
BEGIN: Set up for screens with max of 800px
----------------------------------------------------------------------------------*/
@media screen and (max-width: 768px) 
{

  .indent20
  {
    padding: 5px;
  }

  img
  {
    width: 100%;
  }

  #centercontent
  {
    width: 100%;
    margin-left: 0px;
  }
}

/*--------------------------------------------------------------------------------
BEGIN: Set up for screens with max of 640px
----------------------------------------------------------------------------------*/
@media screen and (max-width: 640px)
{
   #sMsg
   {
      width: 100%;
   }
}
</style>

	<script src="scripts/jquery-1.7.2.min.js"></script>
	<script src="scripts/modules.js"></script>
	<script src="../scripts/layers.js"></script>
	  
	<script>
	<!--
function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}
<%
 'Build the Attachment URL
  lcl_attachment_url = Application("common_url")
  lcl_attachment_url = lcl_attachment_url & "/public_documents300"
  lcl_attachment_url = lcl_attachment_url & "/" & GetVirtualDirectyName()
  lcl_attachment_url = lcl_attachment_url & "/attachments/"

  response.write "function viewAttachment(iFileName) { " & vbcrlf
  response.write "  lcl_width  = 800;" & vbcrlf
  response.write "  lcl_height = 700;" & vbcrlf
  response.write "  lcl_left   = (screen.availWidth/2)-(lcl_width/2);" & vbcrlf
  response.write "  lcl_top    = (screen.availHeight/2)-(lcl_height/2);" & vbcrlf

  response.write "  window.open('" & lcl_attachment_url & "' + iFileName, '_attachment', 'width=' + lcl_width + ',height=' + lcl_height + ',resizable=1,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + lcl_left + ',top=' + lcl_top);" & vbcrlf
  response.write "}"
%>

	function validateForm()
	{

		if($("#frmsubjecttext").val() != '') 
		{
			alert("Please remove any input from the Internal Only field at the bottom of the form.");
			$("#frmsubjecttext").focus();
			return false;
		}

		document.frmPost.submit();

	}

	//-->
	</script>
</head>

<!--#Include file="include_top.asp"-->
<!--BODY CONTENT-->
<%
  RegisteredUserDisplay("")


 'GET INFORMATION FOR THIS REQUEST
  iTrackID = request("REQUEST_ID") 

  if IsNumeric(iTrackID) and  len(iTrackID) > 4 then
    	iTrackID         = CStr(iTrackID)
     iRequestIDLength = len(iTrackID)
   		iTime            = Right(iTrackID,4)
   		iHour            = Left(iTime,2)
   		iMinute          = Right(iTime,2)

   		if iHour = "" OR iMinute = "" then
  		   	iHour = "99"
   	  		iMinute = "99"
   		end if

     if iRequestIDLength > 3 then
        iID = left(iTrackID, iRequestIDLength - 4)
     end if

   		if iID = "" then iID = "000000"

		on error resume next
			iID = clng(iID)
			if err.number <> 0 then
				iID = "000000"
			end if
		on error goto 0


   		if lcl_orghasfeature_action_line_substatus then
   	    sSQL = "SELECT r.*, (select IsNull(s.status_name,NULL) "
   	    sSQL = sSQL &      " from egov_actionline_requests_statuses s "
       	sSQL = sSQL &      " where r.sub_status_id = s.action_status_id) AS sub_status_name "
   		   sSQL = sSQL & " FROM egov_actionline_requests r "
   		else
	      	sSQL = "SELECT r.*, NULL AS sub_status_name "
   		   sSQL = sSQL & " FROM egov_actionline_requests r "
   	 end if

   		sSQL = sSQL & " WHERE (r.action_autoid = '" & iID & "') AND action_autoid != 327320 "
   	 sSQL = sSQL & " AND r.orgid = " & iorgid

   		set oRequest = Server.CreateObject("ADODB.Recordset")
   		oRequest.Open sSQL, Application("DSN"), 3, 1

	   'CHECK FOR INFORMATION
   		if not oRequest.eof then

   			 'REQUEST FOUND GET INFORMATION	
   	  		blnFound             = True
 	  				sTitle               = oRequest("category_title")
 	  				sStatus              = oRequest("status")
 	  				sSubStatus           = oRequest("sub_status_name")
 	  				datSubmitDate        = oRequest("submit_date")
 	  				sComment             = oRequest("comment")
 	  				iFormID              = oRequest("action_autoid")
 	  				iUserID              = oRequest("userid")
 	  				iCategoryId          = oRequest("category_id")

        if oRequest("public_actionline_pdf") <> "" then
           sPublicActionLinePDF = oRequest("public_actionline_pdf")
        else
           sPublicActionLinePDF = getDefaultPublicPDF(iCategoryID)
        end if

       'BEGIN: Action Line Request List ---------------------------------------
       'Get Contact Information
        if sTitle = "" then
     						'response.write "<font color=""#ff0000"">!No action line category name provided!</font><br />" & vbcrlf
     						response.write "<font color=""#ff0000"">!No " & lcl_actionlinelabel & " category name provided!</font><br />" & vbcrlf
        end if

        if sComment <> "" then
           sComment = replace(sComment,"default_novalue","")
        else
   			  			sComment = "<font color=""#ff0000"">No comment/description provided</font>"
        end if

       'Determine if the "View PDF" button is displayed.
       ' 1. The org must be assigned the "requestmergeforms" feature
       ' 2. The form on the request has a PDF associated to it.
        sDisplayButtonPDF = ""

        if lcl_orghasfeature_requestmergeforms AND sPublicActionLinePDF <> "" then
           sDisplayButtonPDF = "<div style=""text-align: right;"">"
           sDisplayButtonPDF = sDisplayButtonPDF & "<input type=""button"" onclick=""window.open('viewXMLPDF.asp?iRequestID=" & iID & "');"" value=""View Request in PDF Format"" />"
           sDisplayButtonPDF = sDisplayButtonPDF & "</div>" & vbcrlf
        end if

       'BEGIN: Issue/Problem Location -----------------------------------------
        sDisplayIssueLocation = ""

        if lcl_orghasfeature_issue_location then
           lcl_featurename_issuelocation = getOrgFeatureName("issue location")

           sSQL = "SELECT il.streetnumber, "
           sSQL = sSQL & " il.streetprefix, "
           sSQL = sSQL & " il.streetaddress, "
           sSQL = sSQL & " il.streetsuffix, "
           sSQL = sSQL & " il.streetdirection, ear.mobileoption_latitude, ear.mobileoption_longitude "
           sSQL = sSQL & " FROM egov_action_response_issue_location il "
		   sSQL = sSQL & " LEFT JOIN egov_actionline_requests ear ON ear.action_autoid = il.actionrequestresponseid "
           sSQL = sSQL & " WHERE actionrequestresponseid = " & iFormID
		   'response.write "<!--" & sSQL & "-->"

         		set oGetIssueLocation = Server.CreateObject("ADODB.Recordset")
         		oGetIssueLocation.Open sSQL, Application("DSN"), 3, 1

           if not oGetIssueLocation.eof then
              lcl_streetnumber    = oGetIssueLocation("streetnumber")
              lcl_streetprefix    = oGetIssueLocation("streetprefix")
              lcl_streetaddress   = oGetIssueLocation("streetaddress")
              lcl_streetsuffix    = oGetIssueLocation("streetsuffix")
              lcl_streetdirection = oGetIssueLocation("streetdirection")
			  lcl_lat 			  = oGetIssueLocation("mobileoption_latitude")
			  lcl_lng 			  = oGetIssueLocation("mobileoption_longitude")
           else
              lcl_streetnumber    = ""
              lcl_streetprefix    = ""
              lcl_streetaddress   = ""
              lcl_streetsuffix    = ""
              lcl_streetdirection = ""
			  lcl_lat 			  = ""
			  lcl_lng 			  = ""
           end if

           oGetIssueLocation.close
           set oGetIssueLocation = nothing

          'Build the street name
           lcl_street_name = buildStreetAddress(lcl_streetnumber, _
                                                lcl_streetprefix, _
                                                lcl_streetaddress, _
                                                lcl_streetsuffix, _
                                                lcl_streetdirection)

           sDisplayIssueLocation = "<fieldset class=""fieldset"">" & vbcrlf
           sDisplayIssueLocation = sDisplayIssueLocation & "<legend>" & lcl_featurename_issuelocation & "</legend> " & lcl_street_name & vbcrlf
			if not isnull(lcl_lat) and lcl_lat <> "" and not isnull(lcl_lng) and lcl_lng <> "" then
				sDisplayIssueLocation = sDisplayIssueLocation & "<div id=""map-canvas""></div>"
				sDisplayIssueLocation = sDisplayIssueLocation & "<style>" & vbcrlf
      				sDisplayIssueLocation = sDisplayIssueLocation & "#map-canvas {" & vbcrlf
        				sDisplayIssueLocation = sDisplayIssueLocation & "height: 200px;" & vbcrlf
        				sDisplayIssueLocation = sDisplayIssueLocation & "width: 200px;" & vbcrlf
        				sDisplayIssueLocation = sDisplayIssueLocation & "margin: 0px;" & vbcrlf
        				sDisplayIssueLocation = sDisplayIssueLocation & "padding: 0px" & vbcrlf
      				sDisplayIssueLocation = sDisplayIssueLocation & "}" & vbcrlf
    				sDisplayIssueLocation = sDisplayIssueLocation & "</style>" & vbcrlf
				sDisplayIssueLocation = sDisplayIssueLocation & "<script src=""https://maps.googleapis.com/maps/api/js?v=3.exp&sensor=false&callback=initialize"" type=""text/javascript""></script>" & vbcrlf
    				sDisplayIssueLocation = sDisplayIssueLocation & "<script>" & vbcrlf
        				sDisplayIssueLocation = sDisplayIssueLocation & "var map;" & vbcrlf
        				sDisplayIssueLocation = sDisplayIssueLocation & "function initialize() {" & vbcrlf
            				'sDisplayIssueLocation = sDisplayIssueLocation & "geocoder = new google.maps.Geocoder();" & vbcrlf
            				sDisplayIssueLocation = sDisplayIssueLocation & "var myLatlng = new google.maps.LatLng(" & lcl_lat & ", " & lcl_lng & ");" & vbcrlf
            				sDisplayIssueLocation = sDisplayIssueLocation & "var mapOptions = {" & vbcrlf
                				sDisplayIssueLocation = sDisplayIssueLocation & "zoom: 17," & vbcrlf
								sDisplayIssueLocation = sDisplayIssueLocation & "center: myLatlng"
            				sDisplayIssueLocation = sDisplayIssueLocation & "}" & vbcrlf
            				sDisplayIssueLocation = sDisplayIssueLocation & "map = new google.maps.Map(document.getElementById('map-canvas'), mapOptions);" & vbcrlf
							sDisplayIssueLocation = sDisplayIssueLocation & "var marker = new google.maps.Marker({ " & vbcrlf
      							sDisplayIssueLocation = sDisplayIssueLocation & "position: myLatlng, " & vbcrlf
      							sDisplayIssueLocation = sDisplayIssueLocation & "map: map " & vbcrlf
  							sDisplayIssueLocation = sDisplayIssueLocation & "}); " & vbcrlf
        				sDisplayIssueLocation = sDisplayIssueLocation & "}" & vbcrlf
    				sDisplayIssueLocation = sDisplayIssueLocation & "</script>" & vbcrlf
				
			end if
           sDisplayIssueLocation = sDisplayIssueLocation & "</fieldset>" & vbcrlf
        end if
       'END: Issue/Problem Location -------------------------------------------

       'BEGIN: Mobile Options -------------------------------------------------
        if lcl_orghasfeature_actionline_formcreator_mobileoptions then
           lcl_displayMobileOptionsTakePic = displayMobileOptions(iCategoryId, "display_mobileoptions_takepic")

          'BEGIN: Mobile Pics -------------------------------------------------
           if lcl_displayMobileOptionsTakePic then
              sMobilePicFile = Application("DocumentsDrive") & "/"
              sMobilePicFile = sMobilePicFile & Application("DocumentsRootDirectory")
              sMobilePicFile = sMobilePicFile & "/custom/pub/"
              sMobilePicFile = sMobilePicFile & sorgVirtualSiteName
              sMobilePicFile = sMobilePicFile & "/mobile_uploads"
              sMobilePicFile = sMobilePicFile & "/" + iTrackID

              dim fso, ObjFolder, ObjOutFile, ObjFiles, ObjFile
 
             'Creating File System Object
              Set fso = CreateObject("Scripting.FileSystemObject")

              if fso.FolderExists(sMobilePicFile) then
                'Getting the Folder Object
                 Set ObjFolder = fso.GetFolder(sMobilePicFile)

                'Getting the list of Files
                 Set ObjFiles = ObjFolder.Files

                 sDisplayMobilePic = ""
                 sDisplayMobilePicSection = ""

                 For Each ObjFile In ObjFiles
                    sDisplayMobilePic = sDisplayMobilePic & "<div class=""mobilePicImg"">" & vbcrlf
                    sDisplayMobilePic = sDisplayMobilePic & "   <img src=""" & sMobilePicFile & "/" & ObjFile.Name & """ class=""mobileImg"" />" & vbcrlf
                    sDisplayMobilePic = sDisplayMobilePic & "</div>" & vbcrlf
                    sDisplayMobilePic = sDisplayMobilePic & "<div class=""mobilePicButtons"">" & vbcrlf
                    sDisplayMobilePic = sDisplayMobilePic &     "<a href=""" & sMobilePicFile & "/" & ObjFile.Name & """ target=""_blank"" border=""0"">"
                    sDisplayMobilePic = sDisplayMobilePic &       "<input type=""button"" name=""buttonMobilePicView"" id=""buttonMobilePicView"" value=""View"" />" & vbcrlf
                    sDisplayMobilePic = sDisplayMobilePic &     "</a>" & vbcrlf
                    sDisplayMobilePic = sDisplayMobilePic &     ObjFile.Name  & vbcrlf
                    sDisplayMobilePic = sDisplayMobilePic & "</div>" & vbcrlf
                 Next

                 sDisplayMobilePic = replace(sDisplayMobilePic,Application("DocumentsDrive") & "/" & Application("DocumentsRootDirectory") & "/custom/pub", Application("common_url") & "/public_documents300")
                 sDisplayMobilePic = replace(sDisplayMobilePic,"\","/")
              end if

              sDisplayMobilePicSection = "<fieldset class=""fieldset"">" & vbcrlf
              sDisplayMobilePicSection = sDisplayMobilePicSection & "  <legend>Mobile Pics</legend>" & vbcrlf
              sDisplayMobilePicSection = sDisplayMobilePicSection & sDisplayMobilePic & vbcrlf
              sDisplayMobilePicSection = sDisplayMobilePicSection & "</fieldset>" & vbcrlf
           end if
          'END: Mobile Pics ---------------------------------------------------

        end if
       'END: Mobile Options ---------------------------------------------------

       'BEGIN: Due Date -------------------------------------------------------
        sDisplayDueDate = ""

        if lcl_orghasfeature_actionline_display_duedate then
           lcl_due_date = getRequestDueDate(iFormID)

           if lcl_due_date <> "" then
              sDisplayDueDate = "<fieldset class=""fieldset"">"
              sDisplayDueDate = sDisplayDueDate & "<legend>Due Date</legend>"
              sDisplayDueDate = sDisplayDueDate & lcl_due_date
              sDisplayDueDate = sDisplayDueDate & "</fieldset>" & vbcrlf
           end if
        end if
       'END: Due Date ---------------------------------------------------------

       'BEGIN: Hiding of contact info added 10/13/06 - Steve Loar -------------
        sDisplayContactInfo = ""

    				if not lcl_orghasfeature_hide_email_actionline then
          'Get the "Assigned To" email address associated to the request.
           lcl_assigned_email = ""

      					sSQLa = "SELECT assigned_email "
           sSQLa = sSQLa & " FROM egov_action_request_view "
           sSQLa = sSQLa & " WHERE action_autoid=" & iID

      					set oAssigned = Server.CreateObject("ADODB.Recordset")
      					oAssigned.Open sSQLa, Application("DSN") , 3, 1

           if not oAssigned.eof then
              lcl_assigned_email = oAssigned("assigned_email")
           end if

           oAssigned.close
           set oAssigned = nothing

				     	'Show CITY CONTACT INFORMATION
      					sDisplayContactInfo = "<fieldset class=""fieldset"">" & vbcrlf
           sDisplayContactInfo = sDisplayContactInfo & "  <legend>Email Contact</legend>" & vbcrlf
           sDisplayContactInfo = sDisplayContactInfo & "  <strong>" & lcl_assigned_email & "</strong> has been assigned to this request. " & vbcrlf
           sDisplayContactInfo = sDisplayContactInfo & "   Please contact via email - <a href=""mailto:" & lcl_assigned_email & """>" & lcl_assigned_email & "</a>" & vbcrlf
           sDisplayContactInfo = sDisplayContactInfo & "   - for further information regarding this request." & vbcrlf
           sDisplayContactInfo = sDisplayContactInfo & "</fieldset>" & vbcrlf
        end if
       'END: Hiding of contact info added 10/13/06 - Steve Loar ---------------

       'BEGIN: Online Dialog Response -----------------------------------------
        sDisplayDialogResponse = ""

        sDisplayDialogResponse = sDisplayDialogResponse & "<form name=""frmPost"" action=""#"" method=""POST"">" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  <input type=""" & lcl_hidden & """ name=""iFormID"" value="""     & iFormID     & """ />" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  <input type=""" & lcl_hidden & """ name=""iCategoryId"" value=""" & iCategoryId & """ />" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  <input type=""" & lcl_hidden & """ name=""iUserID"" value="""     & iUserID     & """ />" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  <input type=""" & lcl_hidden & """ name=""sStatus"" value="""     & sStatus     & """ />" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  <input type=""" & lcl_hidden & """ name=""sSubStatus"" value="""  & sSubStatus  & """ />" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  <input type=""" & lcl_hidden & """ name=""REQUEST_ID"" value="""  & iTrackID    & """ />" & vbcrlf

        sDisplayDialogResponse = sDisplayDialogResponse & "<fieldset class=""fieldset"">" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  <legend>Post a response/question</legend>" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  <textarea name=""sMsg"" id=""sMsg"" rows=""5"" cols=""80"" onMouseOut=""this.style.backgroundColor='#ffffff';"" onMouseOver=""this.style.backgroundColor='#ffffcc';""></textarea>" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  <div>" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "    <input type=""button"" value=""POST MESSAGE"" onclick=""validateForm( )"" />" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "    <input type=""reset"" value=""CLEAR MESSAGE"" />" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  </div>" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "<div id=""problemtextfield1"">" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  Internal Use Only, Leave Blank: <input type=""text"" name=""frmsubjecttext"" id=""frmsubjecttext"" value="""" size=""6"" />" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  <input type=""" & lcl_hidden & """ name=""problemorg"" value=""" & iorgid & """ /><br />" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "  <strong>Please leave this field blank and remove any values that have been populated for it.</strong>" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "</div>" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "</fieldset>" & vbcrlf
        sDisplayDialogResponse = sDisplayDialogResponse & "</form>" & vbcrlf
       'END: Online Dialog Response -------------------------------------------

       'BEGIN: Action Request Log ---------------------------------------------
       'Display Action Request Status
   		   lcl_display_substatus = ""

        if lcl_orghasfeature_action_line_substatus then
			   		   lcl_display_substatus = "<em>(Sub-Status)</em>"
   					end if

     			sListComments = List_Comments(iID, sOrgName)

        sDisplayActionRequestLog = "<fieldset class=""fieldset"">" & vbcrlf
        sDisplayActionRequestLog = sDisplayActionRequestLog & "  <legend>" & lcl_actionlinelabel & " Activity</legend>" & vbcrlf
        sDisplayActionRequestLog = sDisplayActionRequestLog & sListComments
        sDisplayActionRequestLog = sDisplayActionRequestLog & "</fieldset>" & vbcrlf
       'END: Action Request Log -----------------------------------------------

       'BEGIN: Attachments ----------------------------------------------------
      	 sSQL = "SELECT attachmentid, "
        sSQL = sSQL & " submitted_request_id, "
        sSQL = sSQL & " attachment_name, "
        sSQL = sSQL & " attachment_desc, "
        sSQL = sSQL & " adminuserid, "
        sSQL = sSQL & " date_added, "
        sSQL = sSQL & " firstname, "
        sSQL = sSQL & " lastname, "
        sSQL = sSQL & " isSecure, "
        sSQL = sSQL & " displayToPublic "
        sSQL = sSQL & " FROM egov_submitted_request_attachments "
        sSQL = sSQL &      " INNER JOIN users on adminuserid = userid "
        sSQL = sSQL & " WHERE submitted_request_id='" & iID & "'"
        sSQL = sSQL & " AND displayToPublic = 1 "

   					set oAttachments = Server.CreateObject("ADODB.Recordset")
   					oAttachments.Open sSQL, Application("DSN"), 3, 1

        if not oAttachments.eof then

           lcl_bgcolor = "#eeeeee"

      					sDisplayAttachments = "<fieldset class=""fieldset"">" & vbcrlf
           sDisplayAttachments = sDisplayAttachments & "  <legend>Attachments</legend>" & vbcrlf

           do while not oAttachments.eof

              lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
              lcl_attachment_info = ""

              if oAttachments("attachment_name") <> "" then
                 if lcl_attachment_info <> "" then
                    lcl_attachment_info = lcl_attachment_info & " - " & oAttachments("attachment_name")
                 else
                    lcl_attachment_info = oAttachments("attachment_name")
                 end if
              end if

              if oAttachments("attachment_desc") <> "" then
                 if lcl_attachment_info <> "" then
                    lcl_attachment_info = lcl_attachment_info & " - " & oAttachments("attachment_desc")
                 else
                    lcl_attachment_info = oAttachments("attachment_desc")
                 end if
              end if

             'Build Attachment URL
              lcl_file_ext  = lcase(right(oAttachments("attachment_name"), len(oAttachments("attachment_name")) - instrrev(oAttachments("attachment_name"),".")))

              lcl_attachment_url = oAttachments("attachmentid")
              lcl_attachment_url = lcl_attachment_url & "." & lcl_file_ext

              sDisplayAttachments = sDisplayAttachments & "  <div style=""background-color: " & lcl_bgcolor & "; padding: 4px;"">" & vbcrlf
              sDisplayAttachments = sDisplayAttachments & "    <input type=""button"" name=""viewAttachment" & oAttachments("attachmentid") & """ id=""viewAttachment" & oAttachments("attachmentid") & """ value=""View"" onclick=""viewAttachment('" & lcl_attachment_url & "');"" />" & vbcrlf
              sDisplayAttachments = sDisplayAttachments & "    &nbsp;&nbsp;&nbsp;" & lcl_attachment_info & vbcrlf
              sDisplayAttachments = sDisplayAttachments & "  </div>" & vbcrlf

              oAttachments.movenext

           loop

        end if

        sDisplayAttachments = sDisplayAttachments & "</fieldset>" & vbcrlf
       'END: Attachments ------------------------------------------------------

       'BEGIN: Display description of Action Line Requests --------------------
        response.write "<fieldset class=""fieldset"">" & vbcrlf
        response.write "  <legend>" & lcl_actionlinelabel & " Item: " & sTitle & "</legend>" & vbcrlf
        response.write sDisplayButtonPDF
        response.write "<div>" & vbcrlf
        response.write "  <strong>Your initial message:</strong><br /><br />" & vbcrlf
        response.write "  <em>" & sComment & "</em>" & vbcrlf
        response.write "</div>" & vbcrlf
        response.write sDisplayIssueLocation
        response.write sDisplayMobilePicSection
        response.write sDisplayDialogResponse
        response.write sDisplayDueDate
        response.write sDisplayActionRequestLog
        response.write sDisplayAttachments
        response.write sDisplayContactInfo
        response.write "</fieldset>" & vbcrlf
       'END: Display description of Action Line Requests ----------------------
   		else

        blnFound = False
        displayRequestNotFound iTrackID, _
                               sDefaultEmail, _
                               sOrgName

     end if

     set oRequest = nothing

  else

		  'TrackID is non-numeric
   		blnFound = False
     displayRequestNotFound iTrackID, _
                            sDefaultEmail, _
                            sOrgName
		
  end if
%>
<!--#Include file="include_bottom.asp"-->  
<%
'------------------------------------------------------------------------------
function List_Comments(iID, iOrgName)
  dim lcl_return

  lcl_return = ""
  sBGColor   = "#ffffff"

 	sSQL = "SELECT * "
 	sSQL = sSQL & " FROM egov_action_responses egr "
 	sSQL = sSQL & " LEFT OUTER JOIN egov_users ON egr.action_userid = egov_users.userid "
 	sSQL = sSQL & " LEFT OUTER JOIN egov_actionline_requests_statuses AS es "
 	sSQL = sSQL &               "ON egr.action_sub_status_id = es.action_status_id "
 	sSQL = sSQL & " WHERE egr.action_autoid = " & iID
  sSQL = sSQL & " AND egr.action_orgid = " & iorgid
 	sSQL = sSQL & " ORDER BY egr.action_editdate DESC"

 	set oCommentList = Server.CreateObject("ADODB.Recordset")
 	oCommentList.Open sSQL, Application("DSN") , 3, 1
	
 	if not oCommentList.eof then
	  	 do while not oCommentList.eof
        sBGColor           = changeBGColor(sBGColor,"#eeeeee","#ffffff")

        lcl_substatus_name = oCommentList("status_name")
        lcl_comment_label  = ""
        lcl_comment_text   = ""

     	  if lcl_substatus_name <> "" then
	  		      lcl_substatus_name = " <em>(" & lcl_substatus_name & ")</em>"
     	  end If

        'response.write oCommentList("action_externalcomment") <> "" 
        'response.write "#" & oCommentList("action_externalcomment") & "#"
        if oCommentList("action_externalcomment") <> "" then
          lcl_comment_label = iOrgName
          lcl_comment_text  = oCommentList("action_externalcomment")
        elseif oCommentList("action_citizen") <> "" then
          lcl_comment_label = oCommentList("userfname") & " " & oCommentList("userlname")
          lcl_comment_text  = oCommentList("action_citizen")
        else
          lcl_comment_label = iOrgName
          'lcl_comment_text  = "Your request was reviewed and/or status was updated." ' commented out to use a display now, 6/4/2014'
          lcl_comment_text  = lcl_defaultComment
        end if

        if trim(lcl_comment_label) <> "" then
           lcl_comment_label = "<strong>[" & lcl_comment_label & "]: </strong>"
        else
           lcl_comment_label = "<strong>Citizen Comment: </strong>"
        end if

        if lcl_comment_text <> "" then
           lcl_comment_text = formatActivityLogComment(lcl_comment_text)
           lcl_comment_text = "<em>" & lcl_comment_text & "</em>"
        end if

        if lcl_return = "" then
           lcl_return = "<div style=""background-color: #ffffff; border-bottom: 1pt solid #808080; padding:4px;""><strong>Status " & lcl_display_substatus & " - Date of Activity</strong></div>" & vbcrlf
        end if

        lcl_return = lcl_return & "<div style=""padding:6px 6px 0px 6px; background-color: " & sBGColor & ";"">" & ucase(oCommentList("action_status")) & lcl_substatus_name & " - " &  oCommentList("action_editdate") & "</div>" & vbcrlf
        lcl_return = lcl_return & "<div style=""padding:0px 6px 6px 6px; background-color: " & sBGColor & "; border-bottom: 1pt solid #808080;"">&nbsp;&nbsp;&nbsp;" & lcl_comment_label & lcl_comment_text & "</div>" & vbcrlf

     			oCommentList.movenext
     loop
  else
   		lcl_return = "<div style=""color:#ff0000;font-style:italic;"">No activity</div>" & vbcrlf
  end if

  oCommentList.close
  set oCommentList = nothing

  List_Comments = lcl_return

end function

'------------------------------------------------------------------------------
Function CheckSelected(sValue, sValue2)
	sReturnValue = ""
	If sValue = sValue2 Then
		sReturnValue = "SELECTED"
	End If

	CheckSelected = sReturnValue
End Function

'------------------------------------------------------------------------------
Function AddCommentTaskComment(sStatus, _
                               sCitizenMsg, _
                               iFormID, _
                               iUserID, _
                               iOrgId, _
                               iCreateDate)

		sSQL = "INSERT egov_action_responses ("
  sSQL = sSQL & "action_status, "
  sSQL = sSQL & "action_citizen, "
  sSQL = sSQL & "action_userid, "
  sSQL = sSQL & "action_orgid, "
  sSQL = sSQL & "action_autoid, "
  sSQL = sSQL & "action_editdate "
  sSQL = sSQL & ") VALUES ( "
  sSQL = sSQL & "'" & sStatus             & "', "
  sSQL = sSQL & "'" & DBsafe(sCitizenMsg) & "', "
  sSQL = sSQL & "'" & iUserID             & "', "
  sSQL = sSQL & "'" & iOrgID              & "', "
  sSQL = sSQL & "'" & iFormID             & "', "
  sSQL = sSQL & "'" & iCreateDate         & "' "
  sSQL = sSQL & ")"
		Set oComment = Server.CreateObject("ADODB.Recordset")
		oComment.Open sSQL, Application("DSN"), 3, 1
		Set oComment = Nothing
End Function

'------------------------------------------------------------------------------
Function DBsafe( strDB )
	Dim sNewString
	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
	sNewString = Replace( strDB, "'", "''" )
	sNewString = Replace( sNewString, "<", "&lt;" )
	DBsafe = sNewString
End Function

'------------------------------------------------------------------------------
Sub EmailComment(sCitizenMsg, _
                 iActionFormid, _
                 iActionId, _
                 iTrackingNo, _
                 iUserID, _
                 iCreateDate, _
                 iOrgName )

	Dim sSQLadmin, oAdmin, sSQLaddress, oAddress, sMsg2, objMail2

	sMsg2              = ""
 adminid            = 0
 adminDelegateID    = 0
 adminEmailAddr     = ""
 adminDelegateEmail = ""
 lcl_featurename_actionline = GetOrgFeatureName("action line")

'Get the user assigned, and delegate (if available), to this request
	sSQLadmin = "SELECT assigned_userid, assigned_email, delegateid, delegate_email "
 sSQLadmin = sSQLadmin & " FROM egov_rpt_actionline "
 sSQLadmin = sSQLadmin & " WHERE [Tracking Number] = '" & iTrackingNo & "' "

	set oAdmin = Server.CreateObject("ADODB.Recordset")
	oAdmin.Open sSQLadmin, Application("DSN"), 0, 1

	if NOT oAdmin.EOF then
	  	if oAdmin("assigned_userid") = "" or isNull(oAdmin("assigned_userid")) then
       'NOTHING
  		else
  		   if iorgid = 18 then
		   			 'This handles Vandalia's inability to receive email from themselves
         	'adminFromAddr = "webmaster@eclink.com"
         	adminFromAddr = "noreply@eclink.com"
       else 
          adminFromAddr    = oAdmin("assigned_email")  'ASSIGNED ADMIN USER EMAIL
          adminDeleteEmail = oAdmin("delegate_email")
       end if

       adminEmailAddr     = oAdmin("assigned_email")   'ASSIGNED ADMIN USER EMAIL
       adminid            = oAdmin("assigned_userid")  'ASSIGNED ADMIN USER ID
       adminDelegateEmail = oAdmin("delegate_email")
       adminDelegateID    = oAdmin("delegateid")
    end if
 end if

 oAdmin.Close
	Set oAdmin = Nothing

'BEGIN: Build message and send email to administrator(s) ----------------------
	sMsg2 = "This automated message was sent by the " & iOrgName & " E-Gov web site.  Do not reply to this message.  "

'Check to see if the org wants to hide their admin emails or not.
 if not lcl_orghasfeature_hide_email_actionline then
    sMsg2 = sMsg2 & "Contact " & adminFromAddr & " for inquiries regarding this email.  " & vbcrlf
 end if

	sMsg2 = sMsg2 & "A " & iOrgName & " " & lcl_actionlinelabel & " issue was updated on " & iCreateDate & "." & vbcrlf 
	sMsg2 = sMsg2 & "<br /><br />" & vbcrlf 

 sMsg2 = sMsg2 & "<p><strong>Click the following link to view this Action Line Request:</strong><br />" & vbcrlf
	sMsg2 = sMsg2 & "<a href=""" & sEgovWebsiteURL & "/admin/action_line/action_respond.asp?control=" & iActionId & "&e=Y"">" & vbcrlf
	sMsg2 = sMsg2 & sEgovWebsiteURL & "/admin/action_line/action_respond.asp?control=" & iActionId & "&e=Y</a></p>" & vbcrlf

	sMsg2 = sMsg2 & UCASE(lcl_actionlinelabel) & " DETAILS<br />" & vbcrlf
	sMsg2 = sMsg2 & "UPDATED BY: "      & GetCitizenName( iUserID ) & "<br />" & vbcrlf
	sMsg2 = sMsg2 & "TRACKING NUMBER: " & iTrackingNo & "<br />" & vbcrlf
	sMsg2 = sMsg2 & "COMMENT: "         & sCitizenMsg & "<br />" & vbcrlf

 lcl_message = BuildHTMLMessage(sMsg2)

'Prepare email to send
	if iorgid <> "7" then
    'lcl_from    = iOrgName & " " & lcl_featurename_actionline & " <webmaster@eclink.com>"
    lcl_from    = iOrgName & " " & lcl_featurename_actionline & " <noreply@eclink.com>"
    lcl_subject = iOrgName & " " & lcl_featurename_actionline & ": User Comment Added"
	else
    'lcl_from    = iOrgName & " ECLINK HELPDESK <webmaster@eclink.com>"
    lcl_from    = iOrgName & " ECLINK HELPDESK <noreply@eclink.com>"
    lcl_subject = "ECLINK HELPDESK - User Comment Added"
	end if

'Remove the name from the email address
 lcl_validate_email          = formatSendToEmail(adminEmailAddr)
 lcl_validate_delegate_email = formatSendToEmail(adminDelegateEmail)

'BEGIN: Send the email -----------------------------------------------------
 if isValidEmail(lcl_validate_email) then
   'If there's a valid delegate then send the email to the delegate and "CC" the assigned person.
   'Otherwise, send it to the assigned person
    if isValidEmail(lcl_validate_delegate_email) then
       sendEmail "", adminDelegateEmail, adminEmailAddr, lcl_subject, lcl_message, "", "Y"
    else
       sendEmail "",adminEmailAddr,"",lcl_subject,lcl_message,"","Y"
    end if
 else
   'check for a delegate (may happen if someone leaves and their email is cleared off their record.)
    if isValidEmail(lcl_validate_delegate_email) then
       sendEmail "", adminDelegateEmail, "", lcl_subject, lcl_message, "", "Y"
    else
       ErrorCode = 1
    end if
 end if
'END: Send the email -------------------------------------------------------


'Add to email queue if unsuccessful
	if ErrorCode <> 0 then
				'sMsg      = Left(sMsg,5000)
	   'SendToAdd = adminEmailAddr
   	'fnPlaceEmailinQueue Application("SMTP_Server"),sOrgName & " E-GOV WEBSITE",adminFromAddr,SendToAdd,sOrgName & " E-GOV MSG - " & UCASE(lcl_featurename_actionline) & " REQUEST",1,sMsg2,1,-1

		  response.write "The request has been logged but there was an error sending an email notice to you.  "
    response.write "You will not receive an email notice.<br /><br /><br />" & vbcrlf

				bMailSent1 = False
	end if

'END: Build message and send email to citizen --------------------------------

End Sub 

'------------------------------------------------------------------------------
Function GetCitizenName( iUserID )
	Dim sSql, oName

	sSql = "Select userfname, userlname from egov_users where userid = "  & iUserID

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF Then 
		GetCitizenName = oName("userfname") & " " & oName("userlname")
	Else 
		GetCitizenName = "Unknown Citizen"
	End If 

	oName.close
	Set oName = Nothing
End Function

'------------------------------------------------------------------------------
sub displayRequestNotFound(iTrackID, _
                           iDefaultEmail, _
                           iOrgName)

  'response.write "<div style=""margin-left:20px;"" class=""box_header2"">Action Line Request Lookup</div>" & vbcrlf
  'response.write "<p>We could not locate an action line request using <strong>TRACKING NUMBER</strong> <strong>(" & iTrackID & ")</strong>.</p>" & vbcrlf
  response.write "<div style=""margin-left:20px;"" class=""box_header2"">" & lcl_actionlinelabel & " Lookup</div>" & vbcrlf
  response.write "<div class=""groupsmall"" style=""margin-left:20px;"">" & vbcrlf
  response.write "<p>We could not locate an " & lcl_actionlinelabel & " using <strong>TRACKING NUMBER</strong> <strong>(" & iTrackID & ")</strong>.</p>" & vbcrlf

  if iDefaultEmail = "" then
     'sActionDefaultEmail = "webmaster@eclink.com"
     sActionDefaultEmail = "noreply@eclink.com"
  else
     sActionDefaultEmail = iDefaultEmail
  end if

  response.write "<p>" & vbcrlf
  response.write "   Please press <strong>BACK</strong> on your browser, check the <strong>TRACKING NUMBER</strong> and try again. "
  response.write "   If you continue to receive this message please contact "
  response.write "   <a href=""mailto:""" & sActionDefaultEmail & """>" & sActionDefaultEmail & "</a> for further assistance with "
  response.write "   this request." & vbcrlf
  response.write "</p>" & vbcrlf
  response.write "<p>Thank you for using " & iOrgName & " E-gov website.</p>" & vbcrlf
  response.write "</div>" & vbcrlf
end sub

'------------------------------------------------------------------------------
function getDefaultPublicPDF(iFormID)

  lcl_return = ""

  if iFormID <> "" then
     sSQL = "SELECT public_actionline_pdf "
     sSQL = sSQL & " FROM egov_action_request_forms "
     sSQL = sSQL & " WHERE action_form_id=" & iFormID

    	set oPDF = Server.CreateObject("ADODB.Recordset")
   	 oPDF.Open sSQL, Application("DSN"), 0, 1

     if not oPDF.eof then
        lcl_return = oPDF("public_actionline_pdf")
     end if

     oPDF.close
     set oPDF = nothing

  end if

  getDefaultPublicPDF = lcl_return

end function
%>
