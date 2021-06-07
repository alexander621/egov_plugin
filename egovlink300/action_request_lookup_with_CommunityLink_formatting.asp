<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="class/classOrganization.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<!-- #include file="communitylink_global_functions.asp" //-->
<!-- #include file="communitylink_global_styles.asp" //-->
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
' 2.4 06/18/09  David Boyer - Converted to "CommunityLink" layout.
' 2.5 08/04/09  David Boyer - Added "delegate"
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'To help prevent hacks
 if NOT isnumeric(request("REQUEST_ID")) then
    response.redirect "action.asp"
 end if

'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 iStartSecs = timer
 sUserName  = ""
 lcl_onload = ""

 Dim sError, sActionDefaultEmail, oOrg, lcl_scripts

 set oOrg = New classOrganization

 lcl_hidden = "hidden"  'Show/Hides all hidden fields.  HIDDEN = Hide, TEXT = Show

 datOrgDateTime = ConvertDateTimetoTimeZone(iorgid)

'If users supplied comments then update them
 if Request.ServerVariables("REQUEST_METHOD") = "POST" AND request("sMsg") <> "" then
   	sCitizenMsg = request("sMsg")
   	iFormID     = CLng(request("iFormID"))
   	iUserID     = CLng(request("iUSerID"))
   	sStatus     = request("sStatus")
   	iOrgID      = iorgid

   	AddCommentTaskComment sStatus,sCitizenMsg,iFormID,iUserID,iOrgID, datOrgDateTime

  	'Email the comment to those who get the message. - Steve Loar - 4/10/2006
   	EmailComment sCitizenMsg, request("iCategoryId"), iFormID, request("REQUEST_ID"), iUserID, datOrgDateTime
 end if

'Check for org features
 lcl_orghasfeature_requestmergeforms     = orghasfeature(iorgid,"requestmergeforms")
 lcl_orghasfeature_action_line_substatus = orghasfeature(iorgid,"action_line_substatus")
 lcl_orghasfeature_hide_email_actionline = orghasfeature(iorgid,"hide email actionline")

'Set the "Action Line Request" label
 lcl_actionlinelabel = "Action Line Request"

 if OrgHasDisplay(iOrgID,"actionlinelabel_publicrequestlookup") then
    lcl_actionlinelabel = GetOrgDisplayWithId(iOrgID,getDisplayID("actionlinelabel_publicrequestlookup"),False)
 end if

'Capture current path
 session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
 session("RedirectLang") = "Return to " & lcl_actionlinelabel

'Determine if the user is accessing site from desktop or mobile device (iPhone/Blackberry)
 session("accessdevice") = checkAccessMethod(request.servervariables("http_user_agent"))

'S = Standard, M = Mobile
 if request("setDeviceViewMode") <> "" then
    session("deviceViewMode") = request("setDeviceViewMode")
 else
    if session("deviceViewMode") = "" then
       if session("accessdevice") = "BLACKBERRY" OR session("accessdevice") = "IPHONE" then
          session("deviceViewMode") = "M"
       else
          session("deviceViewMode") = "S"
       end if
    end if
 end if

'Check for cookies
 lcl_cookie_userid = request.cookies("userid")

'Check for a CommunityLink record for the org.
'If one DOES exist then pull all of the values.
'If one does NOT exist then create it and enter the default values.
 lcl_communitylinkid = getCommunityLinkID(iorgid,lcl_cookie_userid)

'Retrieve the CommunityLink record.
 getCommunityLinkInfo lcl_communitylinkid, iorgid, lcl_isEgovHomePage, lcl_website_size, lcl_website_size_customsize, _
                      lcl_website_alignment, lcl_website_bgcolor, lcl_showlogo, lcl_logo_filename, lcl_logo_filenamebg, _
                      lcl_logo_alignment, lcl_showtopbar, lcl_topbar_bgcolor, lcl_topbar_fonttype, lcl_topbar_fontcolor, _
                      lcl_topbar_fontcolorhover, lcl_showsidemenubar, lcl_sidemenubar_alignment, lcl_sidemenuoption_bgcolor, _
                      lcl_sidemenuoption_bgcolorhover, lcl_sidemenuoption_alignment, lcl_sidemenuoption_fonttype, _
                      lcl_sidemenuoption_fontcolor, lcl_sidemenuoption_fontcolorhover, lcl_showpageheader, lcl_pageheader_alignment, _
                      lcl_pageheader_fontsize, lcl_pageheader_fontcolor, lcl_pageheader_fonttype, lcl_pageheader_bgcolor, _
                      lcl_showfooter, lcl_footer_bgcolor, lcl_footer_fonttype, lcl_footer_fontcolor, lcl_footer_fontcolorhover, _
                      lcl_showRSS, lcl_url_twitter, lcl_url_facebook, lcl_url_myspace, lcl_url_blogger

'Check for org features
 lcl_orghasfeature_administrationlink = orghasfeature(iorgid,"AdministrationLink")

'Build the Title
 lcl_title = sOrgName

 if iorgid <> 7 then
    lcl_title = "E-Gov Services " & lcl_title
 end if
%>
<html>
<head>

 	<title><%=lcl_title%></title>
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />

 	<link rel="stylesheet" type="text/css" href="css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="global.css" />
 	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />
<%
 'BEGIN: CommunityLink Styles -------------------------------------------------
  response.write "<style type=""text/css"">" & vbcrlf

 'Do not build the following styles if the user is accessing the screen from a mobile device.
  if session("deviceViewMode") <> "M" then
     buildStyles_CL_SideMenuBar blnMenuOn, lcl_showsidemenubar, lcl_sidemenuoption_bgcolor, lcl_sidemenuoption_fonttype, _
                                lcl_sidemenuoption_fontcolor, lcl_sidemenuoption_fontcolorhover
  else
    'Modify the "topbar/footer" styles if the user is accessing the screen from a mobile device.
     lcl_topbar_bgcolor        = "#ffffff"
     lcl_topbar_fontcolor      = "#0000ff"
     lcl_topbar_fontcolorhover = "#0000ff"

     lcl_footer_bgcolor        = "#ffffff"
     lcl_footer_fontcolor      = "#0000ff"
     lcl_footer_fontcolorhover = "#0000ff"
  end if

  buildStyles_CL_Misc
  buildStyles_CL_TopBar     lcl_topbar_bgcolor, lcl_topbar_fontcolor, lcl_topbar_fonttype, lcl_topbar_fontcolorhover
  buildStyles_CL_PageHeader lcl_pageheader_bgcolor, lcl_pageheader_fonttype, lcl_pageheader_fontcolor, lcl_pageheader_fontsize
  buildStyles_CL_Footer     lcl_footer_bgcolor, lcl_footer_fontcolor, lcl_footer_fonttype, lcl_footer_fontcolorhover
  buildStyles_CL_LinkOptions

  response.write "</style>" & vbcrlf
 'END: CommunityLink Styles ---------------------------------------------------

 'BEGIN: Javascripts ----------------------------------------------------------
  if session("deviceViewMode") <> "M" then
     response.write "<script type=""text/javascript"" src=""https://s7.addthis.com/js/200/addthis_widget.js""></script>" & vbcrlf
     response.write "<script type=""text/javascript"">var addthis_pub=""cschappacher"";</script>" & vbcrlf
  end if
%>
 	<script language="javascript" src="scripts/modules.js"></script>
 	<!-- <script language="javascript" src="scripts/layers.js"></script> -->
	  
<script language="javascript">
<!--
function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}
<%
'------------------------------------------------------------------------------
'Only build these javascript functions if the user is accessing from a non-mobile device or is in STANDARD view mode
'------------------------------------------------------------------------------
 if session("deviceViewMode") <> "M" then
%>
function setupMenuOption(iType,iRowID) {
		var rege = /^[0-9a-f]{3,6}$/i;

  if(iType=="OVER") {
     lcl_optionbg      = '<%=lcl_sidemenuoption_bgcolorhover%>';
     lcl_textcolor     = '<%=lcl_sidemenuoption_fontcolorhover%>';
     lcl_showUnderLine = 'underline';
  }else{
     lcl_optionbg      = '<%=lcl_sidemenuoption_bgcolor%>';
     lcl_textcolor     = '<%=lcl_sidemenuoption_fontcolor%>';
     lcl_showUnderLine = 'none';
  }

  if(lcl_optionbg!="") {
   		var Ok = rege.exec(lcl_optionbg);
     if ( Ok ) {
         document.getElementById("sideMenuBar"       + iRowID).style.backgroundColor="#" + lcl_optionbg;
         document.getElementById("sideMenuBarOption" + iRowID).style.backgroundColor="#" + lcl_optionbg;
     }
  }

  if(lcl_textcolor!="") {
   		var Ok = rege.exec(lcl_textcolor);
     if ( Ok ) {
         document.getElementById("sideMenuBarOption" + iRowID).style.color="#" + lcl_textcolor;
     }
  }

  document.getElementById("sideMenuBarOption" + iRowID).style.textDecoration = lcl_showUnderLine;
}

function openWin(p_page, p_field_id, p_wintype, p_width, p_height) {
  if ((p_wintype=="")||(p_wintype==undefined)) {
       s_wintype="_picker";
  }else{
       s_wintype=p_wintype;
  }
  if ((p_width=="")||(p_width==undefined)) {
       s_width=600;
  }else{
       s_width=p_width;
  }
  if ((p_height=="")||(p_height==undefined)) {
       s_height=470;
  }else{
       s_height=p_height;
  }

  s_left = (screen.availWidth/2) - (s_width/2);
  s_top  = (screen.availHeight/2) - (s_height/2);

  if((p_field_id=="")||(p_field_id!=undefined)) {
      lcl_url = p_page;
  }else{
      lcl_url = p_page + "?fieldid=" + p_field_id;
  }

		eval('window.open("' + lcl_url + '", "' + s_wintype + '", "width=' + s_width + ',height=' + s_height + ',left=' + s_left + ',top=' + s_top + 'status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes")');
}
<%
'------------------------------------------------------------------------------
 end if
'------------------------------------------------------------------------------
%>

function changeElementStyles(p_viewLinkID, p_fontsize, p_fonttype, p_fontcolor, p_underline, p_backgroundcolor) {
		var rege = /^[0-9a-f]{3,6}$/i;

  if(p_viewLinkID!="") {
     var lcl_viewlink = document.getElementById(p_viewLinkID);

     if(p_fontsize!="") {
        lcl_viewlink.style.fontSize = p_fontsize+'px';
     }

     if(p_fonttype!="") {
        lcl_viewlink.style.fontFamily = p_fonttype;
     }

     if(p_fontcolor!="") {
      		var Ok = rege.exec(p_fontcolor);

        if ( Ok ) {
            lcl_viewlink.style.color = '#' + p_fontcolor;
        }
     }

     if(p_underline!="") {
        lcl_viewlink.style.textDecoration = p_underline;
     }

     if(p_backgroundcolor!="") {
      		var Ok = rege.exec(p_backgroundcolor);

        if ( Ok ) {
            lcl_viewlink.style.backgroundColor = '#' + p_backgroundcolor;
        }
     }
  }
}

window.onload = function(){
  <%=lcl_onload%>
}
//-->
</script>
</head>
<body bgcolor="#<%=lcl_website_bgcolor%>" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
 'BEGIN: CommunityLink --------------------------------------------------------
  lcl_website_width = getWebsiteWidth(lcl_website_size, lcl_website_size_customsize)

  response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"" align=""" & lcl_website_alignment & """ bgcolor=""" & lcl_website_bgcolor & """>" & vbcrlf
  response.write "          <table border=""0"" bordercolor=""#ff0000"" cellspacing=""0"" cellpadding=""0"" width=""" & lcl_website_width & """ bgcolor=""#ffffff"" style=""border:1pt solid #000000;"">" & vbcrlf

 'BEGIN: Show Logo ------------------------------------------------------------
  if lcl_showlogo then
    'Build the Logo URLs
     lcl_orgLogoURL = sEgovWebsiteURL
     lcl_orgLogoURL = lcl_orgLogoURL & "/admin/custom/pub/"
     lcl_orgLogoURL = lcl_orgLogoURL & sorgVirtualSiteName
     lcl_orgLogoURL = lcl_orgLogoURL & "/unpublished_documents"

     if lcl_logo_filename <> "" then
        lcl_logo_filename = lcl_orgLogoURL & lcl_logo_filename
     else
        lcl_logo_filename = getDefaultLogo("LEFT",iorgid)
     end if

     if lcl_logo_filenamebg <> "" then
        lcl_logo_filenamebg = lcl_orgLogoURL & lcl_logo_filenamebg
     else
        lcl_logo_filenamebg = getDefaultLogo("RIGHT",iorgid)
     end if

     lcl_orgLogo = "<img src=""" & lcl_logo_filename & """ name=""orgLogo"" id=""orgLogo"" border=""0"" />"

    'If the logofilenamebg is NULL then display the logo bgcolor
     if lcl_logo_filenamebg <> "" then
        lcl_orgLogoBGstyle = "background-image:url('" & lcl_logo_filenamebg & "');"
     end if

    'Get the Org's website url
     lcl_egovSiteURL = oOrg.GetEgovURL

     response.write "            <tr>" & vbcrlf
     response.write "                <td colspan=""3"" class=""orgLogo"" align=""" & lcl_logo_alignment & """ onclick=""location.href='" & lcl_egovSiteURL & "'"">" & vbcrlf
     response.write "                    <div style=""" & lcl_orgLogoBGstyle & """><a href=""" & lcl_egovSiteURL & """>" & lcl_orgLogo & "</a></div>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if
 'END: Show Logo --------------------------------------------------------------

 'BEGIN: Show TopBar ----------------------------------------------------------
  if lcl_showtopbar AND session("deviceViewMode") <> "M" then

    'Get the citizen's name if they are logged in.
 				sUserName = ""

   		if sOrgRegistration AND lcl_cookie_userid <> "" AND lcl_cookie_userid <> "-1" then
     			sSQL = "SELECT userfname + ' ' + userlname as username "
        sSQL = sSQL & " FROM egov_users "
        sSQL = sSQL & " WHERE userid = '" & lcl_cookie_userid & "'"

    			 set oCitizenName = Server.CreateObject("ADODB.Recordset")
     			oCitizenName.Open sSQL, Application("DSN"), 3, 1

     			if not oCitizenName.eof then
       				sUserName = ", <strong>" & trim(ucase(oCitizenName("username"))) & "</strong>"
        end if

     			oCitizenName.close 
     			set oCitizenName = nothing
     end if

     response.write "            <tr>" & vbcrlf
     response.write "                <td colspan=""3"" class=""topBar"">" & vbcrlf
     response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" class=""topBar"">" & vbcrlf
     response.write "                      <tr>" & vbcrlf
     response.write "                          <td align=""left"">" & vbcrlf
     'response.write "                              <img class=""accountmenu"" src=""" & sEgovWebsiteURL & "/images/accountmenu.jpg"" style=""float:left;"" />" & vbcrlf
  			response.write "                              <i>Today is " & FormatDateTime(Date(), vbLongDate) & ".</i>&nbsp;&nbsp;Welcome" & sUserName & "!" & vbcrlf
     response.write "                          </td>" & vbcrlf
     response.write "                          <td align=""right"">" & vbcrlf

    'If the user has logged in then show the account links.
  			if sOrgRegistration AND lcl_cookie_userid <> "" AND lcl_cookie_userid <> "-1" then
        ShowLoggedinLinks ""
     else
        buildTopBarLink "LOGIN", sPath & "user_login.asp"
     end if

     response.write "                          </td>" & vbcrlf
     response.write "                      </tr>" & vbcrlf
     response.write "                    </table>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if
 'END: Show TopBar ------------------------------------------------------------

  response.write "            <tr valign=""top"">" & vbcrlf

 'BEGIN: Build the column widths ----------------------------------------------
  if blnMenuOn AND lcl_showsidemenubar AND session("deviceViewMode") <> "M" then
     lcl_sidemenubar_width = 200
     lcl_pagecontent_width = lcl_website_width - lcl_sidemenubar_width
     'lcl_pageheader_width  = lcl_website_width - lcl_sidemenubar_width
  else
     lcl_sidemenubar_width = 0
     lcl_pagecontent_width = lcl_website_width - lcl_sidemenubar_width
     'lcl_pageheader_width  = lcl_website_width - lcl_sidemenubar_width
  end if

  'lcl_column1_width = lcl_pageheader_width * 0.55
  'lcl_column2_width = lcl_pageheader_width * 0.45
 'END: Build the column widths ------------------------------------------------

 'BEGIN: Side Menubar (LEFT) --------------------------------------------------
  if blnMenuOn AND lcl_showsidemenubar AND lcl_sidemenubar_alignment = "LEFT" AND session("deviceViewMode") <> "M" then
     response.write "                <td rowspan=""2"" nowrap=""nowrap"" style=""width:" & lcl_sidemenubar_width & "px; background-color:#" & lcl_sidemenuoption_bgcolor & """>" & vbcrlf

     displaySideMenubar iorgid, lcl_sidemenuoption_bgcolor, lcl_sidemenuoption_bgcolorhover, lcl_sidemenuoption_alignment, lcl_cookie_userid, lcl_isEgovHomePage

     response.write "                </td>" & vbcrlf
  end if
 'END: Side Menubar (LEFT) ----------------------------------------------------

 'BEGIN: Page Header ----------------------------------------------------------
  'lcl_orgname_label = sOrgName

  'if getState(iorgid) <> "" then
  '   lcl_orgname_label = lcl_orgname_label & ", " & getState(iorgid)
  'end if

  'lcl_tagline = getOrgTagLine(iorgid)

  'if lcl_tagline <> "" then
  '   lcl_orgname_label = lcl_orgname_label & ", " & lcl_tagline
  'end if

 'Find the length of the page header minus the AddThis button width
  'lcl_pageheadertext_width = lcl_pageheader_width - 125

  'response.write "                <td colspan=""2"" style=""width:" & lcl_pageheader_width & "px;"" align=""left"" class=""pageHeader"">" & vbcrlf
  'response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"" class=""pageHeader_homePageMsg"">" & vbcrlf
  'response.write "                      <tr valign=""top"">" & vbcrlf
  'response.write "                          <td width=""" & lcl_pageheadertext_width & """ align=""" & lcl_pageheader_alignment & """>" & vbcrlf
  'response.write "                              <div class=""pageHeader_welcome"">" & lcl_orgname_label & " - CommunityLink</div>" & vbcrlf
  'response.write "                              <div class=""pageHeader_welcomeSubMsg"">Your connection to " & sOrgName & "</div><br />" & vbcrlf

 'Display the "page header" if the org has an "Edit Display" for the "homepage message".
  'if oOrg.OrgHasDisplay( "homepage message" ) then
  '   response.write "                           <span class=""pageHeader_homePageMsg"">" & vbcrlf
	 '			response.write                                oOrg.GetOrgDisplay("homepage message")
  '   response.write "                           </span>" & vbcrlf
  'end if

  'response.write "                          </td>" & vbcrlf
  'response.write "                          <td align=""right"" style=""padding-right:5px;"">" & vbcrlf
  '                                              if session("deviceViewMode") <> "M" then
  '                                                 displayAddThisButton iorgid
  '                                              end if

  '                                              getSocialSiteIcons "H", lcl_showRSS, lcl_url_twitter, lcl_url_facebook, _
  '                                                                 lcl_url_myspace, lcl_url_blogger

  'response.write "                          </td>" & vbcrlf
  'response.write "                      </tr>" & vbcrlf
  'response.write "                    </table>" & vbcrlf
  'response.write "                </td>" & vbcrlf
 'END: Page Header ------------------------------------------------------------

 'BEGIN: Side Menubar (RIGHT) -------------------------------------------------
  if blnMenuOn AND lcl_showsidemenubar AND lcl_sidemenubar_alignment = "RIGHT" AND session("deviceViewMode") <> "M" then
     response.write "                <td rowspan=""2"" nowrap=""nowrap"" style=""width:" & lcl_sidemenubar_width & "px; background-color:#" & lcl_sidemenuoption_bgcolor & """>" & vbcrlf

     displaySideMenubar lcl_sidemenuoption_bgcolor, lcl_sidemenuoption_bgcolorhover, lcl_sidemenuoption_alignment, lcl_cookie_userid, lcl_isEgovHomePage

     response.write "                </td>" & vbcrlf
  end if

  response.write "            </tr>" & vbcrlf
 'END: Side Menubar (RIGHT) ---------------------------------------------------

 'BEGIN: Page Content ---------------------------------------------------------
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td align=""left"" style=""width:" & lcl_pagecontent_width & "px; padding:5px;"">" & vbcrlf

 'BEGIN: Page specific code ---------------------------------------------------
  response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "                      <tr>" & vbcrlf
  response.write "                          <td>" & vbcrlf
  response.write "                              <font class=""pagetitle"">" & sOrgName & " " & lcl_actionlinelabel & " Status</font>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                    </table>" & vbcrlf

 'Get information for this request
  iTrackID = request("REQUEST_ID") 

  if IsNumeric(iTrackID) then
    	iTrackID = CStr(CDbl(iTrackID))
   		iTime    = Right(iTrackID,4)
   		iHour    = Left(iTime,2)
   		iMinute  = Right(iTime,2)

   		if iHour = "" OR iMinute = "" then
  		   	iHour = "99"
   	  		iMinute = "99"
   		end if

   		iID = replace(iTrackID,iTime,"")

   		if iID = "" then
   	  		iID = "000000"
   		end if

   		if lcl_orghasfeature_action_line_substatus then
   	    sSQL = "SELECT r.*, (select IsNull(s.status_name,NULL) "
   	    sSQL = sSQL &      " from egov_actionline_requests_statuses s "
       	sSQL = sSQL &      " where r.sub_status_id = s.action_status_id) AS sub_status_name "
   		   sSQL = sSQL & " FROM egov_actionline_requests r "
   		else
	      	sSQL = "SELECT r.*, NULL AS sub_status_name "
   		   sSQL = sSQL & " FROM egov_actionline_requests r "
   	 end if

   		sSQL = sSQL & " WHERE (r.action_autoid='" & iID & "') "
   	 sSQL = sSQL & " AND r.orgid = " & iorgid
   	 'sSQL = sSQL & " AND (DATEPART(hh, r.submit_date) = '"& iHour &"') "
   	 'sSQL = sSQL & " AND (DATEPART(mi, r.submit_date) = '"& iMinute &"')"

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
   			  			sComment = "<font color=""red"">No comment/description provided</font>"
        end if

       'Display description of Action Line Requests
        'response.write "<div class=""box_header4"">Action Line Item: " & sTitle & "</div>" & vbcrlf
        response.write "<div class=""box_header4"">" & lcl_actionlinelabel & " Item: " & sTitle & "</div>" & vbcrlf
        response.write "<div class=""group"">" & vbcrlf
        response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
        response.write "  <tr valign=""top"">" & vbcrlf
        response.write "      <td><strong>Your initial message:</strong><br /><br /></td>" & vbcrlf
        response.write "      <td align=""right"">" & vbcrlf

       'Determine if the "View PDF" button is displayed.
       ' 1. The org must be assigned the "requestmergeforms" feature
       ' 2. The form on the request has a PDF associated to it.
        if lcl_orghasfeature_requestmergeforms AND sPublicActionLinePDF <> "" then
           response.write "<input type=""button"" class=""button"" onClick=""window.open('viewPDF.asp?iRequestID=" & iID & "');"" value=""View Request in PDF Format"" />" & vbcrlf
        else
           response.write "&nbsp;" & vbcrlf
        end if

        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "  <tr>" & vbcrlf
        response.write "      <td colspan=""2""><i>" & sComment & "</i></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "</table>" & vbcrlf

       'BEGIN: Online Dialog Response -----------------------------------------
        response.write "<form name=""frmPost"" action=""action_request_lookup.asp"" method=""POST"">" & vbcrlf
        response.write "  <input type=""" & lcl_hidden & """ name=""iFormID"" value="""     & iFormID     & """ />" & vbcrlf
        response.write "  <input type=""" & lcl_hidden & """ name=""iCategoryId"" value=""" & iCategoryId & """ />" & vbcrlf
        response.write "  <input type=""" & lcl_hidden & """ name=""iUserID"" value="""     & iUserID     & """ />" & vbcrlf
        response.write "  <input type=""" & lcl_hidden & """ name=""sStatus"" value="""     & sStatus     & """ />" & vbcrlf
        response.write "  <input type=""" & lcl_hidden & """ name=""sSubStatus"" value="""  & sSubStatus  & """ />" & vbcrlf
        response.write "  <input type=""" & lcl_hidden & """ name=""REQUEST_ID"" value="""  & iTrackID    & """ />" & vbcrlf
        response.write "<div id=""post_form"" style=""padding:5px;margin-top:5px;border:solid 1px #000000;background-color:#E0E0E0;"">" & vbcrlf
        response.write "<table>" & vbcrlf
        response.write "  <tr>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <strong>Post a response/question:</strong><br />" & vbcrlf
        response.write "          <textarea onMouseOut=""this.style.backgroundColor='#ffffff';"" onMouseOver=""this.style.backgroundColor='#FFFFCC';"" name=""sMsg"" rows=""5"" cols=""80""></textarea>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "  <tr>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""submit"" value=""POST MESSAGE"" class=""button"" />" & vbcrlf
        response.write "          <input type=""reset"" value=""CLEAR MESSAGE"" class=""button"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "</table>" & vbcrlf
        response.write "</form>" & vbcrlf
        response.write "</div>" & vbcrlf
       'END: Online Dialog Response -------------------------------------------

       'Display Action Request Status
        if lcl_orghasfeature_action_line_substatus then
			   		   lcl_display_substatus = "<i>(Sub-Status)</i>"
   					else
			   		   lcl_display_substatus = ""
   					end if

        response.write "<p>" & vbcrlf
        'response.write "   <strong>Action Request Activity:</strong>" & vbcrlf
        response.write "   <strong>" & lcl_actionlinelabel & " Activity:</strong>" & vbcrlf
        response.write "   <div style=""margin-top:5px;border-top:solid 1px #000000;border-bottom:solid 1px #000000;background-color:#FFFFFF"">" & vbcrlf
        response.write "   <table>" & vbcrlf
        response.write "     <tr>" & vbcrlf
        response.write "         <td><strong>Status " & lcl_display_substatus & " - Date of Activity</strong></td>" & vbcrlf
        response.write "     </tr>" & vbcrlf
        response.write "   </table>" & vbcrlf
        response.write "   </div>" & vbcrlf

       'List History
     			List_Comments(iID)

        response.write "</p>" & vbcrlf

       'Hiding of contact info added 10/13/06 - Steve Loar
    				if not lcl_orghasfeature_hide_email_actionline then
				     	'Show CITY CONTACT INFORMATION
      					response.write "<p>" & vbcrlf
           response.write "   <strong>Email Contact:</strong>" & vbcrlf
           response.write "   <div style=""padding: 5px; margin-top:5px;border:solid 1px #000000;background-color:#ffffff"">" & vbcrlf

          'Get the "Assigned To" email address associated to the request.
      					sSQLa = "SELECT assigned_email FROM egov_action_request_view where action_autoid=" & iID
      					set oAssigned = Server.CreateObject("ADODB.Recordset")
      					oAssigned.Open sSQLa, Application("DSN") , 3, 1

           if not oAssigned.eof then
              lcl_assigned_email = oAssigned("assigned_email")
           else
              lcl_assigned_email = ""
           end if

           oAssigned.close
           set oAssigned = nothing

           response.write "   <strong>" & lcl_assigned_email & "</strong> has been assigned to this request. " & vbcrlf
           response.write "   Please contact via email - <a href=""mailto:" & lcl_assigned_email & """>" & lcl_assigned_email & "</a>" & vbcrlf
           response.write "   - for further information regarding this request." & vbcrlf
           response.write "   </div>" & vbcrlf
           response.write "</p>" & vbcrlf
        end if

        response.write "  </div>" & vbcrlf
        'response.write "</blockquote>
        response.write "</div>" & vbcrlf

   		else

        blnFound = False
        displayRequestNotFound iTrackID, sDefaultEmail, sorgname

     end if

     set oRequest = nothing

  else

		  'TrackID is non-numeric
   		blnFound = False
     displayRequestNotFound iTrackID, sDefaultEmail, sorgname
		
  end if
 'END: Page specific code -----------------------------------------------------

  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
 'END: CommunityLink Columns --------------------------------------------------

 'BEGIN: Display "switch to standard viewing" button --------------------------
  if session("accessdevice") = "BLACKBERRY" OR session("accessdevice") = "IPHONE" then
     response.write "            <tr>" & vbcrlf
     response.write "                <td align=""center"">" & vbcrlf
                                         displaySwitchViewModeLink sOrgName, session("deviceViewMode")
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if
 'END: Display "switch to standard viewing" button ----------------------------

 'BEGIN: Footer ---------------------------------------------------------------
  lcl_cityhome_label = oOrg.GetOrgDisplayName("homewebsitetag")

  if lcl_cityhome_label = "" then
     lcl_cityhome_label = "City Home"
  end if

  response.write "            <tr>" & vbcrlf
  response.write "                <td colspan=""3"" class=""topBar"">" & vbcrlf
  response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "                      <tr>" & vbcrlf
  response.write "                          <td>&nbsp;</td>" & vbcrlf
  response.write "                          <td align=""center"" class=""topBar"">" & vbcrlf
  response.write "                              <a href=""" & oOrg.GetOrgURL()  & """ class=""topBarOption"" target=""_top"">" & lcl_cityhome_label & "</a> |" & vbcrlf
  response.write "                              <a href=""" & oOrg.GetEgovURL() & """ class=""topBarOption"" target=""_top"">E-Gov Home</a>" & vbcrlf
                                                ShowPublicDefaultFooterNav iorgid, 2, lcl_isEgovHomePage
  response.write "                              <br />" & vbcrlf

  if oOrg.OrgHasDisplay("privacy policy") then
     response.write "<a href=""" & oOrg.GetEgovURL() & "/privacy_policy_display.asp"" class=""topBarOption"" target=""_top""><strong>Privacy Policy</strong></a> | " & vbcrlf
  end if

  if oOrg.OrgHasDisplay("refund policy") then
     response.write "<a href=""" & oOrg.GetEgovURL() & "/refund_policy.asp"" class=""topBarOption"" target=""_top"">Refund Policy</a> | " & vbcrlf
  end if

  response.write "                              <a href=""user_login.asp"" class=""topBarOption"" target=""_top"">Login</a>	| " & vbcrlf
  response.write "                              <a href=""register.asp"" class=""topBarOption"" target=""_top"">Register</a>" & vbcrlf

  response.write "                              <p style=""font-size:10px; color:#" & lcl_footer_fontcolor & """>" & vbcrlf
  response.write "                              Copyright &copy; 2004-" & year(now) & " Electronic Commerce Link, Inc. dba <a href=""http://www.egovlink.com"" target=""_NEW"" class=""topBarOption"">E-Gov Link</a>&nbsp;" & iDisplayTime & vbcrlf

 'Demo check to add admin link
  if lcl_orghasfeature_administrationlink then
     response.write "                           &nbsp;&nbsp;&nbsp;<a target=""_new"" href=""" & sEgovWebsiteURL & "/admin"" class=""topBarOption"">Administrator</a>" & vbcrlf
  end if

  response.write "                              </p>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                    </table>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
 'END: Footer -----------------------------------------------------------------

  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
 '-----------------------------------------------------------------------------

 'Determine if there are any inline javascripts to run
  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf
  end if

  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
Function List_Comments(iID)
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
    sBGColor = "#FFFFFF"
	
	if not oCommentList.eof then
	  	do while not oCommentList.eof
       sBGColor = changeBGColor(sBGColor,"#eeeeee","#ffffff")

       lcl_substatus_name = oCommentList("status_name")

    	  if lcl_substatus_name <> "" then
	 		      lcl_substatus_name = " <i>(" & lcl_substatus_name & ")</i>"
    	  end If

    			response.write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & """>" & vbcrlf
       response.write "<table>" & vbcrlf
     		response.write "  <tr>" & vbcrlf
       response.write "      <td>" & UCASE(oCommentList("action_status")) & lcl_substatus_name & " - " &  oCommentList("action_editdate") & "</td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
			
    			if oCommentList("action_externalcomment") <> "" then
      				lcl_comment_label = sOrgName
          lcl_comment_text  = oCommentList("action_externalcomment")
    			elseif oCommentList("action_citizen") <> "" then
      				lcl_comment_label = oCommentList("userfname") & " " & oCommentList("userlname")
          lcl_comment_text  = oCommentList("action_citizen")
       else
      				lcl_comment_label = sOrgName
          lcl_comment_text  = "Your request was reviewed and/or status was updated."
       end if

      	response.write "  <tr><td>&nbsp;&nbsp;&nbsp;<strong>" & lcl_comment_label & ": </strong><i>" & lcl_comment_text & "</i></td></tr>" & vbcrlf
    			response.write "</table>" & vbcrlf
       response.write "</div>" & vbcrlf

    			oCommentList.movenext
    loop
 else
  		response.write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & """>" & vbcrlf
    response.write "<table>" & vbcrlf
    response.write "  <tr>" & vbcrlf
    response.write "      <td><font color=red>&nbsp;&nbsp;&nbsp;<i>No activity</i></td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
    response.write "</table>" & vbcrlf
    response.write "</div>" & vbcrlf
 end if

 oCommentList.close
 set oCommentList = nothing

end function

'------------------------------------------------------------------------------
Function CheckSelected(sValue,sValue2)
	sReturnValue = ""
	If sValue = sValue2 Then
		sReturnValue = "SELECTED"
	End If

	CheckSelected = sReturnValue
End Function

'------------------------------------------------------------------------------
Function AddCommentTaskComment(sStatus,sCitizenMsg,iFormID,iUserID,iOrgId,iCreateDate)
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
Sub EmailComment( sCitizenMsg, iActionFormid, iActionId, iTrackingNo, iUserID, iCreateDate )
	Dim sSQLadmin, oAdmin, sSQLaddress, oAddress, sMsg2, objMail2

	sMsg2 = ""
 lcl_featurename_actionline = GetOrgFeatureName("action line")

'Get the user assigned to this request
	sSQLadmin = "SELECT assigned_userid, assigned_email "
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
         	adminFromAddr = "webmaster@eclink.com"
       else 
          adminFromAddr = oAdmin("assigned_email")  'ASSIGNED ADMIN USER EMAIL
       end if

       adminEmailAddr = oAdmin("assigned_email")   'ASSIGNED ADMIN USER EMAIL
       adminid        = oAdmin("assigned_userid")  'ASSIGNED ADMIN USER ID       
    end if
 end if

 oAdmin.Close
	Set oAdmin = Nothing

'BEGIN: Build message and send email to administrator(s) ----------------------
	sMsg2 = "This automated message was sent by the " & sOrgName & " E-Gov web site.  Do not reply to this message.  "

'Check to see if the org wants to hide their admin emails or not.
 if not lcl_orghasfeature_hide_email_actionline then
    sMsg2 = sMsg2 & "Contact " & adminFromAddr & " for inquiries regarding this email.  " & vbcrlf
 end if

	sMsg2 = sMsg2 & "A " & sOrgName & " " & lcl_actionlinelabel & " issue was updated on " & iCreateDate & "." & vbcrlf 
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
    lcl_from    = sOrgName & " " & lcl_featurename_actionline & " <webmaster@eclink.com>"
    lcl_subject = sOrgName & " " & lcl_featurename_actionline & ": User Comment Added"
	else
    lcl_from    = sOrgName & " ECLINK HELPDESK <webmaster@eclink.com>"
    lcl_subject = "ECLINK HELPDESK - User Comment Added"
	end if

'Remove the name from the email address
 lcl_validate_email = formatSendToEmail(adminEmailAddr)

 if isValidEmail(lcl_validate_email) then
   'Check for a delegate
    getDelegateInfo adminid, lcl_delegateid, lcl_delegate_username, lcl_delegate_useremail

   'Setup the SENDTO and check for a DELEGATE
    setupSendToAndDelegateEmails adminEmailAddr, lcl_delegate_useremail, lcl_email_sendto, lcl_email_cc

   'Send the email
    sendEmail "",lcl_email_sendto,lcl_email_cc,lcl_subject,lcl_message,"","Y"
 else
    ErrorCode = 1
 end if

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
sub displayRequestNotFound(iTrackID,iDefaultEmail,iOrgName)
  'response.write "<div style=""margin-left:20px;"" class=""box_header2"">Action Line Request Lookup</div>" & vbcrlf
  'response.write "<p>We could not locate an action line request using <strong>TRACKING NUMBER</strong> <strong>(" & iTrackID & ")</strong>.</p>" & vbcrlf
  response.write "<div style=""margin-left:20px;"" class=""box_header2"">" & lcl_actionlinelabel & " Lookup</div>" & vbcrlf
  response.write "<div class=""groupsmall"" style=""margin-left:20px;"">" & vbcrlf
  response.write "<p>We could not locate an " & lcl_actionlinelabel & " using <strong>TRACKING NUMBER</strong> <strong>(" & iTrackID & ")</strong>.</p>" & vbcrlf

  if iDefaultEmail = "" then
     sActionDefaultEmail = "webmaster@eclink.com"
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
