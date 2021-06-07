<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="class/classOrganization.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<!-- #include file="communitylink_global_functions_nostyles.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: communitylink.asp
' AUTHOR:    David Boyer
' CREATED:   04/14/09
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  The new CommunityLink interface!
'
' MODIFICATION HISTORY
' 1.0 04/29/09 David Boyer - Initial Version
' 1.1 05/14/09 David Boyer - Added check for mobile devices.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("communitylink") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 iStartSecs = timer
 sUserName  = ""
 lcl_onload = ""

 Dim sError, oOrg, lcl_scripts

'Capture current path
 session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
 session("RedirectLang") = "Return to CommunityLink"

 set oOrg = New classOrganization

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

 lcl_pagetitle = "CommunityLink"

'Check for cookies
 lcl_cookie_userid = request.cookies("userid")

'Check to see if any Mayor's Blog images exist and if so resize the borders around the image.
 if session("deviceViewMode") <> "M" then
    lcl_onload = lcl_onload & "resizeBlogImgBorders();"
 end if

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
                      lcl_sidemenuoption_fontcolor, lcl_sidemenuoption_fontcolorhover, lcl_pageheader_alignment, _
                      lcl_pageheader_fontcolor, lcl_pageheader_fonttype, lcl_pageheader_bgcolor, lcl_showRSS, lcl_url_twitter, _
                      lcl_url_facebook, lcl_url_myspace, lcl_url_blogger

'Check for org features
 lcl_orghasfeature_administrationlink = orghasfeature(iorgid,"AdministrationLink")

'Set up the base path for images if user is on secure server
 sImgBaseURL = getImgBaseURL(sEgovWebsiteURL)

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

<%

 'BEGIN: Javascripts ----------------------------------------------------------
  if session("deviceViewMode") <> "M" then
     response.write "<script type=""text/javascript"" src=""https://s7.addthis.com/js/200/addthis_widget.js""></script>" & vbcrlf
     response.write "<script type=""text/javascript"">var addthis_pub=""cschappacher"";</script>" & vbcrlf
  end if
%>
<script language="javascript">
<!--
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

function resizeBlogImgBorders() {
  lcl_total_blogimgs = document.getElementsByName("blogimg").length;

  for (i=1; i<=lcl_total_blogimgs; ++ i) {
       if(document.getElementById("blogimg_"+i)) {
          //Get the height of the blog image.
          lcl_total_height = document.getElementById("blogimg_"+i).height;

          //Get the widths of the blog images.
          lcl_blogimgleft_width  = document.getElementById("blogimg_left_"+i).width;
          lcl_blogimgright_width = document.getElementById("blogimg_right_"+i).width;
          lcl_blogimg_width      = document.getElementById("blogimg_"+i).width;
          lcl_total_width        = lcl_blogimgleft_width + lcl_blogimgright_width + lcl_blogimg_width;

          //Adjust the top/bottom blog images with the proper width.
          document.getElementById("blogimg_top_"+i).width    = lcl_total_width;
          document.getElementById("blogimg_bottom_"+i).width = lcl_total_width;

          //Adjust the left/right blog images with the proper height.
          document.getElementById("blogimg_left_"+i).height  = lcl_total_height
          document.getElementById("blogimg_right_"+i).height = lcl_total_height;
       }
  }
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
 'response.write "deviceViewMode: [" & session("deviceViewMode") & "] - iAccessDevice: [" & iAccessDevice & "] "

 'BEGIN: CommunityLink --------------------------------------------------------
  lcl_website_width = getWebsiteWidth(lcl_website_size, lcl_website_size_customsize)

  response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"" align=""" & lcl_website_alignment & """ bgcolor=""" & lcl_website_bgcolor & """>" & vbcrlf
  response.write "          <table border=""0"" bordercolor=""#ff0000"" cellspacing=""0"" cellpadding=""0"" width=""" & lcl_website_width & """ bgcolor=""#ffffff"">" & vbcrlf

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
     response.write "                <td colspan=""3"" align=""" & lcl_logo_alignment & """ onclick=""location.href='" & lcl_egovSiteURL & "'"">" & vbcrlf
     response.write "                    <div><a href=""" & lcl_egovSiteURL & """>" & lcl_orgLogo & "</a></div>" & vbcrlf
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
     response.write "                <td colspan=""3"">" & vbcrlf
     response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
     response.write "                      <tr>" & vbcrlf
     response.write "                          <td align=""left"">" & vbcrlf
     'response.write "                              <img src=""" & sEgovWebsiteURL & "/images/accountmenu.jpg"" />" & vbcrlf
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
     lcl_pageheader_width  = lcl_website_width - lcl_sidemenubar_width
  else
     lcl_sidemenubar_width = 0
     lcl_pageheader_width  = lcl_website_width - lcl_sidemenubar_width
  end if

  lcl_column1_width = lcl_pageheader_width * 0.55
  lcl_column2_width = lcl_pageheader_width * 0.45
 'END: Build the column widths ------------------------------------------------

 'BEGIN: Side Menubar (LEFT) --------------------------------------------------
  if blnMenuOn AND lcl_showsidemenubar AND lcl_sidemenubar_alignment = "LEFT" AND session("deviceViewMode") <> "M" then
     response.write "                <td rowspan=""2"" nowrap=""nowrap"">" & vbcrlf

     displaySideMenubar iorgid, lcl_sidemenuoption_bgcolor, lcl_sidemenuoption_bgcolorhover, lcl_sidemenuoption_alignment, lcl_cookie_userid, lcl_isEgovHomePage

     response.write "                </td>" & vbcrlf
  end if
 'END: Side Menubar (LEFT) ----------------------------------------------------

 'BEGIN: Page Header ----------------------------------------------------------
  lcl_orgname_label = sOrgName

  if getState(iorgid) <> "" then
     lcl_orgname_label = lcl_orgname_label & ", " & getState(iorgid)
  end if

  lcl_tagline = getOrgTagLine(iorgid)

  if lcl_tagline <> "" then
     lcl_orgname_label = lcl_orgname_label & ", " & lcl_tagline
  end if

 'Find the length of the page header minus the AddThis button width
  lcl_pageheadertext_width = lcl_pageheader_width - 125

  response.write "                <td colspan=""2"" align=""left"">" & vbcrlf
  response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
  response.write "                      <tr valign=""top"">" & vbcrlf
  response.write "                          <td width=""" & lcl_pageheadertext_width & """ align=""" & lcl_pageheader_alignment & """>" & vbcrlf
  response.write "                              <div>" & lcl_orgname_label & " - CommunityLink</div>" & vbcrlf
  response.write "                              <div>Your connection to " & sOrgName & "</div><br />" & vbcrlf

 'Display the "page header" if the org has an "Edit Display" for the "homepage message".
  if oOrg.OrgHasDisplay( "homepage message" ) then
     response.write "                           <span>" & vbcrlf
	 			response.write                                oOrg.GetOrgDisplay("homepage message")
     response.write "                           </span>" & vbcrlf
  end if

  response.write "                          </td>" & vbcrlf
  response.write "                          <td align=""right"">" & vbcrlf
                                                if session("deviceViewMode") <> "M" then
                                                   displayAddThisButton iorgid
                                                end if

                                                getSocialSiteIcons "H", lcl_showRSS, lcl_url_twitter, lcl_url_facebook, _
                                                                   lcl_url_myspace, lcl_url_blogger

  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                    </table>" & vbcrlf
  response.write "                </td>" & vbcrlf
 'END: Page Header ------------------------------------------------------------

 'BEGIN: Side Menubar (RIGHT) -------------------------------------------------
  if blnMenuOn AND lcl_showsidemenubar AND lcl_sidemenubar_alignment = "RIGHT" AND session("deviceViewMode") <> "M" then
     response.write "                <td rowspan=""2"" nowrap=""nowrap"">" & vbcrlf

     displaySideMenubar lcl_sidemenuoption_bgcolor, lcl_sidemenuoption_bgcolorhover, lcl_sidemenuoption_alignment, lcl_cookie_userid, lcl_isEgovHomePage

     response.write "                </td>" & vbcrlf
  end if

  response.write "            </tr>" & vbcrlf
 'END: Side Menubar (RIGHT) ---------------------------------------------------

 'BEGIN: CommunityLink Columns ------------------------------------------------
  response.write "            <tr valign=""top"">" & vbcrlf
                                  displayPortalSections "CL", 1, iorgid, sOrgRegistration, lcl_cookie_userid, "Y", lcl_column1_width, "Y" 

 'If user is accessing from a mobile device then put columns in separate rows.
  if session("deviceViewMode") = "M" then
     response.write "            </tr>" & vbcrlf
     response.write "            <tr valign=""top"">" & vbcrlf
  end if

                                  displayPortalSections "CL", 2, iorgid, sOrgRegistration, lcl_cookie_userid, "Y", lcl_column2_width, "Y"
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
  response.write "                <td colspan=""3"">" & vbcrlf
  response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "                      <tr>" & vbcrlf
  response.write "                          <td>&nbsp;</td>" & vbcrlf
  response.write "                          <td align=""center"">" & vbcrlf
  response.write "                              <a href=""" & oOrg.GetOrgURL()  & """ target=""_top"">" & lcl_cityhome_label & "</a> |" & vbcrlf
  response.write "                              <a href=""" & oOrg.GetEgovURL() & """ target=""_top"">E-Gov Home</a>" & vbcrlf
                                                ShowPublicDefaultFooterNav iorgid, 2, lcl_isEgovHomePage
  response.write "                              <br />" & vbcrlf

  if oOrg.OrgHasDisplay("privacy policy") then
     response.write "<a href=""" & oOrg.GetEgovURL() & "/privacy_policy_display.asp"" target=""_top""><strong>Privacy Policy</strong></a> | " & vbcrlf
  end if

  if oOrg.OrgHasDisplay("refund policy") then
     response.write "<a href=""" & oOrg.GetEgovURL() & "/refund_policy.asp"" target=""_top"">Refund Policy</a> | " & vbcrlf
  end if

  response.write "                              <a href=""user_login.asp"" target=""_top"">Login</a>	| " & vbcrlf
  response.write "                              <a href=""register.asp"" target=""_top"">Register</a>" & vbcrlf

  response.write "                              <p>" & vbcrlf
  response.write "                              Copyright &copy; 2004-" & year(now) & " Electronic Commerce Link, Inc. dba <a href=""http://www.egovlink.com"" target=""_NEW"">E-Gov Link</a>&nbsp;" & iDisplayTime & vbcrlf

 'Demo check to add admin link
  if lcl_orghasfeature_administrationlink then
     response.write "                           &nbsp;&nbsp;&nbsp;<a target=""_new"" href=""" & sEgovWebsiteURL & "/admin"">Administrator</a>" & vbcrlf
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
%>
