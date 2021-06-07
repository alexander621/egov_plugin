<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="includes/time.asp" //-->
<!-- #include file="class/classOrganization.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<!-- #include file="communitylink_global_functions.asp" //-->
<!-- #include file="communitylink_global_styles.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: communitylink_section.asp
' AUTHOR:    David Boyer
' CREATED:   10/27/09
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows developers to access a specific CommunityLink section!
'
' MODIFICATION HISTORY
' 1.0 10/27/09 David Boyer - Initial Version
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

 lcl_pagetitle = "CommunityLink"

'Check for cookies
 lcl_cookie_userid = request.cookies("userid")

'Check for a CommunityLink record for the org.
'If one DOES exist then pull all of the values.
'If one does NOT exist then create it and enter the default values.
 lcl_communitylinkid = getCommunityLinkID(iorgid,lcl_cookie_userid)

'Retrieve the CommunityLink record.
 getCommunityLinkInfo lcl_communitylinkid, _
                      iorgid, _
                      lcl_isEgovHomePage, _
                      lcl_website_size, _
                      lcl_website_size_customsize, _
                      lcl_website_alignment, _
                      lcl_website_bgcolor, _
                      lcl_showlogo, _
                      lcl_logo_filename, _
                      lcl_logo_filenamebg, _
                      lcl_logo_alignment, _
                      lcl_showtopbar, _
                      lcl_topbar_bgcolor, _
                      lcl_topbar_fonttype, _
                      lcl_topbar_fontcolor, _
                      lcl_topbar_fontcolorhover, _
                      lcl_showsidemenubar, _
                      lcl_sidemenubar_alignment, _
                      lcl_sidemenuoption_bgcolor, _
                      lcl_sidemenuoption_bgcolorhover, _
                      lcl_sidemenuoption_alignment, _
                      lcl_sidemenuoption_fonttype, _
                      lcl_sidemenuoption_fontcolor, _
                      lcl_sidemenuoption_fontcolorhover, _
                      lcl_showpageheader, _
                      lcl_pageheader_alignment, _
                      lcl_pageheader_fontsize, _
                      lcl_pageheader_fontcolor, _
                      lcl_pageheader_fonttype, _
                      lcl_pageheader_bgcolor, _
                      lcl_showfooter, _
                      lcl_footer_bgcolor, _
                      lcl_footer_fonttype, _
                      lcl_footer_fontcolor, _
                      lcl_footer_fontcolorhover, _
                      lcl_showRSS, _
                      lcl_url_twitter, _
                      lcl_url_facebook, _
                      lcl_url_myspace, _
                      lcl_url_blogger

'Check for org features
 'lcl_orghasfeature_administrationlink = orghasfeature(iorgid,"AdministrationLink")

'Set up the base path for images if user is on secure server
 sImgBaseURL = getImgBaseURL(sEgovWebsiteURL)

'Build the Title
 lcl_title = sOrgName

 if iorgid <> 7 then
    lcl_title = "E-Gov Services " & lcl_title
 end if

'Determine which section to display
 if request("fn") <> "" then
    if containsApostrophe(request("fn")) then
       iFeatureName = "NONE"
    else
       iFeatureName = request("fn")
    end if
 else
    iFeatureName = "NONE"
 end if

'Determine which CommunityLink styling options to display.
'  CL    = CommunityLink (E-Gov)
'  SAVVY = Savvy/IFRAME individual section displays
 if request("ptype") <> "" then
    if not containsApostrophe(request("ptype")) then
       if UCASE(request("ptype")) = "SAVVY" OR UCASE(request("ptype")) = "CL" then  'This defeats a long-time hack attempt
          iPType = UCASE(request("ptype"))
       end if
    end if
 else
    iPType = "CL"
 end if

'Determine if there is a border around the section
'Values: ON/OFF
 iSectionBorder      = "1"
 iSectionBorderColor = "000000"

 if request("sectionborder") <> "" then
    if UCASE(request("sectionborder")) = "OFF" then
       iSectionBorder = "0"
    else
       if request("sectionbordercolor") <> "" then
          iSectionBorderColor = request("sectionbordercolor")
       end if
    end if
 end if

'Determine if the user is overriding the background color for individual sections.
 if request("bgcolor") <> "" then
    lcl_website_bgcolor = request("bgcolor")
 end if

 lcl_website_width = "100%"
%>
<html>
<head>
 	<title><%=lcl_title%></title>
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />

  <!--	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" /> -->
 	<link rel="stylesheet" type="text/css" href="css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="global.css" />
 	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

<script language="javascript">
<!--
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

function changeBodyBG(p_bgcolor) {
  if(p_bgcolor != '' && p_bgcolor != undefined) {
     document.getElementById('sectionBody').style.backgroundColor = p_bgcolor;
  }
}

window.onload = function(){
  <%=lcl_onload%>
}
//-->
</script>

<style>

  #sectionBody {
     background-color: #<%=lcl_website_bgcolor%>;
     margin:           0 0;
  }

  #table_body {
     width:            <%=lcl_website_width%>;
     background-color: #ff0000;
     border:           <%=iSectionBorder%>pt solid #<%=iSectionBorderColor%>;
  }

</style>
</head>
<%
  'response.write "<body bgcolor=""#" & lcl_website_bgcolor & """ leftmargin=""0"" topmargin=""0"" marginheight=""0"" marginwidth=""0"">" & vbcrlf
  response.write "<body id=""sectionBody"">" & vbcrlf

 'response.write "deviceViewMode: [" & session("deviceViewMode") & "] - iAccessDevice: [" & iAccessDevice & "] "

 'BEGIN: CommunityLink --------------------------------------------------------
  'lcl_website_width = getWebsiteWidth(lcl_website_size, lcl_website_size_customsize)

  response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"" align=""" & lcl_website_alignment & """ bgcolor=""" & lcl_website_bgcolor & """>" & vbcrlf
  'response.write "          <table border=""0"" bordercolor=""#ff0000"" cellspacing=""0"" cellpadding=""0"" width=""" & lcl_website_width & """ bgcolor=""#ffffff"" style=""border:" & iSectionBorder & "pt solid #" & iSectionBorderColor & ";"">" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" id=""table_body"">" & vbcrlf

 'BEGIN: CommunityLink Columns ------------------------------------------------
  response.write "            <tr valign=""top"">" & vbcrlf

                                  lcl_column_num   = 1
                                  lcl_wrap_td_tags = "Y"

                                  displayPortalSections iPType, _
                                                        lcl_column_num, _
                                                        iorgid, _
                                                        sOrgRegistration, _
                                                        lcl_cookie_userid, _
                                                        lcl_wrap_td_tags, _
                                                        lcl_column1_width, _
                                                        lcl_showRSS, _
                                                        iFeatureName

 'If user is accessing from a mobile device then put columns in separate rows.
  if session("deviceViewMode") = "M" then
     response.write "            </tr>" & vbcrlf
     response.write "            <tr valign=""top"">" & vbcrlf
  end if

                                  lcl_column_num   = 2
                                  lcl_wrap_td_tags = "Y"

                                  displayPortalSections iPType, _
                                                        lcl_column_num, _
                                                        iorgid, _
                                                        sOrgRegistration, _
                                                        lcl_cookie_userid, _
                                                        lcl_wrap_td_tags, _
                                                        lcl_column2_width, _
                                                        lcl_showRSS, _
                                                        iFeatureName

  response.write "            </tr>" & vbcrlf
 'END: CommunityLink Columns --------------------------------------------------

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