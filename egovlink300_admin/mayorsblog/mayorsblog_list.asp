<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mayorsblog_list.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the entries in the Blog
'
' MODIFICATION HISTORY
' 1.0 03/30/09 David Boyer - Initial Version
' 1.1 07/23/09 David Boyer - Changed the "Post a Comment" dropdown list to rely on Action Line and not CommunityLink feature.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("mayorsblog") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"mayorsblog") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 lcl_rssType   = "MAYORSBLOG"
 lcl_pagetitle = "Blog"
 lcl_success   = request("success")

'Check for a screen message
 lcl_onload = ""

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Used in custom reporting
 session("RSSType") = lcl_rssType

'Check for org features
 lcl_orghasfeature_rssfeeds_mayorsblog                 = orghasfeature("rssfeeds_mayorsblog")
 lcl_orghasfeature_action_line                         = orghasfeature("action line")
 lcl_orghasfeature_maintainactionlineform_postacomment = orghasfeature("maintainactionlineform_postacomment")
 'lcl_orghasfeature_communitylink       = orghasfeature("communitylink")

'Check for user permissions
 lcl_userhaspermission_rssfeeds_mayorsblog                 = userhaspermission(session("userid"),"rssfeeds_mayorsblog")
 lcl_userhaspermission_action_line                         = userhaspermission(session("userid"),"action line")
 lcl_userhaspermission_maintainactionlineform_postacomment = userhaspermission(session("userid"),"maintainactionlineform_postacomment")
 'lcl_userhaspermission_communitylink       = userhaspermission(session("userid"),"communitylink")

'Retrieve the search options
 lcl_sc_fromcreatedate = ""
 lcl_sc_tocreatedate   = ""
 lcl_sc_title          = ""
 lcl_sc_userid         = 0
 lcl_sc_orderby        = "createdate"

 if request("sc_fromcreatedate") <> "" then
    lcl_sc_fromcreatedate = request("sc_fromcreatedate")
 end if

 if request("sc_tocreatedate") <> "" then
    lcl_sc_tocreatedate = request("sc_tocreatedate")
 end if

 if request("sc_title") <> "" then
    lcl_sc_title = request("sc_title")
 end if

 if request("sc_userid") <> "" then
    lcl_sc_userid = request("sc_userid")
 end if

 if request("sc_orderby") <> "" then
    lcl_sc_orderby = request("sc_orderby")
 end if
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
<!--
function confirm_delete(blogid) {
  lcl_blog = document.getElementById("blog"+blogid).innerHTML;

 	if (confirm("Are you sure you want to delete '" + lcl_blog + "' ?")) { 
  				//DELETE HAS BEEN VERIFIED
		  		location.href='mayorsblog_delete.asp?blogid='+ blogid;
		}
}

function viewRSSLog(pID) {
  lcl_width  = 900;
  lcl_height = 400;
  lcl_left   = (screen.availWidth/2) - (lcl_width/2);
  lcl_top    = (screen.availHeight/2) - (lcl_height/2);
		popupWin = window.open("../customreports/customreports.asp?CR=RSSLOG&id=" + pID, "_blank","resizable,width=" + lcl_width + ",height=" + lcl_height + ",left=" + lcl_left + ",top=" + lcl_top);
}

function sendToRSS(pID) {
  var sParameter = 'id=' + encodeURIComponent(pID);
  sParameter    += '&isAjax=Y';

  doAjax('mayorsblog_sendToRSS.asp', sParameter, 'displayScreenMsg', 'post', '0');
}

function validateFields() {
  var lcl_false_count = 0;
		var daterege        = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		var dateFromOk      = daterege.test(document.getElementById("sc_fromcreatedate").value);
		var dateToOk        = daterege.test(document.getElementById("sc_tocreatedate").value);

  if (document.getElementById("sc_tocreatedate").value!="") {
   		if (! dateToOk ) {
         document.getElementById("sc_tocreatedate").focus();
         inlineMsg(document.getElementById("toDateCalPop").id,'<strong>Invalid Value: </strong> The "To Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'toDateCalPop');
         lcl_false_count = lcl_false_count + 1;
     }else{
         clearMsg("toDateCalPop");
     }
  }

  if (document.getElementById("sc_fromcreatedate").value!="") {
   		if (! dateFromOk ) {
         document.getElementById("sc_fromcreatedate").focus();
         inlineMsg(document.getElementById("fromDateCalPop").id,'<strong>Invalid Value: </strong> The "From Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'fromDateCalPop');
         lcl_false_count = lcl_false_count + 1;
     }else{
         clearMsg("fromDateCalPop");
     }
  }

  if(lcl_false_count > 0) {
     return false;
  }else{
     document.getElementById("searchMayorsBlog").submit();
     return true;
  }
}

function openPic(iID,iFile) {
  w = 300;
  h = 300;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);

  eval('window.open("' + iFile + '", "_blogimg", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
}

function doCalendar(ToFrom) {
  w = 350;
  h = 250;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
}

function updatePostComment() {
  lcl_formid = document.getElementById("CL_postcomment_formid").value;

  //Build the parameter string
  var sParameter = 'orgid='    + encodeURIComponent("<%=session("orgid")%>");
  sParameter    += '&feature=' + encodeURIComponent("mayorsblog");
  sParameter    += '&formid='  + encodeURIComponent(lcl_formid);
  sParameter    += '&isAjaxRoutine=Y';

  doAjax('../communitylink/saveCommunityLinkOptions.asp', sParameter, 'displayScreenMsg', 'post', '0');
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
 	<div id="centercontent">

<table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
  <tr>
      <td valign="top">
          <div style="margin-top:20px; margin-left:20px;">
            <table border="0" cellspacing="0" cellpadding="0" width="1000px">
              <tr>
                  <td><font size="+1"><strong><%=Session("sOrgName")%>&nbsp;<%=lcl_pagetitle%>s</strong></font></td>
                  <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
              </tr>
            </table>
            <table border="0" cellspacing="0" cellpadding="0">
              <form name="searchMayorsBlog" id="searchMayorsBlog" action="mayorsblog_list.asp" method="post">
              <tr>
                  <td>
                      <fieldset>
                        <legend>Search Options&nbsp;</legend>
                          <p>
                          <table border="0" cellspacing="0" cellpadding="2" width="600">
                            <tr>
                                <td width="100">Create Date:</td>
                                <td>
                                    From: <input type="text" name="sc_fromcreatedate" id="sc_fromcreatedate" value="<%=lcl_sc_fromcreatedate%>" size="10" maxlength="10" onchange="clearMsg('fromDateCalPop');" />
                                          <a href="javascript:void doCalendar('sc_fromcreatedate');"><img src="../images/calendar.gif" id="fromDateCalPop" border="0" onclick="clearMsg('fromDateCalPop');" /></a>&nbsp;&nbsp;
                                    To: <input type="text" name="sc_tocreatedate" id="sc_tocreatedate" value="<%=lcl_sc_tocreatedate%>" size="10" maxlength="10" onchange="clearMsg('toDateCalPop');" />
                                        <a href="javascript:void doCalendar('sc_tocreatedate');"><img src="../images/calendar.gif" id="toDateCalPop" border="0" onclick="clearMsg('toDateCalPop');" /></a>
                                </td>
                            </tr>
                            <tr>
                                <td>Title:</td>
                                <td><input type="text" name="sc_title" id="sc_title" value="<%=lcl_sc_title%>" size="50" maxlength="500" /></td>
                            </tr>
                            <tr>
                                <td>Blog Owner:</td>
                                <td>
                                    <select name="sc_userid" id="sc_userid">
                                      <option value=""></option>
                                      <% showBlogOwners lcl_sc_userid %>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>Order By:</td>
                                <td>
                                  <%
                                    checkForOrderBySelected lcl_sc_orderby, lcl_selected_createdate, lcl_selected_blogowner, _
                                                                            lcl_selected_createdby, lcl_selected_active

                                    response.write "<select name=""sc_orderby"" id=""sc_orderby"">" & vbcrlf
                                    response.write "  <option value=""createdate""" & lcl_selected_createdate & ">Create Date</option>" & vbcrlf
                                    response.write "  <option value=""blogowner"""  & lcl_selected_blogowner  & ">Blog Owner</option>" & vbcrlf
                                    response.write "  <option value=""createdby"""  & lcl_selected_createdby  & ">Created By</option>" & vbcrlf
                                    response.write "  <option value=""active"""     & lcl_selected_active     & ">Active</option>" & vbcrlf
                                    response.write "</select>" & vbcrlf
                                  %>
                                </td>
                            </tr>
                            <tr><td colspan="2"><input type="button" name="searchButton" id="searchButton" value="Search" onclick="return validateFields();" /></td></tr>
                          </table>
                          </p>
                      </fieldset>
                  </td>
              </tr>
              </form>
            </table>
            <br />
            <table border="0" cellspacing="0" cellpadding="0" style="margin-bottom:5px;">
              <tr>
                  <td><input type="button" name="newButton" id="newButton" value="New Blog Entry" class="button" onclick="window.location='mayorsblog_maint.asp';" /></td>
                  <td align="right">
                  <%
                   'Post a Comment - Action Line form
                    if lcl_orghasfeature_action_line AND lcl_orghasfeature_maintainactionlineform_postacomment then
                    'if lcl_orghasfeature_communitylink AND lcl_userhaspermission_communitylink then
                       lcl_actionline_label = GetFeatureName("action line")
                       lcl_comments_formid  = getCommentsFormID(session("orgid"), "", "mayorsblog")

                       response.write "<span style=""color:#800000"">" & vbcrlf
                       'response.write "CommunityLink: Link for ""Post a Comment"" (" & lcl_actionline_label & " Requests):" & vbcrlf
                       response.write "Action Line Request to be used for ""Post a Comment""" & vbcrlf
                       response.write "</span><br />" & vbcrlf

                      'If the user has the proper permission then allow then to maintain the list.
                      'If not then display the form selected.
                       if lcl_userhaspermission_maintainactionlineform_postacomment then
                          response.write "<select name=""CL_postcomment_formid"" id=""CL_postcomment_formid"" onchange=""updatePostComment();"">" & vbcrlf
                                            displayActionLineForms session("orgid"), lcl_comments_formid, "Y"
                          response.write "</select>" & vbcrlf
                       else
                          lcl_actionline_formname = getActionLineFormName(session("orgid"), lcl_comments_formid)

                          response.write "[" & lcl_actionline_formname & "]" & vbcrlf
                       end if
                    else
                       response.write "&nbsp;" & vbcrlf
                    end if
                  %>
                  </td>
              </tr>
            </table>
            <% listBlogEntries lcl_sc_fromcreatedate, lcl_sc_tocreatedate, lcl_sc_title, lcl_sc_userid, lcl_sc_orderby %>
          </div>
      </td>
  </tr>
</table>

  </div>
</div>
	
<!--#Include file="../admin_footer.asp"--> 

</body>
</html>
<%
'------------------------------------------------------------------------------
sub listBlogEntries(p_sc_fromcreatedate, p_sc_tocreatedate, p_sc_title, p_sc_userid, p_sc_orderby)
 	Dim iRowCount

 	iRowCount = 0

  sSQL = "SELECT b.blogid, b.userid, u.firstname + ' ' + u.lastname AS blogowner, b.title, b.createdbyid, b.createdbydate, "
  sSQL = sSQL & " b.lastmodifiedbyid, b.lastmodifiedbydate, b.isInactive, u.imagefilename, "
  sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS createdbyname, u3.firstname + ' ' + u3.lastname AS lastmodifiedbyname "
  sSQL = sSQL & " FROM egov_mayorsblog b "
  sSQL = sSQL &      " LEFT OUTER JOIN users u ON b.userid = u.userid AND u.orgid = " & session("orgid")
  sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON b.createdbyid = u2.userid AND u2.orgid = " & session("orgid")
  sSQL = sSQL &      " LEFT OUTER JOIN users u3 ON b.lastmodifiedbyid = u3.userid AND u3.orgid = " & session("orgid")
  sSQL = sSQL & " WHERE b.orgid = " & session("orgid")

 'Setup the WHERE clause with the search option values.
  if trim(p_sc_fromcreatedate) <> "" then
     sSQL = sSQL & " AND b.createdbydate >= CAST('" & p_sc_fromcreatedate & "' as datetime) "
  end if

  if trim(p_sc_tocreatedate) <> "" then
     sSQL = sSQL & " AND b.createdbydate <= CAST('" & p_sc_tocreatedate & "' as datetime) "
  end if

  if trim(p_sc_userid) <> "" AND p_sc_userid > 0 then
     sSQL = sSQL & " AND b.userid = " & p_sc_userid
  end if

  if trim(p_sc_title) <> "" then
     sSQL = sSQL & " AND UPPER(b.title) LIKE ('%" & UCASE(p_sc_title) & "%') "
  end if

 'Setup the ORDER BY
  lcl_orderby = "b.createdbydate DESC"

  if trim(p_sc_orderby) <> "" then
     lcl_sc_orderby = trim(UCASE(p_sc_orderby))

     if lcl_sc_orderby = "BLOGOWNER" then
        lcl_orderby = "u.lastname, u.firstname, b.createdbydate DESC"
     elseif lcl_sc_orderby = "CREATEDBY" then
        lcl_orderby = "u2.lastname, u2.firstname, b.createdbydate DESC"
     elseif lcl_sc_orderby = "ACTIVE" then
        lcl_orderby = "b.isInactive DESC, b.createdbydate DESC"
     end if
  end if

  sSQL = sSQL & " ORDER BY " & lcl_orderby

 	set oBlog = Server.CreateObject("ADODB.Recordset")
	 oBlog.Open sSQL, Application("DSN"), 3, 1
	
 	if not oBlog.eof then
   		response.write "<div class=""shadow"">" & vbcrlf
 		  response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" style=""width:1000px"">" & vbcrlf
   		response.write "  <tr align=""left"">" & vbcrlf
     response.write "      <th>Title</th>" & vbcrlf
     response.write "      <th nowrap=""nowrap"" colspan=""2"">Blog Owner</th>" & vbcrlf
     response.write "      <th align=""center"">Active</th>" & vbcrlf
     response.write "      <th>&nbsp;</th>" & vbcrlf

     if lcl_orghasfeature_rssfeeds_mayorsblog AND lcl_userhaspermission_rssfeeds_mayorsblog then
        response.write "      <th align=""center"" nowrap=""nowrap"">Send<br />to RSS</th>" & vbcrlf
        response.write "      <th align=""center""  nowrap=""nowrap"">RSS<br />Send Log</th>" & vbcrlf
     end if

    response.write "      <th nowrap=""nowrap"">Created By</th>" & vbcrlf
    response.write "      <th nowrap=""nowrap"">Last Modified By</th>" & vbcrlf
    response.write "  </tr>" & vbcrlf

    lcl_bgcolor             = "#ffffff"
    lcl_original_categoryid = 0

    do while not oBlog.eof
       lcl_bgcolor  = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
    			iRowCount    = iRowCount + 1

      'Determine if the blog is/isn't active
       if oBlog("isInactive") then
          lcl_active = "&nbsp;"
       else
          lcl_active = "Y"
       end if

      'Determine if the blog owner has a picture.
       lcl_imagefilename = oBlog("imagefilename")

       if lcl_imagefilename <> "" then
          'lcl_blogimage_url = session("egovclientwebsiteurl")
          'lcl_blogimage_url = lcl_blogimage_url & "/admin/custom/pub/"
          'lcl_blogimage_url = lcl_blogimage_url & session("virtualdirectory")
          'lcl_blogimage_url = lcl_blogimage_url & "/unpublished_documents"
          'lcl_blogimage_url = lcl_blogimage_url & oBlog("imagefilename")

          if left(lcl_imagefilename,1) <> "/" then
             lcl_imagefilename = "/" & lcl_imagefilename
          end if

          lcl_blogimage_url = Application("CommunityLink_DocUrl")
          lcl_blogimage_url = lcl_blogimage_url & "/public_documents300/"
          lcl_blogimage_url = lcl_blogimage_url & session("virtualdirectory")
          lcl_blogimage_url = lcl_blogimage_url & "/unpublished_documents"
          lcl_blogimage_url = lcl_blogimage_url & lcl_imagefilename
       else
          lcl_blogimage_url = session("egovclientwebsiteurl")
          lcl_blogimage_url = lcl_blogimage_url & "/admin/images/notavailable_person.jpg"
       end if

       lcl_blog_image = "<img src=""" & lcl_blogimage_url & """ name=""blogimg_" & iRowCount & """ id=""blogimg_" & iRowCount & """ width=""20"" height=""20"" align=""left"" style=""border:1px solid #000000"" onclick=""openPic(" & iRowCount & ",'" & lcl_blogimage_url & "')"" />&nbsp;" & vbcrlf

      'Setup the onclick
       lcl_row_onclick = "location.href='mayorsblog_maint.asp?blogid=" & oBlog("blogid") & "';"

       response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
       response.write "      <td class=""formlist"" title=""click to edit"" onClick=""" & lcl_row_onclick & """ width=""200""><span id=""blog" & oBlog("blogid") & """>" & oBlog("title") & "</span></td>" & vbcrlf
       response.write "      <td class=""formlist"">" & lcl_blog_image & "</td>" & vbcrlf
       response.write "      <td class=""formlist"" title=""click to edit"" onClick=""" & lcl_row_onclick & """ width=""150"">" & trim(oBlog("blogowner")) & "</td>" & vbcrlf
       response.write "      <td class=""formlist"" title=""click to edit"" onClick=""" & lcl_row_onclick & """ align=""center"">" & lcl_active & "</td>" & vbcrlf
       response.write "      <td class=""formlist"" align=""center""><input type=""button"" name=""delete" & iRowCount & """ id=""delete"   & iRowCount & """ value=""Delete"" class=""button"" onclick=""confirm_delete('" & oBlog("blogid") & "');"" /></td>" & vbcrlf

       if lcl_orghasfeature_rssfeeds_mayorsblog AND lcl_userhaspermission_rssfeeds_mayorsblog then
          response.write "      <td class=""formlist"" align=""center""><input type=""button"" name=""sendToRSS" & iRowCount & """ id=""sendToRSS"   & iRowCount & """ value=""Send"" class=""button"" onclick=""sendToRSS('" & oBlog("blogid") & "');"" /></td>" & vbcrlf

         'Check to see if a log exists for this row
          if checkRSSLogExists(session("orgid"),oBlog("blogid"),lcl_rssType) then
             response.write "      <td class=""formlist"" align=""center""><input type=""button"" name=""viewRSSLog" & iRowCount & """ id=""viewRSSLog" & iRowCount & """ value=""View"" class=""button"" onclick=""viewRSSLog('" & oBlog("blogid") & "');"" /></td>" & vbcrlf
          else
             response.write "      <td class=""formlist"" align=""center"">&nbsp;</td>" & vbcrlf
          end if
       end if

       response.write "      <td class=""formlist"" title=""click to edit"" onClick=""" & lcl_row_onclick & """ width=""150"">" & vbcrlf
       response.write            trim(oBlog("createdbyname")) & "<br/>" & vbcrlf
       response.write "          <span style=""color:#800000;"">[" & oBlog("createdbydate") & "]</span>" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "      <td class=""formlist"" title=""click to edit"" onClick=""" & lcl_row_onclick & """ width=""150"" nowrap=""nowrap"">" & vbcrlf
       response.write            trim(oBlog("lastmodifiedbyname")) & "<br />" & vbcrlf
       response.write "          <span style=""color:#800000;"">[" & oBlog("lastmodifiedbydate") & "]</span>" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "  </tr>"  & vbcrlf

       oBlog.movenext
   loop

 		response.write "</table>" & vbcrlf
	  response.write "</div>" & vbcrlf

 else
  		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No blog entries have been created.</p>" & vbcrlf
	end if

	oBlog.close
	set oBlog = nothing 

end sub

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SR" then
        lcl_return = "Successfully Reordered..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "NE" then
        lcl_return = "Blog does not exist..."
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
sub showBlogOwners(iUserID)

  sSQL = "SELECT DISTINCT u.firstname + ' ' + u.lastname as blogowner, b.userid "
  sSQL = sSQL & " FROM egov_mayorsblog b "
  sSQL = sSQL &      " LEFT OUTER JOIN users u ON b.userid = u.userid "
  sSQL = sSQL &      " AND u.orgid = " & session("orgid")
  sSQL = sSQL & " WHERE b.orgid = " & session("orgid")
  sSQL = sSQL & " ORDER BY 1 "

  set oBlogOwners = Server.CreateObject("ADODB.Recordset")
 	oBlogOwners.Open sSQL, Application("DSN"), 3, 1

  if not oBlogOwners.eof then
     do while not oBlogOwners.eof
        if CLng(iUserID) = oBlogOwners("userid") then
           lcl_selected = " selected=""selected"""
        else
           lcl_selected = ""
        end if

        response.write "  <option value=""" & oBlogOwners("userid") & """" & lcl_selected & ">" & oBlogOwners("blogowner") & "</option>" & vbcrlf
        oBlogOwners.movenext
     loop
  end if

  oBlogOwners.close
  set oBlogOwners = nothing

end sub

'----------------------------------------------------------------------------------------
sub checkForOrderBySelected(ByVal lcl_sc_orderby, ByRef lcl_selected_createdate, ByRef lcl_selected_blogowner, _
                                                  ByRef lcl_selected_createdby, ByRef lcl_selected_active)

  if lcl_sc_orderby <> "" then
     lcl_sc_orderby = UCASE(lcl_sc_orderby)

     if lcl_sc_orderby = "BLOGOWNER" then
        lcl_selected_createdate = ""
        lcl_selected_blogowner  = " selected=""selected"""
        lcl_selected_createdby  = ""
        lcl_selected_active     = ""
     elseif lcl_sc_orderby = "CREATEDBY" then
        lcl_selected_createdate = ""
        lcl_selected_blogowner  = ""
        lcl_selected_createdby  = " selected=""selected"""
        lcl_selected_active     = ""
     elseif lcl_sc_orderby = "ACTIVE" then
        lcl_selected_createdate = ""
        lcl_selected_blogowner  = ""
        lcl_selected_createdby  = ""
        lcl_selected_active     = " selected=""selected"""
     else
        lcl_selected_createdate = " selected=""selected"""
        lcl_selected_blogowner  = ""
        lcl_selected_createdby  = ""
        lcl_selected_active     = ""
     end if
  else
     lcl_selected_createdate = " selected=""selected"""
     lcl_selected_blogowner  = ""
     lcl_selected_createdby  = ""
     lcl_selected_active     = ""
  end if

end sub
%>