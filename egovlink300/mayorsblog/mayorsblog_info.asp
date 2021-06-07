 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="mayorsblog_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mayorsblog_info.asp
' AUTHOR:   David Boyer
' CREATED:  04/10/2009
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays a specific Mayor's Blog article
'
' MODIFICATION HISTORY
' 1.0  04/10/09	 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSearchText, lcl_blogMonth, lcl_blogYear, lcl_blogID, lcl_hidden, oBlogOrg, sReturnToURL, re, matches

'Check to see if the feature is offline
If isFeatureOffline("mayorsblog") = "Y" Then 
	response.redirect "outage_feature_offline.asp"
End If 

Set oBlogOrg = New classOrganization
Set re = New RegExp
re.Pattern = "^\d+$"

'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide
lcl_hidden = "HIDDEN"

'Set up local variables based on posting (list) type.
lcl_list_label   = "Mayor's Blog"
lcl_list_title   = "Mayor's Blog"
lcl_feature_name = GetFeatureName("mayorsblog")

'Retrieve the search parameters
lcl_blogMonth = request("blogMonth")
lcl_blogYear  = request("blogYear")

lcl_blogID    = request("id")
Set matches = re.Execute(lcl_blogID)
If matches.Count > 0 Then
	lcl_blogID = CLng(lcl_blogID)
Else
	response.redirect "mayorsblog.asp?blogMonth=" & lcl_blogMonth & "&blogYear=" & lcl_blogYear
End If 

If request("searchtext") <> "" Then 
	sSearchText = "&searchtext=" & request("searchtext")
	sReturnToURL = "blogsearch"
Else
	sSearchText = ""
	sReturnToURL = "mayorsblog"
End If 

If lcl_blogID <> "" Then 
	If Not IsNumeric( lcl_blogID ) Then 
		response.redirect "mayorsblog.asp?blogMonth=" & lcl_blogMonth & "&blogYear=" & lcl_blogYear
	End If 
End If 

'Check for org features
 lcl_orghasfeature_button_addthis   = oBlogOrg.orghasfeature("button_addthis")
 lcl_orghasfeature_button_yahoobuzz = oBlogOrg.orghasfeature("button_yahoobuzz")
 lcl_orghasfeature_action_line      = oBlogOrg.orghasfeature("action line")
%>
<html>
<head>
	<title>E-Gov Services - <%=sOrgName%></title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" type="text/css" href="blog_styles.css" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/easyform.js"></script>
	<script language="javascript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>

	<script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script>
	<script type="text/javascript">var addthis_pub="cschappacher";</script>

</head>
<!--#include file="../include_top.asp"-->

<div id="content">
  <div id="centercontent">

	<table border="0" cellspacing="0" cellpadding="0" width="800">
	  <tr>
		  <td>
			<%
			 'Build the Welcome message
			  lcl_org_name        = oBlogOrg.GetOrgName()
			  lcl_org_state       = oBlogOrg.GetState()
			  lcl_org_featurename = lcl_feature_name

			  oBlogOrg.buildWelcomeMessage iorgid, lcl_orghasdisplay_action_page_title, lcl_org_name, lcl_org_state, lcl_org_featurename
			%>
			  <!--<font class="pagetitle"><%'lcl_feature_name%></font>-->
		  </td>
	  </tr>
	</table>

	<% RegisteredUserDisplay("") %>

	<p>
		<input type="button" class="button" name="returnButton" id="returnButton" onclick="location.href='<%=sReturnToURL%>.asp?blogMonth=<%=lcl_blogMonth%>&blogYear=<%=lcl_blogYear%><%=sSearchText%>'" value="<< Back" />
	</p>


	<table border="0" cellspacing="0" cellpadding="2" width="800">
	  <tr valign="top">
		  <td width="600">
			<% displayBlog iorgid, lcl_blogID %>
		  </td>
		  <td>
			<%
			  response.write "<div align=""center"" style=""font-size:16pt;"">ARCHIVES</div><br />" & vbcrlf
			  response.write "<div align=""center"">" & vbcrlf

			  displayArchives iorgid

			  response.write "</div>" & vbcrlf
			%>
		  </td>
	  </tr>
	</table>

  </div>
</div>
<!-- #include file="../include_bottom.asp" -->
<%
'------------------------------------------------------------------------------
sub displayBlog( ByVal p_orgid, ByVal p_blogID)
	Dim sSql, oBlogInfo

  sSql = "SELECT mb.blogid, mb.userid, u.firstname + ' ' + u.lastname AS blogOwner, mb.title, mb.article, "
  sSql = sSql & " mb.createdbydate, u.imagefilename "
  sSql = sSql & " FROM egov_mayorsblog mb "
  sSql = sSql &      " LEFT OUTER JOIN users u ON mb.userid = u.userid AND u.orgid = " & p_orgid
  sSql = sSql & " WHERE mb.isInactive = 0 "
  sSql = sSql & " AND mb.orgid = "  & p_orgid
  sSql = sSql & " AND mb.blogid = " & p_blogID
  sSql = sSql & " ORDER BY blogid DESC "

 	set oBlogInfo = Server.CreateObject("ADODB.Recordset")
  oBlogInfo.Open sSql, Application("DSN"), 3, 1

  if not oBlogInfo.eof then

    'Determine if there is a form associated to this faq/rumor
     lcl_postcomments_formid = getCommentsFormID(p_orgid, "", "mayorsblog")
     lcl_postcomments_url    = sEgovWebsiteURL & "/action.asp?actionid=" & lcl_postcomments_formid

     do while not oBlogInfo.eof

       'Determine if there is an image associated to the BlogOwner
        lcl_imagefilename = buildBlogImg(oBlogInfo("imagefilename"),sorgVirtualSiteName)

        lcl_article = formatArticle(oBlogInfo("article"))

        response.write "<p>" & vbcrlf
        response.write "  <fieldset class=""articleblock"">" & vbcrlf
        response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
        response.write "    <tr>" & vbcrlf
        response.write "        <td valign=""top"">" & vbcrlf
        response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
        response.write "              <tr>" & vbcrlf
        response.write "                  <td style=""font-size:12pt; color:#800000;"">" & formatdatetime(oBlogInfo("createdbydate"),vbshortdate) & "<br /><br /></td>" & vbcrlf
        response.write "              </tr>" & vbcrlf
        response.write "              <tr>" & vbcrlf
        response.write "                  <td style=""font-size:16px; font-weight:bold;"">" & oBlogInfo("title") & "</td>" & vbcrlf
        response.write "              </tr>" & vbcrlf
        response.write "            </table>" & vbcrlf
        response.write "            <br />" & vbcrlf

        if lcl_imagefilename <> "" then
           response.write "            <img src=""" & lcl_imagefilename & """ style=""float:left; border:1pt solid #000000; margin: 5px;"" />" & vbcrlf
        end if

        response.write "            <i style=""color:#800000"">by " & oBlogInfo("blogOwner") & "</i><br />" & vbcrlf
        response.write "            <p>" & lcl_article & "</p>" & vbcrlf

        response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
        response.write "              <tr>" & vbcrlf
        response.write "                  <td>" & vbcrlf

        if lcl_orghasfeature_action_line AND lcl_postcomments_formid > 0 then
           response.write "                      <input type=""button"" name=""postCommentButton"" id=""postCommentButton"" value=""Post a Comment"" style=""cursor:pointer"" onclick=""location.href='" & lcl_postcomments_url & "'"" />" & vbcrlf
        end if

        response.write "                  </td>" & vbcrlf
        response.write "                  <td align=""right"">" & vbcrlf
        response.write "                      <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
        response.write "                        <tr valign=""top"">" & vbcrlf

       'Yahoo Buzz
        if lcl_orghasfeature_button_yahoobuzz then
           response.write "                            <td>" & vbcrlf
                                                           displayYahooBuzzButton iorgid
           response.write "                            </td>" & vbcrlf
        end if

       'Add This
        if lcl_orghasfeature_button_addthis then
           response.write "                            <td>" & vbcrlf
                                                           displayAddThisButton iorgid
           response.write "                            </td>" & vbcrlf
        end if

        response.write "                        </tr>" & vbcrlf
        response.write "                      </table>" & vbcrlf
        response.write "                      </a>" & vbcrlf
        response.write "                  </td>" & vbcrlf
        response.write "              </tr>" & vbcrlf
        response.write "            </table>" & vbcrlf

        response.write "        </td>" & vbcrlf
        response.write "    </tr>" & vbcrlf
        response.write "  </table>" & vbcrlf
        response.write "  </fieldset>" & vbcrlf
        response.write "</p>" & vbcrlf

        oBlogInfo.movenext
     loop
  end if

  oBlogInfo.close
  set oBlogInfo = nothing

end sub
%>
