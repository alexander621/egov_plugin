 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="mayorsblog_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mayorsblog.asp
' AUTHOR:   David Boyer
' CREATED:  04/06/2009
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the Mayor's Blog
'
' MODIFICATION HISTORY
' 1.0  04/06/09	 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim oBlogOrg, re, matches

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
lcl_list_label = "Mayor's Blog"
lcl_list_title = "Mayor's Blog"
lcl_feature_name = GetFeatureName("mayorsblog")

'Retrieve the search parameters
' lcl_blogMonth = trim(request("blogMonth"))
' lcl_blogYear = trim(request("blogYear"))

lcl_blogMonth = ""
lcl_blogYear  = ""

if request("blogMonth") <> "" then
'	if not containsApostrophe(request("blogMonth")) then
'		lcl_blogMonth = clng(request("blogMonth"))
'	end if

	lcl_blogMonth = request("blogMonth")
	Set matches = re.Execute(lcl_blogMonth)
	If matches.Count > 0 Then
		lcl_blogMonth = CLng(lcl_blogMonth)
	Else
		lcl_blogMonth = ""
	End If 
end if

if request("blogYear") <> "" then
'	if not containsApostrophe(request("blogYear")) then
'		lcl_blogYear = clng(request("blogYear"))
'	end if

	lcl_blogYear = request("blogYear")
	Set matches = re.Execute(lcl_blogYear)
	If matches.Count > 0 Then
		lcl_blogYear = CLng(lcl_blogYear)
	Else
		lcl_blogYear = ""
	End If 
end if

'If BOTH the blogMonth AND blogYear are blank then grab the Month/Year of the latest, active, blog entry.
If (lcl_blogMonth = "" Or IsNull(lcl_blogMonth)) And (lcl_blogYear = "" Or IsNull(lcl_blogYear)) Then 
	getCurrentBlogArchive iorgid, lcl_blogMonth, lcl_blogYear
End If 

'Set to the current month/year if either are blank.
If lcl_blogMonth = "" Or IsNull(lcl_blogMonth) Then 
	lcl_blogMonth = Month(Now())
End If 

If lcl_blogYear = "" Or IsNull(lcl_blogYear) Then 
	lcl_blogYear = Year(Now())
End If 

'Get the local date/time
lcl_local_datetime = ConvertDateTimetoTimeZone(iOrgID)

%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<title>E-Gov Services - <%=sOrgName%></title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/easyform.js"></script>
	<script language="javascript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>

	<script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script>
	<script type="text/javascript">var addthis_pub="cschappacher";</script>

	<script type="text/javascript" src="https://code.jquery.com/jquery-1.6.2.min.js"></script>

	<script language="Javascript">
	<!--

		function goFindIt()
		{
			if ($("#searchtext").val() == "")
			{
				alert( "Missing search phrase. Please enter a phrase, then try your search again.");
				$("#searchtext").focus();
				return false;
			}
			//alert($("#searchtext").val());
			// send off the search
			document.frmSearch.submit();
		}

	//-->
	</script>

</head>
<!--#include file="../include_top.asp"-->

<div id="content">
  <div id="centercontent">

	<table border="0" cellspacing="0" cellpadding="0" style="max-width:800px;">
	  <tr>
		  <td>
			  <font class="pagetitle"><%=lcl_feature_name%></font>
			  <% checkForRSSFeed iorgid, "", "", "MAYORSBLOG", sEgovWebsiteURL %>
		  </td>
		  <td align="right">
			  <% displayAddThisButton iorgid %>
		  </td>
	  </tr>
	</table>
	<% RegisteredUserDisplay("../") %>

	<!-- Search Field here -->
	<form name="frmSearch" method="post" action="blogsearch.asp">
		<input type="hidden" name="blogMonth" value="<%=lcl_blogMonth%>" />
		<input type="hidden" name="blogYear" value="<%=lcl_blogYear%>" />

		<table border="0" cellspacing="2" cellpadding="0" width="300" class="respTable">
		  <tr>
			  <td align="right">
				  <input type="text" id="searchtext" name="searchtext" value="<%=sSearchText%>" placeholder="Search Phrase" size="50" maxlength="50" />
			  </td>
			  <td>
				  <input type="button" class="button" value="Search" onclick="goFindIt()" />
			  </td>
		  </tr>
		</table><br />
	</form>


	<table border="0" cellspacing="0" cellpadding="2" style="max-width:800px;" class="respTable">
	  <form name="mayorsblog" action="mayorsblog.asp" method="post">
	  <tr valign="top">
		  <td style="max-width:600px;">
			<%
			  response.write "<div style=""font-size:16pt;"">" & monthname(lcl_blogMonth) & " " & lcl_blogYear & "</div>" & vbcrlf
			  displayBlogList iorgid, lcl_blogMonth, lcl_blogYear
			%>
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
	  </form>
	</table>

  </div>
</div>
<!-- #include file="../include_bottom.asp" -->


<%
'------------------------------------------------------------------------------
Sub displayBlogList( ByVal p_orgid, ByVal p_blogMonth, ByVal p_blogYear )
	Dim sSql, oRs

	iBlogMonth = Month(Now())
	iBlogYear  = Year(Now())

	If p_blogMonth <> "" Then 
		iBlogMonth = p_blogMonth
	End If 

	If p_blogYear <> "" Then 
		iBlogYear = p_blogYear
	End If 

	sSql = "SELECT mb.blogid, mb.userid, u.firstname + ' ' + u.lastname AS blogOwner, mb.title, mb.article, "
	sSql = sSql & "mb.createdbydate, u.imagefilename "
	sSql = sSql & "FROM egov_mayorsblog mb "
	sSql = sSql & "LEFT OUTER JOIN users u ON mb.userid = u.userid AND u.orgid = " & p_orgid
	sSql = sSql & " WHERE mb.isInactive = 0 "
	sSql = sSql & " AND mb.orgid = "  & p_orgid
	sSql = sSql & " AND month(mb.createdbydate) = " & iBlogMonth
	sSql = sSql & " AND year(mb.createdbydate) = "  & iBlogYear
	sSql = sSql & " ORDER BY blogid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 

		'Determine if there is a form associated to this faq/rumor
		lcl_postcomments_formid = getCommentsFormID(p_orgid, "", "mayorsblog")
		lcl_postcomments_url    = sEgovWebsiteURL & "/action.asp?actionid=" & lcl_postcomments_formid

		Do While Not oRs.EOF

			'Determine if there is an image associated to the BlogOwner
			lcl_imagefilename = buildBlogImg(oRs("imagefilename"),sorgVirtualSiteName)

			'Setup the url to the individual article
			lcl_article_url = "mayorsblog_info.asp?id=" & oRs("blogid") & "&blogMonth=" & iBlogMonth & "&blogYear=" & iBlogYear

			'Display only the first 1000 characters of the blog article.
			lcl_article = oRs("article")

			If Len(lcl_article) > 1000 Then 
				lcl_article = left(lcl_article,1000) & "... [<a href=""" & lcl_article_url & """>more</a>]"
			End If 

			lcl_article = formatArticle(lcl_article)

			response.write "<p>" & vbcrlf
			response.write "<fieldset>" & vbcrlf
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
			response.write "<tr>" & vbcrlf
			response.write "<td valign=""top"">" & vbcrlf
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
			response.write "<tr>" & vbcrlf
			response.write "<td style=""font-size:12pt; color:#800000;"">" & formatdatetime(oRs("createdbydate"),vbshortdate) & "<br /><br /></td>" & vbcrlf
			response.write "</tr>" & vbcrlf
			response.write "<tr>" & vbcrlf
			response.write "<td style=""font-size:16px; font-weight:bold;"">" & oRs("title") & "</td>" & vbcrlf
			response.write "</tr>" & vbcrlf
			response.write "</table>" & vbcrlf
			response.write "<br />" & vbcrlf

			If lcl_imagefilename <> "" Then 
				response.write "<img src=""" & lcl_imagefilename & """ style=""float:left; border:1pt solid #000000; margin: 5px;"" />" & vbcrlf
			End If 

			response.write "<i style=""color:#800000"">by " & oRs("blogOwner") & "</i><br />" & vbcrlf
			response.write "<p>" & lcl_article & "</p>" & vbcrlf
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
			response.write "<tr>" & vbcrlf
			response.write "<td>" & vbcrlf

			If lcl_postcomments_formid > 0 Then 
				lcl_shareComment_label = "Share and Comment"
			Else 
				lcl_shareComment_label = "Share This Article"
			End If 

			response.write "<button type=""button"" name=""viewBlog" & oRs("blogid") & """ id=""viewBlog" & oRs("blogid") & """ class=""button"" style=""cursor:pointer"" onclick=""location.href='" & lcl_article_url & "';"">" & lcl_shareComment_label & "</button>" & vbcrlf
			response.write "</td>" & vbcrlf
			response.write "</tr>" & vbcrlf
			response.write "</table>" & vbcrlf
			response.write "</td>" & vbcrlf
			response.write "</tr>" & vbcrlf
			response.write "</table>" & vbcrlf
			response.write "</fieldset>" & vbcrlf
			response.write "</p>" & vbcrlf

			oRs.MoveNext
		Loop 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Sub getCurrentBlogArchive( ByVal p_orgid, ByRef lcl_blogMonth, ByRef lcl_blogYear )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(MAX(createdbydate),'" & now & "') AS maxCreateDate "
	sSql = sSql & " FROM egov_mayorsblog "
	sSql = sSql & " WHERE isInactive = 0 AND orgid = " & p_orgid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.eof Then 
		lcl_blogMonth = month(oRs("maxCreateDate"))
		lcl_blogYear  = year(oRs("maxCreateDate"))
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


%>
