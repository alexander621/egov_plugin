<!-- #include file="../includes/common.asp" //-->
<!-- #include file="mayorsblog_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: blogsearch.asp
' AUTHOR: Steve Loar
' CREATED: 09/07/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Search of blog items.
'
' MODIFICATION HISTORY
' 1.0   09/07/2011	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearchText, sFeatureName, sBlogMonth, sBlogYear

If request("searchtext") <> "" Then 
	sSearchText = request("searchtext")
Else
	sSearchText = ""
End If 

 sBlogMonth = request("blogMonth")
 sBlogYear  = request("blogYear")

sFeatureName = GetFeatureName("mayorsblog")

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If


%>
<html lang="en">
<head runat="server">
    <meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

    <title><%=sTitle%></title>

    <link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" type="text/css" href="blog_styles.css" />
    
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

		function goBack( NewsMonth, NewsYear )
		{
			location.href = 'mayorsblog.asp?blogMonth=' + NewsMonth + '&blogYear=' + NewsYear;
		}

	//-->
	</script>

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<div id="content">
  <div id="centercontent">

	<table border="0" cellspacing="0" cellpadding="0" width="800">
		<tr><td><font class="pagetitle">Search: <%=sFeatureName%></font></td></tr>
	</table>
	<%	RegisteredUserDisplay( "../" ) %>

	<!-- Back Button -->
	<p><input type="button" class="button" value="<< Back" onclick="goBack( '<%=sBlogMonth%>', '<%=sBlogYear%>' )" /></p>

	<!-- Search Field here -->
	<form name="frmSearch" method="post" action="blogsearch.asp">
		<input type="hidden" name="blogMonth" value="<%=sBlogMonth%>" />
		<input type="hidden" name="blogYear" value="<%=sBlogYear%>" />

		<table border="0" cellspacing="2" cellpadding="0" width="300">
		  <tr>
			  <td align="right">
				  <input type="input" id="searchtext" name="searchtext" value="<%=sSearchText%>" placeholder="Search Phrase" size="50" maxlength="50" />
			  </td>
			  <td>
				  <input type="button" class="button" value="Search" onclick="goFindIt()" />
			  </td>
		  </tr>
		</table><br />
	</form>

<%
	  displayBlogList iorgid, sSearchText, sBlogMonth, sBlogYear
%>

  </div>
</div>

<!-- #include file="../include_bottom.asp" -->


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void displayBlogList iOrgId, sSearchText
'--------------------------------------------------------------------------------------------------
Sub displayBlogList( ByVal iOrgId, ByVal sSearchText, ByVal sBlogMonth, ByVal sBlogYear )
	Dim sSql, oRs, lcl_imagefilename, lcl_article_url, lcl_article, lcl_postcomments_formid
	Dim lcl_shareComment_label

	sSql = "SELECT mb.blogid, mb.userid, u.firstname + ' ' + u.lastname AS blogOwner, mb.title, mb.article, "
	sSql = sSql & "mb.createdbydate, u.imagefilename "
	sSql = sSql & "FROM egov_mayorsblog mb "
	sSql = sSql & "LEFT OUTER JOIN users u ON mb.userid = u.userid AND u.orgid = mb.orgid "
	sSql = sSql & "WHERE mb.orgid = "  & iOrgId & " AND mb.isInactive = 0 "
	sSql = sSql & " AND ( mb.title LIKE '%" & track_dbsafe(sSearchText) & "%'  "
	sSql = sSql & "OR mb.article LIKE '%"  & track_dbsafe(sSearchText) & "%' ) "
	sSql = sSql & " ORDER BY mb.blogid DESC"

	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 

		'Determine if there is a form associated to this faq/rumor
		lcl_postcomments_formid = getCommentsFormID(iOrgId, "", "mayorsblog")
		lcl_postcomments_url = sEgovWebsiteURL & "/action.asp?actionid=" & lcl_postcomments_formid

		Do While Not oRs.EOF

			'Determine if there is an image associated to the BlogOwner
			lcl_imagefilename = buildBlogImg(oRs("imagefilename"),sorgVirtualSiteName)

			'Setup the url to the individual article
			lcl_article_url = "mayorsblog_info.asp?id=" & oRs("blogid") & "&blogMonth=" & sBlogMonth & "&blogYear=" & sBlogYear & "&searchtext=" & sSearchText

			'Display only the first 1000 characters of the blog article.
			lcl_article = oRs("article")

			If Len(lcl_article) > 1000 Then 
				lcl_article = left(lcl_article,1000) & "... [<a href=""" & lcl_article_url & """>more</a>]"
			End If 

			lcl_article = formatArticle(lcl_article)

			'response.write "<p>"
			response.write "<fieldset class=""articleblock"">"
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">"
			response.write "<tr>"
			response.write "<td valign=""top"">"
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
			response.write "<tr>"
			response.write "<td style=""font-size:12pt; color:#800000;"">" & FormatDateTime(oRs("createdbydate"),vbshortdate) & "<br /><br /></td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td style=""font-size:16px; font-weight:bold;"">" & oRs("title") & "</td>"
			response.write "</tr>"
			response.write "</table>"
			response.write "<br />"

			If lcl_imagefilename <> "" Then 
				response.write "<img src=""" & lcl_imagefilename & """ style=""float:left; border:1pt solid #000000; margin: 5px;"" />"
			End If 

			response.write "<i style=""color:#800000"">by " & oRs("blogOwner") & "</i><br />"
			response.write "<p>" & lcl_article & "</p>"
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
			response.write "<tr>"
			response.write "<td>"

			If lcl_postcomments_formid > 0 Then 
				lcl_shareComment_label = "Share and Comment"
			Else 
				lcl_shareComment_label = "Share This Article"
			End If 

			response.write "<button type=""button"" name=""viewBlog" & oRs("blogid") & """ id=""viewBlog" & oRs("blogid") & """ class=""button"" style=""cursor:pointer"" onclick=""location.href='" & lcl_article_url & "';"">" & lcl_shareComment_label & "</button>"
			response.write "</td>"
			response.write "</tr>"
			response.write "</table>"
			response.write "</td>"
			response.write "</tr>"
			response.write "</table>"
			response.write "</fieldset>"
			'response.write "</p>"

			oRs.MoveNext
		Loop 
	Else
		response.write vbcrlf & "<div id=""blogsearchfailed""><p>Nothing was found that matched your search phrase.</p></div>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

%>
