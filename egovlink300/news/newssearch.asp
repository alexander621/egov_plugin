<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: newssearch.asp
' AUTHOR: Steve Loar
' CREATED: 09/07/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Search of news items.
'
' MODIFICATION HISTORY
' 1.0   09/07/2011	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearchText, sFeatureName, sNewsMonth, sNewsYear

If request("searchtext") <> "" Then 
	sSearchText = request("searchtext")
Else
	sSearchText = ""
End If 

 sNewsMonth = request("newsMonth")
 sNewsYear  = request("newsYear")

sFeatureName = GetFeatureName("news_items")

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If


%>
<html lang="en">
<head>
    	<meta charset="UTF-8">

    	<title><%=sTitle%></title>

    	<link rel="stylesheet" href="../css/styles.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" href="news_styles.css" />
    
    <script src="https://code.jquery.com/jquery-1.6.2.min.js"></script>

	<script>
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

		function viewArticle( iNewsID ) 
		{
			var lcl_width = 400;
			var lcl_height = 250;
			var lcl_left = (screen.availWidth/2) - (lcl_width/2);
			var lcl_top = (screen.availHeight/2) - (lcl_height/2);
			var lcl_url = "news_info.asp?id=" + iNewsID;

			popupWin = window.open(lcl_url, "_blank"+ iNewsID,"width=" + lcl_width + ",height=" + lcl_height + ",left=" + lcl_left + ",top=" + lcl_top + ",resizable=yes,scrollbars=yes,status=yes");
		}

		function goBack( NewsMonth, NewsYear )
		{
			location.href = 'news.asp?newsMonth=' + NewsMonth + '&newsYear=' + NewsYear;
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
	<p><input type="button" class="button" value="<< Back" onclick="goBack( '<%=sNewsMonth%>', '<%=sNewsYear%>' )" /></p>

	<!-- Search Field here -->
	<form name="frmSearch" method="post" action="newssearch.asp">
		<input type="hidden" name="newsMonth" value="<%=sNewsMonth%>" />
		<input type="hidden" name="newsYear" value="<%=sNewsYear%>" />

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

<%	displayNewsList iOrgId, sSearchText		%>

  </div>
</div>

<!-- #include file="../include_bottom.asp" -->


<%
'--------------------------------------------------------------------------------------------------
' void displayNewsList iOrgId, sSearchText
'--------------------------------------------------------------------------------------------------
Sub displayNewsList( ByVal iOrgId, ByVal sSearchText )
	Dim sSql, oRs, sBgColor

	sSql = "SELECT newsitemid, itemtitle, itemtext, itemlinkurl, itemdisplay, itemorder, "
	sSql = sSql & "ISNULL(ISNULL(publicationstart,itemdate),GETDATE()) AS articledate "
	sSql = sSql & "FROM egov_news_items "
	sSql = sSql & "WHERE itemdisplay = 1 AND UPPER(newstype) = 'NEWS' "
	sSql = sSql & "AND orgid = "  & iOrgId
	sSql = sSql & " AND ( itemtitle LIKE '%" & track_dbsafe(sSearchText) & "%' "
	sSql = sSql & "OR itemtext LIKE '%" & track_dbsafe(sSearchText) & "%' ) AND publicationstart IS NOT NULL "
	sSql = sSql & "ORDER BY 7 DESC"
	'sSql = sSql & "ORDER BY ISNULL(publicationstart, itemdate) DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""100%"" class=""newsTable"">"
		response.write vbcrlf & "<tr align=""left"">"
		response.write "<td class=""newsTableHeaders"" width=""150"">Article Date</td>"
		response.write "<td class=""newsTableHeaders"">Title</td>"
		response.write "<td class=""newsTableHeaders"">&nbsp;</td>"
		response.write "</tr>"

		sBgColor = "#eeeeee"

		Do While Not oRs.EOF

			sBgColor = changeBGColor( sBgColor, "#eeeeee", "#ffffff" )

			response.write vbcrlf & "<tr bgcolor=""" & sBgColor & """>"
			response.write "<td width=""150"">" & FormatDateTime(oRs("articledate"),vbshortdate) & "</td>"
			response.write "<td>" & oRs("itemtitle") & "</td>"
			response.write "<td align=""center"">"
			response.write "<input type=""button"" name=""viewArticle" & oRs("newsitemid") & """ id=""viewArticle" & oRs("newsitemid") & """ value=""View Article"" class=""button"" onclick=""viewArticle('" & oRs("newsitemid") & "');"" />"
			response.write "</td>"
			response.write "</tr>"

			oRs.MoveNExt
		Loop 

		response.write vbcrlf & "</table>"

	Else
		response.write vbcrlf & "<div id=""newsearchfailed""><p>Nothing was found that matched your search phrase.</p></div>"
	End If 

	oRs.close
	Set oRs = Nothing 

End Sub 



%>
