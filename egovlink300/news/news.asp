 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: news.asp
' AUTHOR:   David Boyer
' CREATED:  04/07/2009
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the News
'
' MODIFICATION HISTORY
' 1.0	04/07/09	David Boyer - Initial Version
' 1.1	09/07/2011	Steve Loar - Added the search field and button
' 1.2	06/10/2013	Steve Loar - Added regular expression match to month and year to block intrusion attack attempts
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("news_items") = "Y" then
	response.redirect "outage_feature_offline.asp"
end if

Dim oActionOrg, re, matches

set oActionOrg = New classOrganization
Set re = New RegExp
re.Pattern = "^\d+$"

if request.querystring("id") <> "" then
	response.redirect "news_info.asp?id=" & request.querystring("id")
end if

'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide
lcl_hidden = "HIDDEN"

'Set up local variables based on posting (list) type.
lcl_list_label   = "News"
lcl_list_title   = "News"
lcl_feature_name = GetFeatureName("news_items")

'Retrieve the search parameters
lcl_newsMonth = request("newsMonth")
lcl_newsYear  = request("newsYear")

'If BOTH the newsMonth AND newsYear are blank then grab the Month/Year of the latest, active, blog entry.
' if lcl_newsMonth = "" AND lcl_newsYear = "" then
'    getCurrentNewsArchive iorgid, lcl_newsMonth, lcl_newsYear
' end if

'Set to the current month/year if either are blank.
If lcl_newsMonth = "" Then 
	lcl_newsMonth = month(now)
Else
	Set matches = re.Execute(lcl_newsMonth)
	If matches.Count > 0 Then
		lcl_newsMonth = CLng(lcl_newsMonth)
	Else
		response.redirect("news.asp")
	End If 
end if

if lcl_newsYear = "" Then 
	lcl_newsYear = year(now)
Else
	Set matches = re.Execute(lcl_newsYear)
	If matches.Count > 0 Then
		lcl_newsYear = CLng(lcl_newsYear)
	Else
		response.redirect("news.asp")
	End If 
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
	<link rel="stylesheet" type="text/css" href="news_styles.css" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/easyform.js"></script>
	<script language="javascript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>

	<script type="text/javascript" src="https://code.jquery.com/jquery-1.6.2.min.js"></script>

	<script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script>
	<script type="text/javascript">var addthis_pub="cschappacher";</script>

	<script language="javascript">
		function viewArticle( iNewsID ) 
		{
			lcl_width  = 400;
			lcl_height = 250;
			lcl_left   = (screen.availWidth/2) - (lcl_width/2);
			lcl_top    = (screen.availHeight/2) - (lcl_height/2);
			lcl_url    = "news_info.asp?id=" + iNewsID;

			popupWin = window.open(lcl_url, "_blank"+ iNewsID,"width=" + lcl_width + ",height=" + lcl_height + ",left=" + lcl_left + ",top=" + lcl_top + ",resizable=yes,scrollbars=yes,status=yes");
		}

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

	</script>

</head>

<!--#include file="../include_top.asp"-->

<div id="content">
  <div id="centercontent">

<table border="0" cellspacing="0" cellpadding="0" style="max-width:800px;">
  <tr>
      <td>
          <font class="pagetitle"><%=lcl_feature_name%></font>
          <% checkForRSSFeed iorgid, "", "", "NEWS", sEgovWebsiteURL %>
      </td>
      <td align="right">
          <% displayAddThisButton iorgid %>
      </td>
  </tr>
</table>
<% RegisteredUserDisplay( "../" ) %>

<!-- Search Field here -->
<form name="frmSearch" method="post" action="newssearch.asp">
	<input type="hidden" name="newsMonth" value="<%=lcl_newsMonth%>" />
	<input type="hidden" name="newsYear" value="<%=lcl_newsYear%>" />
	<table border="0" cellspacing="2" cellpadding="0" width="300">
	  <tr>
		  <td align="right">
			  <input type="text" id="searchtext" name="searchtext" value="" placeholder="Search Phrase" size="50" maxlength="50" />
		  </td>
		  <td>
			  <input type="button" class="button" value="Search" onclick="goFindIt()" />
		  </td>
	  </tr>
	</table><br />
</form>


<table border="0" cellspacing="0" cellpadding="2" style="max-width:800px;">
  <form name="mayorsblog" action="mayorsblog.asp" method="post">
  <tr valign="top">
      <td style="max-width:600px;" class="respCol">
        <%
          response.write "<div style=""font-size:16pt;"">" & monthname(lcl_newsMonth) & " " & lcl_newsYear & "</div><br />" & vbcrlf
          displayNewsList iorgid, lcl_newsMonth, lcl_newsYear
        %>
      </td>
      <td class="respCol">
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
Sub displayNewsList( ByVal p_orgid, ByVal p_newsMonth, ByVal p_newsYear )
	Dim sSql, oNews

	If p_newsMonth = "" Then 
		lcl_newsMonth = month(now)
	End If 

	If p_newsYear = "" Then 
		lcl_newsYear = year(now)
	End If 

	sSql = "SELECT newsitemid, itemtitle, itemtext, itemlinkurl, itemdisplay, itemorder, "
	sSql = sSql & " isnull(publicationstart,itemdate) AS articledate, publicationend "
	sSql = sSql & "FROM egov_news_items "
	sSql = sSql & "WHERE itemdisplay = 1 "
	sSql = sSql & "AND UPPER(newstype) = 'NEWS' "
	sSql = sSql & "AND orgid = "  & p_orgid
	sSql = sSql & " AND month(isnull(publicationstart,itemdate)) = " & lcl_newsMonth
	sSql = sSql & " AND year(isnull(publicationstart,itemdate)) = "  & lcl_newsYear
	sSql = sSql & " ORDER BY isnull(publicationstart,itemdate) DESC "

	Set oNews = Server.CreateObject("ADODB.Recordset")
	oNews.Open sSql, Application("DSN"), 3, 1

	If Not oNews.eof Then 
		response.write "<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""100%"" class=""newsTable liquidtable"">" & vbcrlf
		response.write "<thead>" & vbcrlf
		response.write "<tr align=""left"">" & vbcrlf
		response.write "<td class=""newsTableHeaders"" width=""150"">Article Date</td>" & vbcrlf
		response.write "<td class=""newsTableHeaders"">Title</td>" & vbcrlf
		response.write "<td class=""newsTableHeaders"">&nbsp;</td>" & vbcrlf
		response.write "</tr>" & vbcrlf
		response.write "</thead>" & vbcrlf

		lcl_bgcolor = "#eeeeee"

		Do While Not oNews.EOF

			lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

			response.write "<tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
			response.write "<td class=""repeatheaders"">Article Date</td>" & vbcrlf
			response.write "<td width=""150"">" & formatdatetime(oNews("articledate"),vbshortdate) & "</td>" & vbcrlf
			response.write "<td class=""repeatheaders"">Title</td>" & vbcrlf
			response.write "<td>" & oNews("itemtitle") & "</td>" & vbcrlf
			response.write "<td align=""center"">" & vbcrlf
			response.write "<input type=""button"" name=""viewArticle" & oNews("newsitemid") & """ id=""viewArticle" & oNews("newsitemid") & """ value=""View Article"" class=""button"" onclick=""viewArticle('" & oNews("newsitemid") & "');"" />" & vbcrlf
			response.write "</td>" & vbcrlf
			response.write "</tr>" & vbcrlf

			oNews.movenext
		Loop 

		response.write "</table>" & vbcrlf

	End If 

	oNews.close
	Set oNews = Nothing 

End Sub 


'------------------------------------------------------------------------------
sub displayArchives( ByVal p_orgid )
	Dim sSql, oNewsDates

 'Retreive a distinct list of createdbydates from egov_mayorsblog.
  sSql = "SELECT distinct DATEPART(mm,isnull(publicationstart,itemdate)) AS newsMonth, "
  sSql = sSql & " DATEPART(yyyy,isnull(publicationstart,itemdate)) as newsYear "
  sSql = sSql & " FROM egov_news_items "
  sSql = sSql & " WHERE itemdisplay = 1 "
  sSql = sSql & " AND UPPER(newstype) = 'NEWS' "
  sSql = sSql & " AND orgid = " & p_orgid
  sSql = sSql & " ORDER BY 2 DESC, 1 DESC "

 	set oNewsDates = Server.CreateObject("ADODB.Recordset")
  oNewsDates.Open sSql, Application("DSN"), 3, 1

  if not oNewsDates.eof then
     do while not oNewsDates.eof

        if oNewsDates("newsMonth") <> "" then
           lcl_monthName = monthname(oNewsDates("newsMonth"))
        else
           lcl_monthName = ""
        end if

        response.write "<a href=""news.asp?newsMonth=" & oNewsDates("newsMonth") & "&newsYear=" & oNewsDates("newsYear") & """>" & lcl_monthName & " " & oNewsDates("newsYear") & "</a><br />" & vbcrlf

        oNewsDates.movenext
     loop
  end if

  oNewsDates.close
  set oNewsDates = nothing

end sub

'------------------------------------------------------------------------------
sub getCurrentNewsArchive(ByVal p_orgid, ByRef lcl_newsMonth, ByRef lcl_newsYear)
	Dim sSql, oMaxDate

	sSql = "SELECT max(isnull(publicationstart,itemdate)) AS maxDate "
	sSql = sSql & " FROM egov_news_items "
	sSql = sSql & " WHERE itemdisplay = 1 "
	sSql = sSql & " AND UPPER(newstype) = 'NEWS' "
	sSql = sSql & " AND orgid = " & p_orgid

	set oMaxDate = Server.CreateObject("ADODB.Recordset")
	oMaxDate.Open sSql, Application("DSN"), 3, 1

	if not oMaxDate.eof then
		lcl_newsMonth = month(oMaxDate("maxDate"))
		lcl_newsYear  = year(oMaxDate("maxDate"))
	else
		lcl_newsMonth = month(now)
		lcl_newsYear  = year(now)
	end if

	oMaxDate.close
	set oMaxDate = nothing

end Sub


%>
