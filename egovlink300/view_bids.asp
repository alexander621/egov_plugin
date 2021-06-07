<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="postings_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: view_bids.asp
' AUTHOR: David Boyer
' CREATED: 10/15/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  User can view all bids (s)he has uploaded to org.
'
' MODIFICATION HISTORY
' 1.0	 10/15/08	David Boyer	- Initial Version
' 1.1  03/11/09 David Boyer - Added "user label" to results list.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()

' handle bots that do not have the userid set.
If request.cookies("userid") = "" Then
  response.redirect session("RedirectPage") 
End If 

sTitle = "View Bids"
session("redirectlang") = "Return to Bids History"
%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<title>E-Gov Services <%=sOrgName & " - " & sTitle %></title>

	<link rel="stylesheet" type="text/css" href="css/styles.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />
	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

	<script language="javascript" src="scripts/modules.js"></script>
	<script language="javascript" src="scripts/easyform.js"></script>  

	<script language="javascript">
	<!--
function openWin2(url, name) {
			popupWin = window.open(url, name,"resizable,width=500,height=450");
}
	//-->
	</script>
</head>

<!--#include file="include_top.asp"-->

<!--BODY CONTENT-->
<font class="pagetitle">Welcome to <%=sOrgName%> Bid History</font><br />

<%	RegisteredUserDisplay( "" ) %>

<div id="content">
 	<div id="centercontent">

<div class="transactionreportshadow" style="max-width: 800px">
<table border="0" cellspacing="0" cellpadding="2" class="transactionreport liquidtable" style="max-width: 800px">
<thead>
  <tr>
      <td class="transaction_header" width="200">Submit Date</td>
      <td class="transaction_header" width="120">Bid Number</td>
      <td class="transaction_header" width="120">Title</td>
      <td class="transaction_header" width="120">Status</td>
      <td class="transaction_header" width="120">End Date</td>
      <td class="transaction_header" width="120">Uploaded Bid(s)</td>
  </tr>
  </thead>
<%
  sSQL = "SELECT userbidid, posting_id, posting_type, userid, orgid, submitdate, uploadid, filelocation, filename, userlabel "
  sSQL = sSQL & " FROM egov_jobs_bids_userbids "
  sSQL = sSQL & " WHERE orgid = " & iorgid
  sSQL = sSQL & " AND userid = "  & request.cookies("userid")
  sSQL = sSQL & " ORDER BY submitdate "

 	set oBids = Server.CreateObject("ADODB.Recordset")
 	oBids.Open sSQL, Application("DSN"), 3, 1

  if oBids.eof then
     response.write "<tr><td colspan=""6"" style=""border:0;"">None</td></tr>"
  else
     lcl_bgcolor = "#ffffff"
     while not oBids.eof
        lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        getPostingInfo oBids("posting_id"), lcl_jobbid_id, lcl_title, lcl_description, lcl_enddate, lcl_statusname

       'Build the filepath
        lcl_filelocation = oBids("filelocation")
        lcl_filelocation = replace(lcl_filelocation,"\custom\pub\","")

        lcl_file_url = Application("CommunityLink_DocUrl")
        lcl_file_url = lcl_file_url & "public_documents300/"
        lcl_file_url = lcl_file_url & lcl_filelocation
        lcl_file_url = lcl_file_url & oBids("filename")
        lcl_file_url = replace(lcl_file_url,"\","/")

        response.write "  <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "      <td class=""repeatheader"">Submit Date</td>" & vbcrlf
        response.write "      <td width=""200"">" & oBids("submitdate") & "</td>" & vbcrlf
        response.write "      <td class=""repeatheader"">Bid Number</td>" & vbcrlf
        response.write "      <td width=""120"">" & lcl_jobbid_id       & "</td>" & vbcrlf
        response.write "      <td class=""repeatheader"">Title</td>" & vbcrlf
        response.write "      <td width=""120"">" & lcl_title           & "</td>" & vbcrlf
        response.write "      <td class=""repeatheader"">Status</td>" & vbcrlf
        response.write "      <td width=""120"">" & lcl_statusname      & "</td>" & vbcrlf
        response.write "      <td class=""repeatheader"">End Date</td>" & vbcrlf
        response.write "      <td width=""120"">" & lcl_enddate         & "</td>" & vbcrlf
        response.write "      <td class=""repeatheader"">Uploaded Bid(s)</td>" & vbcrlf
        response.write "      <td width=""120"">" & vbcrlf
        'response.write "          <a href=""" & sEgovWebsiteURL & "/admin" & oBids("filelocation") & oBids("filename") & """ target=""_blank"">" & vbcrlf
        response.write "          <a href=""" & lcl_file_url & """ target=""_blank"">" & vbcrlf
        response.write "          [" & oBids("uploadid") & "]<br />" & vbcrlf
        response.write "          <span style=""color:#800000"">" & oBids("userlabel") & "</span></a>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        oBids.movenext
     wend
  end if

  oBids.close
  set oBids = nothing
%>
</table>
</div>
 	</div>
</div>

<p><br />&nbsp;<br />&nbsp;</p>

<!--#include file="include_bottom.asp"-->
