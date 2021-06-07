<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="postings_global_functions.asp" //-->
<!-- #include file="include_top_functions.asp"-->
<!-- #include file="class/classOrganization.asp" -->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: view_planholders.asp
' AUTHOR: David Boyer
' CREATED: 08/19/09
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  User can view a distinct list of all other citizens that have clicked on a download available link(s) on a posting.
'
' MODIFICATION HISTORY
' 1.0	 08/19/09	David Boyer	- Initial Version
' 1.1  08/20/09 David Boyer - Query changed to pull a distinct list of users who have actually uploaded a bid to the posting.
' 1.2  10/27/09 David Boyer - Changed the query BACK to initial version via Peter Selden request.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sTitle, lcl_posting_id, re, matches

sTitle = "View Plan Holders"

if request("posting_id") <> "" then
lcl_posting_id = request("posting_id")
else
lcl_posting_id = ""
end If

Set re = New RegExp
re.Pattern = "^\d+$"

If lcl_posting_id = "" Then 
	lcl_posting_id = CLng(0)
Else
	Set matches = re.Execute(lcl_posting_id)
	If matches.Count > 0 Then
		lcl_posting_id = CLng(lcl_posting_id)
	Else
		lcl_posting_id = CLng(0)
	End If 
end If

%>
<html>
<head>
	<title>E-Gov Services <%=sOrgName & " - " & sTitle%></title>

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
<body style="margin-top:5px;">
<div id="content">
 	<div id="centercontent">

<font class="pagetitle"><%=sTitle%></font><br />
<input type="button" name="closeWindowButton" id="closeWindowButton" value="Close Window" class="button" style="margin-top:5px;" onclick="parent.close();" />
<div class="transactionreportshadow" style="width: 800px">
<table border="0" cellspacing="0" cellpadding="2" class="transactionreport" style="width:900px">
  <tr>
      <td class="transaction_header" width="200">BID ID</td>
      <td class="transaction_header" width="120">Title</td>
      <td class="transaction_header" width="120">Plan Holder</td>
      <td class="transaction_header" width="120">Business Name</td>
      <td class="transaction_header" width="120">Phone</td>
      <td class="transaction_header" width="120">Email</td>
      <td class="transaction_header" width="120">Fax</td>
  </tr>
<%
  if lcl_posting_id <> "" then
    'WHY 2 QUERIES??? ---------------------------------------------------------
    'This query pulls a distinct list of citizens that have clicked on a "download available" link(s).
    'This was the original version of this screen.  However, it was changed to the query below via phone meeting
    'request by the client.  Once in production, after testing and verification that it was indeed what they wanted (query 2)
    'it was believed that query 2 was not what they wanted, as there was confusion on how to get data into the report.
    'When Peter asked the client they said they wanted query 1, which is what they originally had and asked to be changed.
    'Therefore, it has been changed BACK to the original version on 10/27/09.

    'QUERY 1 ------------------------------------------------------------------
     sSql = "SELECT distinct "
     sSql = sSql & " isnull(eu.userlname,'') + ', ' + isnull(eu.userfname,'') AS username, "
     sSql = sSql & " isnull(jb.jobbid_id,'') AS jobbid_id, "
     sSql = sSql & " isnull(jb.title,'') AS title, "
     sSql = sSql & " eu.userbusinessname, "
     sSql = sSql & " eu.userworkphone, "
     sSql = sSql & " eu.useremail, "
     sSql = sSql & " eu.userfax "
     sSql = sSql & " FROM egov_clickcounter_postings c "
     sSql = sSql &      " LEFT OUTER JOIN egov_jobs_bids jb ON c.posting_id = jb.posting_id "
     sSql = sSql &      " LEFT OUTER JOIN egov_users eu ON c.userid = eu.userid "
     sSql = sSql & " WHERE c.orgid = " & iOrgID
     sSql = sSql & " AND c.posting_id = " & Track_DBsafe( lcl_posting_id )
     'sSql = sSql & " ORDER BY c.clicked_linktext, c.clicked_date DESC, 4 "

    'QUERY 2 ------------------------------------------------------------------
    'This query pulls a distinct list of citizens that have actually uploaded a bid to the posting
    'Disabled and revert back to initial version via Peter Selden (10/27/09)
     'sSql = "SELECT distinct ub.userid, "
     'sSql = sSql & " isnull(eu.userlname,'') + ', ' + isnull(eu.userfname,'') AS username, "
     'sSql = sSql & " jb.jobbid_id, "
     'sSql = sSql & " jb.title, "
     'sSql = sSql & " eu.userbusinessname, "
     'sSql = sSql & " eu.userworkphone, "
     'sSql = sSql & " eu.useremail, "
     'sSql = sSql & " eu.userfax "
     'sSql = sSql & " FROM egov_jobs_bids_userbids ub "
     'sSql = sSql &      " LEFT OUTER JOIN egov_jobs_bids jb ON ub.posting_id = jb.posting_id "
     'sSql = sSql &      " LEFT OUTER JOIN egov_users eu ON ub.userid = eu.userid "
     'sSql = sSql & " WHERE ub.orgid = " & iOrgID
     'sSql = sSql & " AND ub.posting_id = " & lcl_posting_id

    	set oPlanHolders = Server.CreateObject("ADODB.Recordset")
    	oPlanHolders.Open sSql, Application("DSN"), 3, 1

     if not oPlanHolders.eof then
        lcl_bgcolor = "#ffffff"
        while not oPlanHolders.eof
           lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

           response.write "  <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
           response.write "      <td style=""width:150px"">" & oPlanHolders("jobbid_id")                        & "</td>" & vbcrlf
           response.write "      <td style=""width:200px"">" & oPlanHolders("title")                            & "</td>" & vbcrlf
           response.write "      <td nowrap=""nowrap"">"     & oPlanHolders("username")                         & "</td>" & vbcrlf
           response.write "      <td style=""width:150px"">" & oPlanHolders("userbusinessname")                 & "</td>" & vbcrlf
           response.write "      <td nowrap=""nowrap"">"     & formatphonenumber(oPlanHolders("userworkphone")) & "</td>" & vbcrlf
           response.write "      <td style=""width:150px"">" & oPlanHolders("useremail")                        & "</td>" & vbcrlf
           response.write "      <td nowrap=""nowrap"">"     & formatphonenumber(oPlanHolders("userfax"))       & "</td>" & vbcrlf
           response.write "  </tr>" & vbcrlf

           oPlanHolders.movenext
        wend
     end if

     oPlanHolders.close
     set oPlanHolders = nothing
  end if
%>
</table>
</div>
 	</div>
</div>

<p><br />&nbsp;<br />&nbsp;</p>
</body>
</html>
