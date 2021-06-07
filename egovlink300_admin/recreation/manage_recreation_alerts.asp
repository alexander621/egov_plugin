<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: manage_recreation_alerts.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 03/21/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "rec payment alerts" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>
<html>
<head>
  <title>Recreation Payment Alerts</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="facility.css" />

  <script language="Javascript" src="../scripts/selectAll.js"></script>
  <script language="Javascript" src="../scripts/modules.js"></script>

  <script language="Javascript">
	<!--

	String.prototype.trim = function() {
		return this.replace(/^\s+|\s+$/,"");
	}

	//-->
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <!--<td width="151" align="center"><img src="../images/icon_home.jpg"></td>-->
      <td>
			<font size="+1"><b>Recreation Payment Alerts</b></font>
			<!--<br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.back()"><%=langBackToStart%></a>-->
	  </td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2" valign="top">
	  <!--BEGIN: ALERT LIST -->
		
<%
		listAlerts 
%>		
	  
	  <!-- END: ALERT LIST -->
      </td>
       
    </tr>
  </table>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub listAlerts( )
'--------------------------------------------------------------------------------------------------
Sub listAlerts( )
	Dim orderBy, sSql, oRequests

	If request("useSessions") = 1 Then 
		orderBy = Session("orderBy")
	Else 
		orderBy = request("orderBy")

		If orderBy = "" Or IsNull(orderBy) Then  
			orderBy = "paymentservicename"
		Else 
			orderBy = orderBy
		End If 
	End If 
	session("orderBy") = orderBy


	' LIST Alerts 
	sSql = "SELECT DISTINCT paymentserviceid, paymentservicename, assigned_email, egov_paymentservices.assigned_userID, LastName, FirstName "
	sSql = sSql & "FROM egov_paymentservices  LEFT OUTER JOIN Users ON egov_paymentservices.assigned_userID = Users.userID "
	sSql = sSql & "WHERE egov_paymentservices.orgid = " & Session("OrgID") & " AND paymentserviceenabled = 0 ORDER BY paymentservicename"
	
	Set oRequests = Server.CreateObject("ADODB.Recordset")

	' SET PAGE SIZE AND RECORDSET PARAMETERS
	oRequests.PageSize = 10
	oRequests.CacheSize = 10
	oRequests.CursorLocation = 3

	oRequests.Open sSql, Application("DSN"), 3, 1
	'response.write "<!--" & sSql & "-->"

	If oRequests.EOF Then 
		response.write "There are no recreation payments alerts to manage."
	Else 

		If request("useSessions") = 1 Then 
			If Len(Session("pagenum")) <> 0 then
				oRequests.AbsolutePage = clng(Session("pagenum"))	
				Session("pageNum") = clng(Session("pagenum"))		
			Else
				oRequests.AbsolutePage = 1
				Session("pageNum") = 1
			End If
		Else 
			If Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1  Then
				oRequests.AbsolutePage = 1
				Session("pageNum") = 1
				'Response.write "Issue 2"
			Else
				If clng(Request("pagenum")) <= oRequests.PageCount Then
					oRequests.AbsolutePage = Request("pagenum")
					Session("pageNum") = Request("pagenum")
				Else
					oRequests.AbsolutePage = 1
					Session("pageNum") = 1
				End If
			End If
		End If 

		' DISPLAY RECORD STATISTICS
		Dim abspage, pagecnt
		abspage = oRequests.AbsolutePage
		pagecnt = oRequests.PageCount

		sQueryString = replace(request.querystring,"pagenum","HFe301") ' REPLACE PAGENUM FIELD WITH RANDOM FIELD FOR NAVIGATION PURPOSES

		'Response.Write "<b>Page <font color=blue>" & oRequests.AbsolutePage & "</font>  " & vbcrlf
		'Response.Write "of <font color=blue> " & oRequests.PageCount & "</font></b> &nbsp;|&nbsp; " & vbcrlf
		'Response.Write " <b><font color=blue>" & oRequests.RecordCount & "</font> total Online Payment forms"
		'response.write "</b><br /><br />"

		' DISPLAY FORWARD AND BACKWARD NAVIGATION TOP
		'Response.write "<div><table><tr><td valign=top><a href=""manage_action_forms.asp?pagenum=" & abspage - 1 & "&" & sQueryString & """><img border=0 src=""../images/arrow_back.gif""></a> <a href=""manage_action_forms.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a></td><td>&nbsp;&nbsp;"  & "<a href=""manage_action_forms.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a> <a href=""manage_action_forms.asp?pagenum=" & abspage + 1 & "&" & sQueryString & """> <img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td></tr></table></div>"

		response.write "<div class=""shadow"">"
		Response.Write "<table cellspacing=""0"" cellpadding=""5"" class=""tablelist"" width=""100%"">"
		'Response.Write "<tr class=""tablelist""><th align=""left"">ID</th>"
		Response.Write "<th align=""left"">Alert</th>"
		Response.Write "<th align=""left"">Assigned To</th></tr>"
		bgcolor = "#eeeeee"
		response.write "<tbody>"
		iRow = 0

		' LOOP AND DISPLAY THE RECORDS
		For intRec = 1 To oRequests.PageSize
			If Not oRequests.EOF Then
				If bgcolor = "#eeeeee" Then
					bgcolor = "#ffffff" 
				Else
					bgcolor = "#eeeeee"
				End If
				iRow = iRow + 1
				Response.Write vbcrlf & "<tr id=""" & iRow & """ bgcolor=""" & bgcolor & """ onClick=""location.href='edit_form.asp?control=" & oRequests("paymentserviceid") & "';"" onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
				'response.write "<td width=25><b>" & oRequests("paymentserviceid") & "</b></td>"
				response.write "<td><b>" & oRequests("paymentservicename") & "</b></td>"
				response.write "<td>" & oRequests("FirstName") & " " & oRequests("LastName") & " <b>" & oRequests("assigned_email") & "</b></td>"
				response.write "</tr>"
				oRequests.MoveNext 
			End If
		Next
		response.write "</tbody>"
		Response.Write vbcrlf & "</table>"
		response.write "</div><br />"

		' DISPLAY FORWARD AND BACKWARD NAVIGATION BOTTOM
		'Response.write "<div><table><tr><td valign=top><a href=""manage_action_forms.asp?pagenum=" & abspage - 1 & "&" & sQueryString & """><img border=0 src=""../images/arrow_back.gif""></a> <a href=""manage_action_forms.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a></td><td>&nbsp;&nbsp;"  & "<a href=""manage_action_forms.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a> <a href=""manage_action_forms.asp?pagenum=" & abspage + 1 & "&" & sQueryString & """><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td></tr></table></div>"

	End If 

	oRequests.Close
	Set oRequests = Nothing 

End Sub 



%>


