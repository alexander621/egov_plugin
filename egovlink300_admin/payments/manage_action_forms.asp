<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: manage_action_forms.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the payments list
'
' MODIFICATION HISTORY
' 1.0   ???			???? - INITIAL VERSION
' 1.1	10/12/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "payment notifications" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>

<html>
<head>
  <title><%=langBSPayments%></title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

  <script src="../scripts/selectAll.js"></script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%'DrawTabs tabPayments,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <!--<td width="151" align="center"><img src="../images/icon_home.jpg"></td>-->
      <td><font size="+1"><b>Manage Online Payment Forms</b></font>
		  <!--<br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.back()"><%=langBackToStart%></a>-->
	  </td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <!--<td valign="top" nowrap>-->

        <!-- START: QUICK LINKS MODULE //-->
        
        <%
'        sLinks = "<div style=""padding-bottom:8px;""><b>" & langEventLinks & "</b></div>"

'        If bCanEdit Then
'          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/calendar.gif"" align=""absmiddle"">&nbsp;<a href=""newevent.asp"">" & langNewEvent & "</a></div>"
'          bShown = True
'        End If
        
'        If bShown Then
'          Response.Write sLinks & "<br>"
'        End If
        %>

        <% 'DrawQuicklinks("", 1) %>
        <!-- END: QUICK LINKS MODULE //-->

      <!--</td>-->
        
      <td colspan="2" valign="top">
	  <!--BEGIN: ACTION LINE REQUEST LIST -->
		
	    <% List_Forms sSortBy %>		
	  
	  <!-- END: ACTION LINE REQUEST LIST -->
      </td>
       
    </tr>
  </table>
  
  </div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' Sub List_Forms(sSortBy)
'--------------------------------------------------------------------------------------------------
Sub List_Forms( sSortBy )

if request("useSessions")=1 then
		orderBy = Session("orderBy")
else
			orderBy = request("orderBy")
			
			If orderBy = "" or IsNull(orderBy) Then 
					orderBy = "paymentservicename"
			else
					orderBy = orderBy
			end if
end if
session("orderBy") = orderBy


' LIST ACTION REQUESTS 
'sSQL = "SELECT * FROM egov_paymentservices"
sSQL = "SELECT DISTINCT paymentserviceid,paymentservicename,assigned_email,egov_paymentservices.assigned_userID,LastName,FirstName " _
		& " FROM egov_paymentservices " _
		& " INNER JOIN egov_organizations_to_paymentservices  ON egov_paymentservices.paymentserviceid=egov_organizations_to_paymentservices.paymentservice_id  " _
		& " left outer join Users on egov_paymentservices.assigned_userID=Users.userID " _
		& " WHERE (egov_organizations_to_paymentservices.paymentservice_enabled <> 0 AND (egov_organizations_to_paymentservices.orgid=" & Session("OrgID") & "))  " _
		& " order by " & orderBy ''' paymentserviceid"
Set oRequests = Server.CreateObject("ADODB.Recordset")
 
 ' SET PAGE SIZE AND RECORDSET PARAMETERS
 oRequests.PageSize = 10
 oRequests.CacheSize = 10
 oRequests.CursorLocation = 3
 
 ' OPEN RECORDSET
 oRequests.Open sSQL, Application("DSN"), 3, 1
 
 if oRequests.EOF then%>
	 There are no forms to manage.
 <%else

 ' SET PAGE TO VIEW
 'If Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1  Then
'	  oRequests.AbsolutePage = 1
 'Else
'	If clng(Request("pagenum")) <=oRequests.PageCount Then
'		oRequests.AbsolutePage = Request("pagenum")
'	Else
'		oRequests.AbsolutePage = 1
'	End If
 'End If
 
 if request("useSessions")=1 then
	If Len(Session("pagenum")) <> 0 then
		oRequests.AbsolutePage = clng(Session("pagenum"))	
		Session("pageNum") = clng(Session("pagenum"))		
	Else
		oRequests.AbsolutePage = 1
		Session("pageNum") = 1
	End If
	'Response.write "Issue 1"
 else
	 If Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1  Then
		oRequests.AbsolutePage = 1
		Session("pageNum") = 1
		'Response.write "Issue 2"
	 Else
		If clng(Request("pagenum")) <= oRequests.PageCount Then
			oRequests.AbsolutePage = Request("pagenum")
			Session("pageNum") = Request("pagenum")
			'Response.write "Issue 3: " & oRequests.PageCount & "-" & clng(Request("pagenum"))
		Else
			oRequests.AbsolutePage = 1
			Session("pageNum") = 1
			'Response.write "Issue 4"
		End If
							
 	End If
 end if
	 
	 

 ' DISPLAY RECORD STATISTICS
  Dim abspage, pagecnt
  abspage = oRequests.AbsolutePage
  pagecnt = oRequests.PageCount
  
  'If request("selectAssignedto") <> "" Then
	 sQueryString = replace(request.querystring,"pagenum","HFe301") ' REPLACE PAGENUM FIELD WITH RANDOM FIELD FOR NAVIGATION PURPOSES
  'Else
	' sQueryString = "filter=false"
  'End If
 
  Response.Write "<b>Page <font color=blue>" & oRequests.AbsolutePage & "</font>  " & vbcrlf
  Response.Write "of <font color=blue> " & oRequests.PageCount & "</font></b> &nbsp;|&nbsp; " & vbcrlf
  Response.Write " <b><font color=blue>" & oRequests.RecordCount & "</font> total Online Payment forms"
  response.write "</b><br /><br />"
  
  'Response.Write "Total number of requests : <font color=blue>" & oRequests.RecordCount
  'response.write "</font></b>"

 ' DISPLAY FORWARD AND BACKWARD NAVIGATION TOP
  Response.write "<div><table><tr><td valign=""top""><a href=""manage_action_forms.asp?pagenum="&abspage - 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_back.gif""></a> <a href=""manage_action_forms.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a></td><td valign=""top"">&nbsp;&nbsp;"  & "<a href=""manage_action_forms.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a> <a href=""manage_action_forms.asp?pagenum="&abspage + 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td></tr></table></div>"
  response.write "<div class=""shadow"">"
  Response.Write "<table cellspacing=0 cellpadding=5 class=tablelist width=""100%"">"
  Response.Write "<tr class=tablelist><th align=left>ID</th>"
  Response.Write "<th align=left><a href=manage_action_forms.asp?orderBy=paymentservicename>Action Line Form Name</a></th>"
  Response.Write "<th align=left><a href=manage_action_forms.asp?orderBy=assigned_email>Assigned To</a></th></tr>"
  bgcolor = "#eeeeee"
	
  ' LOOP AND DISPLAY THE RECORDS
	 For intRec=1 To oRequests.PageSize
	  If Not oRequests.EOF Then
		
		If bgcolor="#eeeeee" Then
			bgcolor="#ffffff" 
		Else
			bgcolor="#eeeeee"
		End If

		Response.Write "<tr bgcolor=" & bgcolor & " onClick=""location.href='edit_form.asp?control=" & oRequests("paymentserviceid") & "';"" onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';""><td width=25><b>" & oRequests("paymentserviceid") & "</b></td><td><b>" & oRequests("paymentservicename") & "</b></td><td>" & oRequests("FirstName") & " " & oRequests("LastName") & " <b>" & oRequests("assigned_email") & "</b></td></tr>"
		oRequests.MoveNext 

	  End If
	 Next
	 Response.Write "</table>"
	 response.write "</div><br />"

' DISPLAY FORWARD AND BACKWARD NAVIGATION BOTTOM
  Response.write "<div><table><tr><td valign=""top""><a href=""manage_action_forms.asp?pagenum="&abspage - 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_back.gif""></a> <a href=""manage_action_forms.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a></td><td valign=""top"">&nbsp;&nbsp;"  & "<a href=""manage_action_forms.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a> <a href=""manage_action_forms.asp?pagenum="&abspage + 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td></tr></table></div>"

end if

End Sub 


%>


