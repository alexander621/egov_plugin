<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<% Dim sError %>

<html>
<head>
<title>E-Gov Services <%=sOrgName%></title>
<link rel="stylesheet" href="css/styles.css" type="text/css">
<link rel="stylesheet" href="css/style_<%=iorgid%>.css" type="text/css">
<link href="global.css" rel="stylesheet" type="text/css">
<script language="Javascript" src="scripts/modules.js"></script>
<script language=javascript>
function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}
</script>
</head>


<!--#Include file="include_top.asp"-->


<!--BODY CONTENT-->


<TR><TD VALIGN=TOP>
	<div class=main>	
	
	<font class=pagetitle><%=sWelcomeMessage%></font> <BR>
	

	<!--BEGIN:  USER REGISTRATION - USER MENU-->
	<% If sOrgRegistration Then %>
			<%  If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
					RegisteredUserDisplay()
				Else %>
					<font class=datetagline>Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font>
			<% End If %>
	<% Else %>
		<font class=datetagline>Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font>
	<% End If%>
	<!--END:  USER REGISTRATION - USER MENU-->


	</font></div>
	<center>
	<table><tr>
	
	<!--BEGIN: CUSTOM SIDE CONTENT-->
		<td><!--#Include file="custom_left.asp"--></td>
	<!--END: CUSTOM SIDE CONTENT-->


	<td>	
	<table border=0 cellspacing=25 cellpadding=0 align=center>
	<tr>
		<%If blnOrgAction Then%>
		<td valign=top>
			<table cellspacing=0 cellpadding=0 border=0 class=box>
			  <tr><td valign=top class=box_edge><img src="images/left_side.jpg"></td><td height=20 class=box_header><%=sOrgActionName%></td><td valign=top class=box_edge ><img src="images/right_side.jpg"></td></tr>
			  <tr><td valign=top colspan=3 class=box_content><%fnListForms()'fn_DisplayTop5ActionLine(iOrgID)%></td></tr>
			</table>
		</td>
		<%End IF%>
		<%If blnOrgPayment Then%>
		<td valign=top>
			<table cellspacing=0 cellpadding=0 border=0 class=box>
			  <tr><td valign=top class=box_edge><img src="images/left_side.jpg"></td><td height=20 class=box_header><%=sOrgPaymentName%></td><td valign=top class=box_edge><img src="images/right_side.jpg"></td></tr>
			  <tr>
				<td valign=top colspan=3 class=box_content>
			  
					<%
					If iorgid = 1 Then
						
						response.write "<CENTER><BR><BR><BR><H4>COMING SOON!</h4></CENTER>"

					Else
				
					
							fn_DisplayTop5Payments(iOrgID)

					End If
					%>
						  
			   </td></tr>
			</table>
		</td>
		<%End If%>
	</tr>
	<tr>
	<%If blnOrgDocument Then%>
		<td valign=top>
			<table cellspacing=0 cellpadding=0 border=0 class=box>
			  <tr><td><img src="images/left_side.jpg"></td><td height=20 class=box_header><%=sOrgDocumentName%></td><td><img src="images/right_side.jpg"></td></tr>
			  <tr><td valign=top colspan=3 class=box_content>
				<ul style="padding-top:7px;">
					<!--#Include file="docs/menu/menuHome2.asp"-->
				</ul>
			  <span style="margin-left:25px;" ><a  href="docs/menu/home.asp">View all documents...</a><P>&nbsp;</P></span>
			  </td></tr>
			</table>
		</td>
	<%End IF%>


<%If blnOrgCalendar Then%>
		<td valign=top>
			<table cellspacing=0 cellpadding=0 border=0 class=box>
			  <tr><td><img src="images/left_side.jpg"></td><td height=20 class=box_header><%=sOrgCalendarName%></td><td><img src="images/right_side.jpg"></td></tr>
			  <tr><td valign=top colspan=3 class=box_content><!--#Include file="events/calendarSmall.asp"--></td></tr>
			</table>
			<% If blnCalRequest Then %>
				<center><a href="action.asp?actionid=<%=iCalForm%>">CLICK HERE<br>To Request to add Calendar Item</a></center>
			<% End If %>
		</td>
<%End If%>


	</tr>
	
	<!--JS REMOVED 12/2/2004--->
	<% If iorgid <> 99 Then response.write "<!--" %>
	<tr>
		<td colspan=2>
			<table cellspacing=0 cellpadding=0 border=0 class=box>
			  <tr><td><img src="images/left_side.jpg"></td><td height=20 class=box_header>Log In</td><td><img src="images/right_side.jpg"></td></tr>
			  <tr>
			  	<form name=frmLogin action=login.asp method=post>
			  	<td valign=top colspan=3 class=box_content>
					<% if request.cookies("userid") <> "" and request.cookies("userid") <> "-1" then
						if request.cookies("userid") = "-1" then response.write "Invalid Login<br>"
						sSQL = "SELECT useremail FROM egov_users WHERE userid = " & request.cookies("userid")
						Set name = Server.CreateObject("ADODB.Recordset")
						name.Open sSQL, Application("DSN") , 3, 1%>
						
						<div class=plainbox>
							You are logged in as: <% = name("useremail") %> <a href="logout.asp">click here</a> to logout</a><br>
							or <a href="account.asp">click here</a> to manage your account
						</div>
					<% else %>
						<table cellspacing=0 cellpadding=0 border=0 style="margin:5px;">
							<tr>
								<td>Email:</td>
								<td><input type=text name=email></td>
							</tr>
							<tr>
								<td>Password:</td>
								<td><input type=password name=password></td>
							</tr>
							<tr>
								<td colspan=2>
									<table cellpadding=0 cellspacing=5 border=0>
										<tr>
											<td rowspan=2><input class=loginbtn type=submit value=GO></td>
											<td>&nbsp;&nbsp;<a href="forgot_password.asp">Forgot Password?</a></td>
										</tr>
										<tr>	
											<td>&nbsp;&nbsp;<a href="register.asp">New User?</a></td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					<% end if %>
			  	</td>
				</form>
			  </tr>
			</table>
		</td>
	</tr>
	<% If iorgid <> 99 Then response.write "-->" %>


	</table>
	</td>
	
	
	<!--BEGIN: CUSTOM SIDE CONTENT-->
	<Td><!--#Include file="custom_right.asp"--></td>
	<!--END: CUSTOM SIDE CONTENT-->


	</tR></table>
	</center>


   
<!--#Include file="include_bottom.asp"-->    


<%
'--------------------------------------------------------------------------------------------------
' BEGIN: VISITOR TRACKING
'--------------------------------------------------------------------------------------------------
	iSectionID = 1
	sDocumentTitle = "MAIN"
	sURL = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
	datDate = Date()	
	datDateTime = Now()
	sVisitorIP = request.servervariables("REMOTE_ADDR")
	Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
'--------------------------------------------------------------------------------------------------
' END: VISITOR TRACKING
'--------------------------------------------------------------------------------------------------
%>



<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' FUNCTION FN_DISPLAYTOP5ACTIONLINE()
'--------------------------------------------------------------------------------------------------
Function fn_DisplayTop5ActionLine(iorgID)

	sSQL = "SELECT DISTINCT TOP 5 action_form_id,action_form_name FROM dbo.egov_form_list_200 where form_category_name='Top Requested Forms' and (orgid=" & iorgID & ") ORDER BY action_form_name"
	Set oActions = Server.CreateObject("ADODB.Recordset")
	oActions.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oActions.EOF Then

		response.write "<ul style=""padding-top:2px;"">"

		Do while NOT oActions.EOF 
			sTopic = Server.URLEncode("Top Requested Forms" & " > " & oActions("action_form_name"))
			
			response.write "<li><a href=""action.asp?actionid=" & oActions("action_form_id") & "&actiontitle=" & sTopic & "&list=true"">" & oActions("action_form_name") &  "</a>"

			oActions.MoveNext
		Loop
	

		response.write "</ul><span style=""margin-left:25px;"" ><a href=""action.asp"">View all Action Line Forms...</a></span>"
		response.write "</ul>"
	Else
		response.write "<P style=""padding-top:10px;""><center><font  color=red><B><I>No Action Forms enabled.</I></B></font></P>"
	End If

	Set oActions = Nothing 

End Function



'------------------------------------------------------------------------------------------------------------
' FUNCTION FN_DISPLAYTOP5PAYMENTS()
'------------------------------------------------------------------------------------------------------------
Function fn_DisplayTop5Payments(iOrgID)
	response.write "<ul style=""padding-top:2px;"">"
	sSQL = "SELECT DISTINCT TOP 5 paymentserviceid,paymentservicename, (select Count(*) FROM egov_paymentservices LEFT OUTER JOIN egov_organizations_to_paymentservices  ON egov_paymentservices.paymentserviceid=egov_organizations_to_paymentservices.paymentservice_id where (egov_organizations_to_paymentservices.paymentservice_enabled <> 0 and (egov_organizations_to_paymentservices.orgid=" & iorgid & " ))) as TotalAvailable FROM egov_paymentservices LEFT OUTER JOIN egov_organizations_to_paymentservices  ON egov_paymentservices.paymentserviceid=egov_organizations_to_paymentservices.paymentservice_id where (egov_organizations_to_paymentservices.paymentservice_enabled <> 0 and (egov_organizations_to_paymentservices.orgid=" & iorgid & " ))"
	Set oPaymentServices = Server.CreateObject("ADODB.Recordset")
	oPaymentServices.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oPaymentServices.EOF Then
		iNumAvailable = oPaymentServices("TotalAvailable") 

		Do while NOT oPaymentServices.EOF 
			response.write "<li><a href=""payment.asp?paymenttype=" & oPaymentServices("paymentserviceid") & """>" & oPaymentServices("paymentservicename") &  "</a>"
			oPaymentServices.MoveNext
		Loop

	End If
	Set oPaymentServices = Nothing 

	If iNumAvailable  < 5 Then
		' LESS THAN 5 DON'T DISPLAY LIST
		response.write "</ul><span style=""margin-left:25px;"" ><a  href=""payment.asp"">Learn more...</a></span>"
	Else
		' 5 OR MORE SO DISPLAY LIST
		response.write "</ul><span style=""margin-left:25px;"" ><a  href=""payment.asp"">View more permits and payments...</a></span>"
	End If 

	response.write "</ul>"
End Function

'--------------------------------------------------------------------------------------------------
' FUNCTION FNLISTFORMS()
'--------------------------------------------------------------------------------------------------
Function fnListForms()
	
	sLastCategory = "NONE_START"
	sSQL = "SELECT * FROM dbo.egov_form_list_200  WHERE ((orgid=" & iorgID & ")) AND (form_category_id <> 6) order by form_category_Sequence,action_form_name"

	Set oForms = Server.CreateObject("ADODB.Recordset")
	oForms.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oForms.EOF Then

	iNumAvailable = oForms.RecordCount
	
	response.write "<ul>"

		Do while NOT oForms.EOF 

			sCurrentCategory = oForms("form_category_name")
			If sLastCategory = "NONE_START" Then
				response.write "<li><a  href=""action.asp#" & oForms("form_category_id") & """>" & sCurrentCategory & "</a>"
			End If

			If (sCurrentCategory <> sLastCategory) AND (sLastCategory <> "NONE_START") Then
				response.write "<li><a  href=""action.asp#" & oForms("form_category_id") & """>" & sCurrentCategory & "</a>"
			End If
			
		
			oForms.MoveNext

			sLastCategory = sCurrentCategory
		Loop

	response.write "</ul>"

	End If
	
	If iNumAvailable > 6 Then 
		response.write "</ul><span style=""margin-left:25px;"" ><a  href=""action.asp"">View all Action Line Forms...</a></span>"
	Else
		response.write "</ul><span style=""margin-left:25px;"" ><a  href=""action.asp"">Learn more...</a></span>"
	End If

	Set oForms = Nothing 

End Function
%>








	   
