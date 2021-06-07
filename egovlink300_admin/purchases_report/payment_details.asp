<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->

<%
'prevAssignedemployeeid


' GET INFORMATION FOR CURRENT INSTITUTION
'Dim sTopGraphicLeftURL,sHomeWebsiteURL,sTopGraphicRighURL,sEgovWebsiteURL,iPaymentGatewayID 
'Dim sWelcomeMessage,sTagline,sActionDescription,sOrgName,iHeaderSize,sPaymentDescription
iOrgID = SetOrganizationParameters()


' IF UPDATE PROCESS ITEMS
If Request.ServerVariables("REQUEST_METHOD") = "POST" THEN
	Update_Action(request("TrackID"))
	iTrackID = request("TrackID")
	blnUpdate = True
End If


' GET INFORMATION FOR THIS REQUEST
If iTrackID = "" Then
	iTrackID = request("iPaymentid") 
End If

sSQL = "SELECT * FROM dbo.egov_payment_list where paymentid=" & iTrackID & " ORDER BY paymentdate"

Set oRequest = Server.CreateObject("ADODB.Recordset")
oRequest.Open sSQL, Application("DSN") , 3, 1

' CHECK FOR INFORMATION
If NOT oRequest.EOF Then
	' REQUEST FOUND GET INFORMATION	
	blnFound = True
	sTitle = oRequest("paymentservicename")
	sStatus = oRequest("paymentstatus")
	datSubmitDate = oRequest("paymentdate")
	sDetails = oRequest("paymentsummary")
	sPaymentInfo = oRequest("payment_information")
	iUserid = oRequest("userid")
	sTheUserid = oRequest("userid")
	iemployeeid =  oRequest("assigned_userid")
	cPaymentAmount = oRequest("paymentamount")

	If datSubmitDate <> "" Then
		lngTrackingNumber = iTrackID  & replace(FormatDateTime(cdate(datSubmitDate),4),":","")
	Else
		lngTrackingNumber = "000000000"
	End If
Else
	' REQUEST NOT FOUND
	blnFound = False
End If

Set oRequest = Nothing 
%>

<html>
<head>
  <title><%=langBSPayments%></title>

  <link href="../global.css" rel="stylesheet" type="text/css">
  <link type="text/css" media="print" rel="stylesheet" href="receiptprint.css" />
  
  <script src="../scripts/selectAll.js"></script>
  <script src="../scripts/layers.js"></script>
  
  <script>
	function update_user_display(){
	
	if (document.getElementById('user_on').value == 'on'){
		document.getElementById('user_expand').innerHTML = '<b>+ <u>Contact Information:</u></b>';
		document.getElementById('user_on').value = 'off';}
	
	else {
		document.getElementById('user_expand').innerHTML = '<b>- <u>Contact Information:</u></b>';
		document.getElementById('user_on').value = 'on';
		}
	}
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="toggleDisplay('contact_user');update_user_display();toggleDisplay('info');toggleDisplay('details')">

<%DrawTabs tabRecreation, 1%>


<div id="content">
	<div id="centercontent">
	<div id="receiptlinks">
		<img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)""><%=langBackToStart%></a><span id="printbutton"><input type="button" onclick="javascript:window.print();" value="Print" /></span>
	</div>

	<h3><%=Session("sOrgName")%> Online Payment Details</h3><br />

	<table border="0" cellpadding="10" cellspacing="0" class="start">
    <tr>
      <td></td>
    </tr>
    <tr>
      <td valign="top">
	  <!--BEGIN: PAYMENT LIST -->
      
			<%' GET CONTACT INFORMATION
				fnDisplayUserInfo( iUserid )
			%>
			
			<%'DISPLAY TRANSACTION DETAILS
			If sDetails <> "" Then
				response.write vbcrlf & "<div class=""purchasereportshadow"">"
				response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
				response.write "<tr><th colspan=""2"" align=""left"">Transaction Details</th></tr>"
				response.write "<tr><td width=""20%"">Payment Date: </td><td>" & DateValue(datSubmitDate) & "</td></tr>"
				response.write "<tr><td>Payment Method: </td><td>Credit Card</td></tr>"
				response.write "<tr><td>Payment Location: </td><td>Online</td></tr>"
				response.write "<tr><td>Amount: </td><td>" & FormatCurrency(cPaymentAmount,2) & "</td></tr>"
				response.write "</table></div>"
			Else
				response.write "<p><br><font color=red>!No transaction details available!</font></p>"
			End If
			%>

			<%'DISPLAY PAYMENT DETAILS
			If sPaymentInfo <> "" Then
				response.write vbcrlf & "<div class=""purchasereportshadow"">"
				response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
				response.write "<tr><th colspan=""2"" align=""left"">Payment Information</th></tr>"
				response.write "<tr><td width=""20%"">Payment Type:</td><td>" & sTitle & "</td></tr>"
				response.write "<tr><td>Tracking No:</td><td>" & lngTrackingNumber & "</td></tr>"
				response.write "<tr><td valign=""top"">Details:</td><td>" & replace(sPaymentInfo,"custom_","") & "</td></tr>"
				response.write "</table></div>"
			Else
				response.write "<p><font color=red>!No payment information available!</font></p>"
			End If
			%>

			</p>
			</div>
	  
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

Function Update_Action(iID)

	' UPDATE EMPLOYEE ASSIGNED TO TASK
	arrEmployee = split(request("assignedemployeeid"),",")
	iEmployeeID = arrEmployee(0)
	sEmployeeName = arrEmployee(1)


	sSQL = "SELECT * FROM dbo.egov_payment_list where paymentid=" & iID
	Set oUpdate = Server.CreateObject("ADODB.Recordset")
	oUpdate.CursorLocation = 3
	oUpdate.Open sSQL, Application("DSN") , 1, 2
	If oUpdate("assigned_userid") <> iEmployeeID OR ISNULL(oUpdate("assigned_userid")) Then
		oUpdate("assigned_userid") = iEmployeeID
				if iEmployeeID <> request("prevAssignedemployeeid") then
						sCommentLine = " This item has been re-assigned to " & sEmployeeName & "."
						'''AddCommentTaskComment sCommentLine,NULL,request("selStatus"), iID,Session("UserID"),Session("OrgID")
						newCommentTask = 1
						''**
				end if
				''**
		oUpdate.Update
	End If
	Set oUpdate = Nothing

	' UPDATE STATUS FOR REQUEST
	If request("selStatus") <> request("currentStatus") Then
		sSQL = "UPDATE egov_payments SET paymentstatus='" & request("selStatus") & "' where paymentid=" & iID
		Set oUpdate = Server.CreateObject("ADODB.Recordset")
		oUpdate.Open sSQL, Application("DSN") , 3, 1
		Set oUpdate = Nothing
	End If


	' ADD COMMENTS ROW
	' IF STATUS HAS CHANGED OR THERE ARE EXTERNAL OR INTERNAL COMMENTS TO BE ADDED
	''If (request("selStatus") <> request("currentStatus")) Or Trim(request("internal_comment")) <> "" Or Trim(request("external_comment")) <> ""  Then
	If (request("selStatus") <> request("currentStatus")) Or Trim(request("internal_comment")) <> "" Or Trim(request("external_comment")) <> "" OR newCommentTask = 1 Then
			  if request("internal_comment") <> "" and sCommentLine <> "" then
			  			intComment = request("internal_comment") & "<br>" & sCommentLine
			  elseif request("internal_comment") <> "" then
			  			intComment = request("internal_comment")
			  else
			  			intComment = sCommentLine
			  end if
				AddCommentTaskComment intComment,request("external_comment"),request("selStatus"), iID,Session("UserID"),Session("OrgID")
				''AddCommentTaskComment request("internal_comment"),request("external_comment"),request("selStatus"), iID,Session("UserID"),Session("OrgID")
	End If

	' SEND EMAIL TO Submitter
	If request("sendemail") = "yes" Then
		iTrackID = request("TrackID")

		sSQLs = "SELECT * FROM dbo.egov_payment_list where paymentid=" & iTrackID
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.Open sSQLs, Application("DSN") , 3, 1
		
		lngTrackingNumber = iTrackID  & replace(FormatDateTime(cdate(oRS("paymentdate")),4),":","")
		
		' SEND EMAIL TO USER
			sMsg = sMsg & "This automated message was sent by the  " & sOrgName & "  E-Gov web site. Do not reply to this message.  Follow the instruction below or contact " & oRS("assigned_email") & " for inquiries regarding this email." & vbcrlf 
			sMsg = sMsg & " " & vbcrlf 
			sMsg = sMsg & "The status of your Online Payment has been updated, or new information has been added." & vbcrlf  & vbcrlf
			sMsg = sMsg & "---------------------------------------------------------------------------------------------------" & vbcrlf 
			sMsg = sMsg & " " & vbcrlf 
			
			sMsg = sMsg & "The current status or your Online Payment is: '" & UCASE(request("selStatus")) & "'" & vbcrlf 
			if Trim(request("external_comment")) <> "" then
				sMsg = sMsg & "Newest information available: '" & replace(request("external_comment"),"'","`") & "'" & vbcrlf 
			end if
			sMsg = sMsg & "Payment Service Name: " & oRS("paymentservicename")   & vbcrlf & vbcrlf
			sMsg = sMsg & "Your Tracking Number is: " & lngTrackingNumber & " " & vbcrlf 
			sMsg = sMsg & "Payment was submitted on: " & oRS("paymentdate") & " " & vbcrlf 
			sMsg = sMsg & "---------------------------------------------------------------------------------------------------" & vbcrlf 
			sMsg = sMsg & " " & vbcrlf
			sMsg = sMsg & "Thank you for using the  " & sOrgName & "  messaging service.  We want to understand what you want and expect, make it easier for you to do business with the City, and respond as quickly as practical to your requests." & vbcrlf 
			sMsg = sMsg & " " & vbcrlf 

			' CREATE MAIL OBJECT
      If NOT ISNULL(oRS("useremail")) Then
				sEmail = oRS("useremail")
			Else
				sEmail = "UNKNOWN@EGOVLINK.COM"
			End If

			'DON'T SEND EMAIL
			'sendEmail "", sEmail, "", sOrgName & "  E-GOV MSG - ONLINE PAYMENT UPDATE", "", sMsg, "Y"

	End If
End Function


Function List_Comments(iID)

	sSQL = "SELECT * FROM egov_payment_responses INNER JOIN users on egov_payment_responses.action_userid=users.userid where action_autoid=" & iID & " ORDER BY action_editdate DESC"
	Set oCommentList = Server.CreateObject("ADODB.Recordset")
	oCommentList.Open sSQL, Application("DSN") , 3, 1
    sBGColor = "#FFFFFF"
	

	If NOT oCommentList.EOF Then
		Do While NOT oCommentList.EOF 
			If sBGColor = "#FFFFFF" Then
				sBGColor = "#eeeeee"
			Else
				sBGColor = "#FFFFFF"
			End If
			Response.Write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & """><table>"
			response.write "<tr><td>" & oCommentList("firstname") & " " & oCommentList("lastname") & " - " & UCASE(oCommentList("action_status")) & " - " &  oCommentList("action_editdate") & "</td></tr>"
			
			If oCommentList("action_externalcomment") <> "" Then
				response.write "<tr><td>&nbsp;&nbsp;&nbsp;<b>External Note: </b><i>" & oCommentList("action_externalcomment")  & "</i></td></tr>" 
			End If

			If oCommentList("action_internalcomment") <> "" Then
				response.write "<tr><td>&nbsp;&nbsp;&nbsp;<b>Internal Note: </b><i>" & oCommentList("action_internalcomment")  & "</i></td></tr>" 
			End If
			Response.Write "</table></div>"

			oCommentList.MoveNext
		Loop
	Else
		Response.Write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & """><table>"
		response.write "<tr><td><font color=red>&nbsp;&nbsp;&nbsp;<i>No activity</i></td></tr></table></div>"
		
	End If
 
End Function


Function CheckSelected(sValue,sValue2)
	sReturnValue = ""
	If uCase(sValue) = sValue2 Then
		sReturnValue = "SELECTED"
	End If

	CheckSelected = sReturnValue
End Function

Function fnDisplayUserInfo(iID)

	' GET INFORMATION FOR SPECIFIED USER
	sSQL = "Select * From egov_users where userid= " & iID
'	if iID="" or IsNull(iID) then
'		response.write "<div style=""display:none;"" id=contact_user><font color=red><i>No information available for specified user.</i></font></div>"
'	else
		' OPEN RECORDSET
		Set oUser = Server.CreateObject("ADODB.Recordset")
		oUser.Open sSQL, Application("DSN"), 3, 1
		
		If NOT oUser.EOF Then
			'response.write "<div id=contact_user style=""margin-top:5px;border:solid 1px #000000;background-color:#eeeeee;display:none;""><table>"
			response.write vbcrlf & "<div class=""purchasereportshadow"">"
			response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
			response.write "<tr><th colspan=""2"" align=""left"">Contact Information</th></tr>"
			response.write "<tr><td width=""20%"">Name:</td><td>" & oUser("userfname") & " " & oUser("userlname") & "</td></tr>"
			response.write "<tr><td>Business Name:</td><td>" & oUser("userbusinessname") & "</td></tr>"
			response.write "<tr><td>Email:</td><td>" & oUser("useremail") & "</td></tr>"
			response.write "<tr><td>Phone:</td><td>" & oUser("userhomephone") & "</td></tr>"
			response.write "<tr><td>Address:</td><td>" & oUser("useraddress") & "</td></tr>"
			response.write "<tr><td>&nbsp;</td><td>" & oUser("usercity") & ", " & oUser("userstate") & " " & oUser("userzip") & "</td></tr>"
			response.write "</table></div>"
		Else
			response.write "<div style=""display:none;"" id=contact_user><font color=red><i>No information available for specified user.</i></font></div>"
		End If
'	end if
End Function


'----------------------------------------------------------------------------------------------------------------------
' ADDCOMMENTTASKCOMMENT(SINTERNALMSG,SEXTERNALMSG)
'----------------------------------------------------------------------------------------------------------------------
Function AddCommentTaskComment(sInternalMsg,sExternalMsg,sStatus,iFormID,iUserID,iOrgID)
		sSQL = "INSERT egov_payment_responses (action_status,action_internalcomment,action_externalcomment,action_userid,action_orgid,action_autoid) VALUES ('" & sStatus & "','" & sInternalMsg & "','" & sExternalMsg & "','" & iUserID & "','" & iOrgID & "','" &iFormID & "')"
		Set oComment = Server.CreateObject("ADODB.Recordset")
		oComment.Open sSQL, Application("DSN") , 3, 1
		Set oComment = Nothing
End Function



'------------------------------------------------------------------------------------------------------------
' FUNCTION SETORGANIZATIONPARAMETERS()
'------------------------------------------------------------------------------------------------------------
Function SetOrganizationParameters()
	' SET DEFAULT RETURN VALUE
	iReturnValue = 1


	' LOOKUP CURRENT URL IN DATABASE
	sSQL = "SELECT * FROM Organizations WHERE OrgID='" & Session("OrgID") & "'"
	Set oOrgInfo = Server.CreateObject("ADODB.Recordset")
	oOrgInfo.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oOrgInfo.EOF Then
		iOrgID = oOrgInfo("OrgID")
		sOrgName = oOrgInfo("OrgName")
		sHomeWebsiteURL = oOrgInfo("OrgPublicWebsiteURL")
		sEgovWebsiteURL = oOrgInfo("OrgEgovWebsiteURL")
		sTopGraphicLeftURL = oOrgInfo("OrgTopGraphicLeftURL")
		sTopGraphicRighURL = oOrgInfo("OrgTopGraphicRightURL")
		sWelcomeMessage = oOrgInfo("OrgWelcomeMessage")
		sActionDescription = oOrgInfo("OrgActionLineDescription")
		sPaymentDescription = oOrgInfo("OrgPaymentDescription")
		iHeaderSize = oOrgInfo("OrgHeaderSize")
		sTagline = oOrgInfo("OrgTagline")
		iPaymentGatewayID = oOrgInfo("OrgPaymentGateway")
	End If
	Set oOrgInfo = Nothing 

	If NOT ISNULL(iOrgID) Then 
		iReturnValue = iOrgID
	End If

	' RETURN VALUE
	SetOrganizationParameters = iReturnValue
	
End Function
%>
