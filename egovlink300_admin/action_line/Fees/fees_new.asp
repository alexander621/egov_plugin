<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">


<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->


<% 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: FEES_NEW.ASP
' AUTHOR: TERRY FOSTER
' CREATED: 04/19/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
'
' MODIFICATION HISTORY
' 1.0	04/19/07	TERRY FOSTSER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------


' INITIALIZE AND DECLARE VARIABLES
Dim sError
sLevel = "../../" ' OVERRIDE OF VALUE FROM COMMON.ASP

' SET TIMEZONE INFORMATION INTO SESSION
Session("iUserOffset") = request.cookies("tz")
%>

<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" THEN
	'Insert Fee
	Amount = request.form("Amount") * -1
	Description = request.form("Description")
	EnteredBy = session("userID")
	ActionID = request.querystring("irequestid")

	sSQL = "INSERT INTO egov_action_fees (action_autoid,FeeDescription,FeeAmount,FeeCalculatedByID) VALUES(" & ActionID & ",'" & Description & "'," & Amount & "," & EnteredBy & ")"
	Set oNewFee = Server.CreateObject("ADODB.Connection")
	oNewFee.Open Application("DSN")
	oNewFee.Execute(sSQL)
	oNewFee.Close
	Set oNewFee = Nothing


	'EMAIL FEE Notice to Action Line Submitter's email
	sSQL = "Select useremail,userid,submit_date,ASSIGNED_EMAIL From egov_action_request_View WHERE useremail <> '' AND NOT Useremail IS NULL AND action_autoid=" & ActionID
	'response.write sSQL
	'response.end
	Set oUser2 = Server.CreateObject("ADODB.Recordset")
	oUser2.Open sSQL, Application("DSN"), 3, 1
	If NOT oUser2.EOF Then
		'response.write oUser2("Userid")
		'response.end
		sUserEmail = trim(oUser2("useremail"))
		'response.write sUserEmail
		sFromEmail = oUser2("ASSIGNED_EMAIL")
		'response.write sFromEmail
		'response.end

		SubmitDate = oUser2("Submit_Date")

		lngTrackingNumber = ActionID  & replace(FormatDateTime(cdate(SubmitDate),4),":","")
		sMsgNew = sMsgNew & "<p>This automated message was sent by the EGOVLINK web site. Do not reply to this message.  Please follow the instructions below "
		sMsgNew = sMsgNew & "<p>A new fee has been added to the Action Line with the tracking number of: " & lngTrackingNumber
		sMsgNew = sMsgNew & "<br>Description: " & Description
		sMsgNew = sMsgNew & "<br>Fee: $" & request.form("Amount")
		sMsgNew = sMsgNew & "<p>To pay this fee follow this link: <a href=http://dev4.egovlink.com/eclink/payment.asp?paymenttype=40&TrackingNumber=" & lngTrackingNumber & "&PaymentAmount=" & request.form("Amount") & ">Pay Fee</a>"

		sendEmail "", sUserEmail, "", UCase(sOrgName) & " E-GOV MSG - ACTION LINE NEW FEE", sMsgNew, "", "Y"

	end if
	oUser2.Close
	Set oUser2 = Nothing
	'response.end

	response.redirect "../action_respond.asp?control=" & ActionID & "&retrunfromfrees#Fees"

End If
%>


<html>

<head>

  <title><%=langBSHome%></title>

  <link rel="stylesheet" type="text/css" href="../../global.css" />
  <link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />

  <script language="Javascript" src="../../scripts/modules.js"></script>
  <script language="Javascript" src="../../scripts/easyform.js"></script>  

  <script language="Javascript" > 
  <!--

	//Set timezone in cookie to retrieve later
	var d=new Date();
	if (d.getTimezoneOffset)
	{
		var iMinutes = d.getTimezoneOffset();
		document.cookie = "tz=" + iMinutes;
	}

  //-->
  </script>

  <STYLE>
		div.correctionsbox {border: solid 1px #336699;padding: 4px 0px 0px 4px ;}
		div.correctionsboxnotfound  {background-color:#e0e0e0;border: solid 1px #000000;padding: 10px;color:red;font-weight:bold;}
		td.correctionslabel {font-weight:bold;}
		th.corrections {background-color:#93bee1;font-size:12px;padding:5px;color:#000000; }
		th.correctionsinternal{background-color:#e0e0e0;font-size:12px;padding:5px;color:#000000; }
		input.correctionstextbox {border: solid 1px #336699;width:400px;}
		textarea.correctionstextarea {border: solid 1px #336699;width:600px;height:100px;}
		.savemsg {font-size:12px;padding:5px;color:#0000ff;font-weight:bold; }
  </STYLE>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >


<% ShowHeader sLevel %>


<!--#Include file="../../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		

		<h3>New Fee</h3>
		<img align=absmiddle src="../../../admin/images/arrow_2back.gif"> <a href="../action_respond.asp?control=<%=request("irequestid")%>">Return to Request</a> 
		<p>
			This is where the fee calculator would go
		<p>
		<form action=# method=post name=addfee>
			Fee Description: <input type=text name=Description value=""><br>
			Fee Amount: $<input type=text name=Amount value=""><br>
			<input type=hidden name=ef:Description-text/req value="Fee Description">
			<input type=hidden name=ef:Amount-text/currency/req value="Fee Amount">
			<input type=button class=button onClick="if(validateForm('addfee')) {document.addfee.submit();};" value="Add Fee">
		</form>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../../admin_footer.asp"-->  

</body>
</html>



<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB SUBSENDEMAIL(SEMAILBODY,SFROMADDRESS,SEMAILTOADDRESS)
'--------------------------------------------------------------------------------------------------
Sub subSendEmail(sEmailBody,sFromAddress,sEmailToAddress)

	sendEmail "", sEmailToAddress, "", sOrgName & " E-GOV MSG - ACTION LINE REQUEST ASSIGNMENT", "", sEmailBody, "Y"

End Sub



'------------------------------------------------------------------------------------------------------------
' FUNCTION GETEMPLOYEEEMAIL(IVALUE)
'------------------------------------------------------------------------------------------------------------
Function GetEmployeeEmail(iValue)

	sSQL = "SELECT * FROM Users WHERE userid='" & iValue & "'"

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oUser.EOF Then
		iReturnValue = oUser("Email")
	Else
		iReturnValue = "NOT SPECIFIED"
	End If


	Set oUser = Nothing
	
	GetEmployeeEmail = iReturnValue
	
End Function
%>
