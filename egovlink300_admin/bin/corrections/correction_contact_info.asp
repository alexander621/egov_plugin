<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">


<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->


<% 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CORRECTION_CONTACT_INFO.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/2/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  DELETE ATTACHMENT
'
' MODIFICATION HISTORY
' 1.0	02/02/07	JOHN STULLENBERGER - INITIAL VERSION
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



<html>

<head>

  <title><%=langBSHome%></title>

  <link rel="stylesheet" type="text/css" href="../../global.css" />
  <link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />

  <script language="Javascript" src="../../scripts/modules.js"></script>

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
		input.correctionstextbox {border: solid 1px #336699;width:400px;}
		.savemsg {font-size:12px;padding:5px;color:#0000ff;font-weight:bold; }
  </STYLE>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >


<% ShowHeader sLevel %>


<!--#Include file="../../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		

		<h3>Edit Request Contact Information</h3>		
		<a href="../action_respond.asp?control=<%=request("irequestid")%>">Return to Request Details</a> | <a href="../action_line_list.asp">View All Requests List</a> 

		
		<%
		' DISPLAY TO USER THAT VALUES WERE SAVED
		If request("r") = "save" Then 
			response.write "<P><span class=""savemsg"">Saved " & Now() & ".</span></P>"
		End If
				
		
		' GET CONTACT INFORMATION
		fnDisplayUserInfo(request("irequestid"))

		%>


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
' FUNCTION FNDISPLAYUSERINFO(IID)
'--------------------------------------------------------------------------------------------------
Function fnDisplayUserInfo(iID)

	' CHECK FOR EMPTY OR MISSING USERID
	If IsNull(iID) or iID="" then
		response.write "<P><div class=correctionsboxnotfound>No information available for this request.</div></P>"
	Else
			' GET INFORMATION FOR SPECIFIED USER
			sSQL = "Select * From  egov_actionline_requests INNER JOIN egov_users ON  egov_actionline_requests.userid = egov_users.userid where egov_actionline_requests.action_autoid= " & iID
			
			' OPEN RECORDSET
			Set oUser = Server.CreateObject("ADODB.Recordset")
			oUser.Open sSQL, Application("DSN"), 3, 1
			
			If NOT oUser.EOF Then
				' FOUND USER - DISPLAY DETAILS
				sUserEmail = trim(oUser("useremail"))
		
				response.write "<form action=""correction_contact_info_cgi.asp"" method=""POST"" >"
				response.write "<div class=shadow>"
				response.write "<table class=tablelist cellspacing=0 cellpadding=2 >"
				response.write "<tr><th class=corrections colspan=2 align=left>&nbsp;Request Contact Information</th></tr>"
				response.write "<tr><td class=correctionslabel align=""right"">First Name:</td><td><input name=userfname type=text class=correctionstextbox value=""" & oUser("userfname") & """></td></tr>"
				response.write "<tr><td class=correctionslabel align=""right"">Last Name:</td><td><input name=userlname type=text class=correctionstextbox value=""" & oUser("userlname") & """</td></tr>"
				response.write "<tr><td class=correctionslabel align=""right"">Business Name:</td><td><input name=userbusinessname type=text class=correctionstextbox value=""" & oUser("userbusinessname") & """</td></tr>"
				response.write "<tr><td class=correctionslabel align=""right"">Email:</td><td><input name=useremail type=text class=correctionstextbox value=""" & oUser("useremail") & """</td></tr>"
				response.write "<tr><td class=correctionslabel align=""right"">Daytime Phone:</td><td><input name=userhomephone type=text class=correctionstextbox value=""" & FormatPhone(oUser("userhomephone")) & """</td></tr>"
				response.write "<tr><td class=correctionslabel align=""right"">Fax:</td><td><input name=userfax type=text class=correctionstextbox value=""" & FormatPhone(oUser("userfax")) & """</td></tr>"
				response.write "<tr><td class=correctionslabel align=""right"">Address:</td><td><input name=useraddress type=text class=correctionstextbox value=""" & oUser("useraddress") & """</td></tr>"
				response.write "<tr><td class=correctionslabel align=""right"">City:</td><td><input name=usercity type=text class=correctionstextbox value=""" & oUser("usercity") & """</td></tr>"
				response.write "<tr><td class=correctionslabel align=""right"">State / Province:</td><td><input name=userstate type=text class=correctionstextbox value=""" & oUser("userstate") & """</td></tr>"
				response.write "<tr><td class=correctionslabel align=""right"">Zip / Postal Code:</td><td><input name=userzip type=text class=correctionstextbox value=""" & oUser("userzip") & """</td></tr>"
				response.write "<tr><td class=correctionslabel align=""right"">Preferred Contact Method:</td><td>"
				Call subListContactMethods(oUser("contactmethodid")) 
				response.write "</td></tr>"

				' BUTTON ROW
				response.write "<tr><td class=correctionslabel align=""left"">&nbsp;&nbsp;<input  type=button value=""Cancel"" onClick=""location.href='../action_respond.asp?control=" & request("irequestid")  & "';""></td><td align=""right""><input type=submit value=""Save"">&nbsp;&nbsp;</td></tr>"


				response.write "</table>"
				response.write "</div>"
				response.write "<input type=hidden value=""" & request("status") & """ name=""status"">"
				response.write "<input type=hidden value=""" & iID & """ name=""irequestid"">"
				response.write "</form>"

			Else
				' NO MATCHING USER FOUND
				response.write "<P><div class=correctionsboxnotfound>No information available for this request.</div></P>"
			End If
		End If

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION FORMATPHONE( NUMBER )
'--------------------------------------------------------------------------------------------------
Function FormatPhone( Number )
  If Len(Number) = 10 Then
    FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
  Else
    FormatPhone = Number
  End If
End Function


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYQUESTIONS(IFORMID)
'------------------------------------------------------------------------------------------------------------
Sub subListContactMethods(iSelected)

	sSQL = "SELECT * FROM egov_contactmethods ORDER BY contactdescription"

	Set oMethods = Server.CreateObject("ADODB.Recordset")
	oMethods.Open sSQL, Application("DSN"), 3, 1

	response.write "<select name=contactmethodid>"
	response.write "<option value=0>Please select a contact method...</option>"

	If NOT oMethods.EOF Then
	
		Do While NOT oMethods.EOF 

			If clng(iSelected) = oMethods("rowid") Then
				sSelected = " SELECTED "
			Else
				sSelected = "  "
			End If
			
			response.write "<option " & sSelected & " value=""" &  oMethods("rowid") & """>" & oMethods("contactdescription") & "</option>"
			oMethods.MoveNext
		
		Loop

	End If

	response.write "</select>"

	oMethods.close
	Set oMethods = Nothing 

End Sub

%>


