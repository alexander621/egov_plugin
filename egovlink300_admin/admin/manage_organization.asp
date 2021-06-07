<%
' IF POST PERFORM DATABASE OPERATIONS
If request.servervariables("REQUEST_METHOD") = "POST" Then
	Select Case request("post_type")

		Case "NEW"
			' ADD NEW ORGANIZATION
			Call subNewOrganization()

		Case "UPDATE"
			' EDIT EXIST ORGANIZATION
			Call subEditOrganization(request("iorgid"))

		Case Else
			' DEFAULT ACTION
			' NO DATABASE OPERATIONS TO PERFORM
	End Select 

Else
	' IF NOT ORG ID MUST BE NEW ORG OTHERWISE SET ORG TO EDIT MODE
	If request("IORGID") <> "" Then
		sPostType = "UPDATE"
	Else
		sPostType = "NEW"
	End If

End If


' LOAD VALUES IF WE KNOW ORGID
Dim OrgName,OrgPublicWebsiteURL,OrgEgovWebsiteURL,OrgTopGraphicRightURL,OrgTopGraphicLeftURL,OrgWelcomeMessage,OrgTagline,OrgActionlinedescription,Orgheadersize,orgpaymentdescription,orgPaymentGateway,orgactionlinedisplayoption
If request("iorgid") <> "" Then
	Call subGetOrganizationValues(request("iorgid"))
End If
%>


<HTML>
<HEAD>
<TITLE> E-GovLink - New Organization </TITLE>
<link href="../global.css" rel="stylesheet" type="text/css">
</HEAD>

<BODY>
<!--BEGIN: ORGANIZATION EDIT FORM-->
<form name="frmManageOrg" action="manage_organization.asp" METHOD="POST">
<input type=hidden value="<%=sPostType%>" name="post_type">
<input type=hidden value="<%=request("iorgid")%>" name="iorgid">

<div style="margin-top:20px; margin-left:20px;" >

<font class=label>Organization Parameters - <small>(<a href="list_organization.asp">Return To List</a>)</small></font>

<div class="orgadminbox">
	<table cellspacing="0" cellpadding=0 border="0" >
		<tr><td valign="top"><b>Organization Name:</b><br><input type="text" name="org_name" class="new_org_form" value="<%=orgname%>" ></td></tr>
		<tr><td valign="top"><b>Public Website URL:</b><br><input type="text" name="org_website_url" class="new_org_form" value="<%=OrgPublicWebsiteURL%>" ></td></tr>
		<tr><td valign="top"><b>E-GovLink URL:</b><br><input type="text" name="org_egov_url" class="new_org_form" value="<%=OrgEgovWebsiteURL%>"></td></tr>
		<tr><td valign="top"><b>E-GovLink Top Right Logo Image URL:</b><br><input type="text" name="org_right_logo_url" class="new_org_form" value="<%=OrgTopGraphicRightURL%>"></td></tr>
		<tr><td valign="top"><b>E-GovLink Top Left Logo Image URL:</b><br><input type="text" name="org_left_logo_url" class="new_org_form" value="<%=OrgTopGraphicLeftURL%>"></td></tr>
		<tr><td valign="top"><b>E-GovLink Logo Height in Pixels (Left and Right must match):</b><br><input type="text" name="org_logo_height" class="new_org_form" value="<%=Orgheadersize%>"></td></tr>
		<tr><td valign="top"><b>E-GovLink Welcome Message:</b><br><textarea class="new_org_form" name="org_welcome_msg"  ><%=OrgWelcomeMessage%></textarea></td></tr>
		<tr><td valign="top"><b>E-GovLink Tagline Message:</b><br><textarea class="new_org_form" name="org_tagline_msg" ><%=OrgTagline%></textarea></td></tr>
		<tr><td valign="top"><b>E-GovLink Action Line Intro Text:</b><br><textarea class="new_org_form" name="org_action_text"  ><%=OrgActionlinedescription%></textarea></td></tr>
		<tr><td valign="top"><b>E-GovLink Action Display Option:</b><br><select name="org_actionlineoption"><% Call subListActionDisplayOptions(orgactionlinedisplayoption) %></select><br>&nbsp;</td></tr>
		<tr><td valign="top"><b>E-GovLink Payment Intro Text:</b><br><textarea class="new_org_form" name="org_payment_text"  ><%=orgpaymentdescription%></textarea></td></tr>
		<tr><td valign="top"><b>E-GovLink Payment Gateway:</b><br><select name="org_payment_gateway"><% Call subListGateways(orgPaymentGateway) %></select> <br> <a href="javascript:alert('Feature not yet available. Payment gateway information must be configured manually via SQL Enterprise Manager.');"><B> - CONFIGURE GATEWAY - </B></a></td></tr>
		<tr><td valign="top" align=right><input type=submit value="SAVE" class=submitbtn ></td></tr>
	</table>
</div>
</div>
</form>
<!--END: ORGANIZATION EDIT FORM-->


<!--#include file="bottom_include.asp"-->


</BODY>
</HTML>


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' SUB SUBNEWORGANIZATION()
'------------------------------------------------------------------------------------------------------------
Sub subNewOrganization()
	' INSERT NEW ORGANIZATION
	sSQL = "INSERT INTO Organizations (OrgName,OrgPublicWebsiteURL,OrgEgovWebsiteURL,OrgTopGraphicRightURL,OrgTopGraphicLeftURL,OrgWelcomeMessage,OrgTagline,OrgActionlinedescription,Orgheadersize,orgpaymentdescription,orgPaymentGateway,orgactionlinedisplayoption) VALUES ('" & DBsafe(request("org_name")) & "','" & DBsafe(request("org_website_url")) & "','" & DBsafe(request("org_egov_url")) & "','" & DBsafe(request("org_right_logo_url")) & "','" & DBsafe(request("org_left_logo_url"))  & "','" & DBsafe(request("org_welcome_msg"))  & "','" & DBsafe(request("org_tagline_msg")) & "','" & DBsafe(request("org_logo_height"))  & "','" &  DBsafe(request("org_action_text")) & "','" & DBsafe(request("org_payment_text")) & "','" & DBsafe(request("org_payment_gateway")) & "','" & DBsafe(request("org_actionlineoption"))& "')"
	Set oNewOrg = Server.CreateObject("ADODB.Recordset")
	oNewOrg.Open sSQL, Application("DSN") , 3, 1
	Set oNewOrg = Nothing

	' GET NEW ID
	sSQL = "SELECT MAX(OrgID) as orgid From Organizations"
	Set oNewOrg = Server.CreateObject("ADODB.Recordset")
	oNewOrg.Open sSQL, Application("DSN") , 3, 1
	If NOT oNewOrg.EOF Then
		iOrgID = oNewOrg("orgid")
	End If
	Set oNewOrg = Nothing

	' REDIRECT TO ORGANIZATION EDIT PAGE
	response.redirect("manage_organization.asp?IORGID=" & iOrgID)

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBGETORGANIZATIONVALUES(IORGID)
'------------------------------------------------------------------------------------------------------------
Sub subGetOrganizationValues(iorgid)
	' GET ORGANIZATIONS VALUES
	sSQL = "SELECT * From Organizations WHERE ORGID=" & IORGID
	Set oOrg = Server.CreateObject("ADODB.Recordset")
	oOrg.Open sSQL, Application("DSN") , 3, 1
	If NOT oOrg.EOF Then
		OrgName = oOrg("OrgName")
		OrgPublicWebsiteURL = oOrg("OrgPublicWebsiteURL")
		OrgEgovWebsiteURL = oOrg("OrgEgovWebsiteURL")
		OrgTopGraphicRightURL = oOrg("OrgTopGraphicRightURL")
		OrgTopGraphicLeftURL = oOrg("OrgTopGraphicLeftURL") 
		OrgWelcomeMessage = oOrg("OrgWelcomeMessage")
		OrgTagline = oOrg("OrgTagline")
		OrgActionlinedescription = oOrg("OrgActionlinedescription")
		Orgheadersize = oOrg("Orgheadersize")
		orgpaymentdescription = oOrg("orgpaymentdescription")
		orgPaymentGateway = oOrg("orgPaymentGateway")
		orgactionlinedisplayoption = oOrg("orgactionlinedisplayoption")
	End If
	Set oOrg = Nothing
End Sub


'------------------------------------------------------------------------------------------------------------
'  SUB SUBEDITORGANIZATION(IORGID)
'------------------------------------------------------------------------------------------------------------
 Sub subEditOrganization(iorgid)

 	' INSERT NEW ORGANIZATION
	sSQL = "UPDATE Organizations SET "
	sSQL = sSQL & "OrgName='"& DBsafe(request("org_name")) & "',"
	sSQL = sSQL & "OrgPublicWebsiteURL='" & DBsafe(request("org_website_url")) & "',"
	sSQL = sSQL & "OrgEgovWebsiteURL='" & DBsafe(request("org_egov_url")) & "',"
	sSQL = sSQL & "OrgTopGraphicRightURL='" & DBsafe(request("org_right_logo_url")) & "',"
	sSQL = sSQL & "OrgTopGraphicLeftURL='" & DBsafe(request("org_left_logo_url"))  & "',"
	sSQL = sSQL & "OrgWelcomeMessage='" & DBsafe(request("org_welcome_msg"))  & "',"
	sSQL = sSQL & "OrgTagline='" & DBsafe(request("org_tagline_msg")) & "',"
	sSQL = sSQL & "OrgActionlinedescription='" & DBsafe(request("org_action_text"))  & "',"
	sSQL = sSQL & "Orgheadersize='" &  DBsafe(request("org_logo_height")) & "',"
	sSQL = sSQL & "orgpaymentdescription='" & DBsafe(request("org_payment_text")) & "',"
	sSQL = sSQL & "orgactionlinedisplayoption='" & DBsafe(request("org_actionlineoption")) & "',"
	sSQL = sSQL & "orgPaymentGateway='" & DBsafe(request("org_payment_gateway")) & "' WHERE ORGID=" & iorgid
	Set oEditOrg = Server.CreateObject("ADODB.Recordset")
	oEditOrg.Open sSQL, Application("DSN") , 3, 1
	Set oEditOrg = Nothing

	' REDIRECT TO ORGANIZATION EDIT PAGE
	response.redirect("manage_organization.asp?IORGID=" & iOrgID)

 End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBLISTGATEWAYS(ISELECTED)
'------------------------------------------------------------------------------------------------------------
Sub subListGateways(iSelected)

	sSQL = "SELECT * FROM egov_payment_gateways"
	Set oPaymentGateways = Server.CreateObject("ADODB.Recordset")
	oPaymentGateways.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oPaymentGateways.EOF Then
		Do while NOT oPaymentGateways.EOF 
			If iSelected = oPaymentGateways("paymentgatewayid") Then
				sSelected = "SELECTED"
			Else
				sSelected = ""
			End If 
			
			response.write "<option " & sSelected & " value=""" & oPaymentGateways("paymentgatewayid") & """>" & oPaymentGateways("paymentgatewayname") 
			oPaymentGateways.MoveNext
		Loop

	End If
	Set oPaymentGateways = Nothing 

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBLISTACTIONDISPLAYOPTIONS(ISELECTED)
'------------------------------------------------------------------------------------------------------------
Sub subListActionDisplayOptions(iSelected)

	sSQL = "SELECT * FROM egov_actionlinedisplayoption"
	Set oOption = Server.CreateObject("ADODB.Recordset")
	oOption.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oOption.EOF Then
		Do while NOT oOption.EOF 
			If iSelected = oOption("optionid") Then
				sSelected = "SELECTED"
			Else
				sSelected = ""
			End If 
			
			response.write "<option " & sSelected & " value=""" & oOption("optionid") & """>" & oOption("optionname") 
			oOption.MoveNext
		Loop

	End If
	Set oOption = Nothing 

End Sub


'------------------------------------------------------------------------------------------------------------
' FUNCTION DBSAFE( STRDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

%>
