<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="../pool_pass/poolpass_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: poolpass_receipt.asp
' AUTHOR: Steve Loar
' CREATED: 02/7/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0 02/07/06 Steve Loar - Code added
' 2.0 08/01/06 Steve Loar - Changed to be part of the purchases report
' 2.1 12/08/09 David Boyer - Now check for the membership name to be used instead of hard-coded "pool"
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim iPoolPassId, iUserId, oOrganization

 Set oOrganization = New classOrganization

 iPoolPassId        = request("iPoolPassId")
 iUserId            = GetPurchaserId( iPoolPassId )

 if iUserId <> request.cookies("userid") then response.redirect "purchases_list.asp"

 lcl_membershipdesc = getMembershipDesc(iPoolPassID)
%>
<html>
<head>
	<title><%=oOrganization.GetOrgName()%> E-Gov <%=lcl_membershipdesc%> Membership Purchase Details</title>

	<link type="text/css" rel="stylesheet" href="../global.css" />
	<link type="text/css" rel="stylesheet" href="../css/style_<%=iorgid%>.css" />
	<link type="text/css" media="print" rel="stylesheet" href="receiptprint.css" />
</head>

<!--#Include file="../include_top.asp"-->

<%	RegisteredUserDisplay( "../" ) %>

<div id="content">
	 <div id="centercontent">

<div id="receiptlinks">
 		<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Back</a><span id="printbutton"><input type="button" onclick="javascript:window.print();" value="Print" /></span>
</div>

<h3><%=GetCityName()%>&nbsp;<%=lcl_membershipdesc%> Membership</h3> 

<% ShowUserInfo iUserId  %>

<div class="purchasereportshadow">
  <table border="0" cellpadding="5" cellspacing="0" class="purchasereport">
		  <tr><th colspan="2" align="left">Transaction Details</th></tr>
  		<% ShowPassInfo iPoolPassId %>
 	</table>
</div>
<div class="purchasereportshadow">
		<table border="0" cellpadding="5" cellspacing="0" class="purchasereport">
  		<tr><th align="left" colspan="2"><%=lcl_membershipdesc%> Membership Details</th></tr>
		  <% ShowPassMembers iPoolPassId %>
		</table>
</div>

 	</div>
</div>

<%	Set oOrganization = Nothing %>

<!--#Include file="../include_bottom.asp"-->  

<%
'------------------------------------------------------------------------------
Sub ShowPassInfo( iPoolPassId )
	Dim sSQL, oName 

	' Ths is the transaction details 
	sSQL = "SELECT U.userfname, U.userlname, U.useraddress, U.useraddress2, U.usercity, U.userstate, U.userzip, "
	sSQL = sSQL & "P.paymentamount, P.paymenttype, P.paymentdate, P.paymentlocation, R.description, T.description as residenttype, "
	sSQL = sSQL & "P.paymentauthcode, P.paymentpnref, P.paymentrespmsg, P.paymentresult "
	sSQL = sSQL & "FROM egov_poolpasspurchases P, egov_users U, egov_poolpassrates R, egov_poolpassresidenttypes T "
	sSQL = sSQL & "WHERE P.orgid = " & iorgid
	sSQL = sSQL & " AND P.poolpassid = " & iPoolPassId 
	sSQL = sSQL & " AND P.userid = " & iUserId
	sSQL = sSQL & " AND U.userid = P.userid "
	sSQL = sSQL & " AND P.rateid = R.rateid "
	sSQL = sSQL & " AND R.residenttype = T.resident_type "
	sSQL = sSQL & " AND T.orgid = P.orgid "

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 3, 1

	If Not oName.EOF Then 
		response.write "  <tr>" & vbcrlf
  response.write "      <td width=""20%"">Purchase Date: </td>" & vbcrlf
  response.write "      <td>" & DateValue(oName("paymentdate")) & "</td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
		response.write "  <tr>" & vbcrlf
  response.write "      <td>Payment Method: </td>" & vbcrlf
  response.write "      <td>" & MakeProper(oName("paymenttype")) & "</td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
		response.write "  <tr>" & vbcrlf
  response.write "      <td>Payment Location: </td>" & vbcrlf
  response.write "      <td>" & MakeProper(oName("paymentlocation")) & "</td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
		response.write "  <tr>" & vbcrlf
  response.write "      <td>Amount: </td>" & vbcrlf
  response.write "      <td>" & FormatCurrency(oName("paymentamount"),2) & "</td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
'		If LCase(oName("paymenttype")) = "creditcard" Then
'			ShowPassInfo = ShowPassInfo & "<tr><td align=""right"">Authcode: </td><td>" & oName("paymentauthcode") & "</td></tr>"
'			ShowPassInfo = ShowPassInfo & "<tr><td align=""right"">PnRef: </td><td>" & oName("paymentpnref") & "</td></tr>"
'			ShowPassInfo = ShowPassInfo & "<tr><td align=""right"">RespMsg: </td><td>" & oName("paymentrespmsg") & "</td></tr>"
'			ShowPassInfo = ShowPassInfo & "<tr><td align=""right"">Status: </td><td>" & oName("paymentresult") & "</td></tr>"
'		End If 
'		ShowPassInfo = ShowPassInfo & "<tr><td valign=""top"" align=""right"">Purchaser: </td><td>" & oName("userfname") & " " & oName("userlname") & "<br />"
'		ShowPassInfo = ShowPassInfo & oName("useraddress") & "<br />"
'		If oName("useraddress2") <> "" Or IsNull(oName("useraddress2")) = False then
'			ShowPassInfo = ShowPassInfo & oName("useraddress2") & "<br />"
'		End if
'		ShowPassInfo = ShowPassInfo & oName("usercity") & ", " & oName("userstate") & " " & oName("userzip") 
'		ShowPassInfo = ShowPassInfo & "</td></tr>"
	End If
		
	oName.close
	Set oName = Nothing
End Sub  

'------------------------------------------------------------------------------
Sub ShowPassMembers(iPoolPassId)
	Dim sSQL, oMembers, oName 

	' Get the pass info
	sSQL = "SELECT U.userfname, U.userlname, U.useraddress, U.useraddress2, U.usercity, U.userstate, U.userzip, "
	sSQL = sSQL & "P.paymentamount, P.paymenttype, P.paymentdate, P.paymentlocation, R.description, T.description as residenttype, "
	sSQL = sSQL & "P.paymentauthcode, P.paymentpnref, P.paymentrespmsg, P.paymentresult "
	sSQL = sSQL & "FROM egov_poolpasspurchases P, egov_users U, egov_poolpassrates R, egov_poolpassresidenttypes T "
	sSQL = sSQL & "WHERE P.orgid = " & iOrgId
 sSQL = sSQL & " AND P.poolpassid = " & iPoolPassId 
 sSQL = sSQL & " AND U.userid = P.userid "
 sSQL = sSQL & " AND P.rateid = R.rateid "
 sSQL = sSQL & " AND R.residenttype = T.resident_type "
 sSQL = sSQL & " AND T.orgid = P.orgid "

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 3, 1

	if not oName.eof then
		response.write  "  <tr>" & vbcrlf
  response.write "       <td width=""20%"">Member No: </td>" & vbcrlf
  response.write "       <td>" & iPoolPassId & "</td>" & vbcrlf
  response.write "   </tr>" & vbcrlf
		response.write  "  <tr>" & vbcrlf
  response.write "       <td width=""20%"">Season: </td>" & vbcrlf
  response.write "       <td>" & Year(oName("paymentdate")) & "</td>" & vbcrlf
  response.write "   </tr>" & vbcrlf
		response.write  "  <tr>" & vbcrlf
  response.write "       <td>Membership Type: </td>" & vbcrlf
  response.write "       <td>" & oName("residenttype") & " &mdash; " & oName("description") & "</td>" & vbcrlf
  response.write "   </tr>" & vbcrlf
	End If 
	oName.close
	Set oName = Nothing

	' Get the pass members
	sSQL = "SELECT F.firstname, F.lastname, F.relationship "
 sSQL = sSQL & " FROM egov_familymembers F, egov_poolpassmembers P "
 sSQL = sSQL & " WHERE P.poolpassid = " & iPoolPassId
 sSQL = sSQL & " AND F.familymemberid = P.familymemberid "
 sSQL = sSQL & " ORDER BY birthdate, lastname, firstname"

	Set oMembers = Server.CreateObject("ADODB.Recordset")
	oMembers.Open sSQL, Application("DSN"), 3, 1

	response.write "  <tr><td colspan=""2""><strong>This Membership includes</strong></td></tr>" & vbcrlf

	do while not oMembers.eof
  		response.write "  <tr>" & vbcrlf
    response.write "      <td width=""20%"">" & oMembers("firstname") & " " & oMembers("lastname") & "</td>" & vbcrlf
    response.write "      <td>" & TranslateMember(oMembers("relationship")) &"</td>" & vbcrlf
    response.write "  </tr>" & vbcrlf

		  oMembers.movenext
	loop 
		
	oMembers.close
	Set oMembers = Nothing
End Sub 

'------------------------------------------------------------------------------
Function GetCityName()
	Dim sSQL
	GetCityName = ""

	sSQL = "SELECT orgname FROM organizations WHERE orgid = " & iOrgId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN") , 3, 1

	If Not oName.eof Then 
		  GetCityName = oName("orgname")
	End If
		
	oName.close
	Set oName = Nothing
End Function 

'------------------------------------------------------------------------------
Function MakeProper( sString )
	If sString = "" Then
		MakeProper = ""
	Else
		MakeProper = UCase(Left(sString,1)) & LCase(Mid(sString,2))
	End If 
End Function 

'------------------------------------------------------------------------------
Function TranslateMember( sRelationship )
	If UCase(sRelationship) = "YOURSELF" Then
		TranslateMember = "Purchaser"
	Else 
		TranslateMember = sRelationship
	End If 
	
End Function 

'------------------------------------------------------------------------------
Sub ShowUserInfo( iUserId )
	Dim oCmd, sResidentDesc, sUserType

	sUserType = GetUserResidentType(iUserid)
	' If they are not one of these (R, N), we have to figure which they are
	If sUserType <> "R" And sUserType <> "N" Then
		' This leaves E and B - See if they are a resident, also
		sUserType = GetResidentTypeByAddress(iUserid, iOrgId)
	End If 

	sResidentDesc = GetResidentTypeDesc(sUserType)

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserInfoList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iUserId", 3, 1, 4, iUserId)
	    Set oUser = .Execute
	End With

	response.write "<div class=""purchasereportshadow"">" & vbcrlf
	response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">" & vbcrlf
	response.write "  <tr><th colspan=""2"" align=""left"">Purchaser Contact Information</th></tr>" & vbcrlf
	response.write "  <tr>" & vbcrlf
 response.write "      <td width=""20%"" valign=""top"">Name:</td>" & vbcrlf
 response.write "      <td>" & oUser("userfname") & " " & oUser("userlname") & "<br /><strong>" & sResidentDesc & "</strong></td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
	response.write "  <tr>" & vbcrlf
 response.write "      <td>Email:</td>" & vbcrlf
 response.write "      <td>" & oUser("useremail") & "</td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
	response.write "  <tr>" & vbcrlf
 response.write "      <td>Phone:</td>" & vbcrlf
 response.write "      <td>" & FormatPhone(oUser("userhomephone")) & "</td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
	response.write "  <tr>" & vbcrlf
 response.write "      <td valign=""top"">Address:</td>" & vbcrlf
 response.write "      <td>" & oUser("useraddress") & "<br />"  & vbcrlf

	if oUser("useraddress2") = "" then
  		response.write oUser("useraddress2") & "<br />" & vbcrlf
 end if

	response.write oUser("usercity") & ", " & oUser("userstate") & " " & oUser("userzip") & "</td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
	response.write "</table>" & vbcrlf
 response.write "</div>" & vbcrlf

	oUser.close
	set oUser = nothing
	set oCmd  = nothing

end sub

'------------------------------------------------------------------------------
Function GetPurchaserId( iPoolPassId )
	Dim sSql, oName

	GetPurchaserId = 0

	sSQL = "Select P.userid FROM egov_poolpasspurchases P "
	sSQL = sSQL & " WHERE P.orgid = " & iOrgId & " and P.poolpassid = " & iPoolPassId 

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 3, 1

	If Not oName.eof Then
		GetPurchaserId = oName("userid")
	End If 

	oName.close
	Set oName = Nothing 

end function

'------------------------------------------------------------------------------
function getMembershipDesc(iPoolPassID)
  lcl_return = "Pool"

  sSQL = "SELECT membershipdesc "
  sSQL = sSQL & " FROM egov_memberships m "
  sSQL = sSQL & " WHERE membershipid = (select membershipid "
  sSQL = sSQL &                       " from egov_poolpasspurchases "
  sSQL = sSQL &                       " where poolpassid = " & iPoolPassID & ") "

  set oMemberDesc = Server.CreateObject("ADODB.Recordset")
  oMemberDesc.Open sSQL, Application("DSN"), 3, 1

  if not oMemberDesc.eof then
     lcl_return = oMemberDesc("membershipdesc")
  end if

  oMemberDesc.close
  set oMemberDesc = nothing

  getMembershipDesc = lcl_return


end function
%>
