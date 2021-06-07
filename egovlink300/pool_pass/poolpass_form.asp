<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: poolpass_form.asp
' AUTHOR: Steve Loar
' CREATED: 01/27/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  01/27/06 Steve Loar - Code added to template
' 1.1  03/05/09 David Boyer - Added "Alternate Layout"
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim oMembership

 set oMembership = New classMembership 

 response.write "<html>" & vbcrlf
 response.write "<head>" & vbcrlf
 %>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
  <%

 if iorgid = 7 then
    lcl_title = sOrgName
 else
    lcl_title = "E-Gov Services " & sOrgName & " Purchase Membership"
 end if

 response.write "<title>" & lcl_title & "</title>" & vbcrlf

	session("RedirectPage") = "pool_pass/poolpass_form.asp" 
	session("RedirectLang") = "Return to Membership Purchase"
%>

<link rel="stylesheet" type="text/css" href="../css/styles.css" />
<link rel="stylesheet" type="text/css" href="../global.css" />
<link rel="stylesheet" type="text/css" href="./style_pool.css" />
<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

<script language="javascript" src="../scripts/modules.js"></script>
<script language="javascript" src="../scripts/easyform.js"></script>

<script language="javascript">
<!--

	function ContinuePurchase() {
 		location.href='./poolpass_select.asp';
	}

	function GotoLogin() {
	 	location.href='../user_login.asp';
	}

	function GotoRegister() {
 		location.href='../register.asp';
	}


//-->
</script>
</head>

<!--#Include file="../include_top.asp"-->

<%
  RegisteredUserDisplay( "../" )

 'BEGIN: Page Content ---------------------------------------------------------
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "    <div id=""poolformtitle"">Purchase Membership</div>" & vbcrlf

  sSQL = "SELECT membership,membershipdesc FROM egov_memberships WHERE orgid = " & iorgid
  Set oM = Server.CreateObject("ADODB.RecordSet")
  oM.Open sSQL, Application("DSN"), 3, 1
  Do While Not oM.EOF
  	response.write "    <div id=""poolforminputarea"">" & vbcrlf
	
  	oMembership.ShowMembershipIntro(oM("membership"))
	
  	response.write "    </div>" & vbcrlf
	
  	if oMembership.PublicCanPurchase(oM("membership")) then
     	response.write "    <div id=""poolfooter"">" & vbcrlf
     	response.write "      <input type=""button"" class=""reserveformbutton"" style=""width:200px;text-align:center;"" name=""continue"" id=""continueButton"" value=""Continue with " & oM("membershipdesc") & " Pass Purchase"" onclick=""location.href='./poolpass_select.asp?mtype=" & oM("membership") & "';"" />" & vbcrlf
     	response.write "    </div>" & vbcrlf
  	end if
	

  	oM.MoveNext
  loop
  oM.Close
  Set oM = Nothing
response.write "  </div>" & vbcrlf




  response.write "</div>" & vbcrlf
 'END: Page Content -----------------------------------------------------------

 'BEGIN: Spacing Code ---------------------------------------------------------
  response.write "<p><br />&nbsp;<br />&nbsp;</p>" & vbcrlf
 'END: Spacing Code -----------------------------------------------------------
%>
<!--#Include file="../include_bottom.asp"-->  
<%
 set oMembership = nothing 

'------------------------------------------------------------------------------
Function GetPoolPassRates( iOrgid, iUserid, sResidenttype, sUserType)
	Dim sDisabled, sSQL, oRates, bPreSelect, iRow
	sDisabled        = ""
	GetPoolPassRates = ""
	bPreSelect       = False
	iRow             = 0

	if sUserType <> sResidentType then
  		sDisabled = " disabled=""disabled"""
	else
  		bPreSelect = True 
	end if
	
	' Get the Pool Pass Rates for the Orgid and residenttype
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetPoolPassRates"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, iOrgid)
		.Parameters.Append oCmd.CreateParameter("@sResidentType", 129, 1, 1, sResidenttype)
	    Set oRates = .Execute
	End With

	do while not oRates.eof
  		iRow = iRow + 1

  		if bPreSelect = True and iRow = 1 then
    			lcl_checked_rateid = " checked=""checked"""
    else
       lcl_checked_rateid = ""
    end if

  		GetPoolPassRates = GetPoolPassRates & "<tr>" & vbcrlf
    GetPoolPassRates = GetPoolPassRates & "    <td><input type=""radio"" name=""rateid"" value=""" & oRates("rateid") & """" & sDisabled & lcl_checked_rateid & " />" & oRates("description") & "</td>" & vbcrlf
    GetPoolPassRates = GetPoolPassRates & "    <td class=""pickprice"">$" & oRates("amount") & "</td>" & vbcrlf
    GetPoolPassRates = GetPoolPassRates & "</tr>" & vbcrlf

    oRates.movenext
 loop
	oRates.close
	set oRates = nothing
	set oCmd   = nothing
	
End function

'------------------------------------------------------------------------------
' Function CanPurchaseOnline( iUserId )
'------------------------------------------------------------------------------
Function CanPurchaseOnline( iUserId )
	Dim sSQL, oType, sUserType

	If iUserid = "" Then
		CanPurchaseOnline = False
	Else
		' Get the resident type
		sUserType = GetUserResidentType(iUserId)

		' Get the Pool Pass Rates for the Orgid and residenttype
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
		    .CommandText = "CanPurchasePoolPassOnline"
		    .CommandType = 4
			.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, iOrgid)
			.Parameters.Append oCmd.CreateParameter("@sResidentType", 129, 1, 1, sUserType)
			.Parameters.Append oCmd.CreateParameter("@PurchaseAllowed", 11, 2, 1)
		    .Execute
		End With
		
		CanPurchaseOnline = oCmd.Parameters("@PurchaseAllowed").Value

		Set oCmd = Nothing
	End If
	
End Function 

'------------------------------------------------------------------------------
' Function GetUserResidentType( iUserId )
'------------------------------------------------------------------------------
Function GetUserResidentType( iUserId )
	Dim sSQL, oType, sResType
	sResType = ""

	If iUserid = "" Then
		GetUserResidentType = ""
	Else
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
		    .CommandText = "GetUserResidentType"
		    .CommandType = 4
			.Parameters.Append oCmd.CreateParameter("@iUserid", 3, 1, 4, iUserId)
			.Parameters.Append oCmd.CreateParameter("@ResidentType", 129, 2, 1)
		    .Execute
		End With
		
		GetUserResidentType = oCmd.Parameters("@ResidentType").Value

		Set oCmd = Nothing

		If IsNull(GetUserResidentType) Or GetUserResidentType = "" Then
			GetUserResidentType = "N"
		End if
		
	End If 

End Function 

'------------------------------------------------------------------------------
' Function ShowPoolPassIntro( iOrgid )
'------------------------------------------------------------------------------
Sub ShowPoolPassIntro( iOrgid )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetPoolPassIntro"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, iOrgid)
		.Parameters.Append oCmd.CreateParameter("@IntroText", 200, 2, 3000)
	    .Execute
	End With
		
	response.write oCmd.Parameters("@IntroText").Value
	Set oCmd = Nothing

End Sub 
%>
