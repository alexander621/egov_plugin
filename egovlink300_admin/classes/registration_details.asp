<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_signup.asp
' AUTHOR: Steve Loar
' CREATED: 03/16/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This handles the signup process for classes and events.
'
' MODIFICATION HISTORY
' 1.0   03/16/06	Steve Loar - Initial version
' 1.1	05/10/06	Steve Loar - Added citizen search
' 1.2	10/17/06	Steve Loar - Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "registration" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

'classlistid
Dim sStudentName
sStudentName = ""

sStudentName = GetStudentName( request("classlistid") )

%>

<html>
<head>

	<title>E-Gov Registration Details</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css">
	<link rel="stylesheet" type="text/css" href="classes.css">

</head>
<body>

<%'DrawTabs tabRecreation,1%>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->

<div id="content">
	<div id="centercontent">
	<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
	<!--<a href="view_roster_list.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Class/Event List</a><br /><br />-->

	<fieldset><legend><strong> Registration Details </strong></legend>
		<p><h3><% If sStudentName <> "" Then
					response.write sStudentName & " &mdash; "
				  End If %>
				<%=GetClassName( request("classlistid") )%></h3></p>
		<% ShowRegistrationInfo request("classlistid")  %>
	</fieldset>
	
	</div>
</div>

<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>

<!--#Include file="class_global_functions.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetStudentName( iClassListid )
'--------------------------------------------------------------------------------------------------
Function GetStudentName( iClassListid )
	Dim sSql, oName

	sSql = "select firstname + ' ' + lastname as name from egov_familymembers F, egov_class_list L "
	sSql = sSql & " where F.familymemberid = L.familymemberid and L.classlistid = " & iClassListid

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF Then 
		GetStudentName = oName("name") 
	End If 

	oName.close
	Set oName = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassName( iClassListid )
'--------------------------------------------------------------------------------------------------
Function GetClassName( iClassListid )
	Dim sSql, oName

	sSql = "select classname from egov_class C, egov_class_list L "
	sSql = sSql & " where C.classid = L.classid and L.classlistid = " & iClassListid

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF Then 
		GetClassName = oName("classname") 
	End If 

	oName.close
	Set oName = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowRegistrationInfo( iClassListid )
'--------------------------------------------------------------------------------------------------
Sub ShowRegistrationInfo( iClassListid )
	Dim sSql, oInfo

	sSql = "select isnull(amount,0) as amount, isnull(refundamount,0) as refundamount, isnull(quantity,0) as quantity, "
	sSql = sSql & " signupdate, paymentlocationname, paymenttypename, userfname + ' ' + userlname as purchaser, status "
	sSql = sSql & " from egov_class_list L, egov_class_payment P, egov_paymentlocations PL, egov_paymenttypes PT, egov_users U "
	sSql = sSql & " where L.paymentid = P.paymentid and "
	sSql = sSql & " P.paymentlocationid = PL.paymentlocationid and "
	sSql = sSql & " P.paymenttypeid = PT.paymenttypeid and L.userid = U.userid and"
	sSql = sSql & " classlistid = " & iClassListid

	Set oInfo = Server.CreateObject("ADODB.Recordset")
	oInfo.Open sSQL, Application("DSN"), 0, 1

	If Not oInfo.EOF Then 
		response.write "<p><strong>Status:</strong> " & oInfo("status") & "</p>"
		response.write "<p><strong>Purchased By:</strong> " & oInfo("purchaser") & "</p>"
		response.write "<p><strong>Registration Date:</strong> " & oInfo("signupdate") & "</p>"
		response.write "<p><strong>Registration Location:</strong> " & oInfo("paymentlocationname") & "</p>"
		response.write "<p><strong>Payment Type:</strong> " & oInfo("paymenttypename") & "</p>"
		response.write "<p><strong>Quantity:</strong> " & oInfo("quantity") & "</p>"
		response.write "<p><strong>Payment Amount:</strong> " & FormatCurrency(oInfo("amount"),2) & "</p>"
		response.write "<p><strong>Refund Amount:</strong> " & FormatCurrency(oInfo("refundamount"),2) & "</p>"
	End If 

	oInfo.close
	Set oInfo = Nothing
End Sub 


%>
