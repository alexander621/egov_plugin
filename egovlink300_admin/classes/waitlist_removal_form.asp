<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: waitlist_removal_form.asp
' AUTHOR: Steve Loar
' CREATED: 05/02/07
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   05/02/06   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sClassName, sResidentTypeDesc, sHeadName, sName, iMemberCount, iFamilyMemberId, iMembershipId
Dim iOldPaymentId, sResidentType

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "registration" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iclassid = request("classid")
itimeid = request("timeid")
iclasslistid = request("classlistid")
sClassName = GetClassName( iclassid, itimeid )

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script language="Javascript" src="tablesort.js"></script>
	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/layers.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>

	<script language="JavaScript">
	<!--

		function validate()
		{
			if (document.frmStatus.notes.value == "")
			{
				alert('Please enter a reason for the removal in the notes.');
				document.frmStatus.notes.focus();
				return;
			}
			else
			{
				document.frmStatus.submit();
			}
		}

	//-->
	</script>

</head>

<body>
 
<%'DrawTabs tabRecreation,1%>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

		<!--BEGIN: PAGE TITLE-->
		<p>
			<font size="+1"><strong>Recreation: Remove Person from the Wait List</strong></font><br /><br />
			<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
		</p>
		<!--END: PAGE TITLE-->


		<!--BEGIN: Removal FORM-->
		<%

		' GET INFORMATION FOR THIS REGISTRANT
		sSQL = "Select * FROM egov_class_roster where classid = " & iclassid & " and classtimeid = " & itimeid & " and classlistid = " & iclasslistid &" ORDER BY status, userlname"
		Set oRegistrant = Server.CreateObject("ADODB.Recordset")
		oRegistrant.Open sSQL, Application("DSN"), 3, 1

		If NOT oRegistrant.EOF Then
			sHeadName = oRegistrant("userfname") & " " & oRegistrant("userlname")
			sName = oRegistrant("firstname") & " " & oRegistrant("lastname")
			iUserId = oRegistrant("userid")
			sResidentTypeDesc = oRegistrant("description")
			sResidentType = oRegistrant("residenttype")
			iFamilyMemberId = oRegistrant("familymemberid")
			If IsNull(iFamilyMemberId) Then
				iFamilyMemberId = 0
			End If 
			'iMemberCount = GetMemberCount( iFamilyMemberId, iUserId, iMembershipId )
			iOldPaymentId = oRegistrant("paymentid")
			iQuantity = oRegistrant("quantity")
		Else
			' Something is wrong
			sHeadName = ""
			sName = ""
			iUserId = 0
			sResidentTypeDesc = ""
			iFamilyMemberId = 0
			'iMemberCount = 0
			iOldPaymentId = 0
			sResidentType = "N"
			iQuantity = 1
		End If

		oRegistrant.Close 
		Set oRegistrant = Nothing
		%>

		<form name="frmStatus" action="waitlist_removal.asp" method="post">

		<p><strong>Class: </strong><%=sClassName%> </p>

		<p><strong>Name: </strong><%=sName%> ( <%=sResidentTypeDesc%> )</p>

		<p><strong>Head of Household: </strong><%=sHeadName%></p>

		<p><strong>Quantity: </strong><%=iQuantity%>
			<input type="hidden" name="quantity" value="<%=iQuantity%>" />
		</p>

		<p>
			<input type="hidden" name="iclasslistid" value="<%=iclasslistid%>" />
			<input type="hidden" name="classid" value="<%=iclassid%>" />
			<input type="hidden" name="timeid" value="<%=itimeid%>" />
			<input type="hidden" name="oldpaymentid" value="<%=iOldPaymentId%>" />
			<input type="hidden" name="paymentamount" value="0.00" />
			<input type="hidden" name="iUserId" value="<%=iUserId%>" />
			<input type="hidden" name="paymenttotal" value="0.00" />
			<input type="hidden" name="PaymentLocationId" value="0" />
			<input type="hidden" name="familymemberid" value="<%=iFamilyMemberId%>" />
			<strong>Notes:</strong><br />
			<textarea name="notes" class="removalnotes"></textarea>
			
		</p>
		<p>
			<input class="button" type="button" onClick="validate();" name="complete" value="Remove" />

		</p>
		</form>
		<!--END: DROP FORM-->

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
' Function CheckResTypeExists(iClassid, iorgid, sResidentType)
'--------------------------------------------------------------------------------------------------
Function CheckResTypeExists(iClassid, iorgid, sResidentType)
	Dim sSql, oCheck

	CheckResTypeExists = False 
	sSql = "Select count(T.pricetype) as hits "
	sSql = sSql & " from egov_price_types T, egov_class_pricetype_price P "
	sSql = sSql & " where T.pricetypeid = P.pricetypeid "
	sSql = sSql & " and orgid = " & iorgid & " and P.classid = " & iClassid & " and T.pricetype = '" & sResidentType & "'"

	Set oCheck = Server.CreateObject("ADODB.Recordset")
	oCheck.Open sSQL, Application("DSN"), 3, 1

	If clng(oCheck("hits")) > 0 Then 
		CheckResTypeExists = True 
	End If 

	oCheck.close
	Set oCheck = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassName( iClassId, iTimeId )
'--------------------------------------------------------------------------------------------------
Function GetClassName( iClassId, iTimeId )
	Dim sSql, oItem

	sSql = "Select classname, activityno from egov_class C, egov_class_time T Where C.classid = T.classid and C.classid = " & iClassId & " and T.timeid = " & iTimeId

	Set oItem = Server.CreateObject("ADODB.Recordset")
	oItem.Open sSQL, Application("DSN"), 3, 1

	If Not oItem.EOF Then 
		GetClassName = oItem("classname") & " &nbsp; ( " & oItem("activityno") & " )"
	Else
		GetClassName = ""
	End If 

	oItem.Close
	Set oItem = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowPaymentLocations()
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentLocations()
	Dim sSql, oLocations

	sSql = "Select paymentlocationid, paymentlocationname from egov_paymentlocations Where isadminmethod = 1 order by paymentlocationid"

	Set oLocations = Server.CreateObject("ADODB.Recordset")
	oLocations.Open sSQL, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""PaymentLocationId"">"
	Do While Not oLocations.EOF
		response.write vbcrlf & "<option value=""" & oLocations("paymentlocationid") & """>" & oLocations("paymentlocationname") & "</option>"
		oLocations.movenext 
	Loop
	response.write vbcrlf & "</select>"

	oLocations.close
	Set oLocations = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function CheckResTypeExists(iClassid, iorgid, sResidentType)
'--------------------------------------------------------------------------------------------------
Function CheckResTypeExists(iClassid, iorgid, sResidentType)
	Dim sSql, oCheck

	CheckResTypeExists = False 
	sSql = "Select count(T.pricetype) as hits "
	sSql = sSql & " from egov_price_types T, egov_class_pricetype_price P "
	sSql = sSql & " where T.pricetypeid = P.pricetypeid "
	sSql = sSql & " and orgid = " & iorgid & " and P.classid = " & iClassid & " and T.pricetype = '" & sResidentType & "'"

	Set oCheck = Server.CreateObject("ADODB.Recordset")
	oCheck.Open sSQL, Application("DSN"), 3, 1

	If clng(oCheck("hits")) > 0 Then 
		CheckResTypeExists = True 
	End If 

	oCheck.close
	Set oCheck = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetMemberCount( ByVal iFamilyMemberId, ByVal iUserid, ByRef iMembershipId )
'--------------------------------------------------------------------------------------------------
Function GetMemberCount( ByVal iFamilyMemberId, ByVal iUserid, ByRef iMembershipId )
	Dim sSql, oMembers, sMembershipstatus

	sMembershipstatus = "O"
	GetMemberCount = 0
	iMembershipId = ""

	sSql = "Select poolpassid From egov_poolpassmembers where familymemberid = " & iFamilyMemberId

	Set oMembers = Server.CreateObject("ADODB.Recordset")
	oMembers.Open sSQL, Application("DSN"), 3, 1

	Do while Not oMembers.EOF 
		sMembershipstatus = DetermineMembership( iFamilyMemberId, iUserid, oMembers("poolpassid") )
		If sMembershipstatus = "M" Then 
			GetMemberCount = 1
			iMembershipId = oMembers("poolpassid")
			Exit Do 
		End If 
		oMembers.MoveNext
	Loop 

	oMembers.Close
	Set oMembers = Nothing 

End Function 



%>

