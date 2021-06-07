<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!--#Include file="facility_functions.asp"-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: FACILITY_TERMS.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/20/06   JOHN STULLENBERGER - INITIAL VERSION
'       01/20/06   Steve Loar - Added code to do move up and down in display order
' 1.1	10/06/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFacilityId
Dim sFacilityName
Dim iMaxOrder

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit facilities" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

If request("facilityid") = "" Then
	response.redirect( "facility_management.asp" )
Else 
	iFacilityId = request("facilityid")
End If

sFacilityName = GetFacilityName(iFacilityId)

iMaxOrder = GetMaxDisplayOrder(iFacilityId)


' IF POST PROCESS SAVE/ADD REQUEST
If request.servervariables("REQUEST_METHOD") = "POST" Then
	SaveTerm request("sdescription"),ifacilityid,request("itermid"),request("displayorder")
End If

Dim oFacilities
Dim iRowCount
%>

<html>
<head>
	<title>E-Gov Facility Rental Terms &amp; Conditions</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="facility.css" />

<script language="Javascript">
  <!--
	function ConfirmDelete(sDesc, iTermId, iFacilityId) 
	{
		var msg = "Do you wish to delete " + sDesc + "?"
		if (confirm(msg))
		{
			location.href='term_delete.asp?itermId='+ iTermId + '&iFacilityId=' + iFacilityId;
		}
	}

	function SaveTerm(passForm)
	{

		if (passForm.sDescription.value == "") {
			alert("Please enter a description.");
			passForm.sDescription.focus();
			return;
		}

		passForm.submit();
	}

	function ChangeOrder(iFacilityId,iDisplayOrder,iDirection)
	{
		location.href='term_move.asp?iDisplayOrder='+ iDisplayOrder + '&iFacilityId=' + iFacilityId + '&iDirection=' + iDirection;
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
	<p>
	<font size="+1"><strong>Recreation: Facility Term/Conditions Management - <%=sFacilityName%></strong></font><br />
	<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
	</p>

	<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
		<tr>
			<th>Description</th><th>&nbsp;</th><th>Display Order</th>
		</tr>

		<!--  Always start with a blank row for adding -->
		<td><form name="rateterm0" method="post" action="facility_terms.asp?facilityid=<%=iFacilityId%>">
				<input type="hidden" name="itermId" value="0" />
				<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>" />
				<input type="hidden" name="displayorder" value="<%=(iMaxOrder + 1)%>" />
				<textarea name="sDescription"></textarea></td>
				<td class="action">
				<a href="javascript:SaveTerm(document.rateterm0);">Add</a>
			</form>	
		</td>
<%
		sSQL = "Select * FROM egov_recreation_terms where facilityid = " & iFacilityId & " order by displayorder"
		Set oTerms = Server.CreateObject("ADODB.Recordset")
		oTerms.Open sSQL, Application("DSN"), 3, 1
		
		If Not oTerms.EOF Then
			iRowCount = 0
			Do While Not oTerms.EOF
				' print out the lines here
				iRowCount = iRowCount + 1
				If iRowCount Mod 2 = 1 Then
					response.write "<tr class=" & Chr(34) & "alt_row" & Chr(34) & ">"
				Else
					response.write "<tr>"
				End If
				
%>
				<td><form name="rateterm<%=iRowCount%>" method="post" action="facility_terms.asp?facilityid=<%=iFacilityId%>">
				<input type="hidden" name="itermid" value="<%=oTerms("termid")%>" />
				<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>" />
				<input type="hidden" name="displayorder" value="<%=oTerms("displayorder")%>" />
				<textarea name="sDescription"><%=oTerms("termdescription")%></textarea></td>
				<td class="action">
					<a href="javascript:SaveTerm(document.rateterm<%=iRowCount%>);">Save</a>&nbsp;&nbsp;
					<a href="javascript:ConfirmDelete('<%=oTerms("termdescription")%>',<%=oTerms("termid")%>,<%=iFacilityId%>);">Delete</a>
				</td>
				<td>
					<!--<%' =oTerms("displayorder")%><br /> -->
					<% If iRowCount <> 1 Then %>
						<a href="javascript:ChangeOrder(<%=iFacilityId%>, <%=oTerms("displayorder")%>, -1);"><img src="../images/ieup.gif" align="absmiddle" border="0">&nbsp;Move Up</a><br />
					<% End If %>
					<% If oTerms("displayorder") <> iMaxOrder Then %>
						<a href="javascript:ChangeOrder(<%=iFacilityId%>, <%=oTerms("displayorder")%>, 1);"><img src="../images/iedown.gif" align="absmiddle" border="0">&nbsp;Move Down</a>
					<% End If %>
					</form>
				</td>
				</tr>
<%
				oTerms.MoveNext
			Loop 
		End If 
		oTerms.close
		Set oTerms = nothing
%>
	</table>
	</div>
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
' PUBLIC SUB SAVETERM(SDESCRIPTION,IFACILITYID,ITERMID)
'--------------------------------------------------------------------------------------------------
Public Sub SaveTerm(sDescription,ifacilityid,itermid, iDisplayorder)

	If itermid = "0" Then
		sSql = "INSERT INTO egov_recreation_terms (facilityid, termdescription, displayorder) Values ('" & iFacilityId & "', '" &  DBsafe( sDescription) & "', " & iDisplayorder & ")"
	Else 
		sSQL = "UPDATE egov_recreation_terms SET termdescription = '" &  DBsafe( sDescription ) & "', displayorder = " & iDisplayorder & " WHERE termid = '" & itermId & "'"
	End If
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

End Sub
%>


