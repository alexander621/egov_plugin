<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CLIENT_TEMPLATE_PAGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.0   01/18/06   Steve Loar - Code added
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFacilityId
Dim sFacilityName

If request("facilityid") = "" Then
	response.redirect( "facility_management.asp" )
Else 
	iFacilityId = request("facilityid")
End If

sFacilityName = GetFacilityName(iFacilityId)

' location.href='rate_save.asp?iRateId='+ iRateId + ',sRateDesc=' + passForm.Description.value + ',iFacilityId=' + iFacilityId;
%>


<!-- #include file="../includes/common.asp" //-->

<!--#Include file="facility_functions.asp"-->

<%
Dim oFacilities
Dim iRowCount
%>

<html>
<head>
	<title>E-Gov Facility Rates</title>

	<link rel="stylesheet" type="text/css" href="../global.css">
	<link rel="stylesheet" type="text/css" href="facility.css">

<script language="Javascript">
  <!--
	function ConfirmDelete(sRate, iRateId, iFacilityId) 
	{
		var msg = "Do you wish to delete " + sRate + "?"
		if (confirm(msg))
		{
			location.href='rate_delete.asp?iRateId='+ iRateId + '&iFacilityId=' + iFacilityId;
		}
	}

	function SaveRate(passForm)
	{
		//alert(passForm.sDescription.value);

		if (passForm.sDescription.value == "") {
			alert("Please enter a description.");
			passForm.sDescription.focus();
			return;
		}

		if (passForm.iRateValue.value == "") {
			alert("Please enter a rate.");
			passForm.iRateValue.focus();
			return;
		}

		var rege = /^\d+.?\d*$/;
		var Ok = rege.exec(passForm.iRateValue.value);

		if (! Ok) {
			alert ("Rates must be in money format.");
			passForm.iRateValue.focus();
			passForm.iRateValue.select();
			return;
		}
		passForm.submit();
	}

  //-->
 </script>
</head>


<body>

 
<%DrawTabs tabRecreation,1%>


<!--BEGIN PAGE CONTENT-->
<div id="content">
	
	<p>
	<font size="+1"><strong>Recreation: Facility Rate Management - <%=sFacilityName%></strong></font><br />
	<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
	</p>

	<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
		<tr>
			<th>Description</th><th>Amount</th><th>&nbsp;</th>
		</tr>

		<!--  Always start with a blank row for adding -->
		<td><form name="rateform0" method="POST" action="rate_save.asp">
				<input type="hidden" name="iRateId" value="0">
				<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>">
				<input type="text" name="sDescription" value="" size="80" maxlength="250"></td>
				<td><input type="text" name="iRateValue" value="" size="10" maxlength="10"></td>
				<td class="action">
				<a href="javascript:SaveRate(document.rateform0);">Add</a>
			</form>	
		</td>
<%
	sSQLb = "Select facilityid, rateid, ratedescription, ratevalue from egov_rate where facilityid = " & iFacilityId & " order by ratedescription"
		Set oRates = Server.CreateObject("ADODB.Recordset")
		oRates.Open sSQLb, Application("DSN"), 3, 1
		
		If Not oRates.EOF Then
			iRowCount = 0
			Do While Not oRates.EOF
				' print out the lines here
				iRowCount = iRowCount + 1
				If iRowCOunt Mod 2 = 1 Then
					response.write "<tr class=" & Chr(34) & "alt_row" & Chr(34) & ">"
				Else
					response.write "<tr>"
				End If
				
%>
				<td><form name="rateform<%=iRowCount%>" method="post" action="rate_save.asp">
				<input type="hidden" name="iRateId" value="<%=oRates("rateid")%>">
				<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>">
				<input type="text" name="sDescription" value="<%=oRates("ratedescription")%>" size="80" maxlength="250"></td>
				<td><input type="text" name="iRateValue" value="<%=oRates("ratevalue")%>" size="10" maxlength="10"></td>
				<td class="action">
					<a href="javascript:SaveRate(document.rateform<%=iRowCount%>);">Save</a>&nbsp;&nbsp;
					<a href="javascript:ConfirmDelete('<%=oRates("ratedescription")%>',<%=oRates("rateid")%>,<%=iFacilityId%>);">Delete</a>
					</form>	
				</td>

				</tr>
<%
				oRates.MoveNext
			Loop 
		End If 
		oRates.close
		Set oRates = nothing
%>
	</table>
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


%>


