<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!--#Include file="facility_functions.asp"-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: FACILITY_AVAILABILITY.ASP
' AUTHOR: STEVE LOAR
' CREATED: 01/17/06
' COPYRIGHT: COPYRIGHT 2006 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
'
' MODIFICATION HISTORY
' 1.?   01/18/2006	STEVE LOAR - CODE ADDED
' 2.0	01/22/2007	JOHN STULLENBERGER - ADD FACILITY RATE SELECTION OPTION
' 2.1	06/11/2010	Steve Loar - Modified to add note about the role the description plays in public side
'								 availability displaying correctly
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFacilityId, sFacilityName, sSql, oRs, oFacilities, iRowCount, x

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit facilities" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

If request("facilityid") = "" Then
	response.redirect( "facility_management.asp" )
Else 
	iFacilityId = CLng(request("facilityid"))
End If

sFacilityName = GetFacilityName( iFacilityId )

iPriceTypeGroupId = GetFacilityPTG( iFacilityId )

%>

<html lang="en">
<head>
	<meta charset="UTF-8">

	<title>E-Gov Facility Availability</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="facility.css" />

	<script src="../scripts/jquery-1.9.1.min.js"></script>

	<script>
	<!--

		function ConfirmDelete( iFacilitytimepartid, iFacilityId ) 
		{
			var msg = "Do you wish to delete this time part?"
			if ( confirm(msg) )
			{
				location.href='availability_delete.asp?iFacilitytimepartid='+ iFacilitytimepartid + '&iFacilityId=' + iFacilityId;
			}
		}

		function SaveAvailability( passForm )
		{
			//alert(passForm.sDescription.value);

			if (passForm.beginhour.value == "") 
			{
				alert("Please enter a start time.");
				passForm.beginhour.focus();
				return;
			}

			if (passForm.endhour.value == "") 
			{
				alert("Please enter an ending time.");
				passForm.endhour.focus();
				return;
			}

			var rege = /^\d{1,2}\:{1}\d{2}$/;
			var Ok = rege.exec(passForm.beginhour.value);

			if (! Ok) 
			{
				alert ("The Begin Time must be input in the format of HH:MM.");
				passForm.beginhour.focus();
				passForm.beginhour.select();
				return;
			}

			Ok = rege.exec(passForm.endhour.value);
	
			if (! Ok) 
			{
				alert ("The End Time must be input in the format of HH:MM.");
				passForm.endhour.focus();
				passForm.endhour.select();
				return;
			}

			if (passForm.description.value == "") 
			{
				alert("Please enter a description.");
				passForm.description.focus();
				return;
			}

			passForm.submit();
		}

		function savePTG( facilityId )
		{
			var priceTypeGroupId = $("#pricetypegroupid").val();
			var orgId = <%= session("orgid") %>;

			// fire off jQuery ajax to update this
			var request = jQuery.ajax({  
				type: "POST", 
				url: "updatefacilityPTG.asp",
				data: { 
						orgid : orgId,
						facilityid : facilityId,
						pricetypegroupid : priceTypeGroupId
					  },  
				dataType: "html"
			}); 

			request.done( function( data, textStatus, jqXHR ) { 
				// show the success message
				displayScreenMsg( "#groupchangemessage", "Changes Saved." )

				//$("#groupchangemessage").show();
				//$("#groupchangemessage").hide(10000);
			});

			request.fail( function(jqXHR, textStatus, errorThrown) {
				alert( "Failed to update the pricing offered: " + textStatus );
			});

		}

		function displayScreenMsg( containerId, sMsg ) 
		{
			if( sMsg != "" ) 
			{
				$(containerId).html( sMsg );
				$(containerId).show();
				setTimeout(function() {
					clearScreenMsg(containerId);
				}, 4000);
			}
		}

		function clearScreenMsg( containerId ) 
		{
			$(containerId).hide(2000);
			$(containerId).html( "" );
		}

	//-->
	</script>

</head>
<body>

 <% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
<div id="centercontent">	
	<p>
		<h3>Recreation: Facility Availability - <%=sFacilityName%></h3><br /><br />
		<input type="button" class="button" value="<< <%=langBackToStart%>" onclick="location.href='facility_management.asp';" />
		<br /><br />
	</p>

	<section id="priceTypeGroupPicks">
		<label for="pricetypegroupid">Offer Pricing For:</label> <% ShowFacilityPrictTypeGroups( iPriceTypeGroupId ) %>
		<input type="button" class="button" value="Save" onclick="savePTG(<%=iFacilityId%>);" />
		<span id="groupchangemessage"></span>
	</section>

	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin" id="facilityavailability">
		<tr>
			<th>Week Day</th><th>Begin Time</th><th>End Time</th><th>Description</th><th>Rate - (<a href="facility_rates.asp?facilityid=<%=iFacilityId%>">New Rate</a>)</th><th>&nbsp;</th>
		</tr>

		<!--  Always start with a blank row for adding -->
		<tr>
		<td><form name="availform0" method="POST" action="availability_save.asp">
				<input type="hidden" name="facilitytimepartid" value="0">
				<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>">
				<select name="weekday">
					<% For x = 1 To 7
						response.write "<option value=" & x & ">" & WeekDayName(x) & "</option>"
					   Next 
					%>
				</select>
				</td>
				<td nowrap="nowrap"><input type="text" name="beginhour" value="" size="5" maxlength="5">&nbsp;
					<select name="beginampm">
						<option value="AM">AM</option>
						<option value="PM">PM</option>
					</select>
				</td>
				<td nowrap="nowrap"><input type="text" name="endhour" value="" size="5" maxlength="5">&nbsp;
					<select name="endampm">
						<option value="AM">AM</option>
						<option value="PM">PM</option>
					</select>
				</td>
				<td><input type="text" name="description" value="" size="20" maxlength="20"></td>
				<td><!--<input type="text" name="rate" value="" size="10" maxlength="10">--->
				
				<% SubDisplayRates 0 %>
				
				</td>
				<td class="action">
				<!--<a href="javascript:SaveAvailability(document.availform0);">Add</a> -->
				<input type="button" class="button" value="Add" onclick="javascript:SaveAvailability(document.availform0);" />
			</form>	
		</td>
		</tr>
<%
		sSql = "SELECT rateid,facilityid, facilitytimepartid, beginhour, beginampm, endhour, endampm, weekday, "
		sSql = sSql & "description, rate FROM egov_facilitytimepart "
		sSql = sSql & "WHERE facilityid = " & iFacilityId & " ORDER BY weekday, beginampm, beginhour"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1
		
		If Not oRs.EOF Then
			iRowCount = 0
			Do While Not oRs.EOF
				' print out the lines here
				iRowCount = iRowCount + 1
				If iRowCOunt Mod 2 = 1 Then
					response.write "<tr class=""alt_row"">"
				Else
					response.write "<tr>"
				End If
				
%>
				<td><form name="availform<%=iRowCount%>" method="post" action="availability_save.asp">
				<input type="hidden" name="facilitytimepartid" value="<%=oRs("facilitytimepartid")%>">
				<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>">
				<select name="weekday">
					<% For x = 1 To 7
						response.write "<option value=""" & x & """"
						If oRs("weekday") = x Then
							response.write " selected=""selected"" "
						End If 
						response.write ">" & WeekDayName(x) & "</option>"
					   Next 
					%>
				</select>
				
				</td>
				<td nowrap="nowrap"><input type="text" name="beginhour" value="<%=oRs("beginhour")%>" size="5" maxlength="5">&nbsp;
					<select name="beginampm">
						<option value="AM" <% If oRs("beginampm") = "AM" Then 
													response.write " selected=""selected"" "
												End If %> >AM</option>
						<option value="PM" <% If oRs("beginampm") = "PM" Then 
													response.write " selected=""selected"" "
												End If %> >PM</option>
					</select>
				</td>
				<td nowrap="nowrap"><input type="text" name="endhour" value="<%=oRs("endhour")%>" size="5" maxlength="5">&nbsp;
					<select name="endampm">
						<option value="AM" <% If oRs("endampm") = "AM" Then 
													response.write " selected=""selected"" "
												End If %>>AM</option>
						<option value="PM" <% If oRs("endampm") = "PM" Then 
													response.write " selected=""selected"" "
												End If %> >PM</option>
					</select>
				</td>
				<td><input type="text" name="description" value="<%=oRs("description")%>" size="20" maxlength="20"></td>
				<td>
		
				<%
					' GET RATES
					SubDisplayRates oRs("rateid")
				%>
				</td>
				<td class="action">
					<!--<a href="javascript:SaveAvailability(document.availform<%=iRowCount%>);">Save</a> -->
					<input type="button" class="button" value="Save" onclick="javascript:SaveAvailability(document.availform<%=iRowCount%>);" />&nbsp;&nbsp;
					<!-- <a href="javascript:ConfirmDelete(<%=oRs("facilitytimepartid")%>,<%=iFacilityId%>);">Delete</a> -->
					<input type="button" class="button" value="Delete" onclick="javascript:ConfirmDelete(<%=oRs("facilitytimepartid")%>,<%=iFacilityId%>);" />
					</form>	
				</td>
				</tr>
<%
				oRs.MoveNext
			Loop 
		End If 

		oRs.Close
		Set oRs = Nothing 
%>
	</table>

	<p>
<%		If OrgHasDisplay( session("orgid"), "facilityavailabilitynotes" ) Then	
			response.write GetOrgDisplay( session("orgid"), "facilityavailabilitynotes" )
		Else 		%>		
			*note: In order for the public display of available times to display and work correctly
			the Description fields MUST be in the correct alphabetical order. First must be the 
			&quot;All Day&quot; time period, then the &quot;Daytime&quot; time period, and finally
			the &quot;Evening&quot; time period.
<%		End If		%>
	</p>

</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' integer = GetFacilityPTG( iFacilityId )
'--------------------------------------------------------------------------------------------------
Function GetFacilityPTG( ByVal iFacilityId )
	Dim sSql, oRs

	sSql = "SELECT pricetypegroupid FROM egov_facility WHERE facilityid = " & iFacilityId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetFacilityPTG = oRs("pricetypegroupid")
	Else
		GetFacilityPTG = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' ShowFacilityPrictTypeGroups iPriceTypeGroupId
'--------------------------------------------------------------------------------------------------
Sub ShowFacilityPrictTypeGroups( ByVal iPriceTypeGroupId )
	Dim sSql, oRs

	sSql = "SELECT DISTINCT G.pricetypegroupid, G.pricetypegroup, G.displayorder "
	sSql = sSql & "FROM egov_price_type_groups G, egov_price_types P "
	sSql = sSql & "WHERE G.pricetypegroupid = P.pricetypegroupid AND P.orgid = " & session("orgid")
	sSql = sSql & " ORDER BY G.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""pricetypegroupid"" id=""pricetypegroupid"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("pricetypegroupid") & """"
			If CLng(oRs("pricetypegroupid")) = CLng(iPriceTypeGroupId) Then
				response.write " selected=""selected"""
			End If 
			response.write ">" & oRs("pricetypegroup") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 

%>