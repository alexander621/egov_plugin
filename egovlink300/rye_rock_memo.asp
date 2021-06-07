<body onload="window.print()">
<%

sSQL = "SELECT u.userfname + ' ' + u.userlname as Name, ad.answer as Address, a.parcelidnumber as ParcelID, " _
		& " sd.answer as StartDate, DATEADD(d,29,CONVERT(datetime,sd.answer)) as EndDate, rm.answer as RemovalType " _
	& " FROM egov_actionline_requests ar " _
	& " INNER JOIN egov_users u ON u.userid = ar.userid " _
	& " INNER JOIN action_submitted_questions_and_answers ad ON ad.action_autoid = ar.action_autoid and ad.question = 'Address' " _
	& " INNER JOIN egov_residentaddresses a ON ad.answer = a.residentstreetnumber + ' ' + a.residentstreetname " _
	& " INNER JOIN action_submitted_questions_and_answers sd ON sd.action_autoid = ar.action_autoid and sd.question = 'Start Date' " _
	& " INNER JOIN action_submitted_questions_and_answers rm ON rm.action_autoid = ar.action_autoid and rm.question = 'Type of Removal' " _
	& " WHERE ar.action_autoid = '" & DBSafe(request.querystring("id")) & "'" 
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN") , 3, 1
if oRs.EOF Then%>
	Your memo could not be found.
<%else%>
<!--p>
<h1 style="margin:0;">Rock Removal Notification</h1>
<h2 style="margin:0;">City of Rye, New York</h2>
1051 Boston Post Road, Rye, New York 10580<br />
www.ryeny.gov<br />
</p>
<p>
<%=oRs("Name")%> or the property owner’s agent is hereby authorized to engage in mechanical rock removal
or blasting at property located at <%=oRs("Address")%>, (<%=oRs("ParcelID")%>) Rye, New York. Mechanical
rock removal or blasting is authorized from 9:00 AM <%=oRs("StartDate")%> to no later than 5:00 PM  <%=oRs("EndDate")%> subject to the following restrictions:
<ul>
	<li>This notification must be posted at the property so that it is visible from the street.</li>
	<li>Pursuant to §133-8 of the Rye City Code, no person shall conduct mechanical rock removal or blasting operations using explosives, within the City of Rye after the hour of 5:00 p.m. and before 9:00 a.m. nor at any time on Saturdays, Sundays or any of the following holidays: New Year's Day, Presidents' Day, Memorial Day, Independence Day, Labor Day, Thanksgiving Day and Christmas Day.</li>
	<li>If the owner of a property or the owner’s agent fails to engages in rock removal activities for more than 30 calendar days they shall be guilty of an offense and shall, upon conviction thereof, be subject to a fine of not more than $1,000, an order to suspend construction work on the site, or by imprisonment not exceeding 15 days, or any combination of such fine, suspension and imprisonment. Each day of mechanical rock removal and/or use of explosives prior to sending in notice of the commencement date or in violation of the thirty (30) day limit shall be construed as a separate offense.</li>
</ul>
</p-->

<table>
	<tr>
		<td><img src="img/RyeLogo.gif" /></td>
		<td style="padding-left:50px;font-size:12pt;">
			<span style="font-weight:bold;font-size:20pt;">Rock Removal Registration</span><br />
			<span style="font-weight:bold;font-size:14pt;">City of Rye, New York</span><br />
			1051 Boston Post Road, Rye, New York 10580<br />
			<span style="font-size:10pt;"><a href="http://www.ryeny.gov">www.ryeny.gov</a></span>
		</td>
	</tr>
</table>
<hr noshade />

<style>
	body {font-size:12pt;}
	td
	{
		padding: 10px 0 10px 0;
		font-size: 14pt;
	}
</style>

<table width="100%">
	<tr>
		<td valign="top"><b>Property Owner/Agent:</b></td>
		<td><%=oRs("Name")%></td>
	</tr>
	<tr>
		<td valign="top"><b>Address:</b></td>
		<td>
			<%=oRs("Address")%>
			<br />
			Rye, New York 10580

		</td>
	</tr>
	<tr>
		<td valign="top"><b>Parcel ID:</b></td>
		<td><%=oRs("ParcelID")%></td>
	</tr>
	<tr>
		<td valign="top"><b>Permitted Start Date:</b></td>
		<td><%=oRs("StartDate")%>&nbsp;(9:00 AM)</td>
	</tr>
	<tr>
		<td valign="top"><b>Termination Date:</b></td>
		<td><%=oRs("EndDate")%>&nbsp;(5:00 PM)</td>
	</tr>
				
</table>
<hr noshade style="margin-bottom:2px; border: 1px solid black;" />
<center><i><b>This notification shall be posted at the property so that it is visible from the street</b></i></center>
<hr noshade style="margin-top:2px;" />

<p>The above-referenced property is hereby authorized to engage in mechanical rock removal or blasting subject to the following restrictions as per §133-8 of the Rye City Code:</p>

<b>
<u>RESTRICTIONS:</u>
<ul>
<li>There shall be no mechanical rock removal or blasting operations after the hour of 5:00 p.m. and before 9:00 a.m.</li>

<li>There shall be no mechanical rock removal or blasting operations at any time on Saturdays, Sundays or any of the following holidays: New Year's Day, Presidents' Day, Memorial Day, Independence Day, Labor Day, Thanksgiving Day and Christmas Day.</li>
</ul>
</b>

<u>PENALTIES:</u>

<ul>
<li>If the owner of a property or the owner’s agent fails to engages in rock removal activities for more than 30 calendar days they shall be guilty of an offense and shall, upon conviction thereof, be subject to a fine of not more than $1,000, an order to suspend construction work on the site, or by imprisonment not exceeding 15 days, or any combination of such fine, suspension and imprisonment. Each day of mechanical rock removal and/or use of explosives prior to sending in notice of the commencement date or in violation of the thirty (30) day limit shall be construed as a separate offense.</li>
</ul>

<%end if
oRs.Close
Set oRs = Nothing

Function DBsafe( ByVal strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then 
		sNewString = strDB
	Else 
		sNewString = Replace( strDB, "'", "''" )
		sNewString = Replace( sNewString, "<", "&lt;" )
	End If 

	DBsafe = sNewString
End Function
%>
