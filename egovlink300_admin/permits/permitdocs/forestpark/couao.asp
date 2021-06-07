<style>
	*, body {font-size:16px !important;line-height:26px}
	h1 {font-size:34px !important;}
	h2 {font-size:20px !important;font-weight:bold;}
	h3 {font-size:18px !important;font-weight:normal;}

	.row
	{
		width:100%;
		display:flex;
		flex-direction:row;
		flex-wrap:nowrap;
		justify-content:space-between;
	}
	.row div
	{
		margin-left:5px;
	}
	.row div:first-child
	{
		margin-left:0;
	}
	.row .value
	{
		flex-grow:4;
		border-bottom: 1px solid black;
		text-align:center;
		margin:0;
	}
	.row .spacer
	{
		flex-grow:4;
		text-align:center;
	}

	.row.signatures div div:first-child
	{
		border-bottom: 1px solid black;
		width:90%;
		margin-left:auto;
		margin-right:auto;
	}
	.row.signatures div
	{
		flex-grow:4;
		text-align:center;
		margin:0;
	}

</style>
<br />
<br />
<br />
<center>
	<h1>
	<%=strTITLE%>
	</h1>
	<h2>
	Department of Planning, Building & Zoning
	</h2>
	<h3>
	City of Forest Park<br />
	Clayton County, Georgia
	</h3>
	<i>This Certificate issued pursuant to the requirements of the City of Forest Park Building Code certifying that at the time of issuance this structure was in compliance with the various ordinances of the Jurisdiction regulating building construction and/or zoning use. For the following:</i>
</center>
<div class="row">
	<div class="field">Use Classification</div>
	<div class="value">&nbsp;</div>
	<div class="field">Bldg. Permit No.</div>
	<div class="value"><%=strPermitNumber%>&nbsp;</div>
</div>
<div class="row">
	<div class="field">Group</div>
	<div class="value"><%=oRs("usegroupcode")%></div>
	<div class="field">Type Construction</div>
	<div class="value"><%=oRs("constructiontype")%>&nbsp;</div>
	<div class="field">Fire District</div>
	<div class="value">&nbsp;</div>
</div>
<div class="row">
	<div class="field">Owner of Building / Business</div>
	<div class="value"><%=strOwnerName%></div>
</div>
<div class="row">
	<div class="field">Building Address:</div>
	<div class="value">
	<%
		response.write oRs("residentstreetnumber")
		If oRs("residentstreetprefix") <> "" Then
			response.write " " & oRs("residentstreetprefix")
		End If
		response.write " " & oRs("residentstreetname")
		If oRs("streetsuffix") <> "" Then
			response.write " " & oRs("streetsuffix")
		End If
		If oRs("streetdirection") <> "" Then
			response.write " " & oRs("streetdirection")
		End If
		If oRs("residentunit") <> "" Then
			response.write ", " & oRs("residentunit")
		End If
		%>&nbsp;</div>
	<div class="field">Locality</div>
	<div class="value">Forest Park, GA 30297</div>
</div>
<div class="row">
	<div class="spacer"></div>
	<div class="spacer"></div>
	<div class="spacer"></div>
	<div class="field">By:</div>
	<div class="value">&nbsp;</div>
</div>
<div class="row">
	<div class="spacer"></div>
	<div class="spacer"></div>
	<div class="spacer"></div>
	<div class="field">Date:</div>
	<div class="value"><%=FormatDateTime(date())%>&nbsp;</div>
</div>
<br />
<div class="row signatures">
	<div>
		<div>&nbsp;</div>
		<div>Building Official</div>
	</div>
	<div>
		<div>&nbsp;</div>
		<div>Zoning Administrator</div>
	</div>
	<div>
		<div>&nbsp;</div>
		<div>Fire Marshal</div>
	</div>
</div>
<center>POST IN A CONSPICUOUS PLACE</center>
<!--
Owner Address: <%=strOwnerAddress%>
Edition of the OBC used for review: <%=strBLDQCODE%>
<br />
Use Group(s) 302.0 & Specific Occupancies &mdash; 401.0: <%=oRs("usegroupcode")%>
<br />
Description of Occupancy &mdash; 303.0 to 312.0: <%=strDESCOCC%>
<br />
Construction Type &mdash; Chapter 6: 
<br />
Automatic Sprinklers &mdash; 903.0:  <%=strAutoSprinklers%>
<blockquote style="line-height:1px;margin: 0 0 18px 40px;">Hazard Classification: <%=strHazard%></blockquote>
Special Conditions/ Variances Granted:
<br />
<%=strPermitConditions%>
		-->
