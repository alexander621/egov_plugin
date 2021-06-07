<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: inspectionreport.asp
' AUTHOR: Terry Foster
' CREATED: 12/11/2019
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Edits permit inspections
'
' MODIFICATION HISTORY
' 1.0   12/11/2019	Terry Foster - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
intPermitInspectionReportID = CLng(request("permitinspectionreportid"))
intPermitID = CLng(request("permitid"))
blnNewRecord = false
blnPrint = false
blnCopy = false


if request("print") = "true" then blnPrint = true
blnPrint = true
if request.querystring("copy") = "true" and request.servervariables("REQUEST_METHOD") <> "POST" then blnCopy = true
if request.servervariables("REQUEST_METHOD") = "POST" then
	SaveChanges
end if

'GET FORM DATA FROM DATABASE BY INSPECTION REPORT ID
sSQL = "SELECT permitinspectionreportid, ir.permitid, datecreated, inspectiontype, approved, disapproved, approvedwcorr, coc, remarks, p.orgid, u.firstname, u.lastname " _
	& " FROM egov_permitinspectionreports ir " _
	& " INNER JOIN egov_permits p ON p.permitid = ir.permitid " _
	& " LEFT JOIN users u ON u.userid = ir.permitinspectorid " _
	& " WHERE permitinspectionreportid = '" & intPermitInspectionReportID & "'"
set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1
if not oRs.EOF then
	strInspectionType = oRs("inspectiontype")
	if oRs("approved") and not blnCopy then strApproved = "checked"
	if oRs("disapproved") and not blnCopy then strDisapproved = "checked"
	if oRs("approvedwcorr") and not blnCopy then strApprovedWCorr = "checked"
	if oRs("coc") and not blnCopy then strCOC = "checked"
	if not blnCopy then strRemarks = oRs("remarks")
	intPermitID = oRs("permitid")
	strDateCreated = oRs("datecreated")
	if blnPrint then session("orgid") = oRs("orgid")
	strInspectorName = oRs("firstname") & " " & oRs("lastname")
end if
oRs.Close
Set oRs = Nothing
strPermitNumber = GetPermitNumber(intPermitID)

bPermitIsCompleted = GetPermitIsCompleted( intPermitID ) '	in permitcommonfunctions.asp

bIsOnHold = GetPermitIsOnHold( intPermitID ) '	in permitcommonfunctions.asp

Set iType=Server.CreateObject("Scripting.Dictionary")
if intPermitInspectionReportID <> "0" then
	sSQL = "SELECT inspectiontype FROM egov_permitinspectionreporttypes WHERE permitinspectionreportid = " & intPermitInspectionReportID
	'response.write sSQL
	'response.end

	set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	Do While not oRs.EOF
		iType.Add oRs("inspectiontype") & "","checked"
		oRs.MoveNext
	loop
	oRs.Close
	Set oRs = Nothing
end if
if blnCopy then intPermitInspectionReportID = 0
%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="JavaScript" src="../scripts/layers.js"></script>
	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
  <style>
  	table th {font-weight:bold !important;}
	.tblinspectiontypes td {text-align:center;}
	.tblinspectiontypes tr.approvalchecks td 
	{
		text-align:left;
		font-weight:bold !important;
	}
  </style>

	<script language="Javascript">
	<!--
	<% if blnNewRecord then %>
		var maxirs = parseInt(parent.document.getElementById("maxirs").value) + 1;
		var myHtmlContent = '<tr id="IRrow' + maxirs + '"><td align="center"><input type="hidden" id="permitinspectionreportid' + maxirs +'" name="permitinspectionreportid' + maxirs + '" value="<%=intPermitInspectionReportID%>" /><input type="checkbox" name="removeIR' + maxirs +'" id="removeIR' + maxirs +'" /></td><td onclick="ViewIR(<%=intPermitInspectionReportID%>);"><%=strDateCreated%></td><td><a href="javascript:CopyIR(<%=intPermitInspectionReportID%>)">Create Copy</a></td><td><a href="javascript:PrintIR(<%=intPermitInspectionReportID%>)">Print</a></td><td><a href="javascript:EmailIR(<%=intPermitInspectionReportID%>)">Email</a></td></tr>';
		var tableRef = parent.document.getElementById('inspectionreportlist').getElementsByTagName('tbody')[0];

		var newRow = tableRef.insertRow(tableRef.rows.length);
		newRow.id = 'IRrow' + maxirs;
		newRow.innerHTML = myHtmlContent;

		parent.document.getElementById("maxirs").value = maxirs;
		
	<% end if %>
		
		var bHasInspectedDate = false;

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function doValidate()
		{
			document.frmInspectionReport.submit();

 	 	}



		function doLoad()
		{
			setMaxLength();
			<% if blnPrint then %>
				window.print();
			<% end if %>
		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>

  $( function() {
    $( ".datepicker" ).datepicker({
      changeMonth: true,
      showOn: "both",
      buttonText: "<i class=\"fa fa-calendar\"></i>",
      changeYear: true
    });
  } );
	//-->
	</script>


</head>

<body onload="doLoad();">

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	

	<!--BEGIN: EDIT FORM-->
	<form name="frmInspectionReport" method="post">
	<input type="hidden" name="permitid" value="<%=intPermitID%>" />
	<input type="hidden" name="permitinspectionreportid" value="<%=intPermitInspectionReportID%>" />
	<input type="hidden" name="copy" value="false" />

	<p>
		<table>
			<tr>
				<td>Building Permit # <%=strPermitNumber %></td>
				<td align="right">Inspection Type: <% if blnPrint then %><%=strInspectionType%><%else%><input type="text" size=40 name="inspectiontype" value="<%=strInspectionType%>" /><%end if%></td>
			</tr>
		</table>
		<%
			sSQL = "SELECT o.* " _
				& " FROM egov_permitinspectionreporttypes t " _
				& " INNER JOIN egov_permitinspectionreporttypeoptions o ON o.code = t.inspectiontype " _
				& " WHERE t.permitinspectionreportid = " & intPermitInspectionReportID
			Set oIT = Server.CreateObject("ADODB.Recordset")
			oIT.Open sSql, Application("DSN"), 3, 1
			Do While Not oIT.EOF
				if strGroup <> oIT("optgroup") then
					if strGroup <> "" then response.write "<br />"
					strGroup = oIT("optgroup")
					response.write "<b>" & strGroup & "</b><br />"
			
			
				end if
				response.write "<input type=""checkbox"" checked />" & oIT("description") & "<br />" & vbcrlf
				oIT.MoveNext
			loop
		%>
		<!--table class="tblinspectiontypes">
			<tr><th>FOOTER</th><th>FRAMING</th><th>PLUMBING</th><th>ELECTRICAL</th></tr>
			<tr><td><input type="checkbox" name="it10" <%=iType.Item("it10")%> />Termite Protection R318</td><td><input type="checkbox" name="it20" <%=iType.Item("it20")%> />Fire Blocking R302.11</td><td><input type="checkbox" name="it30" <%=iType.Item("it30")%> />Rodent Proofing 304</td><td><input type="checkbox" name="it40" <%=iType.Item("it40")%> />Elec. Connections (De-ox) 110.14</td></tr>
			<tr><td><input type="checkbox" name="it50" <%=iType.Item("it50")%> />Soil Testing R401.4</td><td><input type="checkbox" name="it60" <%=iType.Item("it60")%> />Tempered Glass R308.4</td><td><input type="checkbox" name="it70" <%=iType.Item("it70")%> />Pipe Protection 305</td><td><input type="checkbox" name="it80" <%=iType.Item("it80")%> />Crawl space lighting outlets 210E</td></tr>
			<tr><td><input type="checkbox" name="it90" <%=iType.Item("it90")%> />Roof Drainage R801.3</td><td><input type="checkbox" name="it100" <%=iType.Item("it100")%> />Egress Windows R310.1.1</td><td><input type="checkbox" name="it110" <%=iType.Item("it110")%> />AVX &amp; Sec Drain Sys 307.2.3</td><td><input type="checkbox" name="it120" <%=iType.Item("it120")%> />Branch Circuits Required 210.11</td></tr>
			<tr><td><input type="checkbox" name="it130" <%=iType.Item("it130")%> />All Thread Spacing R4504.1, R4504.2.1</td><td><input type="checkbox" name="it140" <%=iType.Item("it140")%> />Stair Headroom R311.7.2</td><td><input type="checkbox" name="it150" <%=iType.Item("it150")%> />Interval of Support 308.5</td><td><input type="checkbox" name="it160" <%=iType.Item("it160")%> />Arc Fault 210.12</td></tr>
			<tr><td><input type="checkbox" name="it170" <%=iType.Item("it170")%> />Footing Size R4503.1.1</td><td><input type="checkbox" name="it180" <%=iType.Item("it180")%> />Treads &amp; Risers R311.7.5</td><td><input type="checkbox" name="it190" <%=iType.Item("it190")%> />Porta-Potty 311</td><td><input type="checkbox" name="it200" <%=iType.Item("it200")%> />Required Outlets 210.52</td></tr>
			<tr><td><input type="checkbox" name="it210" <%=iType.Item("it210")%> />Rebar/threaded rods 3" From Bottom -</td><td><input type="checkbox" name="it220" <%=iType.Item("it220")%> />Handrails R311.7.8.1</td><td><input type="checkbox" name="it230" <%=iType.Item("it230")%> />Drain and Vent Tests 312.2, 3</td><td><input type="checkbox" name="it240" <%=iType.Item("it240")%> />H/AC/R Outlet 25' of Unit 210.63</td></tr>
			<tr><td><input type="checkbox" name="it250" <%=iType.Item("it250")%> />(3" from dirt R4503.1.2 and ACI 318)</td><td><input type="checkbox" name="it260" <%=iType.Item("it260")%> />Guardrails R312.1.1</td><td><input type="checkbox" name="it270" <%=iType.Item("it270")%> />Water supply system Test 312.5</td><td><input type="checkbox" name="it280" <%=iType.Item("it280")%> />Lighting Outlets Required 210.70</td></tr>
			<tr><td><input type="checkbox" name="it290" <%=iType.Item("it290")%> />25" Lap Splice R4503.1.2</td><td><input type="checkbox" name="it300" <%=iType.Item("it300")%> />Floor Truss Dwgs R502.11</td><td><input type="checkbox" name="it310" <%=iType.Item("it310")%> />Shower liner or pan test - 312.9</td><td><input type="checkbox" name="it320" <%=iType.Item("it320")%> />GFCI Protection 210.8</td></tr>
			<tr><td><input type="checkbox" name="it330" <%=iType.Item("it330")%> />Spacing of Piles 4603.5</td><td><input type="checkbox" name="it340" <%=iType.Item("it340")%> />Floor Joist Spans R502.3.1</td><td><input type="checkbox" name="it350" <%=iType.Item("it350")%> />W/H Garage 18" 502.1.1</td><td><input type="checkbox" name="it360" <%=iType.Item("it360")%> />Perm install generators - 250.35</td></tr>
			<tr><td></td><td><input type="checkbox" name="it370" <%=iType.Item("it370")%> />Drilling Joist R502.8.1</td><td><input type="checkbox" name="it380" <%=iType.Item("it380")%> />Water Hammer 604.9</td><td><input type="checkbox" name="it390" <%=iType.Item("it390")%> />Grounding Electrode 250.53</td></tr>
			<tr><th>SLAB</th><td><input type="checkbox" name="it400" <%=iType.Item("it400")%> />Notching Joist R502.8.1</td><td><input type="checkbox" name="it410" <%=iType.Item("it410")%> />PVC primer - purple 605.22.3, 705.11.2</td><td><input type="checkbox" name="it420" <%=iType.Item("it420")%> />Bushings 300.4G</td></tr>
			<tr><td><input type="checkbox" name="it430" <%=iType.Item("it430")%> />Pipe Protection 305</td><td><input type="checkbox" name="it440" <%=iType.Item("it440")%> />Max Stud Height R4505(a),R602.3</td><td><input type="checkbox" name="it450" <%=iType.Item("it450")%> />Hose vacuum breaker 608.15.4.2</td><td><input type="checkbox" name="it460" <%=iType.Item("it460")%> />Bundling 310.15B2</td></tr>
			<tr><td><input type="checkbox" name="it470" <%=iType.Item("it470")%> />Drain Test 312.2</td><td><input type="checkbox" name="it480" <%=iType.Item("it480")%> />Top plate(bear&amp;braced walls)R602.3.2</td><td><input type="checkbox" name="it490" <%=iType.Item("it490")%> />PVC primer sanit.- clear exp.705.11.2</td><td><input type="checkbox" name="it500" <%=iType.Item("it500")%> />Service Wire Size SFD 310.15B7</td></tr>
			<tr><td><input type="checkbox" name="it510" <%=iType.Item("it510")%> />Drain Test Air 312.3</td><td><input type="checkbox" name="it520" <%=iType.Item("it520")%> />Stud Full Bearing R602.3.4</td><td><input type="checkbox" name="it530" <%=iType.Item("it530")%> />Dishwasher waste line secured 802.1.6</td><td><input type="checkbox" name="it540" <%=iType.Item("it540")%> />Smoke Alarms - R314</td></tr>
			<tr><td><input type="checkbox" name="it550" <%=iType.Item("it550")%> />Water Test 312.5</td><td><input type="checkbox" name="it560" <%=iType.Item("it560")%> />Drill &amp; Notch Studs R602.6</td><td><input type="checkbox" name="it570" <%=iType.Item("it570")%> />Washer Standpipes 802.3.3</td><td><input type="checkbox" name="it580" <%=iType.Item("it580")%> />Box Depth 314.24</td></tr>
			<tr><td><input type="checkbox" name="it590" <%=iType.Item("it590")%> />Vapor Retarder R506.2.3</td><td><input type="checkbox" name="it600" <%=iType.Item("it600")%> />Drill &amp; Notch Top Plate R602.6.1</td><td><input type="checkbox" name="it610" <%=iType.Item("it610")%> />Trap Arm Length 909</td><td><input type="checkbox" name="it620" <%=iType.Item("it620")%> />Carbon Monoxide Alarms - R315</td></tr>
			<tr><td></td><td><input type="checkbox" name="it630" <%=iType.Item("it630")%> />Girder and Header Spans R602.7</td><td><input type="checkbox" name="it640" <%=iType.Item("it640")%> />Prohibited Traps 1002.3</td><td><input type="checkbox" name="it650" <%=iType.Item("it650")%> />Ground Switches 404.9A</td></tr>
			<tr><td></td><td><input type="checkbox" name="it660" <%=iType.Item("it660")%> />King and Jack Studs R602.7(1),602.7.5</td><td><input type="checkbox" name="it670" <%=iType.Item("it670")%> />Cut &amp; Notch Appendix C</td><td><input type="checkbox" name="it680" <%=iType.Item("it680")%> />Outlets (Bath Area) 406.8C</td></tr>
			<tr><td></td><td><input type="checkbox" name="it690" <%=iType.Item("it690")%> />Open Porch Spans R602.7(3)</td><th>MECHANICAL</th><td><input type="checkbox" name="it700" <%=iType.Item("it700")%> />"Extra Duty" Covers 406.9B</td></tr>
			<tr><th>FOUNDATION</th><td><input type="checkbox" name="it710" <%=iType.Item("it710")%> />Box Panel Header Spans R602.7.3</td><td><input type="checkbox" name="it720" <%=iType.Item("it720")%> />Modifications approval L&amp;A 105.1</td><td><input type="checkbox" name="it730" <%=iType.Item("it730")%> />Lights (Closet) 410.16</td></tr>
			<tr><td><input type="checkbox" name="it740" <%=iType.Item("it740")%> />Pier/ftg/girder bearing R403.1(2)note c</td><td><input type="checkbox" name="it750" <%=iType.Item("it750")%> />Continuous Sheathing R602.10.3</td><td><input type="checkbox" name="it760" <%=iType.Item("it760")%> />Anchor Equipment 301.15</td><td><input type="checkbox" name="it770" <%=iType.Item("it770")%> />Lights (Bath Area) 410.4D</td></tr>
			<tr><td><input type="checkbox" name="it780" <%=iType.Item("it780")%> />All Thread Spacing R403.1.6, R4504.1</td><td><input type="checkbox" name="it790" <%=iType.Item("it790")%> />Roof Truss Dwgs R802.10</td><td><input type="checkbox" name="it800" <%=iType.Item("it800")%> />Flood Hazard 301.16</td><td><input type="checkbox" name="it810" <%=iType.Item("it810")%> />Double Insulated Pool Pump 680.21</td></tr>
			<tr><td><input type="checkbox" name="it820" <%=iType.Item("it820")%> />Pier Height/Pier Fill/Eng.Fill R404.1.5</td><td><input type="checkbox" name="it830" <%=iType.Item("it830")%> />Ridge board &amp; hip/valley detail R802.3</td><td><input type="checkbox" name="it840" <%=iType.Item("it840")%> />Rodent Proofing 301.17</td><td><input type="checkbox" name="it850" <%=iType.Item("it850")%> />Pool Outlets 680.22</td></tr>
			<tr><td><input type="checkbox" name="it860" <%=iType.Item("it860")%> />Foundation vent sizing R408.1.1</td><td><input type="checkbox" name="it870" <%=iType.Item("it870")%> />Clg Joist and Rafter Connect R802.3.1</td><td><input type="checkbox" name="it880" <%=iType.Item("it880")%> />Cut &amp; Notch 302.3</td><td><input type="checkbox" name="it890" <%=iType.Item("it890")%> />Pool Equipotential Bonding 680.26</td></tr>
			<tr><td><input type="checkbox" name="it900" <%=iType.Item("it900")%> />Ground vapor retarder R408.2</td><td><input type="checkbox" name="it910" <%=iType.Item("it910")%> />Ceiling Joist Overlap R802.3.2</td><td><input type="checkbox" name="it920" <%=iType.Item("it920")%> />Protection Equip. 303.4</td><td></td></tr>
			<tr><td><input type="checkbox" name="it930" <%=iType.Item("it930")%> />Crawl Access Min R408.8</td><td><input type="checkbox" name="it940" <%=iType.Item("it940")%> />Cut &amp; Notch Ceiling members R802.7</td><td><input type="checkbox" name="it950" <%=iType.Item("it950")%> />Manufact Instruct on site 304.1</td><td></td></tr>
			<tr><td><input type="checkbox" name="it960" <%=iType.Item("it960")%> />Removal of Debris R408.9</td><td><input type="checkbox" name="it970" <%=iType.Item("it970")%> />DP Rating R4502</td><td><input type="checkbox" name="it980" <%=iType.Item("it980")%> />Exterior Grade Installations 304.10.1</td><th>MISCELLANEOUS</th></tr>
			<tr><td><input type="checkbox" name="it990" <%=iType.Item("it990")%> />Site Prep R504.2</td><td><input type="checkbox" name="it1000" <%=iType.Item("it1000")%> />High wind nailing R4508.4(b)</td><td><input type="checkbox" name="it1010" <%=iType.Item("it1010")%> />Access and service space 306</td><td><input type="checkbox" name="it1020" <%=iType.Item("it1020")%> />Double Key Dead Bolt R311.2</td></tr>
			<tr><td><input type="checkbox" name="it1030" <%=iType.Item("it1030")%> />3 Nails Joist/Girder R602.3(1)</td><td><input type="checkbox" name="it1040" <%=iType.Item("it1040")%> />Gable End Walls R4506.3, R4506.7b</td><td><input type="checkbox" name="it1050" <%=iType.Item("it1050")%> />Condensate Disposal 307</td><td><input type="checkbox" name="it1060" <%=iType.Item("it1060")%> />Under Stair Protection R314.8</td></tr>
			<tr><td><input type="checkbox" name="it1070" <%=iType.Item("it1070")%> />Piers and Pier Caps R606.7,R606.7.1</td><td><input type="checkbox" name="it1080" <%=iType.Item("it1080")%> />Piles Notched&gt;50% 4603.6</td><td><input type="checkbox" name="it1090" <%=iType.Item("it1090")%> />Paint exterior ferrous gas line FG 404.9</td><td><input type="checkbox" name="it1100" <%=iType.Item("it1100")%> />Rental - smoke and CO alarms</td></tr>
			<tr><td><input type="checkbox" name="it1110" <%=iType.Item("it1110")%> />Crawl w/ Equipment M1305.1.4</td><td><input type="checkbox" name="it1120" <%=iType.Item("it1120")%> />X Bracing 4603.6</td><td><input type="checkbox" name="it1130" <%=iType.Item("it1130")%> />Gas Test FG 406</td><td></td></tr>
			<tr><td></td><th>ENERGY CODE</th><td><input type="checkbox" name="it1140" <%=iType.Item("it1140")%> />Dryer Exhaust 504</td><td><input type="checkbox" name="it1150" <%=iType.Item("it1150")%> />Modifications approval L&amp;A 105.1</td></tr>
			<tr><th>NAILING</th><td><input type="checkbox" name="it1160" <%=iType.Item("it1160")%> />Prot. Exp found. wall insul. R303.2.1</td><td><input type="checkbox" name="it1170" <%=iType.Item("it1170")%> />Duct Insulation 604.1</td><td><input type="checkbox" name="it1180" <%=iType.Item("it1180")%> />Work w/o Permit L&amp;A 305.2</td></tr>
			<tr><td><input type="checkbox" name="it1190" <%=iType.Item("it1190")%> />Over Hang End Walls R4506.7</td><td><input type="checkbox" name="it1200" <%=iType.Item("it1200")%> />R-19 Flr R-15 Wall R-38 Ceil R402.1.2</td><td></td><td><input type="checkbox" name="it1210" <%=iType.Item("it1210")%> />Plans on Site L&amp;A 304.4</td></tr>
			<tr><td><input type="checkbox" name="it1220" <%=iType.Item("it1220")%> />Roof Anchorage R4508.3</td><td><input type="checkbox" name="it1230" <%=iType.Item("it1230")%> />R-30 Ceiling w Attic space R402.2.1</td><th>MANUFACTURED HOME</th><td><input type="checkbox" name="it1240" <%=iType.Item("it1240")%> />Site Address R325.1</td></tr>
			<tr><td><input type="checkbox" name="it1250" <%=iType.Item("it1250")%> />Nail Spacing R4508.4</td><td><input type="checkbox" name="it1260" <%=iType.Item("it1260")%> />Closed crawl space wall R402.2.11</td><td><input type="checkbox" name="it1270" <%=iType.Item("it1270")%> />Manufacturers Set Up Book</td><td><input type="checkbox" name="it1280" <%=iType.Item("it1280")%> />Dumpster - Local ordinance</td></tr>
			<tr><td><input type="checkbox" name="it1290" <%=iType.Item("it1290")%> />3" Stagger R4508.4</td><td><input type="checkbox" name="it1300" <%=iType.Item("it1300")%> />Soffit baffles R402.2.3</td><td><input type="checkbox" name="it1310" <%=iType.Item("it1310")%> />Marriage Wall Bond</td><td><input type="checkbox" name="it1320" <%=iType.Item("it1320")%> />Site Card - Local requirement</td></tr>
			<tr><td><input type="checkbox" name="it1330" <%=iType.Item("it1330")%> />12" Overlap R4508.4</td><td><input type="checkbox" name="it1340" <%=iType.Item("it1340")%> />Access hatches and doors R402.2.4</td><td><input type="checkbox" name="it1350" <%=iType.Item("it1350")%> />Tie Downs</td><td><input type="checkbox" name="it1360" <%=iType.Item("it1360")%> />Permit Posted L&amp;A (Local)</td></tr>
			<tr><td><input type="checkbox" name="it1370" <%=iType.Item("it1370")%> />Purling Blocks R4508.4</td><td><input type="checkbox" name="it1380" <%=iType.Item("it1380")%> />Air Leakage R402.4</td><td><input type="checkbox" name="it1390" <%=iType.Item("it1390")%> />Infiltration Barrier</td><td></td></tr>
			<tr><td colspan="4">&nbsp;</td></tr>
			<tr class="approvalchecks">
				<td><input type="checkbox" name="approved" <%=strApproved%> />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;APPROVED</td>
				<td><input type="checkbox" name="disapproved" <%=strDisapproved%>/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;DISAPPROVED</td>
				<td colspan="2"><input type="checkbox" name="approvedwcorr" <%=strApprovedWCorr%> />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;APPROVED WITH NOTED CORRECTIONS</td>
			</tr>
			<tr class="approvalchecks">
				<td colspan="2"><input type="checkbox" name="coc" <%=strCOC%> />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CERTIFICATE OF COMPLIANCE</td>
			</tr>
		</table-->
		<br />
		<input type="checkbox" name="approved" <%=strApproved%> />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;APPROVED
		<br />
		<input type="checkbox" name="disapproved" <%=strDisapproved%>/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;DISAPPROVED
		<br />
		<input type="checkbox" name="approvedwcorr" <%=strApprovedWCorr%> />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;APPROVED WITH NOTED CORRECTIONS
		<br />
		<input type="checkbox" name="coc" <%=strCOC%> />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CERTIFICATE OF COMPLIANCE
		<br />
		<br />
		Remarks:<br />
		<%if blnPrint then%><%=replace(strRemarks,vbcrlf,"<br />")%><%else%><textarea style="width:100%" rows="10" name="remarks"><%=strRemarks%></textarea><%end if%>
		<br />
		<br />
		<br />
		<br />
		<br />
		<br />
		<table>
			<tr>
				<td colspan="3"><%=strInspectorName%></td>
			</tr>
			<tr>
				<td style="border:0;border-top:1px solid black;">INSPECTOR</td>
				<td style="border:0;">&nbsp;&nbsp;</td>
				<td style="border:0;border-top:1px solid black;">DATE</td>
			</tr>
		</table>
	</p>
	<p>
	<% if not blnPrint then %>

<%					
	tooltipclass=""
	tooltip = ""
	disabled = ""
	If bPermitIsCompleted or bIsOnHold Then		' in permitcommonfunctions.asp
		tooltipclass="tooltip"
		disabled = " disabled "
		tooltip = "<span class=""tooltiptext"">You cannot save because:<br />"
		if bPermitIsCompleted then tooltip = tooltip & "The permit is complete.<br />"
		if bIsOnHoldthen then tooltip = tooltip & "The permit is on hold.<br />"
		tooltip = tooltip & "</span>"
	end if
%>
	<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" id="savebutton" onclick="doValidate();">Save Changes<%=tooltip%></button> &nbsp; &nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> &nbsp; &nbsp;
	<% end if %>
	</p>

	</form>
	<!--END: EDIT FORM-->

	</div>
</div>

<!--END: PAGE CONTENT-->


<%	If request("success") <> "" Then 
		SetupMessagePopUp request("success")
	End If	
%>

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

sub SaveChanges()
	'Need to Insert if new
	strApproved = "0"
	if request.form("approved") = "on" then strApproved = "1"
	strDisapproved = "0"
	if request.form("disapproved") = "on" then strDisapproved = "1"
	strApprovedWCorr = "0"
	if request.form("approvedwcorr") = "on" then strApprovedWCorr = "1"
	strCOC = "0"
	if request.form("coc") = "on" then strCOC = "1"
	
	intPermitInspectionReportID = clng(dbsafe(request.form("permitinspectionreportid")))
	intPermitID = clng(dbsafe(request.form("permitid")))

	if intPermitInspectionReportID = "0" or intPermitInspectionReportID = "" then
		sSQL = "INSERT INTO egov_permitinspectionreports (permitid,inspectiontype,approved,disapproved,approvedwcorr,coc,remarks) " _
			& " VALUES(" _
			& "'" & dbsafe(request.form("permitid")) & "'," _
			& "'" & dbsafe(request.form("inspectiontype")) & "'," _
			& strApproved & "," _
			& strDisapproved & "," _
			& strApprovedWCorr & "," _
			& strCOC & "," _
			& "'" & dbsafe(request.form("remarks")) & "')"
		'response.write sSQL & "<br />"
		intPermitInspectionReportID = RunInsertStatement(sSQL)

		blnNewRecord = true
	else
		'Update if other
		sSQL = "UPDATE egov_permitinspectionreports SET " _
			& " inspectiontype = '" & dbsafe(request.form("inspectiontype")) & "'," _
			& " remarks = '" & dbsafe(request.form("remarks")) & "'," _
			& " approved = " & strApproved & "," _
			& " disapproved = " & strDisapproved & "," _
			& " approvedwcorr = " & strApprovedWCorr & "," _
			& " coc = " & strCOC _
			& " WHERE permitid = '" & intPermitID & "' AND permitinspectionreportid = '" & intPermitInspectionReportID & "'"
		'response.write sSQL & "<br />"
		'response.flush
		RunSQLStatement(sSQL)



		'delete any inspectionreporttypes
		sSQL = "DELETE FROM egov_permitinspectionreporttypes WHERE permitinspectionreportid = '" & intPermitInspectionReportID & "'"
		'response.write sSQL & "<br />"
		RunSQLStatement(sSQL)
	end if

	'loop through fields and insert new inspectionreporttypes
	for each item in request.form
		if instr(item,"it") = 1 then
			sSQL = "INSERT INTO egov_permitinspectionreporttypes (permitinspectionreportid,inspectiontype) VALUES('" & intPermitInspectionReportID & "','" & item & "')"
			'response.write sSQL & "<br />"
			RunSQLStatement(sSQL)
		end if
	next

end sub
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
Sub RunSQLStatement( ByVal sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush
	session("RunSQLStatement") = sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing
	
	session("RunSQLStatement") = ""

End Sub
Function RunInsertStatement( ByVal sInsertStatement )
	Dim sSql, iReturnValue, oInsert

	iReturnValue = 0

'	response.write "<p>" & sInsertStatement & "</p><br /><br />"
'	response.flush
	session("InsertSQL") = sInsertStatement

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSql, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.Close
	Set oInsert = Nothing

	RunInsertStatement = iReturnValue
	session("InsertSQL") = ""

End Function


%>
