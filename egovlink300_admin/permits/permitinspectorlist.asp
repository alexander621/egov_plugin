<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectionlist.asp
' AUTHOR: Steve Loar
' CREATED: 08/08/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit reviews
'
' MODIFICATION HISTORY
' 1.0   08/08/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sFrom, iPageSize, bInitialLoad, sPermitNo, iInspectorUserId, iStatusItemCount, sStatusItem
Dim sParcelIdNumber, sStreetNumber, sStreetName, iPermitId, fromDate, toDate, today, sPermitLocation
Dim iPermitCategoryId

ReDim aInspectionStatuses(0)

sLevel = "../" ' Override of value from common.asp

' Check page availability and user access rights in one call
PageDisplayCheck "edit permit inspection", sLevel	' In common.asp

sSearch = ""
bInitialLoad = False 

If request("pagesize") <> "" Then 
	iPageSize = CLng(request("pagesize"))
Else
	iPageSize = GetUserPageSize( Session("UserId") ) ' In common.asp
	bInitialLoad = True 
End If 


If request("permitno") <> "" Then 
	sPermitNo = Trim(request("permitno"))
	sSearch = sSearch & BuildPermitNoSearch( sPermitNo )	' in permitcommonfunctions.asp
End If 

If request("inspectoruserid") <> "" Then 
	iInspectorUserId = CLng(request("inspectoruserid"))
	If iInspectorUserId > CLng(0) Then 
		sSearch = sSearch & " AND I.inspectoruserid = " & iInspectorUserId
	End If 
Else
	If bInitialLoad Then
		' If they are an inspector, default to their pick
		iInspectorUserId = GetDefaultInspectorUserId( Session("UserId") )
		If iInspectorUserId > CLng(0) Then 
			sSearch = sSearch & " AND I.inspectoruserid = " & iInspectorUserId
		End If
	Else 
		iInspectorUserId = CLng(0)
	End If 
End If 

If CLng(request("inspectionstatusid").count) > CLng(0) Then 
	sSearch = sSearch & " AND ( "
	iStatusItemCount = CLng(0) 
	For Each sStatusItem In request("inspectionstatusid") 
		If iStatusItemCount > UBound(aInspectionStatuses) Then 
			Redim Preserve aInspectionStatuses(iStatusItemCount)
		End If 
		aInspectionStatuses(iStatusItemCount) = sStatusItem
		iStatusItemCount = iStatusItemCount + CLng(1)
		If iStatusItemCount > CLng(1) Then 
			sSearch = sSearch & " OR "
		End If 
		If CLng(sStatusItem) > CLng(0) Then
			sSearch = sSearch & " I.inspectionstatusid = " & CLng(sStatusItem)
		End If 
	Next 
	sSearch = sSearch & " ) "
Else
	' None selected, or first time page displays
	If bInitialLoad Then 
		sSearch = sSearch & " AND T.isneedsinspection = 1 "
	Else 
		sSearch = sSearch & " AND I.inspectionstatusid = 0 "
	End If 
	aInspectionStatuses(0) = 0
End If 

If request("residentstreetnumber") <> "" Then 
	sStreetNumber = request("residentstreetnumber")
	sSearch = sSearch & "AND A.residentstreetnumber = '" & dbsafe(request("residentstreetnumber")) & "' "
End If 
If request("streetname") <> "" And request("streetname") <> "0000" Then 
	sStreetName = request("streetname")
	sSearch = sSearch & " AND (A.residentstreetname = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix + ' ' + A.streetdirection = '" & dbsafe(sStreetName) & "' )"
End If 

If request("parcelidnumber") <> "" Then 
	sParcelIdNumber = request("parcelidnumber")
	sSearch = sSearch & " AND A.parcelidnumber = '" & dbsafe(request("parcelidnumber")) & "' "
End If 

If request("permitlocation") <> "" Then
	sPermitLocation = request("permitlocation")
	sSearch = sSearch & " AND P.permitlocation LIKE '%" & dbsafe(request("permitlocation")) & "%' "
End If 

If request("permitcategoryid") <> "" Then
	iPermitCategoryId = CLng(request("permitcategoryid"))
	If CLng(iPermitCategoryId) > CLng(0) Then 
		sSearch = sSearch & " AND PT.permitcategoryid = " & iPermitCategoryId & " "
	End If 
End If 

' Handle the date range for scheduled dates
fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

' IF EMPTY DEFAULT TO CURRENT TO DATE
If toDate = "" or IsNull(toDate) Then
	toDate = today 
End If

If fromDate = "" or IsNull(fromDate) Then 
	fromDate = CDate( "1/1/2000" ) ' & Year(today) )
End If

sSearch = sSearch & " AND (I.scheduleddate >= '" & fromDate & "' AND I.scheduleddate <= '" & DateAdd("d",1,toDate) & "') "


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/getdates.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script type="text/javascript" src="../scripts/fastinit.js"></script>
	<script language="Javascript" src="../scripts/tablesort2.js"></script>

	<script language="javascript" src="../scripts/tablednd.js"></script>

	<script language="Javascript">
	<!--
  		$( function() {

    			$( ".datepicker" ).datepicker({
      				changeMonth: true,
      				showOn: "both",
      				buttonText: "<i class=\"fa fa-calendar\"></i>",
      				changeYear: true
    			}); 
			$("#inspectorlist tbody").sortable({
    				helper: fixHelperModified,
    				stop: updateIndex,
				axis: "y",
				handle: ".DRAGIMG"
			}).disableSelection();

			$("#inspectorlist td a").click(function(e) {
   				// Do something
   				e.stopPropagation();
			});
  		} );

		var fixHelperModified = function(e, tr) {
    			var $originals = tr.children();
    			var $helper = tr.clone();
    			$helper.children().each(function(index) {
        			$(this).width($originals.eq(index).width())
    			});
    			return $helper;
		},
    		updateIndex = function(e, ui) {
        		$('td.index', ui.item.parent()).each(function (i) {
				var rowno = i+1;
            			$(this).html(rowno);
				var permitinspectionid = $(this).attr("data-id");

				//alert(rowno + ' = ' + permitinspectionid);
				doAjax('changeinspectionroute.asp', 'permitinspectionid=' + permitinspectionid + '&routeorder=' + rowno, '', 'get', '0');
        		});
    		};




		function ChangeInspector( iRowId, iInspectorUserId )
		{
			//alert($("permitinspectionid" + iRowId).value);
			//alert($("inspectoruserid" + iRowId).options[$("inspectoruserid" + iRowId).selectedIndex].value);
			//doAjax('changeinspector.asp', 'permitinspectionid=' + $("permitinspectionid" + iRowId).value + '&inspectoruserid=' + $("inspectoruserid" + iRowId).options[$("inspectoruserid" + iRowId).selectedIndex].value, '', 'get', '0');
			
			var w = (screen.width - 640)/2;
			var h = (screen.height - 480)/2;
			//winHandle = eval('window.open("inspectorpicker.asp?permitinspectionid=' + $("permitinspectionid" + iRowId).value + '&inspectoruserid=' + iInspectorUserId + '&rowid=' + iRowId + '", "_contact", "width=600,height=400,location=1,toolbar=1,statusbar=0,scrollbars=1,menubar=1,left=' + w + ',top=' + h + '")');
			showModal('inspectorpicker.asp?permitinspectionid=' + $("#permitinspectionid" + iRowId).val() + '&inspectoruserid=' + iInspectorUserId + '&rowid=' + iRowId, 'Inspector Selection', 25, 25);
		}

		function Init()
		{
			/*
			var table = $('#inspectorlist');
			var tableDnD = new TableDnD();
			tableDnD.init(table);

			// Redefine the onDrop so that we can update things
			tableDnD.onDrop = function(table, row) 
			{
				var iRowNo = -1;
				var rows = this.table.tBodies[0].rows;
				var debugStr = 'rows now: ';
				for (var i=0; i<rows.length; i++) 
				{
					iRowNo += 1;
					// skip the header row
					if (iRowNo > 0)
					{
						debugStr += iRowNo + ' = ' + rows[i].id + '\n';
						rows[i].cells[1].innerHTML = iRowNo; // change the displayed route order
						rows[i].className = iRowNo & 1? '':'altrow';  // set the row background class
						//alert($("oldrow" + rows[i].id).value + ": " + $("permitinspectionid" + rows[i].id).value + " now " + iRowNo);
						// Fire off ajax routine here to reorder the rows to this order
						doAjax('changeinspectionroute.asp', 'permitinspectionid=' + $("#permitinspectionid" + rows[i].id).val() + '&routeorder=' + iRowNo, '', 'get', '0');
						$("#oldrow" + rows[i].id).val(iRowNo);
					}
				}
			}
			*/

		}

		function PrintList( )
		{
			var w = (screen.width - 680)/2;
			var h = (screen.height - 480)/2;
			//winHandle = eval('window.open("permitinspectorroute.asp", "_details", "width=1100,height=600,location=0,resizable=1,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=' + h + '")');
			showModal('permitinspectorroute.asp', 'Inspector Route', 70, 90);
		}

		function MapList()
		{
			//winHandle = eval('window.open("permitinspectormap.asp", "_details", "width=1100,height=750,location=0,resizable=1,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=100")');
			showModal('permitinspectormap.asp', 'Permit Inspection Map', 70, 90);
		}

	//-->
	</script>

</head>

<body onload="Init();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Permit Inspections</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Search Options</legend>
					<p>
						<form name="frmPermitSearch" method="post" action="permitinspectorlist.asp">
							<input type="hidden" id="pagenum" name="pagenum" value="1" />
							<input type="hidden" name="rq" value="1" />
							<table cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td class="reviewerlistlabel">Category:</td><td><% ShowPermitCategoryPicks iPermitCategoryId %></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Inspector:</td><td><% ShowPermitInspectors iInspectorUserId %></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Permit #:</td><td><input type="text" name="permitno" size="20" maxlength="20" value="<%=sPermitNo%>" /></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Inspection Status:</td><td><% ShowInspectionStatuses aInspectionStatuses, bInitialLoad %></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Address:</td><td><%  DisplayLargeAddressList sStreetNumber, sStreetName %></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Location Like:</td><td><input type="text" name="permitlocation" size="100" maxlength="100" value="<%=sPermitLocation%>" /></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Parcel Id #:</td><td><input type="text" name="parcelidnumber" size="20" maxlength="20" value="<%=sParcelIdNumber%>" /></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Scheduled Date:</td>
									<td>
										From:
										<input type="text" class="datepicker" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />
										&nbsp; To:
										<input type="text" class="datepicker" id="toDate" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />
										&nbsp;
										<%DrawDateChoices "Date" %>
									</td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Records per Page:</td><td><input type="text" name="pagesize" size="10" maxlength="10" value="<%=iPageSize%>" /></td>
								</tr>
								<tr>
			    					<td colspan="2"><input class="button ui-button ui-widget ui-corner-all" type="submit" value="Refresh Results" /></td>
  								</tr>
							</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->
<%				ShowInspections sSearch, sFrom, iPageSize %>			
		</div>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  
	<!--#Include file="modal.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowInspections sSearch, iSearchItem, sYearPick 
'--------------------------------------------------------------------------------------------------
Sub ShowInspections( ByVal sSearch, ByVal sFrom, ByVal iPageSize )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT P.permitid, I.routeorder, I.permitinspectionid, I.permitinspectiontype, I.isreinspection, I.isfinal, ISNULL(I.contact,'') AS contact, ISNULL(I.contactphone,'') AS contactphone, "
	sSql = sSql & " I.scheduleddate, ISNULL(I.scheduledtime,'') AS scheduledtime, ISNULL(I.scheduledampm,'') AS scheduledampm, I.requesteddate, "
	sSql = sSql & " T.inspectionstatus, U.FirstName, U.LastName, ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, I.inspectoruserid, "
	sSql = sSql & " A.residentstreetnumber, ISNULL(A.residentunit,'') AS residentunit, A.residentstreetname, ISNULL(A.streetsuffix,'') AS streetsuffix, ISNULL(A.streetdirection,'') AS streetdirection, "
	sSql = sSql & " ISNULL(A.residentcity,'') AS residentcity, ISNULL(PT.permittype,'') AS permittype, ISNULL(I.schedulingnotes,'') AS schedulingnotes, "
	sSql = sSql & " ISNULL(A.latitude,0.00) AS latitude, ISNULL(A.longitude,0.00) AS longitude, ISNULL(permitlocation,'') AS permitlocation, R.locationtype "
	sSql = sSql & " FROM egov_permits P, egov_permitinspections I, egov_permitstatuses S, egov_inspectionstatuses T, Users U, "
	sSql = sSql & " egov_permitaddress A, egov_permitpermittypes PT, egov_permitlocationrequirements R "
	sSql = sSql & " WHERE (I.scheduleddate IS NOT NULL) AND P.permitid = I.permitid AND P.permitstatusid = S.permitstatusid"
	sSql = sSql & " AND P.permitid = PT.permitid AND P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & " AND P.isonhold = 0 and P.isvoided = 0 AND iscompletedstatus = 0 AND I.inspectionstatusid = T.inspectionstatusid "
	sSql = sSql & " AND I.inspectoruserid = U.userid AND A.permitid = P.permitid AND P.orgid = " & session("orgid") & sSearch
	
	session("sSql") = sSql

	sSql = sSql & " ORDER BY routeorder, I.scheduleddate"

	'response.write sSql & "<br />"
	'response.End 
	

	Set oRs = Server.CreateObject("ADODB.Recordset")
	'oRs.PageSize = iPageSize
	'oRs.CacheSize = iPageSize
	'oRs.CursorLocation = 3
	oRs.Open sSql, Application("DSN"), 3, 1

	If request("pagenum") <> "" Then 
		pagenum = CLng(request("pagenum"))
	Else 
		pagenum = CLng(0)
	End If 

	If (Len(pagenum) = 0 or CLng(pagenum) < CLng(1)) And Not oRs.EOF Then 
		oRs.AbsolutePage = 1
	ElseIf Not oRs.EOF Then 
		iPageCount = CLng(oRs.PageCount)
		If CLng(Request("pagenum")) <= CLng(oRs.PageCount) Then 
			oRs.AbsolutePage = Request("pageNum")
		Else 
			oRs.AbsolutePage = 1
		End If 
	End If 

	Dim abspage, pagecnt
	abspage = oRs.AbsolutePage
	pagecnt = oRs.PageCount
%>
	<div style='font-size:10px; padding-bottom:10px;'>
		<input type="button" class="button ui-button ui-widget ui-corner-all" value="Printable Route" onclick="PrintList();" /> &nbsp; &nbsp;
		<input type="button" class="button ui-button ui-widget ui-corner-all" value="View on Map" onclick="MapList();" />
	</div>

		<table id="inspectorlist" cellpadding="2" cellspacing="0" border="0" class="sortable">
		<thead>
			<tr noDrop="true" noDrag="true"><th class="nosort">&nbsp;</th><th class="number">Route<br />Order</th><th>Permit #</th><th>Permit Type</th><th>Address/Location</th><th>Scheduled<br />Date</th><th class="time">Scheduled<br />Time</th><th>Inspection</th><th>Inspection<br />Status</th><th>Reinspection</th><th>Final</th><th>Inspector</th></tr>
		</thead>
		<tbody>

<%
	'For intRec = 1 To oRs.PageSize
	Do while not oRs.EOF
		'If Not oRs.EOF Then
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr id=""" & iRowCount & """"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			' Drag and Drop Icon
			response.write "<td align=""center""><img src=""..\images\up_down_arrow.gif"" class=""DRAGIMG"" width=""13"" height=""19"" border=""0"" alt=""drag and drop"" />"
			response.write "<input type=""hidden"" id=""permitinspectionid" & iRowCount & """ name=""permitinspectionid" & iRowCount & """ value=""" & oRs("permitinspectionid") & """ />"
			response.write "<input type=""hidden"" id=""oldrow" & iRowCount & """ name=""oldrow" & iRowCount & """ value=""" & iRowCount & """ />"
			response.write "</td>"

			' Route Order
			response.write "<td align=""center"" class=""index"" data-id=""" & oRs("permitinspectionid") & """ title=""click to edit"" onClick=""location.href='permitinspectoredit.asp?permitinspectionid=" & oRs("permitinspectionid") & "';"">" & iRowCount & "</td>"
			
			' Permit No
			response.write "<td nowrap=""nowrap"" title=""click to edit"" onClick=""location.href='permitinspectoredit.asp?permitinspectionid=" & oRs("permitinspectionid") & "';"">"
			response.write GetPermitNumber( oRs("permitid") )
			response.write "</td>"

			' Permit Type
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitinspectoredit.asp?permitinspectionid=" & oRs("permitinspectionid") & "';"">" & oRs("permittype") & "</td>"
			
			' Location
			response.write "<td nowrap=""nowrap"" title=""click to edit"" onClick=""location.href='permitinspectoredit.asp?permitinspectionid=" & oRs("permitinspectionid") & "';"">"
			Select Case oRs("locationtype")

				Case "address"
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
					response.write " " & oRs("residentunit")
					response.write "<br />" & oRs("residentcity")
				Case "location"
					response.write Replace(oRs("permitlocation"),Chr(10),"<br />")

				Case Else 
					response.write "&nbsp;"

			End Select  

			response.write "</td>"
			
			' Scheduled Date
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitinspectoredit.asp?permitinspectionid=" & oRs("permitinspectionid") & "';"">" & FormatDateTime( oRs("scheduleddate"),2 )
			response.write "</td>"

			' Scheduled Time
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitinspectoredit.asp?permitinspectionid=" & oRs("permitinspectionid") & "';"">"
			If oRs("scheduledtime") <> "" Then
				response.write oRs("scheduledtime") & " " & oRs("scheduledampm")
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"
			
			' Inspection Type
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitinspectoredit.asp?permitinspectionid=" & oRs("permitinspectionid") & "';"">" & oRs("permitinspectiontype") & "</td>"
			
			' Status
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitinspectoredit.asp?permitinspectionid=" & oRs("permitinspectionid") & "';"">" & oRs("inspectionstatus") & "</td>"
			
			' Reinspection
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitinspectoredit.asp?permitinspectionid=" & oRs("permitinspectionid") & "';"">" 
			If oRs("isreinspection") Then
				response.write "Yes" 
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"

			' Final Insp
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitinspectoredit.asp?permitinspectionid=" & oRs("permitinspectionid") & "';"">" 
			If oRs("isfinal") Then
				response.write "Yes" 
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"

			' Inspector
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitinspectoredit.asp?permitinspectionid=" & oRs("permitinspectionid") & "';"">" 
			response.write "<span class=""detaildata"" id=""Inspector" & iRowCount & """>"
			response.write oRs("FirstName") & " " & oRs("LastName") 
			response.write "</span>&nbsp;<a href=""javascript:ChangeInspector(" & iRowCount & ", " & oRs("inspectoruserid") & ");"" title=""Click to Change Inspector"">"
			response.write "<i class=""fa fa-pencil""></a>"
			response.write "</td>"

			response.write "</tr>"
			oRs.MoveNext
		'End If 
	loop
	'Next 

	If sSearch <> "" Then 
		If CLng(iRowCount) = CLng(0) Then
			response.write vbcrlf & "<tr><td colspan=""11"">&nbsp;No Permits could be found that match your search criteria.</td></tr>"
		End If 
	Else 
		If CLng(iRowCount) = CLng(0) Then
			response.write vbcrlf & "<tr><td colspan=""11"">&nbsp;No Permits could be found.</td></tr>"
		End If 
	End If 
	%>
		</tbody>
	</table>
		<%

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayLargeAddressList sResidenttype, sStreetNumber, sStreetName, bFound 
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oRs, sCompareName

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & session( "orgid" )
	sSql = sSql & " AND residentstreetname IS NOT NULL "
	sSql = sSql & " ORDER BY sortstreetname "
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If NOT oRs.EOF Then 
		response.write "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write "<select name=""streetname"">" 
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown...</option>"

		Do While Not oRs.EOF
			sCompareName = ""
			If oRs("residentstreetprefix") <> "" Then 
				sCompareName = oRs("residentstreetprefix") & " " 
			End If 

			sCompareName = sCompareName & oRs("residentstreetname")

			If oRs("streetsuffix") <> "" Then 
				sCompareName = sCompareName & " "  & oRs("streetsuffix")
			End If 

			If oRs("streetdirection") <> "" Then 
				sCompareName = sCompareName & " "  & oRs("streetdirection")
			End If 

			response.write vbcrlf & "<option value=""" & sCompareName & """"

			If sStreetName = sCompareName Then 
				response.write " selected=""selected"" "
			End If 

			response.write " >"
			response.write sCompareName & "</option>" 
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetDefaultInspectorUserId( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetDefaultInspectorUserId( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(userid) AS hits FROM users WHERE orgid = " & session("orgid")
	sSql = sSql & " AND ispermitinspector = 1 AND userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			' If they are an inspector, set the default inspector to them
			GetDefaultInspectorUserId = iUserId
		Else
			GetDefaultInspectorUserId = CLng(0)
		End If 
	Else 
		GetDefaultInspectorUserId = CLng(0)
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowPermitInspectors iInspectorUserId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitInspectors( ByVal iInspectorUserId )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname FROM users WHERE orgid = " & session("orgid") & " AND ispermitinspector = 1 "
	sSql = sSql & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""inspectoruserid"">"
		response.write vbcrlf & "<option value=""0"""
		If CLng(iInspectorUserId) = CLng(0) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">All Inspectors</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option "
			If CLng(iInspectorUserId) = CLng(oRs("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " value=""" & oRs("userid") & """>" & oRs("firstname") & " " & oRs("lastname") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowInspectionStatuses aInspectionStatuses, bInitialLoad 
'--------------------------------------------------------------------------------------------------
Sub ShowInspectionStatuses( ByRef aInspectionStatuses, ByVal bInitialLoad )
	Dim sSql, oRs

	sSql = "SELECT inspectionstatusid, inspectionstatus, isinitialstatus FROM egov_inspectionstatuses "
	sSql = sSql & " WHERE isforpermits = 1 AND orgid = " & session("orgid")
	sSql = sSql & " ORDER BY inspectionstatusorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While NOT oRs.EOF
			response.write vbcrlf & "<input type=""checkbox"" name=""inspectionstatusid"" value=""" & oRs("inspectionstatusid") & """"
			If bInitialLoad Then
				If oRs("isinitialstatus") Then 
					response.write " checked=""checked"" "
				End If 
			Else
				For Each Item In aInspectionStatuses
					If CLng(Item) = CLng(oRs("inspectionstatusid")) Then
						response.write " checked=""checked"" "
					End If 
				Next 
			End If 
			response.write " />" & oRs("inspectionstatus") & " &nbsp; "
			oRs.MoveNext
		Loop
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitInspectorPicks iInspectorUserId, iRowId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitInspectorPicks( ByVal iInspectorUserId, ByVal iRowId )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname FROM users WHERE orgid = " & session("orgid") & " AND ispermitinspector = 1 "
	sSql = sSql & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""inspectoruserid" & iRowId & """ name=""inspectoruserid" & iRowId & """ onchange=""ChangeInspector(" & iRowId & ");"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option "
			If CLng(iInspectorUserId) = CLng(oRs("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " value=""" & oRs("userid") & """>" & oRs("firstname") & " " & oRs("lastname") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>
