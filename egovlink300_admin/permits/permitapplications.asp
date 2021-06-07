<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitapplications.asp
' AUTHOR: Terry Foster
' CREATED: 08/18/2020
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit applications
'
' MODIFICATION HISTORY
' 1.0   08/18/2020	Terry Foster - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, iSearchItem, iYearPick, sYearPick, sStreetNumber, sStreetName, sListedOwner, sPermitNo
Dim sContactname, iPermitStatusId, sFrom, sDistinct, iPermitTypeId, iStatusItemCount, sParcelIdNumber
Dim iPageSize, iCurrentPage, bInitialLoad, sFromActivityDate, sToActivityDate, iInvoiceNo, sLegalDescription
Dim iPermitCategoryId, sPermitLocation, sArchiveSearch, sPermitType, sPermitTypeDesc, bFoundArchiveStatus
Dim sTempArchiveSearch, iInvoiceNumber

ReDim aStatuses(0)

sLevel = "../" ' Override of value from common.asp
sSearch = ""
sFrom = ""
sArchiveSearch = ""
sDistinct = ""
bInitialLoad = False 

PageDisplayCheck "edit permits", sLevel	' In common.asp

If request("pagesize") <> "" Then 
	iPageSize = CLng(request("pagesize"))
Else
	iPageSize = GetUserPageSize( Session("UserId") ) ' In common.asp
	bInitialLoad = True 
End If 

' Handle status picks
If CLng(request("statusid").count) > CLng(0) Then 
	sSearch = sSearch & " AND ( "
	iStatusItemCount = CLng(0) 
	For Each sStatusItem In request("statusid") 
		If iStatusItemCount > UBound(aStatuses) Then 
			Redim Preserve aStatuses(iStatusItemCount)
		End If 
		aStatuses(iStatusItemCount) = sStatusItem
		iStatusItemCount = iStatusItemCount + CLng(1)
		If iStatusItemCount > CLng(1) Then 
			sSearch = sSearch & " OR "
		End If 
		If CLng(sStatusItem) > CLng(0) Then
			sSearch = sSearch & " workflowid = " & CLng(sStatusItem) & ""
		End If 
	Next 
	sSearch = sSearch & " ) "
Else
	sSearch = sSearch & " AND workflowid = 1 "

	aStatuses(0) = 0
End If 

If request("todate") <> "" And request("fromdate") <> "" Then
	sFromDate = request("fromdate")
	sToDate = request("todate")
	sSearch = sSearch & " AND (submitteddate >= '" & request("fromdate") & "' AND submitteddate < '" & DateAdd("d",1,request("todate")) & "' ) "
	' No Archive last activity date
End If 

If request("permittypeid") <> "" Then
	iPermitTypeId = CLng(request("permittypeid"))
	If CLng(iPermitTypeId) > CLng(0) Then
		sSearch = sSearch & " AND pa.permittypeid = " & iPermitTypeId
	End If 
	
End If 
%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<style>


	</style>


	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/getdates.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--
	
function setCookie()
{
	var state = $("#accordion h3").hasClass("ui-state-active");
	var d = new Date();
	var days = 7;
    	d.setTime(d.getTime() + (days*24*60*60*1000));
    	var expires = "expires="+ d.toUTCString();
    	document.cookie = "paso=" + state + ";" + expires + ";";
}
function toggleOptions()
{
	$("#searchform").toggle();
}
$( function() {
    $( "#accordion" ).accordion({
	    <% if not request.cookies("paso") = "true" then %>
	    active:false,
	    <% end if %>
      collapsible: true
    });
    $( "#accordion2" ).accordion({
	    active:false,
      collapsible: true
    });
    $( "#defaultsearchaccord" ).accordion({
	    active:false,
      collapsible: true
    });
  } );


		function GoToPage( iPageNum )
		{
			$("#pagenum").val(iPageNum);
			document.frmPermitSearch.submit();
		}

		function RefreshResults()
		{
			var i;
			var hasonecontacttype = false;

			// check the from date
			if ($("#fromdate").val() != '')
			{
				if (! isValidDate($("#fromdate").val()))
				{
					alert("The From Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#fromdate").focus();
					return;
				}
			}
			// check the to date
			if ($("#todate").val() != '')
			{
				if (! isValidDate($("#todate").val()))
				{
					alert("The To Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#todate").focus();
					return;
				}
			}
			// check the last activity from date
			if ($("#fromdate").val() != '')
			{
				if (! isValidDate($("#fromdate").val()))
				{
					alert("The Last Activity From Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#fromdate").focus();
					return;
				}
			}
			// check the last activity to date
			if ($("#todate").val() != '')
			{
				if (! isValidDate($("#todate").val()))
				{
					alert("The Last Activity To Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#todate").focus();
					return;
				}
			}


			// Dates are OK, so post the search
			document.frmPermitSearch.submit();
		}

		function searchcontacts( sFieldId )
		{
			//var w = (screen.width - 640)/2;
			//var h = (screen.height - 480)/2;
			//var winHandle = eval('window.open("contractorpicker.asp?fieldid=' + sFieldId + '", "_contact", "width=600,height=400,location=1,toolbar=1,statusbar=0,scrollbars=1,menubar=1,left=' + w + ',top=' + h + '")');
			showModal('contractorpicker.asp?fieldid=' + sFieldId, 'Contractor Search', 20, 30);
		}

		function ShowLastHour()
		{
			$("#lasthour").val("true");
			document.frmPermitSearch.submit();
		}

		function MapList()
		{
			//var winHandle = eval('window.open("permitlistmap.asp", "_details", "width=1100,height=750,location=0,resizable=1,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=100")');
			showModal('permitlistmap.asp', 'Permit Locations Map', 65, 95);
		}

	//-->
  $( function() {
    $( ".datepicker" ).datepicker({
      changeMonth: true,
      showOn: "both",
      buttonText: "<i class=\"fa fa-calendar\"></i>",
      changeYear: true
    });
  } );
	</script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="notcentercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<strong style="font-size:16px">Permit Applications</strong>
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div id="accordion" onclick="setCookie();">
				   <h3><strong>Search Options</strong></h3>
					<div>
						<form name="frmPermitSearch" method="post" action="permitapplications.asp">
							<input type="hidden" id="pagenum" name="pagenum" value="1" />
							<input type="hidden" name="rq" value="1" />
							<input type="hidden" name="pagesize" value="20" />
							<table cellpadding="2" cellspacing="0" border="0" class="searchoptions">
								<tr>
									<td>Permit Application Type:</td><td><% ShowPermitTypes iPermitTypeId %></td>
								</tr>
								<tr>
									<td>Application Status:</td><td><% ShowPermitApplicationStatuses aStatuses, bInitialLoad %></td>
								</tr>
								<tr>
									<td>Submitted:</td>
									<td nowrap="nowrap">
										From:
										<input type="text" class="datepicker" id="fromdate" name="fromdate" value="<%=sFromDate%>" size="10" maxlength="10" />
										&nbsp; To:
										<input type="text" class="datepicker" id="todate" name="todate" value="<%=sToDate%>" size="10" maxlength="10" />
										&nbsp;
										<%DrawDateChoices "activitydate" %>
									</td>
								</tr>

								<tr>
			    					<td colspan="2"><input class="button ui-button ui-widget ui-corner-all" type="button" value="Refresh Results" onclick="RefreshResults();" /></td>
  								</tr>
							</table>
						</form>
					</div>
			</div>
			<br />
			<!--END: FILTER SELECTION-->

<%				ShowPermitApplications sSearch, sFrom, sDistinct, iPageSize, sArchiveSearch
%>			
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
' void ShowPermitApplications sSearch, sFrom, sDistinct, iPageSize 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitApplications( ByVal sSearch, ByVal sFrom, ByVal sDistinct, ByVal iPageSize, ByVal sArchiveSearch )
	Dim sSql, oRs, iRowCount, sClickLink, sLinkTitle, bIsArchive

	iRowCount = 0

	sSQL = "SELECT pa.*,pt.permittype + ' - ' + pt.permittypedesc AS permittype FROM egov_permitapplication_submitted pa " _
		& " INNER JOIN egov_permittypes pt ON pa.permittypeid = pt.permittypeid" _
		& " WHERE pa.orgid = " & session("orgid") & sSearch _
		& " ORDER BY pa.submitteddate DESC "

	'response.write sSql & "<br /><br />"
	session("PermitApplicationListSql") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.PageSize = iPageSize
	oRs.CacheSize = iPageSize
	oRs.CursorLocation = 3
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
	<div style=''>
     		<input type="button"<% if abspage <= 1 then response.write " disabled "%> name="prevRecordsButton" id="prevRecordsButton" value="<< Back" class="button ui-button ui-widget ui-corner-all" onclick="GoToPage(<%=abspage-1%>);"  />
        	<input type="button"<% if abspage >= pagecnt then response.write " disabled "%> name="nextRecordsButton" id="nextRecordsButton" value="Next >>" class="button ui-button ui-widget ui-corner-all" onclick="GoToPage(<%=abspage+1%>);"  />&nbsp;&nbsp;


	</div>


	<div id="constructiontypesshadow" class="shadow">
		<table id="constructiontypes" cellpadding="1" cellspacing="0" border="0" class="sortable tablelist" width="100%">
			<tr valign="bottom" class="tablelist"><th>Application #</th><th>Permit Type</th><th>Submitted<br />Date</th></tr>

<%
	For intRec = 1 To oRs.PageSize
		If Not oRs.EOF Then
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr id=""" & iRowCount & """ class=""tablelist"
			If iRowCount Mod 2 = 0 Then
				response.write " altrow"
			End If 
			response.write """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" "
			response.write " onclick=""location.href='permitapplication.asp?paid=" & oRs("permitapplication_submittedid") & "'"""
			response.write ">"
			response.write "<td align=""center"">" & oRs("permitapplication_submittedid")  & "</td>"
			response.write "<td align=""center"">" & oRs("permittype")  & "</td>"
			response.write "<td align=""center"">" & oRs("submitteddate")  & "</td>"

			response.write "</tr>"
			oRs.MoveNext
		End If 
	Next 

	If sSearch <> "" Then 
		If CLng(iRowCount) = CLng(0) Then
			response.write vbcrlf & "<tr><td colspan=""6"">&nbsp;No Permits could be found that match your search criteria.</td></tr>"
		End If 
	Else 
		If CLng(iRowCount) = CLng(0) Then
			response.write vbcrlf & "<tr><td colspan=""6"">&nbsp;No Permits could be found.</td></tr>"
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 
	%>
				</table>
     		<input type="button"<% if abspage <= 1 then response.write " disabled "%> name="prevRecordsButton" id="prevRecordsButton" value="<< Back" class="button ui-button ui-widget ui-corner-all" onclick="GoToPage(<%=abspage-1%>);"  />
        	<input type="button"<% if abspage >= pagecnt then response.write " disabled "%> name="nextRecordsButton" id="nextRecordsButton" value="Next >>" class="button ui-button ui-widget ui-corner-all" onclick="GoToPage(<%=abspage+1%>);"  />&nbsp;&nbsp;
	<%

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayLargeAddressList sStreetNumber, sStreetName 
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oRs, sCompareName, sOldCompareName

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & session( "orgid" )
	sSql = sSql & " AND residentstreetname IS NOT NULL "
	sSql = sSql & " ORDER BY sortstreetname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write "<select name=""streetname"">" 
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown...</option>"
		sOldCompareName = "qwerty"

		Do While Not oRs.EOF
			sCompareName = ""
			If oRs("residentstreetprefix") <> "" Then 
				sCompareName = UCase(oRs("residentstreetprefix")) & " " 
			End If 

			sCompareName = sCompareName & UCase(oRs("residentstreetname"))

			If oRs("streetsuffix") <> "" Then 
				sCompareName = sCompareName & " "  & UCase(oRs("streetsuffix"))
			End If 

			If oRs("streetdirection") <> "" Then 
				sCompareName = sCompareName & " "  & UCase(oRs("streetdirection"))
			End If 

			If sOldCompareName <> sCompareName Then 
				' only write out unique values
				sOldCompareName = sCompareName
				response.write vbcrlf & "<option value=""" & sCompareName & """"

				If sStreetName = sCompareName Then 
					response.write " selected=""selected"" "
				End If 

				response.write " >"
				response.write sCompareName & "</option>" 
			End If 

			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitStatuses aStatuses, bInitialLoad 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitApplicationStatuses( ByRef aStatuses, ByVal bInitialLoad )
	Dim sSql, oRs

	sSql = "SELECT permitapplication_workflowid as statusid, workflowstep FROM egov_permitapplication_workflowsteps  ORDER BY workfloworder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While NOT oRs.EOF
			response.write vbcrlf & "<input type=""checkbox"" name=""statusid"" value=""" & oRs("statusid") & """"
			If bInitialLoad and oRs("statusid") = "1" Then
				response.write " checked=""checked"" "
			Else
			For Each Item In aStatuses
				If CLng(Item) = CLng(oRs("statusid")) Then
					response.write " checked=""checked"" "
				End If 
			Next 
			End If 
			response.write " />" & oRs("workflowstep") & " &nbsp; "
			oRs.MoveNext
		Loop

	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitTypes iPermitTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitTypes( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT permittypeid, permittype, permittypedesc FROM egov_permittypes "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY permittype, permittypedesc"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<select name=""permittypeid"">"
		response.write vbcrlf & "<option value=""0"""
		If CLng(iPermitTypeId) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">All Permit Types</option>"
		Do While NOT oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("permittypeid") & """"
			If CLng(iPermitTypeId) = CLng(oRs("permittypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permittype") & " &ndash; " & oRs("permittypedesc") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetOtherIsIssuedStatus( sStatusItem )
'--------------------------------------------------------------------------------------------------
Function GetOtherIsIssuedStatus( ByVal sStatusItem )
	Dim sSql, oRs

	sSql = "SELECT permitstatusid FROM egov_permitstatuses "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND isissued = 1 AND permitstatusid != " & sStatusItem

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetOtherIsIssuedStatus = CLng(oRs("permitstatusid"))
	Else
		GetOtherIsIssuedStatus = CLng(0) 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean StatusIsIssued( sStatusItem )
'--------------------------------------------------------------------------------------------------
Function StatusIsIssued( ByVal sStatusItem )
	Dim sSql, oRs

	sSql = "SELECT isissued FROM egov_permitstatuses "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND permitstatusid = " & sStatusItem

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("isissued") Then 
			StatusIsIssued = True 
		Else
			StatusIsIssued = False 
		End If 
	Else
		StatusIsIssued = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void GetPermitTypeFields iPermitTypeId, sPermitType, sPermitTypeDesc 
'--------------------------------------------------------------------------------------------------
Sub GetPermitTypeFields( ByVal iPermitTypeId, ByRef sPermitType, ByRef sPermitTypeDesc )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(permittype,'') AS permittype, ISNULL(permittypedesc,'') AS permittypedesc "
	sSql = sSql & "FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitType = Trim(oRs("permittype"))
		sPermitTypeDesc = Trim(oRs("permittypedesc"))
	Else
		sPermitType = ""
		sPermitTypeDesc = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub




%>
