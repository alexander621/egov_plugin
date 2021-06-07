<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitlist.asp
' AUTHOR: Steve Loar
' CREATED: 03/20/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permits
'
' MODIFICATION HISTORY
' 1.0   03/20/2008	Steve Loar - INITIAL VERSION
' 1.1	11/19/2009	Steve Loar - Changes to include the completed permits in the initial load for Loveland
' 1.2	08/18/2010	Steve Loar - Changes to display of permit type and description.
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

If request("residentstreetnumber") <> "" Then 
	sStreetNumber = request("residentstreetnumber")
	sSearch = sSearch & "AND A.residentstreetnumber = '" & dbsafe(request("residentstreetnumber")) & "' "
	sArchiveSearch = sArchiveSearch & " AND residentstreetnumber = '" & dbsafe(request("residentstreetnumber")) & "' "
End If 
If request("streetname") <> "" And request("streetname") <> "0000" Then 
	sStreetName = request("streetname")
	sSearch = sSearch & " AND (A.residentstreetname = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix + ' ' + A.streetdirection = '" & dbsafe(sStreetName) & "' )"
	sArchiveSearch = sArchiveSearch & " AND jobaddress LIKE '%" & dbsafe(sStreetName) & "%' "
End If 

If request("parcelidnumber") <> "" Then 
	sParcelIdNumber = request("parcelidnumber")
	sSearch = sSearch & " AND A.parcelidnumber = '" & dbsafe(request("parcelidnumber")) & "' "
	sArchiveSearch = sArchiveSearch & " AND parcelidnumber = '" & dbsafe(request("parcelidnumber")) & "' "
End If 

If request("permitno") <> "" Then 
	sPermitNo = Trim(request("permitno"))
	sSearch = sSearch & BuildPermitNoSearch( sPermitNo )
	sArchiveSearch = sArchiveSearch & " AND actualpermitnumber = '" & dbsafe(sPermitNo) & "' "
End If 

If request("listedowner") <> "" Then
	sListedOwner = request("listedowner")
	sSearch = sSearch & " AND A.listedowner LIKE '%" & dbsafe(request("listedowner")) & "%' "
	sArchiveSearch = sArchiveSearch & " AND listedowner LIKE '%" & dbsafe(request("listedowner")) & "%' "
End If 

If request("contactname") <> "" Then 
	sContactname = request("contactname")
	sFrom = ", egov_permitcontacts C "
	sDistinct = " DISTINCT "
	sSearch = sSearch & " AND P.permitid = C.permitid AND (C.company like '%" & dbsafe(request("contactname")) & "%' OR C.firstname like '%" & dbsafe(request("contactname")) & "%' OR C.lastname like '%" & dbsafe(request("contactname")) & "%')"
	sArchiveSearch = sArchiveSearch & " AND (contractorcompany LIKE '%" & dbsafe(request("contactname")) & "%' OR contractorname LIKE '%" & dbsafe(request("contactname")) & "%') "
	iItemCount = CLng(0) 
	For Each item In request("contacttype")
		If iItemCount > CLng(0) Then
			sSearch = sSearch & " OR "
		Else 
			sSearch = sSearch & " AND ( "
		End If 
		iItemCount = iItemCount + 1
		sSearch = sSearch & item & " = 1 "
	Next 
	If iItemCount > CLng(0) Then
		sSearch = sSearch & " ) "
	End If 
End If 

' Handle status picks
If CLng(request("permitstatusid").count) > CLng(0) Then 
	sSearch = sSearch & " AND ( "
	sArchiveSearch = sArchiveSearch & " AND ( "
	iStatusItemCount = CLng(0) 
	bFoundArchiveStatus = False 
	For Each sStatusItem In request("permitstatusid") 
		If iStatusItemCount > UBound(aStatuses) Then 
			Redim Preserve aStatuses(iStatusItemCount)
		End If 
		aStatuses(iStatusItemCount) = sStatusItem
		iStatusItemCount = iStatusItemCount + CLng(1)
		If iStatusItemCount > CLng(1) Then 
			sSearch = sSearch & " OR "
			sArchiveSearch = sArchiveSearch & " OR "
		End If 
		If CLng(sStatusItem) > CLng(0) Then
			sArchiveSearch = sArchiveSearch & " permitstatus = '" & GetPermitStatusByStatusId( sStatusItem ) & "'"
			bFoundArchiveStatus = True 
			sSearch = sSearch & " (P.permitstatusid = " & CLng(sStatusItem) & " AND P.isvoided = 0 AND P.isonhold = 0 AND P.isexpired = 0) "
			' If status is for issued then get any pushed back to issued status as well
			If StatusIsIssued( CLng(sStatusItem) ) Then 
				' Add the other issued to the query
				iOtherStatusId = GetOtherIsIssuedStatus( CLng(sStatusItem) )
				If iOtherStatusId <> CLng(0) Then 
					sSearch = sSearch & " OR (P.permitstatusid = " & CLng(iOtherStatusId) & " AND P.isvoided = 0 AND P.isonhold = 0 AND P.isexpired = 0) "
				End If 
			End If 
		Else
			If CLng(sStatusItem) = CLng(-1) Then
				sSearch = sSearch & " P.isonhold = 1"
				sArchiveSearch = sArchiveSearch & " permitstatus = 'on hold'"
			End If 
			If CLng(sStatusItem) = CLng(-2) Then
				sSearch = sSearch & " P.isvoided = 1"
				sArchiveSearch = sArchiveSearch & " permitstatus = 'void'"
			End If 
			If CLng(sStatusItem) = CLng(-3) Then
				sSearch = sSearch & " P.isexpired = 1"
				sArchiveSearch = sArchiveSearch & " permitstatus = 'expired'"
			End If 
		End If 
	Next 
	sSearch = sSearch & " ) "
	sArchiveSearch = sArchiveSearch & " ) "
Else
	' None selected, or first time page displays
	If bInitialLoad Then 
		' commented out for Loveland 11/19/2009 - Steve Loar
		'sSearch = sSearch & " AND S.iscompletedstatus = 0 AND P.isvoided = 0 AND P.isonhold = 0 AND P.isexpired = 0 "
		sSearch = sSearch & " AND P.isvoided = 0 AND P.isonhold = 0 AND P.isexpired = 0 "
	Else 
		sSearch = sSearch & " AND P.permitstatusid = 0 "
	End If 
	aStatuses(0) = 0
End If 
 
If request("permittypeid") <> "" Then
	iPermitTypeId = CLng(request("permittypeid"))
	If CLng(iPermitTypeId) > CLng(0) Then
		sSearch = sSearch & " AND P.permittypeid = " & iPermitTypeId
		GetPermitTypeFields iPermitTypeId, sPermitType, sPermitTypeDesc 
		sArchiveSearch = sArchiveSearch & " AND permittype = '" & dbsafe(sPermitType) & "' "
		sArchiveSearch = sArchiveSearch & " AND permittypedesc = '" & dbsafe(sPermitTypeDesc) & "' "
	End If 
	
End If 

fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

' IF EMPTY DEFAULT TO CURRENT TO DATE
If toDate = "" or IsNull(toDate) Then
	toDate = today 
End If

If fromDate = "" or IsNull(fromDate) Then 
	fromDate = today
End If

If request("permitdate") <> "none" And request("permitdate") <> "" Then 
	sSearch = sSearch & " AND (P." & request("permitdate") & " >= '" & fromDate & "' AND P." & request("permitdate") & " < '" & DateAdd("d",1,toDate) & "') "
	sArchiveSearch = sArchiveSearch & " AND (" & request("permitdate") & " >= '" & fromDate & "' AND " & request("permitdate") & " < '" & DateAdd("d",1,toDate) & "') "
End If 

If request("toactivitydate") <> "" And request("fromactivitydate") <> "" Then
	sFromActivityDate = request("fromactivitydate")
	sToActivityDate = request("toactivitydate")
	sSearch = sSearch & " AND (P.lastactivitydate >= '" & request("fromactivitydate") & "' AND P.lastactivitydate < '" & DateAdd("d",1,request("toactivitydate")) & "' ) "
	' No Archive last activity date
End If 

If request("invoiceno") <> "" Then
	iInvoiceNo = request("invoiceno")
	If IsNumeric(iInvoiceNo) Then
		iInvoiceNumber = CLng(iInvoiceNo)
	Else
		iInvoiceNumber = 0
	End If
	sFrom = ", egov_permitinvoices V "
	sSearch = sSearch & " AND P.permitid = V.permitid AND V.invoiceid = " & iInvoiceNumber & " "
	' No Archive Invoices
End If 

If request("lasthour") = "true" Then
	sToLastActivityDate = DateAdd( "h", (session("iTimeOffset") +5), Now ) ' The local time stored
	sFromLastActivityDate = DateAdd( "h", -1, sToLastActivityDate ) ' One hour prior to the above time
	'sSearch = " AND (P.lastactivitydate >= '" & sFromLastActivityDate & "' AND P.lastactivitydate < '" & sToLastActivityDate & "' ) "
	sSearch = " AND P.lastactivitydate >= '" & sFromLastActivityDate & "' "
	'response.write sSearch & "<br />"
End If 

If request("legaldescription") <> "" Then
	sLegalDescription = request("legaldescription")
	sSearch = sSearch & " AND A.legaldescription LIKE '%" & dbsafe(request("legaldescription")) & "%' "
	' No archive legal description
End If 

If request("permitlocation") <> "" Then
	sPermitLocation = request("permitlocation")
	sSearch = sSearch & " AND P.permitlocation LIKE '%" & dbsafe(request("permitlocation")) & "%' "
	sArchiveSearch = sArchiveSearch & " AND jobaddress LIKE '%" & dbsafe(request("permitlocation")) & "%' "
End If 

If request("permitcategoryid") <> "" Then
	iPermitCategoryId = CLng(request("permitcategoryid"))
	If CLng(iPermitCategoryId) > CLng(0) Then 
		sSearch = sSearch & " AND T.permitcategoryid = " & iPermitCategoryId & " "
		' No Archive categories
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
    	document.cookie = "pso=" + state + ";" + expires + ";";
}
function toggleOptions()
{
	$("#searchform").toggle();
}
$( function() {
    $( "#accordion" ).accordion({
	    <% if not request.cookies("pso") = "true" then %>
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

		function doCalendar( sField ) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  //eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmPermitSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');

			showModal('calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmPermitSearch', 'Calendar Picker', 25, 35);
		}

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
			if ($("#fromDate").val() != '')
			{
				if (! isValidDate($("#fromDate").val()))
				{
					alert("The From Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#fromDate").focus();
					return;
				}
			}
			// check the to date
			if ($("#toDate").val() != '')
			{
				if (! isValidDate($("#toDate").val()))
				{
					alert("The To Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#toDate").focus();
					return;
				}
			}
			// check the last activity from date
			if ($("#fromactivitydate").val() != '')
			{
				if (! isValidDate($("#fromactivitydate").val()))
				{
					alert("The Last Activity From Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#fromactivitydate").focus();
					return;
				}
			}
			// check the last activity to date
			if ($("#toactivitydate").val() != '')
			{
				if (! isValidDate($("#toactivitydate").val()))
				{
					alert("The Last Activity To Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#toactivitydate").focus();
					return;
				}
			}

			// check that if contractor name is entered that at least one type is selected
			if ($("#contactname").val() != "")
			{
				for (i = 0; i < document.frmPermitSearch.contacttype.length; i++)
				{
					if ( document.frmPermitSearch.contacttype[i].checked == true )
					{
						hasonecontacttype = true; 
						break;
					}
				}
				if ( hasonecontacttype == false )
				{
					alert("To include a contact name in the search, \nplease include at least one contact type selection.");
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
				<strong style="font-size:16px">Permits</strong>
				<%			
				disabled = ""
				tooltipclass=""
				tooltip = ""
				if not UserHasPermission( Session("UserId"), "create building permits" ) then
					tooltipclass="tooltip"
					disabled = " disabled "
					tooltip = "<span class=""tooltiptext"">You don't have permission to create new permits.</span>"
				End If %>
				<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" style="margin-left:35px;" onclick="window.location='newpermit.asp'"><i class="fa fa-plus"></i> New Permit<%=tooltip%></button>
				<br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div id="accordion" onclick="setCookie();">
				   <h3><strong>Search Options</strong></h3>
					<div>
						<form name="frmPermitSearch" method="post" action="permitlist.asp">
							<input type="hidden" id="pagenum" name="pagenum" value="1" />
							<input type="hidden" name="rq" value="1" />
							<table cellpadding="2" cellspacing="0" border="0" class="searchoptions">
								<tr>
									<td>Category:</td><td><%  ShowPermitCategoryPicks iPermitCategoryId %></td>
								</tr>
								<tr>
									<td>Address:</td><td><%  DisplayLargeAddressList sStreetNumber, sStreetName %></td>
								</tr>
								<tr>
									<td>Location Like:</td><td><input type="text" name="permitlocation" size="100" maxlength="100" value="<%=sPermitLocation%>" /></td>
								</tr>
								<tr>
									<td>Legal Desc:</td><td><input type="text" name="legaldescription" size="100" maxlength="100" value="<%=sLegalDescription%>" /></td>
								</tr>
								<tr>
									<td>Parcel Id #:</td><td><input type="text" name="parcelidnumber" size="25" maxlength="25" value="<%=sParcelIdNumber%>" /></td>
								</tr>
								<tr>
									<td>Permit #:</td><td><input type="text" name="permitno" size="20" maxlength="20" value="<%=sPermitNo%>" /></td>
								</tr>
								<tr>
									<td>Listed Owner:</td><td><input type="text" name="listedowner" size="100" maxlength="100" value="<%=sListedOwner%>" /></td>
								</tr>
								<tr>
									<td>Contact:</td><td nowrap="nowrap">
									<input type="text" id="contactname" name="contactname" size="65" maxlength="100" value="<%=sContactname%>" /> &nbsp; &nbsp;
									<input type="button" class="button ui-button ui-widget ui-corner-all" value="Contractor Search" onclick="searchcontacts('contactname');" />
									</td>
								</tr>
								<tr>
									<td>Contact Types:</td>
									<td>
										<input type="checkbox" name="contacttype" value="isapplicant"
<%										For Each item In request("contacttype")
											If item = "isapplicant" Then
												response.write " checked=""checked"" "
											End If 
										Next	%>
										/>Applicant &nbsp; 
										<input type="checkbox" name="contacttype" value="isprimarycontact" 
<%										For Each item In request("contacttype")
											If item = "isprimarycontact" Then
												response.write " checked=""checked"" "
											End If 
										Next	%>
										/>Primary Contact &nbsp; 
										<input type="checkbox" name="contacttype" value="isbillingcontact" 
<%										For Each item In request("contacttype")
											If item = "isbillingcontact" Then
												response.write " checked=""checked"" "
											End If 
										Next	%>
										/>Billing Contact &nbsp; 
										<input type="checkbox" name="contacttype" value="isprimarycontractor" 
<%										For Each item In request("contacttype")
											If item = "isprimarycontractor" Then
												response.write " checked=""checked"" "
											End If 
										Next	%>
										/>Primary Contractor &nbsp; 
										<input type="checkbox" name="contacttype" value="isarchitect" 
<%										For Each item In request("contacttype")
											If item = "isarchitect" Then
												response.write " checked=""checked"" "
											End If 
										Next	%>
										/>Architect/Engineer <br />
										<input type="checkbox" name="contacttype" value="iscontractor" 
<%										For Each item In request("contacttype")
											If item = "iscontractor" Then
												response.write " checked=""checked"" "
											End If 
										Next	%>
										/>Other Contractors
									</td>
								</tr>
								<tr>
									<td>Permit Status:</td><td><% ShowPermitStatuses aStatuses, bInitialLoad %></td>
								</tr>
								<tr>
									<td>Permit Type:</td><td><% ShowPermitTypes iPermitTypeId %></td>
								</tr>
								<tr>
									<td>Date Range:</td>
									<td nowrap="nowrap">
										<select name="permitdate">
											<option value="none"
<%											If request("permitdate") = "none" Then 
												response.write " selected=""selected"" "
											End If	%>
											>Select a Date...</option>

											<option value="applieddate"
<%											If request("permitdate") = "applieddate" Then 
												response.write " selected=""selected"" "
											End If	%>
											>Applied</option>
											<option value="releaseddate"
<%											If request("permitdate") = "releaseddate" Then 
												response.write " selected=""selected"" "
											End If	%>
											>Released</option>
											<option value="approveddate"
<%											If request("permitdate") = "approveddate" Then 
												response.write " selected=""selected"" "
											End If	%>
											>Approved</option>
											<option value="issueddate"
<%											If request("permitdate") = "issueddate" Then 
												response.write " selected=""selected"" "
											End If	%>
											>Issued</option>
											<option value="completeddate"
<%											If request("permitdate") = "completeddate" Then 
												response.write " selected=""selected"" "
											End If	%>
											>Completed</option>
											<option value="expirationdate"
<%											If request("permitdate") = "expirationdate" Then 
												response.write " selected=""selected"" "
											End If	%>
											>Expired</option>
										</select>
										&nbsp; From:
										<input type="text" class="datepicker" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />
										&nbsp; To:
										<input type="text" class="datepicker" id="toDate" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />
										&nbsp;
										<%DrawDateChoices "Date" %>
									</td>
								</tr>
								<tr>
									<td>Last Activity:</td>
									<td nowrap="nowrap">
										From:
										<input type="text" class="datepicker" id="fromactivitydate" name="fromactivitydate" value="<%=sFromActivityDate%>" size="10" maxlength="10" />
										&nbsp; To:
										<input type="text" class="datepicker" id="toactivitydate" name="toactivitydate" value="<%=sToActivityDate%>" size="10" maxlength="10" />
										&nbsp;
										<%DrawDateChoices "activitydate" %>
									</td>
								</tr>
								<tr>
									<td>Invoice #:</td>
									<td><input type="text" name="invoiceno" size="10" maxlength="10" value="<%=iInvoiceNo%>" /></td>
								<tr>
									<td>Records per Page:</td><td><input type="text" name="pagesize" size="10" maxlength="10" value="<%=iPageSize%>" /></td>
								</tr>
								<tr>
			    					<td colspan="2"><input class="button ui-button ui-widget ui-corner-all" type="button" value="Refresh Results" onclick="RefreshResults();" />
										&nbsp;Or&nbsp;
										<input type="button" class="button ui-button ui-widget ui-corner-all" value="Show Only Recently Active Permits" onclick="ShowLastHour();" />
										<input type="hidden" name="lasthour" id="lasthour" value="" />
									</td>
  								</tr>
							</table>
						</form>
					</div>
			</div>
			<!--END: FILTER SELECTION-->

<%				ShowPermits sSearch, sFrom, sDistinct, iPageSize, sArchiveSearch
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
' void ShowPermits sSearch, sFrom, sDistinct, iPageSize 
'--------------------------------------------------------------------------------------------------
Sub ShowPermits( ByVal sSearch, ByVal sFrom, ByVal sDistinct, ByVal iPageSize, ByVal sArchiveSearch )
	Dim sSql, oRs, iRowCount, sClickLink, sLinkTitle, bIsArchive

	iRowCount = 0

	' pull int the live data
	sSql = "SELECT " & sDistinct & " P.permitid, P.permitnumberprefix, P.permitnumberyear, P.permitnumber, P.isonhold, P.isvoided, P.isexpired, ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, "
	sSql = sSql & " permitnumberdisplay = CASE WHEN P.permitnumber IS NULL THEN '' ELSE P.permitnumberyear+P.permitnumberprefix+CAST(P.permitnumber AS varchar) END, "
	sSql = sSql & " P.applieddate, P.releaseddate, P.approveddate, P.issueddate, P.completeddate, P.expirationdate, ISNULL(T.permittype,'') AS permittype, ISNULL(T.permittypedesc,'') AS permittypedesc, S.permitstatus, A.residentstreetnumber, ISNULL(A.residentunit,'') AS residentunit, "
	sSql = sSql & " A.residentstreetname, A.listedowner, ISNULL(A.streetsuffix,'') AS streetsuffix, ISNULL(A.streetdirection,'') AS streetdirection, ISNULL(A.residentcity,'') AS residentcity, S.statusdatedisplayed, S.permitstatusorder, "
	sSql = sSql & " ISNULL(completeddate,ISNULL(issueddate,ISNULL(approveddate,ISNULL(releaseddate,applieddate)))) AS sortdate, permitsort = CASE WHEN P.permitnumber IS NULL THEN '' ELSE 'z' END, "
	sSql = sSql & " ISNULL(A.latitude,0.00) AS latitude, ISNULL(A.longitude,0.00) AS longitude, ISNULL(permitlocation,'') AS permitlocation, R.locationtype, 0 AS isarchive,descriptionofwork "
	sSql = sSql & " FROM egov_permits P, egov_permitpermittypes T, egov_permitstatuses S, egov_permitaddress A, egov_permitlocationrequirements R " & sFrom
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND T.permitid = P.permitid AND P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & " AND P.permitstatusid = S.permitstatusid AND A.permitid = P.permitid " & sSearch 

	' Join to the archived data, if any
	sSql = sSql & " UNION SELECT permitid, permitnumberprefix, permitnumberyear, permitnumber, 0 AS isonhold,0 AS isvoided,0 AS isexpired, "
	sSql = sSql & " ISNULL(residentstreetprefix,'') AS residentstreetprefix, actualpermitnumber AS permitnumberdisplay, "
	sSql = sSql & " applieddate, releaseddate, approveddate, issueddate, completeddate, expirationdate, ISNULL(permittype,'') AS permittype, "
	sSql = sSql & " ISNULL(permittypedesc,'') AS permittypedesc, permitstatus, residentstreetnumber, ISNULL(residentunit,'') AS residentunit, "
	sSql = sSql & " residentstreetname, ISNULL(listedowner,'') AS listedowner, ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection, "
	sSql = sSql & " ISNULL(residentcity,'') AS residentcity, statusdatedisplayed, permitstatusorder, "
	sSql = sSql & " ISNULL(completeddate,ISNULL(issueddate,ISNULL(approveddate,ISNULL(releaseddate,applieddate)))) AS sortdate, "
	sSql = sSql & " permitsort = CASE WHEN permitnumber IS NULL THEN '' ELSE 'z' END, 0.00000000 AS latitude, 0.00000000 AS longitude, "
	sSql = sSql & " '' AS permitlocation, locationtype, 1 AS isarchive, descriptionofwork  "
	sSql = sSql & " FROM egov_permitarchives WHERE orgid = " & session("orgid") & " " & sArchiveSearch

	' order the combined results
	sSql = sSql & " ORDER BY permitsort, permitstatusorder, isonhold DESC, sortdate DESC"

	'response.write sSql & "<br /><br />"
	session("PermitListSql") = sSql

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


     <div class="dropdown right">
  	<button class="ui-button ui-widget ui-corner-all dd-green"><i class="fa fa-bars" aria-hidden="true"></i> Tools</button>
  	<div class="dropdown-content">
		<a href="javascript:MapList()">View on Map</a>
		<a href="permitlistexport.asp">Download to Excel</a>
	</div>
</div>
	</div>


	<div id="constructiontypesshadow" class="shadow">
		<table id="constructiontypes" cellpadding="1" cellspacing="0" border="0" class="sortable tablelist" width="100%">
			<tr valign="bottom" class="tablelist">
				<th>Permit #</th><th>Permit Type</th><th>Address/Location Owner</th>
				<% 'if session("orgid") = "76" then %>
					<th>Description of Work</th>
				<% 'end if %>
				<th>Applicant</th><th>Status</th><th>Status<br />Date</th>
			</tr>

<%
	For intRec = 1 To oRs.PageSize
		If Not oRs.EOF Then
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr id=""" & iRowCount & """ class=""tablelist"
			If iRowCount Mod 2 = 0 Then
				response.write " altrow"
			End If 
			response.write """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

			If oRs("isarchive") Then
				sClickLink = "location.href='viewpermitarchive.asp?permitid=" & oRs("permitid") & "';"
				sLinkTitle = "click to view"
				bIsArchive = True 
				sIcon = " <img src=""..\images\lock.png"" alt=""archive"" border=""0"" height=""9"" width=""9"" />"
			Else
				sClickLink = "location.href='permitedit.asp?permitid=" & oRs("permitid") & "';"
				sLinkTitle = "click to edit"
				bIsArchive = False 
				sIcon = ""
			End If 

			response.write "<td title=""click to edit"" onClick=""" & sClickLink & """ nowrap=""nowrap"" class=""permitnumbercolumn"">"
			If oRs("permitnumberdisplay") = "" Then
				response.write "&nbsp;"
			Else
				If Not bIsArchive Then 
					response.write GetPermitNumber( oRs("permitid") )
				Else
					response.write oRs("permitnumberdisplay")
				End If 
			End If 
			response.write "</td>"
			response.write "<td align=""left"" title=""" & sLinkTitle & """ onClick=""" & sClickLink & """>" & oRs("permittype")
			If oRs("permittype") <> "" And oRs("permittypedesc") <> "" Then 
				response.write " &ndash; "
			End If 
			response.write oRs("permittypedesc") & "</td>"
			response.write "<td title=""" & sLinkTitle & """ onClick=""" & sClickLink & """>"

			Select Case oRs("locationtype")
				Case "address"
					response.write "&nbsp;" & oRs("residentstreetnumber")
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
					response.write "<br />" & "&nbsp;" & oRs("listedowner")

				Case "location"
					response.write "&nbsp;" & Left(oRs("permitlocation"),25)
					If clng(Len(oRs("permitlocation"))) > clng(25) Then
						response.write "..."
					End If 

				Case Else
					response.write "&nbsp;"

			End Select  

			response.write "</td>"
			'if session("orgid") = "76" then
				response.write "<td title=""" & sLinkTitle & """ onClick=""" & sClickLink & """>" & oRs("descriptionofwork") & "</td>"
			'end if 
			
			response.write "<td title=""" & sLinkTitle & """ onClick=""" & sClickLink & """>"
			If Not bIsArchive Then 
				response.write GetPermitApplicantName( oRs("permitid") )
			Else
				response.write GetArchiveContractor( oRs("permitid") )
			End If 
			response.write "</td>"

			If oRs("isonhold") Or oRs("isvoided") Or oRs("isexpired") Then 
				response.write "<td align=""center"" title=""" & sLinkTitle & """ onClick=""" & sClickLink & """>"
				If oRs("isonhold") Then 
					response.write "On Hold"
				Else
					If oRs("isvoided") Then 
						response.write "Voided"
					Else
						response.write "Expired"
					End If 
				End If 
				response.write "</td>"
				response.write "<td align=""center"" title=""" & sLinkTitle & """ onClick=""" & sClickLink & """>"
				If oRs("isexpired") And Not IsNull(oRs("expirationdate")) Then 
					response.write FormatDateTime(oRs("expirationdate"),2)
				Else 
					response.write GetLastLogDate( oRs("permitid") )   ' in permitcommonfunctions.asp
				End If 
				response.write "</td>"
			Else 
				response.write "<td align=""center"" title=""" & sLinkTitle & """ onClick=""" & sClickLink & """ nowrap=""nowrap"">" & oRs("permitstatus") & sIcon & "</td>"
				response.write "<td align=""center"" title=""" & sLinkTitle & """ onClick=""" & sClickLink & """>"
				Select Case oRs("statusdatedisplayed") 
					Case "applieddate"
						response.write FormatDateTime(oRs("applieddate"),2)
					Case "releaseddate"
						response.write FormatDateTime(oRs("releaseddate"),2)
					Case "approveddate"
						response.write FormatDateTime(oRs("approveddate"),2)
					Case "issueddate"
						response.write FormatDateTime(oRs("issueddate"),2)
					Case "completeddate"
						response.write FormatDateTime(oRs("completeddate"),2)
				End Select 
				response.write "</td>"
			End If 
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
Sub ShowPermitStatuses( ByRef aStatuses, ByVal bInitialLoad )
	Dim sSql, oRs

	sSql = "SELECT permitstatusid, permitstatus, iscompletedstatus FROM egov_permitstatuses "
	sSql = sSql & " WHERE isissuedback = 0 AND orgid = " & session("orgid") & " ORDER BY permitstatusorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While NOT oRs.EOF
			response.write vbcrlf & "<input type=""checkbox"" name=""permitstatusid"" value=""" & oRs("permitstatusid") & """"
			If bInitialLoad Then
				' commented out for Loveland 11/19/2009 - Steve Loar
				'If Not oRs("iscompletedstatus") Then 
					response.write " checked=""checked"" "
				'End If 
			Else
				For Each Item In aStatuses
					If CLng(Item) = CLng(oRs("permitstatusid")) Then
						response.write " checked=""checked"" "
					End If 
				Next 
			End If 
			response.write " />" & oRs("permitstatus") & " &nbsp; "
			oRs.MoveNext
		Loop
		response.write vbcrlf & "<input type=""checkbox"" name=""permitstatusid""  value=""-1"""
		For Each Item In aStatuses
			If CLng(Item) = CLng(-1) Then
				response.write " checked=""checked"" "
			End If 
		Next 
		response.write " />Hold &nbsp; "
		response.write vbcrlf & "<input type=""checkbox"" name=""permitstatusid""  value=""-2"""
		For Each Item In aStatuses
			If CLng(Item) = CLng(-2) Then
				response.write " checked=""checked"" "
			End If 
		Next 
		response.write " />Void &nbsp; "
		response.write vbcrlf & "<input type=""checkbox"" name=""permitstatusid""  value=""-3"""
		For Each Item In aStatuses
			If CLng(Item) = CLng(-3) Then
				response.write " checked=""checked"" "
			End If 
		Next 
		response.write " />Expired &nbsp; "

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
