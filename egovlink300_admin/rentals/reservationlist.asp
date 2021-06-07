<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationlist.asp
' AUTHOR: Steve Loar
' CREATED: 08/13/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of rentals. From here you can create or edit rentals
'
' MODIFICATION HISTORY
' 1.0   08/13/2009	Steve Loar - INITIAL VERSION
' 1.1	05/10/2010	Steve Loar - Added receipt number to search
' 1.2	05/11/2010	Steve Loar - Modified mechanics of date selections
' 1.3	03/24/2011	Steve Loar - hide deactivated rentals
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, iRentalId, sRenterName, iReservationTypeId, sFrom, sReservationTypeSelector, iReservedStatus
Dim sStatusSearch, sReservationId, iSupervisorUserId, sPaymentId, sLocation

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "edit reservations", sLevel	' In common.asp

sSearch = ""
sSelect = ""
sFrom = ""

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

sSearch = sSearch & " AND D.reservationstarttime BETWEEN '" & CDate(fromDate) & " 0:00 AM' AND '" & DateAdd("d",1,CDate(toDate)) &" 0:00 AM' "

If request("rentalid") <> "" Then
	iRentalId = request("rentalid")
	If iRentalId <> "0" Then 
		sRentalIdType = Left(iRentalId, 1)
		iActualId = Mid(iRentalId, 2)
		If sRentalIdType = "R" Then
			sSearch = sSearch & " AND D.rentalid = " & iActualId
		Else
			sSearch = sSearch & " AND L.locationid = " & iActualId
		End If 
	End If 
Else
	iRentalId = "0"	  ' The all selection
End If 

If request("reservationtypeid") <> "" Then
	iReservationTypeId = CLng(request("reservationtypeid"))
	If iReservationTypeId <> CLng(0) Then 
		sSearch = sSearch & " AND R.reservationtypeid = " & iReservationTypeId
		sReservationTypeSelector = GetReservationTypeSelection( iReservationTypeId )  ' in rentalscommonfunctions.asp
	Else
		sReservationTypeSelector = ""
	End If 
Else
	iReservationTypeId = "0"
	sReservationTypeSelector = ""
End If 

If request("rentername") <> "" Then
	sRenterName = request("rentername")
Else
	sRenterName = ""
End If 

If request("location") <> "" Then  
	sLocation = request("location")
	sSearch = sSearch & " AND ( RE.rentalname LIKE '%" & dbsafe(sLocation) & "%' OR L.name LIKE '%" & dbsafe(sLocation) & "%' ) "
Else
	sLocation = ""
End If 

' Set up the choice of what statuses to show
If request("reservationstatus") = "" Then
	iReservedStatus = 1
Else
	iReservedStatus = request("reservationstatus")
End If 

Select Case iReservedStatus
	Case 1
		sStatusSearch = " AND DS.iscancelled = 0 AND RS.iscancelled = 0 AND R.isonhold = 0 "
	Case 2
		sStatusSearch = " "
	Case 3
		sStatusSearch = " AND DS.iscancelled = 1 "
	Case 4
		sStatusSearch = " AND DS.iscancelled = 0 AND RS.iscancelled = 0 AND R.isonhold = 1 "
End Select 

If request("reservationid") <> "" Then
	sReservationId = request("reservationid")
	sSearch = sSearch & " AND R.reservationid = " & sReservationId
Else
	sReservationId = ""
End If 


If request("supervisoruserid") <> "" Then
	iSupervisorUserId = CLng(request("supervisoruserid"))
	If iSupervisorUserId > CLng(0) Then
		sSearch = sSearch & " AND RE.supervisoruserid = " & iSupervisorUserId
	End If 
Else
	iSupervisorUserId = 0
End If 

If request("paymentid") <> "" Then
	sPaymentId = CLng(request("paymentid"))
	sSearch = sSearch & " AND J.reservationid = R.reservationid AND J.paymentid = " & sPaymentId
	sFrom = sFrom & ", egov_class_payment J "
End If 


%>

<html lang="en">
<head>
	<meta charset="UTF-8">
	
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="rentalsstyles.css" />

	<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

	<script src="https://code.jquery.com/jquery-1.9.1.js"></script>
  	<script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
  	
	<script src="../scripts/modules.js"></script>
	<script src="../scripts/getdates.js"></script>
	<script src="../scripts/formvalidation_msgdisplay.js"></script>
	<script src="../scripts/isvaliddate.js"></script>

	<script>
	<!--

		// var doCalendar = function( sField ) {
		// 	w = (screen.width - 350)/2;
		// 	h = (screen.height - 350)/2;
		// 	var sSelectedDate = $(sField).value;

		// 	// Set the to date to the from date
		// 	if (sField == "toDate")
		// 	{
		// 		sSelectedDate = $("fromDate").value;
		// 		//alert( $("fromDate").value );
		// 	}

		// 	//alert( sField + ": " + sSelectedDate );

		// 	eval('window.open("calendarpicker.asp?date=' + sSelectedDate + '&p=1&updatefield=' + sField + '&updateform=frmReservationSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		// }

		var RefreshResults = function() {

			// check the from date
			if ($("#fromDate").val() == "")
			{
				alert("Please enter a From Date");
				$("#fromDate").focus();
				return;
			}
			else
			{
				if (! isValidDate($("#fromDate").val()))
				{
					alert("The From date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#fromDate").focus();
					return;
				}
			}

			// check the to date
			if ($("#toDate").val() == "")
			{
				alert("Please enter a To Date");
				$("#toDate").focus();
				return;
			}
			else
			{
				if (! isValidDate($("#toDate").val()))
				{
					alert("The To date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#toDate").focus();
					return;
				}
			}

			// check that any reservationid is numeric
			if ($("#reservationid").val() != '')
			{
				var rege = /^\d*$/
				var Ok = rege.exec($("#reservationid").val());
				if ( ! Ok )
				{
					$("reservationid").focus();
					inlineMsg($("#reservationid").id,'<strong>Invalid Value: </strong>Reservation Ids can only be numeric.',10,$("#reservationid").id);
					return false;
				}
			}
			// Check that the receipt number is numeric
			if ($("#paymentid").val() != '')
			{
				var rege = /^\d*$/
				var Ok = rege.exec($("#paymentid").val());
				if ( ! Ok )
				{
					$("#paymentid").focus();
					inlineMsg($("#paymentid").id,'<strong>Invalid Value: </strong>Receipt Numbers can only be numeric.',10,$("#paymentid").id);
					return false;
				}
			}
			document.frmReservationSearch.action="reservationlist.asp";
			document.frmReservationSearch.submit();
		};

		var ExportToExcel = function() {
			// check that any reservationid is numeric
			if ($("#reservationid").val() != '')
			{
				var rege = /^\d*$/
				var Ok = rege.exec($("#reservationid").val());
				if ( ! Ok )
				{
					$("#reservationid").focus();
					inlineMsg($("#reservationid").id,'<strong>Invalid Value: </strong>Reservation Ids can only be numeric.',10,$("#reservationid").id);
					return false;
				}
			}
			// Check that the receipt number is numeric
			if ($("#paymentid").val() != '')
			{
				var rege = /^\d*$/
				var Ok = rege.exec($("#paymentid").val());
				if ( ! Ok )
				{
					$("#paymentid").focus();
					inlineMsg($("#paymentid").id,'<strong>Invalid Value: </strong>Receipt Numbers can only be numeric.',10,$("#paymentid").id);
					return false;
				}
			}
			document.frmReservationSearch.action="reservationlistexport.asp";
			document.frmReservationSearch.submit();
		}

		var OverpaymentExport = function() {
			// check that any reservationid is numeric
			if ($("#reservationid").val() != '')
			{
				var rege = /^\d*$/
				var Ok = rege.exec($("#reservationid").val());
				if ( ! Ok )
				{
					$("#reservationid").focus();
					inlineMsg($("#reservationid").id,'<strong>Invalid Value: </strong>Reservation Ids can only be numeric.',10,$("#reservationid").id);
					return false;
				}
			}
			// Check that the receipt number is numeric
			if ($("#paymentid").val() != '')
			{
				var rege = /^\d*$/
				var Ok = rege.exec($("#paymentid").val());
				if ( ! Ok )
				{
					$("#paymentid").focus();
					inlineMsg($("#paymentid").id,'<strong>Invalid Value: </strong>Receipt Numbers can only be numeric.',10,$("#paymentid").id);
					return false;
				}
			}
			document.frmReservationSearch.action="overpaymentexport.asp";
			document.frmReservationSearch.submit();
		};
		
		
		// these function set up the date pickers
		$(function() {
			$( "#toDate" ).datepicker({
				showOn: "button",
				buttonImage: "../images/calendar.gif",
				buttonImageOnly: true,
				changeMonth: true,
				changeYear: true
			});
		});

		$(function() {
			$( "#fromDate" ).datepicker({
				showOn: "button",
				buttonImage: "../images/calendar.gif",
				buttonImageOnly: true,
				changeMonth: true,
				changeYear: true
			});
		});

	//-->
	</script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Reservations</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Search Options</legend>
					<p>
						<form name="frmReservationSearch" method="post" action="reservationlist.asp">
							<table cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td>Date Range:</td>
									<td>
										From:
										<input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />
										&nbsp; To:
										<input type="text" id="toDate" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />
										&nbsp;
										<%DrawDateChoices "Date" %>
									</td>
								</tr>
								<tr>
									<td>Reservation Type:</td>
									<td>
<%										ShowReservationTypePicks iReservationTypeId		' In rentalsguifunctions.asp	%>
									</td>
								</tr>
								<tr>
									<td>Rental:</td><td><% ShowRentalLocationPicks iRentalId, True, True 	' In rentalsguifunctions.asp %></td>
								</tr>
								<tr>
									<td>Location Like:</td><td><input type="text" name="location" size="50" maxlength="50" value="<%=sLocation%>" /></td>
								</tr>
								<tr>
									<td>Supervisor:</td><td><% ShowRentalSupervisors iSupervisorUserId, "All Supervisors" 	' In rentalsguifunctions.asp %></td>
								</tr>
								<tr>
									<td>Renter Name Like:</td><td><input type="text" name="rentername" size="50" maxlength="50" value="<%=sRenterName%>" /></td>
								</tr>
								<tr>
									<td>Status:</td>
									<td>
										<select name="reservationstatus">
											<option value="1"<%	If iReservedStatus = 1 Then 
																	response.write " selected=""selected"" "
																End If	%>
																>Reserved Only</option>
											<option value="2"<%	If iReservedStatus = 2 Then 
																	response.write " selected=""selected"" "
																End if	%>
																>Reserved,  Cancelled and On Hold</option>
											<option value="3"<%	If iReservedStatus = 3 Then 
																	response.write " selected=""selected"" "
																End if	%>
																>Cancelled Only</option>
											<option value="4"<%	If iReservedStatus = 4 Then 
																	response.write " selected=""selected"" "
																End if	%>
																>On Hold Only</option>
										</select>
									</td>
								</tr>
								<tr>
									<td>Reservation Id:</td><td><input type="text" id="reservationid" name="reservationid" size="8" maxlength="8" value="<%=sReservationId%>" /></td>
								</tr>
								<tr>
									<td>Receipt #:</td><td><input type="text" id="paymentid" name="paymentid" size="8" maxlength="8" value="<%=sPaymentId%>" /></td>
								</tr>
								<tr>
			    					<td colspan="2">
										<div id="reservationlistbuttons">
											<input class="button" type="button" value="Refresh Results" onclick="RefreshResults();" />&nbsp;&nbsp;
											<input type="button" class="button" value="Export to Excel" onclick="ExportToExcel();" />&nbsp;&nbsp;
											<input type="button" class="button" value="Overpayment Export" onclick="OverpaymentExport();" />
										</div>
									</td>
  								</tr>
							</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->

<%				ShowReservations sSearch, sFrom, sRenterName, sStatusSearch
%>			

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' void ShowReservations sSearch, sFrom, sRenterName, sStatusSearch
'--------------------------------------------------------------------------------------------------
Sub ShowReservations( ByVal sSearch, ByVal sFrom, ByVal sRenterName, ByVal sStatusSearch )
	Dim sSql, oRs, sClassName, sRenter, sRenterSearch

	'response.write sSearch & "<br /><br />"

	sSql = "SELECT D.reservationid, D.reservationstarttime, D.billingendtime, RE.rentalname, L.name AS locationname, ISNULL(R.timeid,0) AS timeid, "
	sSql = sSql & " (R.totalamount + R.totalrefunded - R.totalpaid + R.totalrefundfees) AS balancedue, R.totalamount, R.totalrefunded, R.isonhold, DS.iscancelled, "
	sSql = sSql & " RT.reservationtype, RT.reservationtypeselector, DS.reservationstatus, R.reserveddate, R.rentaluserid" & sSelect
	sSql = sSql & " FROM egov_rentalreservationdates D, egov_rentalreservations R, egov_rentals RE, egov_class_location L, "
	sSql = sSql & " egov_rentalreservationtypes RT, egov_rentalreservationstatuses DS, egov_rentalreservationstatuses RS" & sFrom
	sSql = sSql & " WHERE D.reservationid = R.reservationid AND D.rentalid = RE.rentalid AND RE.locationid = L.locationid "
	sSql = sSql & " AND R.reservationtypeid = RT.reservationtypeid AND D.statusid = DS.reservationstatusid AND R.orgid = " & session("orgid")
	sSql = sSql & " AND RE.isdeactivated = 0 AND R.reservationstatusid = RS.reservationstatusid" & sStatusSearch & sSearch
	sSql = sSql & " ORDER BY D.reservationstarttime, L.name, RE.rentalname"
	'if session("orgid") = "228" then response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<table id=""reservationlist"" cellpadding=""1"" cellspacing=""0"" border=""0"" class=""sortable"">"
	response.write vbcrlf & "<tr><th>Id</th><th>Date</th><th>Time</th><th>Location</th><th>Type</th><th>Renter</th><th>Status</th><th>Reserved</th></tr>"

	Do While Not oRs.EOF
		If oRs("reservationtypeselector") = "public" Then
			sRenter = GetCitizenName( oRs("rentaluserid") )
			sRenterSearch = sRenter
		Else
			If oRs("reservationtypeselector") = "admin" Then
				sRenter = GetAdminName( oRs("rentaluserid") )
				 sRenterSearch = sRenter
			Else 
				If oRs("reservationtypeselector") = "class" Then
					sRenter = GetActivityNo( oRs("timeid") )	' In rentalscommonfunctions.asp 
					sClassName = GetClassName( oRs("timeid") )	' In rentalscommonfunctions.asp 
					sRenterSearch = sRenter & " " & sClassName
					If Len(sClassName) > 20 Then
						sClassName = Left(sClassName,17) & "..."
					End If 
					sRenter = sRenter & "<br />" & sClassName	
				Else
					' this leaves blocked
					sRenter = "&nbsp;"
					sRenterSearch = ""
				End If 
			End If 
		End If 

		' if the rentername is not empty then they are looking for a match of some sort
		If sRenterName <> "" Then 
			If oRs("reservationtypeselector") <> "block" Then
				' they are looking for someone. 0 is not found
				If InStr(LCase(sRenterSearch), LCase(sRenterName)) > 0 Then
					' the searched for name is in the name
					bOK = True 
				Else
					' the searched for name is not in the name
					bOK = False 
				End If 
			Else
				' we leave blocks out of name searches 
				bOk = False  
			End If 
		Else
			' No name search so the record is OK
			bOk = True 
		End If 

		If bOk Then 
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr id=""" & iRowCount & """"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

			' Reservation #
			response.write "<td class=""firstcol"" align=""left"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
			response.write oRs("reservationid")
			response.write "</td>"

			' Date 
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
			response.write DateValue(CDate(oRs("reservationstarttime"))) 
			response.write "</td>"

			' Time
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
			response.write GetTimePortion( oRs("reservationstarttime") ) & " &ndash; " & GetTimePortion( oRs("billingendtime") )
			response.write "</td>"

			' Location 
			response.write "<td align=""left"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
			response.write oRs("rentalname") & "<br /><span class=""locationname"">" & oRs("locationname") & "</span>"
			response.write "</td>"

			' Type
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
			response.write oRs("reservationtype")
			response.write "</td>"

			' Renter
			response.write "<td align=""left"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
			response.write sRenter
			response.write "</td>"

			' Status
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
			If oRs("isonhold") And Not oRs("iscancelled") Then
				response.write "On Hold"
			Else 
				response.write oRs("reservationstatus")
			End If 
			response.write "</td>"

			' Reserved date
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
			response.write DateValue(oRs("reserveddate"))
			response.write "</td>"

			' Balance
'			If oRs("reservationtypeselector") = "public" Then 
'				response.write "<td align=""right"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
'				response.write FormatNumber(CDbl(oRs("totalamount")),2,,,0)
'				response.write "</td>"
'				response.write "<td align=""right"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
'				response.write FormatNumber(CDbl(oRs("balancedue")),2,,,0)
'			Else
'				response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
'				response.write "&nbsp;</td>"
'				response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" nowrap=""nowrap"">"
'				response.write "&nbsp;"
'			End If 
'			response.write "</td>"

			response.write "</tr>"
		End If 

		response.flush
		oRs.MoveNext
	Loop 

	If CLng(iRowCount) = CLng(0) Then
		response.write vbcrlf & "<tr><td colspan=""8"">&nbsp;No Reservations could be found"
		If sSearch <> "" Then
			response.write " that match your search criteria"
		End If 
		response.write ".</td></tr>"
	End If 

	response.write vbcrlf & "</table>"

	oRs.Close
	Set oRs = Nothing 

End Sub



%>
