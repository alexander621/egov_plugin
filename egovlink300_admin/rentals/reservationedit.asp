<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: reservationedit.asp
' AUTHOR: Steve Loar
' CREATED: 10/28/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Edit reservations.
'
' MODIFICATION HISTORY
' 1.0 10/28/2009	Steve Loar - INITIAL VERSION
' 1.1	03/21/2011	Steve Loar - Adding ability to add dates to existing reservations
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim iReservationId, sSuccessFlag, sLoadMsg, iReservationTypeId, sReservationType, iTotalDue
 dim sReservationStatus, sOrganization, sPointOfContact, sNumberAttending, sPurpose, sReceiptNotes
 dim sPrivateNotes, sReservedDate, sReservationTypeSelector, sRenterName, sAdminName, dTotalPaid
 dim sRenterPhone, sServingAlcohol, iLastRentalId, bReservationIsCancelled, iTimeId, sHoldFlag
 dim bIsReservation, bShowCallSelection, sCallFlag, bFacilityAbuse, sFacilityAbuseNote, iRental 
 dim iRentalUserID


 sLevel = "../"  'Override of value from common.asp

'check if page is online and user has permissions in one call not two
 PageDisplayCheck "edit reservations", sLevel	' In common.asp

 iReservationId          = CLng(request("reservationid"))
 iTotalDue               = CDbl(0.00)
 bReservationIsCancelled = false 
 sSuccessFlag            = request("sf")
 sLoadMsg                = ""

 if sSuccessFlag <> "" then
    sSuccessFlag = lcase(sSuccessFlag)
    sLoadMsg     = setupLoadMsg(sSuccessFlag)
 end if

'Set up the "Days & Fees" order by
 lcl_daysFeesOrderBy = "Date_Rental"

 if request("daysFeesOrderBy") <> "" then
    if not containsApostrophe(request("daysFeesOrderBy")) then
       lcl_daysFeesOrderBy = request("daysFeesOrderBy")
    end if
 end if

'Determine which option is selected in the "Days & Fees" dropdown list.
 lcl_selected_daysFeesOrderBy_daterental = " selected=""selected"""
 lcl_selected_daysFeesOrderBy_rentaldate = ""

 if lcl_daysFeesOrderby = "Rental_Date" then
    lcl_selected_daysFeesOrderBy_rentaldate = " selected=""selected"""
 end if

'Check for org features
 lcl_orghasfeature_rentalsummaryreport             = orghasfeature("rentalsummaryreport")
 lcl_orghasfeature_rentals_require_numberattending = orghasfeature("rentals_require_numberattending")
 bShowCallSelection                                = orghasfeature("set_call_flag")

 GetGeneralReservationData iReservationId
%>
<html lang="en">
<head>
	<meta charset="UTF-8">
	
	<title>E-Gov Administration Console { Edit Reservation }</title>

	<link rel="stylesheet" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="rentalsstyles.css" />
	
	<style type="text/css">
		/*margin and padding on body element
		  can introduce errors in determining
		  element position and are not recommended;
		  we turn them off as a foundation for YUI
		  CSS treatments. */
		body {
			margin:0;
			padding:0;
			}
			
		table {
			top: 0;
			left: 0;
		}
		
		#content table {
			top: 0;
			left: 0;
		}
		
	</style>

	<script src="../prototype/prototype-1.6.0.2.js"></script>

	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script src="../scripts/modules.js"></script>
	<script src="../scripts/textareamaxlength.js"></script>
	<script src="../scripts/formatnumber.js"></script>
	<script src="../scripts/removespaces.js"></script>
	<script src="../scripts/removecommas.js"></script>
	<script src="../scripts/setfocus.js"></script>
	<script src="../scripts/formvalidation_msgdisplay.js"></script>
	<script src="../scripts/ajaxLib.js"></script>

	<script>
	<!--
		
		var tabView;
		var winHandle;
		var w = (screen.width - 640)/2;
		var h = (screen.height - 480)/2;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			tabView.set('activeIndex', 0); 

		})();

		function loader()
		{
			setMaxLength();
			<%=sLoadMsg%>
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("screenMsg").innerHTML = "&nbsp;";
		}

		function addDate( iReservationid )
		{
			winHandle = eval('window.open("addnewdate.asp?reservationid=<%=iReservationid%>&rentalid=' + $("lastrentalid").value + '", "_newdate", "width=900,height=500,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=100,top=200")');
		}

		function addDates( iReservationId )
		{
			location.href='rentalsearch.asp?rid=' + iReservationId;
		}

		function ValidateCharges( oFee )
		{
			var bValid = true;

			// Remove any extra spaces
			oFee.value = removeSpaces(oFee.value);
			//Remove commas that would cause problems in validation
			oFee.value = removeCommas(oFee.value);

			// Validate the format of the charge
			if (oFee.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oFee.value);
				if ( Ok )
				{
					oFee.value = format_number(Number(oFee.value),2);
				}
				else 
				{
					oFee.value = format_number(0,2);
					bValid = false;
				}
			}
			else
			{
				oFee.value = "0.00";
			}

			if ( bValid == false ) 
			{
				$("reservationok").value = 'false';
				oFee.focus();
				inlineMsg(oFee.id,'<strong>Invalid Value: </strong>Charges should be numbers in currency format.',8,oFee.id);
				return false;
			}
			$("reservationok").value = 'true';
			return true;
		}

		function ValidateQuantity( oQty, iItem )
		{
			var bValid = true;

			// Clear any error on this quantity input
			clearMsg(oQty.id);

			// Remove any extra spaces
			oQty.value = removeSpaces(oQty.value);
			//Remove commas that would cause problems in validation
			oQty.value = removeCommas(oQty.value);

			// Validate the format of the charge
			if (oQty.value != '')
			{
				var rege = /^\d*$/
				var Ok = rege.exec(oQty.value);
				if ( Ok )
				{
					oQty.value = Number(oQty.value);
				}
				else 
				{
					oQty.value = 0;
					bValid = false;
				}
			}
			else
			{
				oQty.value = 0;
			}

			if ( bValid == false ) 
			{
				$("reservationok").value = 'false';
				$("itemfeeamount" + iItem).value = "0.00";
				oQty.focus();
				inlineMsg(oQty.id,'<strong>Invalid Value: </strong>Quantities should be positive whole numbers.',10,oQty.id);
				return false;
			}
			else
			{
				// check for qty > than MaxQty
				if (Number(oQty.value) > Number($("maxavailable" + iItem).value))
				{
					$("reservationok").value = 'false';
					oQty.value = "0";
					$("itemfeeamount" + iItem).value = "0.00";
					oQty.focus();
					inlineMsg(oQty.id,'<strong>Invalid Value: </strong>Quantities cannot be greater than the maximum allowed.',10,oQty.id);
					return false;
				}
			}
			
			var dAmount = Number($("amount" + iItem).value);
			$("itemfeeamount" + iItem).value = format_number((dAmount * Number(oQty.value)),2);
			$("reservationok").value = 'true';
			return true;
		}

		function validateReservation()
		{

//Check for numberattending as a required field.
<%
  if lcl_orghasfeature_rentals_require_numberattending then
     response.write "  var lcl_numberattending = document.getElementById(""numberattending"").value;" & vbcrlf
     response.write "  if(lcl_numberattending == '' || lcl_numberattending == '0') {" & vbcrlf
     response.write "     document.getElementById('tab_tab2').click();" & vbcrlf
     response.write "     document.getElementById(""numberattending"").focus();" & vbcrlf
     response.write "     inlineMsg(document.getElementById(""numberattending"").id,'<strong>Required Field Missing: </strong>Number Attending.',10,document.getElementById(""numberattending"").id);" & vbcrlf
     response.write "     document.getElementById(""reservationok"").value = 'false';" & vbcrlf
     response.write "  			return false;" & vbcrlf
     response.write "  } else {" & vbcrlf
     response.write "     clearMsg(""numberattending"");" & vbcrlf
     response.write "     document.getElementById(""reservationok"").value = 'true';" & vbcrlf
     response.write "  }" & vbcrlf
  end if
%>

			if ($("reservationok").value == 'true')
			{
				document.frmReservation.submit();
			}
			else
			{
				$("reservationok").value = 'true';
			}
		}

		function setReservationOk()
		{
			$("reservationok").value = 'true';
		}

		function cancelReservation( iReservationId )
		{
			if (confirm("Are you sure you want to cancel this reservation?"))
			{
				// Fire off job to cancel the reservation
				doAjax('cancelreservation.asp', 'reservationid=' + iReservationId , 'refreshReservation', 'get', '0');
			}
		}

		function reinstateReservation( iReservationId )
		{
			if (confirm("Are you sure you want to reinstate this reservation?"))
			{
				// Fire off job to reinstate the reservation
				doAjax('reservationreinstate.asp', 'reservationid=' + iReservationId , 'refreshReservation', 'get', '0');
			}
		}

		function cancelDate( iReservationDateId )
		{
			if (confirm("Are you sure you want to cancel the selected date(s)?"))
			{
				// Fire off job to cancel the date
				doAjax('cancelreservationdate.asp', 'reservationdateid=' + iReservationDateId , 'refreshReservation', 'get', '0');
			}
		}

		function refreshReservation( sReturn )
		{
			//alert( sReturn );
			location.href='reservationedit.asp?reservationid=<%=iReservationId%>&sf=' + sReturn;
		}

		function viewSummary( iReservationId )
		{
			location.href='viewreservationsummary.asp?reservationid=' + iReservationId;
		}

		function changeAlcoholAmount( iReservationFeeCount ) 
		{
			if ($("servingalcohol" + iReservationFeeCount).checked == true)
			{
				$("reservationfeeamount" + iReservationFeeCount).value = $("alcoholfeeamount" + iReservationFeeCount).value;
			}
			else
			{
				$("reservationfeeamount" + iReservationFeeCount).value = '0.00';
			}
		}

  function reorderDaysFees() {
    var lcl_reservationid   = '<%=iReservationId%>';
    var lcl_daysFeesOrderBy = document.getElementById('daysFeesOrderBy').value;
    var lcl_url;

    lcl_url  = 'reservationedit.asp';
    lcl_url += '?reservationid='   + lcl_reservationid;
    lcl_url += '&daysFeesOrderBy=' + lcl_daysFeesOrderBy;

    location.href = lcl_url;
  }

	//-->
	</script>
</head>
<body class="yui-skin-sam" onload="loader();">
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<p><font size=""+1""><strong>Reservation Edit</strong></font></p>" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "		<input type=""button"" class=""button"" value=""<< Reservation List"" onclick=""location.href='reservationlist.asp';"" />" & vbcrlf
  response.write "		<span id=""screenMsg"">&nbsp;</span>" & vbcrlf
  response.write "</p>" & vbcrlf
  response.write "<form name=""frmReservation"" id=""frmReservation"" action=""reservationeditupdate.asp"" method=""post"">" & vbcrlf
  response.write "		<input type=""hidden"" name=""reservationid"" id=""reservationid"" value=""" & iReservationId & """ />" & vbcrlf
  response.write "		<input type=""hidden"" name=""reservationok"" id=""reservationok"" value=""true"" />" & vbcrlf

 'BEGIN: Display Reservation Info ---------------------------------------------
  ShowReservationInfoContainer sReservationType, sRenterName, sRenterPhone, sReservationStatus, sReservedDate, sAdminName, iReservationId, sReservationTypeSelector, iTimeId

  if bIsReservation then
     response.write "<p class=""flagselection"">" & vbcrlf
	 sAbuseChecked = ""
	 if bFacilityAbuse then sAbuseChecked = "checked"
	 response.write "<input type=""hidden"" name=""abuseuserid"" id=""abuseuserid"" value=""" & iRentalUserId & """ />"
     response.write "<input type=""checkbox"" name=""isabusive"" id=""isabusive"" " & sAbuseChecked & " /> Facility/Rental Abuser" & vbcrlf
	 response.write "<br />"
	 response.write "Abuse Notes: <br />"
	 response.write "<textarea name=""abusenote"" id=""abusenote"" style=""width:740px;height:25px"">" & sFacilityAbuseNote & "</textarea>"
     response.write "</p>" & vbcrlf

     response.write "<p class=""flagselection"">" & vbcrlf
     response.write "<input type=""checkbox"" name=""isonhold"" id=""isonhold""" & sHoldFlag & " /> Reservation Is On Hold" & vbcrlf
     response.write "</p>" & vbcrlf
  else
     response.write "<input type=""hidden"" name=""isonhold"" id=""isonhold"" value=""off"" />" & vbcrlf
  end if
 'END: Display Reservation Info -----------------------------------------------
 
 If bShowCallSelection Then
	response.write "<p class=""flagselection"">" & vbcrlf
	response.write "<input type=""checkbox"" name=""iscall"" id=""iscall""" & sCallFlag & " /> Display Call Message to Public" & vbcrlf
	response.write "</p>" & vbcrlf
 Else
 	response.write "<input type=""hidden"" name=""iscall"" id=""iscall"" value=""off"" />" & vbcrlf
 End If

 'BEGIN: Display Buttons ------------------------------------------------------
  response.write "<p>" & vbcrlf

  if not bReservationIsCancelled then
     response.write "  <input type=""button"" class=""button"" value=""Save Changes"" onclick=""validateReservation();"" />" & vbcrlf
     response.write "  <input type=""button"" class=""button"" value=""Cancel Reservation"" onclick=""cancelReservation(" & iReservationId & ");"" />" & vbcrlf
  else
     response.write "  <input type=""button"" class=""button"" value=""Reinstate Reservation"" onclick=""reinstateReservation(" & iReservationId & ");"" />" & vbcrlf
  end if

  if sReservationTypeSelector <> "block" then
     response.write "  <input type=""button"" class=""button"" value=""View Summary"" onclick=""viewSummary(" & iReservationId & ");"" />" & vbcrlf
  end if

  response.write "</p>" & vbcrlf
 'END: Display Buttons --------------------------------------------------------

  response.write "<div id=""demo"" class=""yui-navset"">" & vbcrlf
  response.write "		<ul class=""yui-nav"">" & vbcrlf
  response.write "				<li><a id=""tab_tab1"" href=""#tab1""><em>Days & Fees</em></a></li>" & vbcrlf
  response.write "				<li><a id=""tab_tab2"" href=""#tab2""><em>Event</em></a></li>" & vbcrlf
  response.write "				<li><a id=""tab_tab3"" href=""#tab3""><em>Venue</em></a></li>" & vbcrlf
  response.write "		</ul>" & vbcrlf
  response.write "		<div class=""yui-content"">" & vbcrlf

 'BEGIN: TAB 1 - Days & Fees --------------------------------------------------
  response.write "  <div id=""tab1"">" & vbcrlf
  
  response.write "    <table cellpadding=""0"" cellspacing=""0"" border=""0"" class=""tabtable"">"
  response.write "    <tr><td>"
  response.write "  		<div id=""adddate"">" & vbcrlf
  response.write "      <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "        <tr>" & vbcrlf
  response.write "            <td>" & vbcrlf

  if not bReservationIsCancelled then
     response.write " <input type=""button"" class=""button"" value=""Add Dates"" onclick=""addDates("   & iReservationId & ")"" />" & vbcrlf
     response.write " <input type=""button"" class=""button"" value=""Add One Date"" onclick=""addDate(" & iReservationId & ")"" />" & vbcrlf
  end if

If sReservationTypeSelector = "public" Then 
	If Not bReservationIsCancelled Then 
	response.write " <input type=""button"" class=""button"" value=""Make A Payment"" onclick=""location.href='reservationpayment.asp?reservationid=" & iReservationId & "';"" />" & vbcrlf
	End If 

	'If GetReservationTotalAmount( iReservationId, "totalamount" ) + GetReservationTotalAmount( iReservationId, "totalrefunded" ) -  GetTotalPaidForReservation( iReservationId ) < 0 then
		' I am showing this so that refunds can be made to memo accounts when days are dropped and new dates are added. SJL 1/15/2013
	If ReservationHasUnRefundedPayments( iReservationId ) Then 
		response.write " <input type=""button"" class=""button"" id=""refundbutton"" value=""Give A Refund"" onclick=""location.href='reservationrefund.asp?reservationid=" & iReservationId & "';"" />" & vbcrlf
	End If 
	'End If 

End If 

  response.write "            </td>" & vbcrlf
  response.write "            <td align=""right"">" & vbcrlf
  response.write "                Order Days & Fees By:" & vbcrlf
  response.write "                <select name=""daysFeesOrderBy"" id=""daysFeesOrderBy"" onchange=""reorderDaysFees()"">" & vbcrlf
  response.write "                  <option value=""Date_Rental""" & lcl_selected_daysFeesOrderBy_daterental & ">Date, Rental</option>" & vbcrlf
  response.write "                  <option value=""Rental_Date""" & lcl_selected_daysFeesOrderBy_rentaldate & ">Rental, Date</option>" & vbcrlf
  response.write "                </select>" & vbcrlf
  response.write "            </td>" & vbcrlf
  response.write "        </tr>" & vbcrlf
  response.write "      </table>" & vbcrlf
  response.write "    </div>" & vbcrlf

		response.flush  
  ShowDatesAndFees iReservationId, _
                   sReservationTypeSelector, _
                   iTotalDue, _
                   lcl_daysFeesOrderBy, _
                   sServingAlcohol, _
                   bReservationIsCancelled

  response.write "    </td></tr></table>"
  response.write "  </div>" & vbcrlf
 'END: TAB 1 - Days & Fees ----------------------------------------------------

 'BEGIN: TAB 2 - Event --------------------------------------------------------
  lcl_required_numberattending = ""

  if lcl_orghasfeature_rentals_require_numberattending then
     lcl_required_numberattending = "<span class=""requiredField"">* </span>"
  end if

  response.write "  <div id=""tab2"">" & vbcrlf
  
  response.write "    <table cellpadding=""0"" cellspacing=""0"" border=""0"" class=""tabtable"">"
  response.write "    <tr><td>"
  response.write "    <p>" & vbcrlf
  response.write "    <table id=""eventinfo"" cellpadding=""2"" cellspacing=""0"" border=""0"">" & vbcrlf
  response.write "    		<tr>" & vbcrlf
  response.write "          <td>Organization:</td>" & vbcrlf
  response.write "          <td><input type=""text"" name=""organization"" id=""organization"" value=""" & sOrganization & """ size=""100"" maxlength=""200"" /></td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "    		<tr>" & vbcrlf
  response.write "          <td>Point Of Contact:</td>" & vbcrlf
  response.write "          <td><input type=""text"" name=""pointofcontact"" id=""pointofcontact"" value=""" & sPointOfContact & """ size=""100"" maxlength=""200"" /></td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "      <tr>" & vbcrlf
  response.write "          <td>" & lcl_required_numberattending & "Number Attending:</td>" & vbcrlf
  response.write "          <td>" & vbcrlf

  if lcl_orghasfeature_rentalsummaryreport then
     buildNumberAttendingOptions sNumberAttending
  else
     response.write "          <input type=""text"" id=""numberattending"" name=""numberattending"" value=""" & sNumberAttending & """ size=""20"" maxlength=""20"" />" & vbcrlf
  end if

  response.write "          </td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "      <tr>" & vbcrlf
  response.write "          <td>Purpose:</td>" & vbcrlf
  response.write "          <td><input type=""text"" id=""purpose"" name=""purpose"" value=""" & sPurpose & """ size=""100"" maxlength=""200"" /></td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "    </table>" & vbcrlf
  response.write "    </p>" & vbcrlf
  response.write "    <p>" & vbcrlf
  response.write "      Reservation Notes (This will display on the receipt.):<br />" & vbcrlf
  response.write "      <textarea id=""receiptnotes"" name=""receiptnotes"" maxlength=""2000"" wrap=""soft"">" & sReceiptNotes & "</textarea>" & vbcrlf
  response.write "    </p>" & vbcrlf
  response.write "    <p>" & vbcrlf
  response.write "      Internal Notes:<br />" & vbcrlf
  response.write "      <textarea id=""privatenotes"" name=""privatenotes"" maxlength=""2000"" wrap=""soft"">" & sPrivateNotes & "</textarea>" & vbcrlf
  response.write "    </p>" & vbcrlf
  response.write "    </td></tr></table>"
  
  response.write "  </div>" & vbcrlf
 'END: TAB 2 - Event ----------------------------------------------------------

 'BEGIN: TAB 3 - Venue --------------------------------------------------------
  response.write "  <div id=""tab3"">" & vbcrlf
  response.write "    <p>" & vbcrlf
  response.write "      <strong>Venue Information and documents</strong><br />" & vbcrlf
  response.write "    </p>" & vbcrlf

                      ShowVenueInformation iReservationId, _
                                           iLastRentalId

  response.write "    <input type=""hidden"" name=""lastrentalid"" id=""lastrentalid"" value=""" & iLastRentalId & """ />" & vbcrlf
  response.write "  </div>" & vbcrlf

 'END: TAB 3 - Venue ----------------------------------------------------------

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf

  if not bReservationIsCancelled then
     response.write "  <p>" & vbcrlf
     response.write "    <input type=""button"" class=""button"" value=""Save Changes"" onclick=""validateReservation();"" />" & vbcrlf
     response.write "  </p>" & vbcrlf
  end if

  response.write "</form>" & vbcrlf
  response.write "		</div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
	<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
Sub ShowDatesAndFees( ByVal iReservationId, ByVal sReservationTypeSelector, ByRef iTotalDue, ByRef iDaysFeesOrderBy, ByVal sServingAlcohol, ByVal  bReservationIsCancelled )
	Dim sSql, oRs, iRowCount, sStartHour, sStartMinute, sStartAmPm, sEndHour, sEndMinute, sEndAmPm
	Dim sAmPm, iDateTotal, iReservationFeeCount, iRentalRateCount, iReservationItemCount, dTotalRefundFees
	Dim iReservationDateCount, dTotalRefunded, dTotalPaid, bIsCancelled, dTotalRefundCombined, sKeyClass
	Dim sDaysFeesOrderBy, dTotalRefundDue

	iRowCount             = 0
	iDateTotal            = CDbl(0.00)
	iReservationFeeCount  = clng(0)
	iRentalRateCount      = clng(0)
	iReservationItemCount = clng(0)
	iReservationDateCount = clng(0)
	sDaysFeesOrderBy      = "Date_Rental"
	dTotalRefundDue = CDbl(0)

	If iDaysFeesOrderBy <> "" Then 
		If Not containsApostrophe(iDaysFeesOrderBy) Then 
			sDaysFeesOrderBy = iDaysFeesOrderBy
		End If 
	End If 

	'Get the reserved dates
	sSql = "SELECT "
	sSql = sSql & " D.reservationdateid, D.reservationid, D.reservationstarttime, D.billingendtime, "
	sSql = sSql & " ISNULL(D.actualstarttime, D.reservationstarttime) AS actualstarttime, "
	sSql = sSql & " ISNULL(D.actualendtime, D.billingendtime) AS actualendtime, "
	sSql = sSql & " D.reserveddate, D.adminuserid, D.rentalid, S.reservationstatus, "
	sSql = sSql & " S.iscancelled, R.rentalname, L.name "
	sSql = sSql & " FROM egov_rentalreservationdates D, egov_rentalreservationstatuses S, egov_rentals R, egov_class_location L "
	sSql = sSql & " WHERE D.statusid = S.reservationstatusid "
	sSql = sSql & " AND D.rentalid = R.rentalid AND R.locationid = L.locationid AND D.orgid = " & session("OrgId")
	sSql = sSql & " AND D.reservationid = " & iReservationId

	If sDaysFeesOrderBy = "Rental_Date" Then 
		sSql = sSql & " ORDER BY L.name, R.rentalname, D.reservationstarttime"
	Else 
		sSql = sSql & " ORDER BY D.reservationstarttime, L.name, R.rentalname"
	End If 

	session("reservationdates") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<table id=""datesandfees"" cellpadding=""2"" cellspacing=""0"" border=""0"">"

	Do While Not oRs.EOF
		' Display the date info for all
		iRowCount             = iRowCount + 1
		iReservationDateCount = iReservationDateCount + 1
		iDateTotal            = CDbl(0.00)
		
		If iRowCount Mod 2 = 0 Then
			sClass    = " class=""altrow"" "
			sKeyClass = " altrow"
		Else
			sClass    = ""
			sKeyClass = ""
		End If 
		response.write vbcrlf & "<tr class=""keyreservationdata" & sKeyClass & """>"

		' Date
		response.write "<td><span class=""reservationdata"">"
		response.write DateValue(oRs("reservationstarttime"))
		response.write "</span></td>"

		' From and to times
		sStartAmPm = "AM"
		sStartHour = Hour(oRs("reservationstarttime"))
		If clng(sStartHour) = clng(0) Then
			sStartHour = 12
			sStartAmPm = "AM"
		Else
			If clng(sStartHour) > clng(12) Then
				sStartHour = clng(sStartHour) - clng(12)
				sStartAmPm = "PM"
			End If 
			If clng(sStartHour) = clng(12) Then
				sStartAmPm = "PM"
			End If 
		End If 
		sStartMinute = Minute(oRs("reservationstarttime"))
		If sStartMinute < 10 Then
			sStartMinute = "0" & sStartMinute
		End If 

		sEndAmPm = "AM"
		sEndHour = Hour(oRs("billingendtime"))
		If clng(sEndHour) = clng(0) Then
			sEndHour = 12
			sEndAmPm = "AM"
		Else
			If clng(sEndHour) > clng(12) Then
				sEndHour = clng(sEndHour) - clng(12)
				sEndAmPm = "PM"
			End If 
			If clng(sEndHour) = clng(12) Then
				sEndAmPm = "PM"
			End If 
		End If 
		sEndMinute = Minute(oRs("billingendtime"))
		If sEndMinute < 10 Then
			sEndMinute = "0" & sEndMinute
		End If 

		response.write "<td><span class=""reservationdata"">" & sStartHour & ":" & sStartMinute & " " & sStartAmPm
		response.write " &mdash; " & sEndHour & ":" & sEndMinute & " " & sEndAmPm
		response.write " (" & FormatNumber(CalculateDurationInHours( oRs("reservationstarttime"), oRs("billingendtime") ),2,,,0) & " hours)"
		response.write "</span></td>"
		
		' Location
		response.write "<td colspan=""2""><span class=""reservationdata"">"
		ShowRentalNameAndLocation oRs("rentalid") 
		response.write "</span></td>"

		response.write "<td>&nbsp;</td>"

		response.write "</tr>"

		' change times button
'		If Not bReservationIsCancelled And Not oRs("iscancelled") Then 
'			response.write  vbcrlf & "<tr" & sClass & ">"
'			response.write "<td>&nbsp;</td>"
'			response.write "<td>"
'			response.write "<input type=""button"" class=""button"" value=""Change Times"" onclick=""location.href='reservationtimechange.asp?rdi=" & oRs("reservationdateid") & "';"" />"
'			response.write "</td>"
'			response.write "<td>&nbsp;</td>"
'			response.write "</tr>"
'		End If 


		response.write vbcrlf & "<tr" & sClass & ">"

		' Status and Cancel button
		response.write "<td valign=""top""><span class=""reservationdata"">"
		response.write oRs("reservationstatus") & "</span>"
		If Not oRs("iscancelled") Then
			response.write "<br /><br /> <input type=""button"" class=""button"" value=""Cancel Date"" name=""canceldate"" onclick=""cancelDate( " & oRs("reservationdateid") & ");"" />"
			response.write "<div style=""margin-top:6px;""><input type=""checkbox"" name=""cancels"" value=""" & oRs("reservationdateid") & """ /> cancel multiple?</div>"
			bIsCancelled = False 
		Else
			bIsCancelled = True 
		End If 
		response.write "</td>"


		' Rental details 
		response.write "<td colspan=""2"">"
		If Not bReservationIsCancelled And Not oRs("iscancelled") Then 
			response.write "<input type=""button"" class=""button"" value=""Change Times"" onclick=""location.href='reservationtimechange.asp?rdi=" & oRs("reservationdateid") & "';"" /><br /><br />"
		End If 
		response.write "<span class=""reservationdata"">" & WeekDayName(Weekday(DateValue(oRs("reservationstarttime")))) & "</span>"
		response.write " &ndash; " & GetRentalSeason( oRs("rentalid"), DateValue(oRs("reservationstarttime")) )
		'response.write GetRentalHours( oRs("rentalid"), DateValue(oRs("reservationstarttime")) )
		response.write "</td>"


		' Arrival and departure
		response.write "<td valign=""bottom"" align=""center"">"
		response.write "<input type=""hidden"" name=""reservationdateid" & iReservationDateCount & """ value=""" & oRs("reservationdateid") & """ />"
		response.write "<input type=""hidden"" name=""reservationarrivaldate" & iReservationDateCount & """ value=""" & DateValue(oRs("reservationstarttime")) & """ />"
		response.write "<input type=""hidden"" name=""reservationdeparturedate" & iReservationDateCount & """ value=""" & DateValue(oRs("billingendtime")) & """ />"
		If sReservationTypeSelector <> "block" And sReservationTypeSelector <> "class" Then
			response.write "<strong>Arrival And Departure Times</strong>"
			response.write "<table id=""comeandgotimes"" cellpadding=""1"" cellspacing=""0"" border=""0""" & sClass & ">"
			response.write "<tr>"
			response.write "<td align=""right"">"
			response.write "Arrival:"
			response.write "</td>"
			response.write "<td>"
			ShowHourPicks "arrivalhour" & iReservationDateCount, GetHourFromDateTime( oRs("actualstarttime"), sAmPm ), ""
			response.write ":"
			ShowMinutePicks "arrivalminute" & iReservationDateCount, Minute(oRs("actualstarttime")), ""
			response.write " "
			ShowAmPmPicks "arrivalampm" & iReservationDateCount, sAmPm, ""
			response.write "</td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td align=""right"">"
			response.write "Departure:"
			response.write "</td>"
			response.write "<td>"
			ShowHourPicks "departurehour" & iReservationDateCount, GetHourFromDateTime( oRs("actualendtime"), sAmPm ), ""
			response.write ":"
			ShowMinutePicks "departureminute" & iReservationDateCount, Minute(oRs("actualendtime")), ""
			response.write " "
			ShowAmPmPicks "departureampm" & iReservationDateCount, sAmPm, ""
			response.write "</td>"
			response.write "</tr>"
			response.write "</table>"
		Else
			response.write "&nbsp;"
			response.write vbcrlf & "<input type=""hidden"" name=""arrivalhour" & iReservationDateCount & """ value=""" & sStartHour & """ />"
			response.write vbcrlf & "<input type=""hidden"" name=""arrivalminute" & iReservationDateCount & """ value=""" & sStartMinute & """ />"
			response.write vbcrlf & "<input type=""hidden"" name=""arrivalampm" & iReservationDateCount & """ value=""" & sStartAmPm & """ />"
			response.write vbcrlf & "<input type=""hidden"" name=""departurehour" & iReservationDateCount & """ value=""" & sEndHour & """ />"
			response.write vbcrlf & "<input type=""hidden"" name=""departureminute" & iReservationDateCount & """ value=""" & sEndMinute & """ />"
			response.write vbcrlf & "<input type=""hidden"" name=""departureampm" & iReservationDateCount & """ value=""" & sEndAmPm & """ />"
		End If 
		response.write "</td>"


		response.write "<td>&nbsp;</td>"

		response.write "</tr>"

		If sReservationTypeSelector = "public" Then '  Or sReservationTypeSelector = "admin" Then
			' Display the hourly rates for public, and internal reservations
			response.write vbcrlf & "<tr" & sClass & "><td colspan=""2"" align=""right""><strong>Rates</strong></td>"
			response.write "<td colspan=""3"">&nbsp;</td>"
			response.write "</tr>"
			ShowRentalRatesForDate oRs("reservationdateid"), sClass, iTotalDue, iDateTotal, iRentalRateCount, bIsCancelled, sReservationTypeSelector
		End If 

		If sReservationTypeSelector <> "block" Then 
			' Display items for all but blocked
			response.write vbcrlf & "<tr" & sClass & "><td colspan=""2"" align=""right""><strong>Items</strong></td>"
			response.write "<td align=""right"" colspan=""2""><strong>Qty</strong></td>"
			response.write "<td>&nbsp;</td>"
			response.write "</tr>"
			ShowItemsForDate oRs("reservationdateid"), sReservationTypeSelector, sClass, iTotalDue, iDateTotal, iReservationItemCount, bIsCancelled
		End If 

'		If sReservationTypeSelector = "public" Then
'			' Sub Total Row
'			response.write vbcrlf & "<tr" & sClass & "><td class=""subtotalscell"" colspan=""2"">&nbsp;</td><td class=""subtotalscell"" align=""right""><strong>Subtotal</strong></td>"
'			response.write "<td class=""subtotalscell"" align=""right"">"
'			response.write FormatNumber(iDateTotal,2,,,0) 
'			response.write "</td></tr>"
'		End If 

		oRs.MoveNext
		response.flush
	Loop

	
	If sReservationTypeSelector = "public" Then
		iRowCount = iRowCount + 1
		If iRowCount Mod 2 = 0 Then
			sClass = " class=""altrow"" "
		Else
			sClass = ""
		End If 
		' Display Reservation Level charges like Deposit and Alcohol Fee
		ShowReservationFees iReservationId, sClass, iTotalDue, iReservationFeeCount, sServingAlcohol, bReservationIsCancelled
	Else
		response.write vbcrlf & "<tr" & sClass & "><td colspan=""5"">&nbsp;</td></tr>"
	End If 

	If sReservationTypeSelector = "public" Or sReservationTypeSelector = "admin" Then
		' Total Charges Row
		iRowCount = iRowCount + 1
		If iRowCount Mod 2 = 0 Then
			sClass = " class=""altrow"" "
		Else
			sClass = ""
		End If 
		If Not bReservationIsCancelled Then		
			response.write vbcrlf & "<tr" & sClass & "><td class=""totalscell"" colspan=""2"">"
			response.write "<input type=""button"" class=""button"" value=""Cancel Checked Reservations"" onclick=""cancelDate(getCanceledVals());"" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			response.write "<script> function getCanceledVals() { var checkboxes = document.getElementsByName('cancels'); var vals = ''; for (var i=0, n=checkboxes.length;i<n;i++) { if (checkboxes[i].checked) { vals += checkboxes[i].value+','; } } return vals.substring(0, vals.length-1); }</script>"
			response.write vbcrlf & "</td><td class=""totalscell"" align=""right"">"
			response.write "<input type=""button"" class=""button"" value=""Save Changes"" onclick=""validateReservation();"" />"
		Else
			response.write vbcrlf & "<tr" & sClass & "><td class=""totalscell"" colspan=""3"" align=""right"">"
			response.write "&nbsp;"
		End If 

		response.write "</td>"
		response.write "<td class=""totalscell"" align=""right""><strong>Total Charges</strong></td>"
		response.write "<td class=""totalscell"" align=""right"">"
		response.write FormatNumber(iTotalDue,2,,,0) 
		response.write "</td></tr>"

		' Total Paid Row
		iRowCount = iRowCount + 1
		If iRowCount Mod 2 = 0 Then
			sClass = " class=""altrow"" "
		Else
			sClass = ""
		End If 
		response.write vbcrlf & "<tr" & sClass & "><td class=""totalscell"" colspan=""5"" align=""center""><strong>Payments</strong></td></tr>"
		' Want to show payments with link to receipt page here
		ShowReservationPayments iReservationId, sClass
		response.write vbcrlf & "<tr" & sClass & ">"
		response.write "<td colspan=""4"" align=""right"" class=""totalscell""><strong>Total Paid</strong></td>"
		response.write "<td align=""right"" class=""totalscell"">"
		'dTotalPaid = GetReservationTotalAmount( iReservationId, "totalpaid" ) ' In rentalscommonfunctions.asp
		dTotalPaid = GetTotalPaidForReservation( iReservationId )  ' In rentalscommonfunctions.asp
		response.write FormatNumber(dTotalPaid,2,,,0) 
		response.write "</td></tr>"

		' Refund Row
		If sReservationTypeSelector = "public" Then 
			iRowCount = iRowCount + 1
			If iRowCount Mod 2 = 0 Then
				sClass = " class=""altrow"" "
			Else
				sClass = ""
			End If 
			response.write vbcrlf & "<tr" & sClass & "><td class=""totalscell"" colspan=""5"" align=""center""><strong>Refunds & Refund Fees</strong></td></tr>"
			' Want to show refunds with link to receipt page here
			dTotalRefunded = ShowReservationRefunds( iReservationId, sClass )
			response.write vbcrlf & "<tr" & sClass & ">"
			response.write "<td class=""totalscell"" colspan=""3"">&nbsp;</td>"
			response.write "<td class=""totalscell"" align=""right""><strong>Total Refunds & Refund Fees</strong></td>"
			response.write "<td class=""totalscell"" align=""right"">"
			response.write FormatNumber(dTotalRefunded,2,,,0) 
			response.write "</td></tr>"
		End If 

		' Refund Due Row
		If sReservationTypeSelector = "public" Then
			iRowCount = iRowCount + 1
			If iRowCount Mod 2 = 0 Then
				sClass = " class=""altrow"" "
			Else
				sClass = ""
			End If 
			' get the refund due amount
			dTotalRefundDue = GetReservationRefundDue( iReservationId )
			response.write vbcrlf & "<tr" & sClass & ">"
			response.write "<td class=""totalscell"" colspan=""3"">&nbsp;</td>"
			response.write "<td class=""totalscell"" align=""right""><strong>Refund Due</strong></td>"
			response.write "<td class=""totalscell"" align=""right"">"
			response.write FormatNumber(dTotalRefundDue,2,,,0) 
			response.write "</td></tr>"
		End If 

		' Balance Due Row
		iRowCount = iRowCount + 1
		If iRowCount Mod 2 = 0 Then
			sClass = " class=""altrow"" "
		Else
			sClass = ""
		End If 
		response.write vbcrlf & "<tr" & sClass & "><td class=""totalscell"" colspan=""4"" align=""right""><strong>Balance Due</strong></td><td class=""totalscell"" align=""right"">"
		dBalanceDue = (iTotalDue + dTotalRefunded) - (dTotalPaid - dTotalRefundDue)
		response.write FormatNumber(dBalanceDue,2,,,0) 
		response.write "</td></tr>"
	else
		' Total Charges Row
		iRowCount = iRowCount + 1
		If iRowCount Mod 2 = 0 Then
			sClass = " class=""altrow"" "
		Else
			sClass = ""
		End If 
		If Not bReservationIsCancelled Then		
			response.write vbcrlf & "<tr" & sClass & "><td class=""totalscell"" colspan=""2"">"
			response.write "<input type=""button"" class=""button"" value=""Cancel Checked Reservations"" onclick=""cancelDate(getCanceledVals());"" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			response.write "<script> function getCanceledVals() { var checkboxes = document.getElementsByName('cancels'); var vals = ''; for (var i=0, n=checkboxes.length;i<n;i++) { if (checkboxes[i].checked) { vals += checkboxes[i].value+','; } } return vals.substring(0, vals.length-1); }</script>"
			response.write vbcrlf & "</td><td class=""totalscell"" align=""right"" colspan=""3"">&nbsp;</td></tr>"
		Else
			response.write vbcrlf & "<tr" & sClass & "><td class=""totalscell"" colspan=""3"" align=""right"">"
			response.write "&nbsp;</td></tr>"
		End If 
	End If 

	response.write vbcrlf & "</table>"

	' Write out the maxcounts
	response.write "<input type=""hidden"" id=""maxreservationfees"" name=""maxreservationfees"" value=""" & iReservationFeeCount & """ />"
	response.write "<input type=""hidden"" id=""maxrentalrates"" name=""maxrentalrates"" value=""" & iRentalRateCount & """ />"
	response.write "<input type=""hidden"" id=""maxreservationitems"" name=""maxreservationitems"" value=""" & iReservationItemCount & """ />"
	response.write "<input type=""hidden"" id=""maxreservationdates"" name=""maxreservationdates"" value=""" & iReservationDateCount & """ />"
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------
' void ShowReservationPayments iReservationId, sClass
'------------------------------------------------------------------------------
Sub ShowReservationPayments( ByVal iReservationId, ByVal sClass )
	Dim sSql, oRs

	sSql = "SELECT A.paymentid, J.paymentdate, SUM(A.amount) AS paidamount "
	sSql = sSql & " FROM egov_accounts_ledger A, egov_class_payment J "
	sSql = sSql & " WHERE A.paymentid = J.paymentid AND A.ispaymentaccount = 1 AND A.entrytype = 'debit' "
	sSql = sSql & " AND A.reservationid = " & iReservationId
	sSql = sSql & " GROUP BY A.paymentid, J.paymentdate ORDER BY J.paymentdate"
	'response.write "<tr><td>" & sSQL & "</td></tr>"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr" & sClass & "><td class=""subheadercell"">Receipt #</td><td class=""subheadercell"">Date</td><td class=""subheadercell"" colspan=""2"">&nbsp;</td><td align=""right"" class=""subheadercell"">Amount</td></tr>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr" & sClass & ">"
			response.write "<td>&nbsp;<a href=""viewpaymentreceipt.asp?paymentid=" & oRs("paymentid") & "&rt=r""><strong>" & oRs("paymentid") & "</strong></a></td>"
			response.write "<td>" & DateValue(oRs("paymentdate")) & "</td>"
			response.write "<td colspan=""2"">&nbsp;</td>"
			response.write "<td align=""right"">"
			response.write FormatNumber(oRs("paidamount"),2,,,0) 
			response.write "</td></tr>"
			oRs.MoveNext 
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------
' void ShowReservationRefunds iReservationId, sClass
'------------------------------------------------------------------------------
function ShowReservationRefunds( ByVal iReservationId, ByVal sClass )
	Dim sSql, oRs, dTotalRefund

	dTotalRefund = CDbl(0)
	
	sSql = "SELECT A.paymentid, J.paymentdate, SUM(A.amount) as refundamount "
	sSql = sSql & " FROM egov_accounts_ledger A, egov_class_payment J "
	sSql = sSql & " WHERE A.paymentid = J.paymentid AND A.ispaymentaccount = 0 AND A.entrytype = 'debit' "
	sSql = sSql & " AND A.reservationid = " & iReservationId
	sSql = sSql & " GROUP BY A.paymentid, J.paymentdate ORDER BY J.paymentdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr" & sClass & "><td class=""subheadercell"">Receipt #</td><td class=""subheadercell"" colspan=""3"">Date</td><td align=""right"" class=""subheadercell"">Amount</td></tr>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr" & sClass & ">"
			response.write "<td>&nbsp;<a href=""viewpaymentreceipt.asp?paymentid=" & oRs("paymentid") & "&rt=r""><strong>" & oRs("paymentid") & "</strong></a></td>"
			response.write "<td>" & DateValue(oRs("paymentdate")) & "</td>"
			response.write "<td colspan=""2"">&nbsp;</td>"
			response.write "<td align=""right"">"
			response.write FormatNumber(oRs("refundamount"),2,,,0) 
			dTotalRefund = dTotalRefund + CDbl(oRs("refundamount"))
			response.write "</td></tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	ShowReservationRefunds = dTotalRefund

End function 


'------------------------------------------------------------------------------
' void ShowRentalRatesForDate iReservationDateId, sClass, iTotalDue, iDateTotal, sReservationTypeSelector
'------------------------------------------------------------------------------
Sub ShowRentalRatesForDate( ByVal iReservationDateId, ByVal sClass, ByRef iTotalDue, ByRef iDateTotal, ByRef iRentalRateCount, ByVal bIsCancelled, ByVal sReservationTypeSelector )
	Dim sSql, oRs

	sSql = "SELECT F.reservationdatefeeid, F.amount, F.feeamount, F.duration, P.pricetypename, F.starthour, "
	sSql = sSql & " dbo.AddLeadingZeros(F.startminute,2) AS startminute, F.startampm, P.isweekendsurcharge, R.ratetype "
	sSql = sSql & " FROM egov_rentalreservationdatefees F, egov_price_types P, egov_rentalratetypes R "
	sSql = sSql & " WHERE F.pricetypeid = P.pricetypeid AND F.ratetypeid = R.ratetypeid AND F.reservationdateid = " & iReservationDateId
	sSql = sSql & " ORDER BY P.displayorder"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRentalRateCount = iRentalRateCount + clng(1)
		response.write vbcrlf & "<tr" & sClass & ">"
		response.write "<td colspan=""4"" align=""right"">"
		response.write oRs("pricetypename")
		If sReservationTypeSelector = "public" Then
			response.write " (" & FormatNumber(oRs("amount"),2,,,0) & " " & oRs("ratetype") & ")"
		End If 
		response.write "</td><td align=""right"">"
		response.write "<input type=""hidden"" name=""reservationdatefeeid" & iRentalRateCount & """ value=""" & oRs("reservationdatefeeid") & """ />"
		If Not bIsCancelled Then 
			response.write "<input type=""text"" id=""datefeeamount" & iRentalRateCount & """ name=""datefeeamount" & iRentalRateCount & """ value=""" & FormatNumber(oRs("feeamount"),2,,,0) & """ size=""7"" maxlength=""7"""
			response.write " onfocus=""setReservationOk();"" onchange=""return ValidateCharges(this);"" />"
			iTotalDue = iTotalDue + CDbl(oRs("feeamount"))
			iDateTotal = iDateTotal + CDbl(oRs("feeamount"))
		Else
			response.write "<input type=""hidden"" id=""datefeeamount" & iRentalRateCount & """ name=""datefeeamount" & iRentalRateCount & """ value=""" & oRs("feeamount") & """ />" & FormatNumber(oRs("feeamount"),2,,,0)
		End If 
		response.write "</td>"
		response.write "</tr>"
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------
' void ShowItemsForDate iReservationDateId, sReservationTypeSelector, sClass, iTotalDue, iDateTotal, iReservationItemCount, bIsCancelled
'------------------------------------------------------------------------------
Sub ShowItemsForDate( ByVal iReservationDateId, ByVal sReservationTypeSelector, ByVal sClass, ByRef iTotalDue, ByRef iDateTotal, ByRef iReservationItemCount, ByVal bIsCancelled )
	Dim oRs, sSql

	sSql = "SELECT reservationdateitemid, rentalitem, ISNULL(maxavailable,0) AS maxavailable, ISNULL(quantity,0) AS quantity, "
	sSql = sSql & " ISNULL(amount,0.00) AS amount, ISNULL(feeamount,0.00) AS feeamount "
	sSql = sSql & " FROM egov_rentalreservationdateitems "
	sSql = sSql & " WHERE reservationdateid = " & iReservationDateId
	sSql = sSql & " ORDER BY rentalitem"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iReservationItemCount = iReservationItemCount + clng(1)
		response.write vbcrlf & "<tr" & sClass & ">"
		response.write "<td colspan=""4"" align=""right"">"
		response.write "<input type=""hidden"" name=""reservationdateitemid" & iReservationItemCount & """ value=""" & oRs("reservationdateitemid") & """ />"
		response.write oRs("rentalitem")
		response.write " ("
		If sReservationTypeSelector = "public" Then
			response.write FormatNumber(oRs("amount"),2,,,0) & " each, "
		End If 
		response.write "Max " & oRs("maxavailable") & ") &nbsp; "
		If Not bIsCancelled Then
			response.write "<input type=""text"" id=""quantity" & iReservationItemCount & """ name=""quantity" & iReservationItemCount & """ value=""" & FormatNumber(oRs("quantity"),0,,,0) & """ size=""7"" maxlength=""7"""
			response.write " onfocus=""setReservationOk();"" onchange=""return ValidateQuantity( this, " & iReservationItemCount & " )"" />"
		Else
			response.write "<input type=""hidden"" id=""quantity" & iReservationItemCount & """ name=""quantity" & iReservationItemCount & """ value=""" & oRs("quantity") & """ />" & FormatNumber(oRs("quantity"),0,,,0)
		End If 
		response.write "<input type=""hidden"" id=""maxavailable" & iReservationItemCount & """ name=""maxavailable" & iReservationItemCount & """ value=""" & oRs("maxavailable") & """ />"
		response.write "</td>"
		response.write "<td align=""right"">"
		response.write "<input type=""hidden"" id=""amount" & iReservationItemCount & """ name=""amount" & iReservationItemCount & """ value=""" & oRs("amount") & """ />"
		If sReservationTypeSelector = "public" Then 'Or sReservationTypeSelector = "admin" Then
			If Not bIsCancelled Then
				response.write "<input type=""text"" id=""itemfeeamount" & iReservationItemCount & """ name=""itemfeeamount" & iReservationItemCount & """ value=""" & FormatNumber(oRs("feeamount"),2,,,0) & """ size=""7"" maxlength=""7"""
				response.write " onfocus=""setReservationOk();"" onchange=""return ValidateCharges( this );"" />"
				iTotalDue = iTotalDue + CDbl(oRs("feeamount"))
				iDateTotal = iDateTotal + CDbl(oRs("feeamount"))
			Else
				response.write "<input type=""hidden"" id=""itemfeeamount" & iReservationItemCount & """ name=""itemfeeamount" & iReservationItemCount & """ value=""" & oRs("feeamount") & """ />" & FormatNumber(oRs("feeamount"),2,,,0)
			End If 
		Else
			response.write "<input type=""hidden"" id=""itemfeeamount" & iReservationItemCount & """ name=""itemfeeamount" & iReservationItemCount & """ value=""0.00"" />"
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "</tr>"
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------
' void ShowReservationFees iReservationId, sClass, iTotalDue, iReservationFeeCount, sServingAlcohol, bReservationIsCancelled
'------------------------------------------------------------------------------
Sub ShowReservationFees( ByVal iReservationId, ByVal sClass, ByRef iTotalDue, ByRef iReservationFeeCount, ByVal sServingAlcohol, ByVal bReservationIsCancelled )
	Dim oRs, sSql, sCellClass

	sSql = "SELECT F.reservationfeeid, P.pricetypename, ISNULL(F.amount,0.00) AS amount, "
	sSql = sSql & " ISNULL(F.feeamount,0.00) AS feeamount, ISNULL(F.prompt,'') AS prompt, P.isalcoholsurcharge "
	sSql = sSql & " FROM egov_rentalreservationfees F, egov_price_types P "
	sSql = sSql & " WHERE F.pricetypeid = P.pricetypeid AND F.reservationid = " & iReservationId
	sSql = sSql & " ORDER BY P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iReservationFeeCount = iReservationFeeCount + clng(1)
		If iReservationFeeCount = clng(1) Then
			sCellClass = " class=""totalscell"""
		Else
			sCellClass = ""
		End If 
		response.write vbcrlf & "<tr" & sClass & ">"
		response.write "<td colspan=""4"" align=""right""" & sCellClass & ">"
		If oRs("isalcoholsurcharge") Then
			' There will only be one per reservation so no need to put the count on this
			response.write "<input type=""checkbox"" id=""servingalcohol" & iReservationFeeCount & """ name=""servingalcohol""" & sServingAlcohol
			If bReservationIsCancelled Then 
				response.write " disabled=""disabled"" "
			Else
				response.write " onclick=""changeAlcoholAmount( " & iReservationFeeCount & " );"" "
			End If 
			response.write " /> " & oRs("prompt") & " &nbsp; &mdash; &nbsp; "
			response.write "<input type=""hidden"" name=""alcoholfeeamount" & iReservationFeeCount & """ id=""alcoholfeeamount" & iReservationFeeCount & """ value=""" & FormatNumber(oRs("amount"),2,,,0) & """ />"
		End If 
		response.write oRs("pricetypename")
		response.write " (" & FormatNumber(oRs("amount"),2,,,0) & ")"
		response.write "</td>"
		response.write "<td align=""right""" & sCellClass & ">"
		response.write "<input type=""hidden"" name=""reservationfeeid" & iReservationFeeCount & """ & value=""" &  oRs("reservationfeeid") & """ />"
		If Not bReservationIsCancelled Then 
			response.write "<input type=""text"" id=""reservationfeeamount" & iReservationFeeCount & """ name=""reservationfeeamount" & iReservationFeeCount & """ value=""" & FormatNumber(oRs("feeamount"),2,,,0) & """ size=""7"" maxlength=""7"""
			response.write " onfocus=""setReservationOk();"" return onchange=""ValidateCharges(this);"" />"
			iTotalDue = iTotalDue + CDbl(oRs("feeamount"))
		Else
			response.write "<input type=""hidden"" id=""reservationfeeamount" & iReservationFeeCount & """ name=""reservationfeeamount" & iReservationFeeCount & """ value=""" & oRs("feeamount") & """ />" & FormatNumber(oRs("feeamount"),2,,,0)
		End If 
		response.write "</td>"
		response.write "</tr>"
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------
' void ShowVenueInformation iReservationId, iLastRentalId
'------------------------------------------------------------------------------
Sub ShowVenueInformation( ByVal iReservationId, ByRef iLastRentalId )
	Dim oRs, sSql

	sSql = "SELECT DISTINCT D.rentalid, R.rentalname, L.name AS locationname "
	sSql = sSql & " FROM egov_rentalreservationdates D, egov_rentals R, egov_class_location L "
	sSql = sSql & " WHERE R.locationid = L.locationid AND R.rentalid = D.rentalid "
	sSql = sSql & " AND D.orgid = " & session("orgid")
	sSql = sSql & " AND D.reservationid = " & iReservationId
	sSql = sSql & " ORDER BY L.name, R.rentalname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<table id=""venueinfo"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
	Do While Not oRs.EOF
		iLastRentalId = oRs("rentalid")
		response.write vbcrlf & "<tr>"
		response.write "<td colspan=""2""><strong>" & oRs("rentalname") & " &ndash; " & oRs("locationname") & "</strong></td>"
		response.write "</tr>"
		response.write vbcrlf & "<tr>"
		response.write "<td class=""firstcol"">&nbsp;</td><td>" & GetRentalShortDescription( oRs("rentalid") ) & "</td>"
		response.write "</tr>"
		ShowRentalDocuments oRs("rentalid")
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</table>"
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------
' void ShowRentalDocuments iRentalId
'------------------------------------------------------------------------------
Sub ShowRentalDocuments( ByVal iRentalId )
	Dim oRs, sSql

	sSql = "SELECT documenturl, documenttitle "
	sSql = sSql & " FROM egov_rentaldocuments WHERE rentalid = " & iRentalId
	sSql = sSql & " ORDER BY documenttitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"
		response.write "<td class=""firstcol"">&nbsp;</td><td><a href=""" & oRs("documenturl") & """ target=""_blank"" ><strong>" & oRs("documenttitle") & "</strong></a>"
		response.write "</tr>"
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 
End Sub 

'------------------------------------------------------------------------------
sub buildNumberAttendingOptions( ByVal iNumberAttending)
  dim sNumberAttending

  sNumberAttending = 0

  if iNumberAttending <> "" then
     if not containsApostrophe(iNumberAttending) then
        sNumberAttending = clng(iNumberAttending)
     end if
  end if

  response.write "<select name=""numberattending"" id=""numberattending"">" & vbrlf

  for i = 0 to 5 step 5
     lcl_selected_0_5 = ""

     if sNumberAttending = i then
        lcl_selected_0_5 = " selected=""selected"""
     end if

     if i = 0 then
        lcl_displayvalue = ""
     else
        lcl_displayvalue = i
     end if

     response.write "  <option value=""" & i & """" & lcl_selected_0_5 & ">" & lcl_displayvalue & "</option>" & vbcrlf
  next

  for i = 10 to 50 step 10
     lcl_selected_10_50 = ""

     if sNumberAttending = i then
        lcl_selected_10_50 = " selected=""selected"""
     end if

     response.write "  <option value=""" & i & """" & lcl_selected_10_50 & ">" & i & "</option>" & vbcrlf
  next

  for i = 75 to 500 step 25
     lcl_selected_75_500 = ""

     if sNumberAttending = i then
        lcl_selected_75_500 = " selected=""selected"""
     end if

     response.write "  <option value=""" & i & """" & lcl_selected_75_500 & ">" & i & "</option>" & vbcrlf
  next

  response.write "</select>" & vbcrlf

end sub

'------------------------------------------------------------------------------
function setupLoadMsg( ByVal iSuccessFlag )
  dim lcl_return, lcl_successflag

  lcl_successflag = ""

  if iSuccessFlag <> "" then
     lcl_successflag = lcase(iSuccessFlag)

     if lcl_successflag = "rc" then
  	     lcl_return = "displayScreenMsg('The reservation has been successfully made.');"
     elseif lcl_successflag = "re" then
  	     lcl_return = "displayScreenMsg('The reservation has been successfully restored.');"
     elseif lcl_successflag = "ru" then
  	     lcl_return = "displayScreenMsg('The reservation has been successfully updated.');"
     elseif lcl_successflag = "cs" then
  	     lcl_return = "displayScreenMsg('Your changes have been saved.');"
     elseif lcl_successflag = "cd" then
  	     lcl_return = "displayScreenMsg('The selected date(s) have been cancelled.');"
     elseif lcl_successflag = "cr" then
  	     lcl_return = "displayScreenMsg('This reservation has been cancelled.');"
     elseif lcl_successflag = "tc" then
  	     lcl_return = "displayScreenMsg('The reservation time has been changed.');"
     end if
  end if

  setupLoadMsg = lcl_return

end Function

'------------------------------------------------------------------------------
' boolean ReservationHasUnRefundedPayments iReservationId
'------------------------------------------------------------------------------
Function ReservationHasUnRefundedPayments( ByVal iReservationId )
	Dim bHasUnRefundedPayments, sSql, oRs

	bHasUnRefundedPayments = False 

	' this is a combination of what is pulled on the refund making page to get what needs a refund
	sSql = "SELECT COUNT(reservationdatefeeid) AS hits "
	sSql = sSql + "FROM egov_rentalreservationdatefees "
	sSql = sSql + "WHERE reservationid = " & iReservationId
	sSql = sSql + " AND paidamount > (feeamount + refundamount) "
	sSql = sSql + "UNION "
	sSql = sSql + "SELECT COUNT(reservationdateitemid) AS hits "
	sSql = sSql + "FROM egov_rentalreservationdateitems "
	sSql = sSql + "WHERE reservationid = " & iReservationId
	sSql = sSql + " AND paidamount > (feeamount + refundamount) "
	sSql = sSql + "UNION "
	sSql = sSql + "SELECT COUNT(reservationfeeid) AS hits "
	sSql = sSql + "FROM egov_rentalreservationfees "
	sSql = sSql + "WHERE reservationid = " & iReservationId
	sSql = sSql + " AND paidamount > (feeamount + refundamount) "
	sSql = sSql + "ORDER BY hits DESC"
	response.write "<!-- " & sSql & " -->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		' because of the order by clause, we just need to look at the first row, if this pulls more than one row
		If clng(oRs("hits")) > clng(0) Then
			bHasUnRefundedPayments = True 
		Else
			bHasUnRefundedPayments = False 
		End If 
	Else
		bHasUnRefundedPayments = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

	ReservationHasUnRefundedPayments = bHasUnRefundedPayments

End Function 



%>
