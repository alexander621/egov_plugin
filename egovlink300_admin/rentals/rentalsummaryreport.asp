<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: rentalsummaryreport.asp
' AUTHOR: David Boyer
' CREATED: 04/17/2012
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: Summary report on rentals reservations
'
' MODIFICATION HISTORY
' 1.0  04/17/2012	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim iSelectedStartYear, iSelectedEndYear, iSelectedStartMonth, iSelectedEndMonth, dStartDate, dEndDate
 dim iTempMonth, iTempYear, iReservationTypeId, sReservationTypeTitle
 dim lcl_sc_rentalname, lcl_sc_renter, lcl_sc_location

 sLevel = "../"  'Override of value from common.asp

' USER SECURITY CHECK
'PageDisplayCheck "rental totals rpt", sLevel	 'In common.asp
	iReservationTypeId = 0
 lcl_today          = date()
 lcl_sc_fromDate    = ""
 lcl_sc_toDate      = ""
 lcl_sc_dateOptions = "0"
 lcl_sc_rentalname  = ""
 lcl_sc_renter      = ""
 lcl_sc_location    = ""

 if request("sc_fromDate") <> "" then
    if not containsApostrophe(request("sc_fromdate")) then
       lcl_sc_fromDate = request("sc_fromDate")
    end if
 end if

 if request("sc_toDate") <> "" then
    if not containsApostrophe(request("sc_toDate")) then
       lcl_sc_toDate = request("sc_toDate")
    end if
 end if

 if lcl_sc_fromDate = "" or IsNull(lcl_sc_fromDate) then
    lcl_sc_fromDate = dateAdd("yyyy",-1,lcl_today)
 end if

 if lcl_sc_toDate = "" or IsNull(lcl_sc_toDate) then
    lcl_sc_toDate = dateAdd("d",0,lcl_today)
 end if

 if request("sc_dateOptions") <> "" then
    if not containsApostrophe(request("sc_dateOptions")) then
       lcl_sc_dateOptions = clng(request("sc_dateOptions"))
    end if
 end if

 if request("sc_rentalname") <> "" then
    lcl_sc_rentalname = request("sc_rentalname")
 end if

 if request("sc_renter") <> "" then
    lcl_sc_renter = request("sc_renter")
 end if

 if request("sc_location") <> "" then
    lcl_sc_location = request("sc_location")
 end if

'If the "end" date (toDate) is earlier then the "start" date (fromDate)
'then switch the dates in the fields.
 if CDATE(lcl_sc_toDate) < CDATE(lcl_sc_fromDate) then
    lcl_sc_fromDate_original = lcl_sc_fromDate
    lcl_sc_toDate_original   = lcl_sc_toDate

    lcl_sc_fromDate = lcl_sc_toDate_original
    lcl_sc_toDate   = lcl_sc_fromDate_original
 end if

 if request("reservationtypeid") <> "" then
    if not containsApostrophe(request("reservationtypeid")) then
      	iReservationTypeId = request("reservationtypeid")
    end if
 end if

'Check for org features
 lcl_orghasfeature_edit_reservations = orghasfeature("edit reservations")
%>
<html>
<head>
  <title>E-Gov Administration Console { Rental Summary Report }</title>

	 <link rel="stylesheet" type="text/css" href="reporting.css" />
	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
	 <link rel="stylesheet" type="text/css" href="rentalsstyles.css" />
	 <link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

	 <script type="text/javascript" src="../scripts/getdates.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
	 <script type="text/javascript" src="../prototype/prototype-1.6.0.2.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.1.min.js"></script>

<script type="text/javascript">
<!--
jQuery.noConflict();

jQuery(document).ready(function(){
  jQuery('#printButton').click(function() {
     window.print();
  });

  jQuery('#exportButton').click(function() {
     location.href = 'rentalsummaryreport_export.asp';
  });

  jQuery('#searchButton').click(function() {
   		var daterege         = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
     var dateFromOk       = true;
     var dateToOk         = true;
     var lcl_return_false = Number(0);

     if(jQuery('#sc_fromDate').val() != '') {
      		dateFromOk = daterege.test(jQuery('#sc_fromDate').val());
     }

     if(jQuery('#sc_toDate').val() != '') {
      		dateToOk = daterege.test(jQuery('#sc_toDate').val());
     }

   		if(! dateFromOk ) {
        jQuery('#sc_fromDate').focus();
        inlineMsg(document.getElementById("sc_fromDateCalPop").id,'<strong>Invalid Value: </strong> The "From Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'sc_fromDateCalPop');
        lcl_return_false = lcl_return_false + 1;
     }else{
        clearMsg("sc_fromDateCalPop");
     }

   		if(! dateToOk ) {
        jQuery('#toDate').focus();
        inlineMsg(document.getElementById("sc_toDateCalPop").id,'<strong>Invalid Value: </strong> The "To Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'sc_toDateCalPop');
        lcl_return_false = lcl_return_false + 1;
     }else{
        clearMsg("sc_toDateCalPop");
     }

     if(lcl_return_false > 0) {
        return false;
     } else {
        jQuery('#rentalSummaryReport').submit();
     }
  });

  jQuery('#sc_fromDateCalPop').click(function() {
     doCalendar('sc_fromDate');
  });

  jQuery('#sc_toDateCalPop').click(function() {
     doCalendar('sc_toDate');
  });
});

function doCalendar(ToFrom) {
  w = 350;
  h = 250;
  l = (screen.width - 350)/2;
  t = (screen.height - 350)/2;
  lcl_url  = 'calendarpicker.asp';
  lcl_url += '?p=1';
  lcl_url += '&updateform=rentalSummaryReport';
  lcl_url += '&updatefield=' + ToFrom;

  eval('window.open("' + lcl_url + '", "_calendar", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
}
//-->
</script>
</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<form name=""rentalSummaryReport"" id=""rentalSummaryReport"" action=""rentalsummaryreport.asp"" method=""post"">" & vbcrlf
  response.write "<div id=""idControls"" class=""noprint"">" & vbcrlf
  response.write "  <input type=""button"" name=""printButton"" id=""printButton"" class=""button"" value=""Print"" />" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><font size=""+1""><strong>Rental Summary Report</strong></font></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "				      <fieldset class=""fieldset"">" & vbcrlf
  response.write "					       <legend>Search Options</legend>" & vbcrlf
  response.write "					       <p>" & vbcrlf
  response.write "					       <table id=""rentalsummary_searchoptionsTable"" border=""0"" cellpadding=""2"" cellspacing=""0"">" & vbcrlf
  response.write "						        <tr>" & vbcrlf
  response.write "							           <td class=""labelcolumn"">Start Date:</td>" & vbcrlf
  response.write "							           <td align=""left"">" & vbcrlf
  response.write "                      <input type=""text"" name=""sc_fromDate"" id=""sc_fromDate"" value=""" & lcl_sc_fromDate & """ size=""10"" maxlength=""10"" onchange=""clearMsg('sc_fromDateCalPop');"" />" & vbcrlf
  response.write "                      <img src=""../images/calendar.gif"" id=""sc_fromDateCalPop"" border=""0"" onclick=""clearMsg('sc_fromDateCalPop');"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
                                        DrawDateChoices "RentalSummaryReport"
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td class=""labelcolumn"">End Date:</td>" & vbcrlf
  response.write "                  <td colspan=""3"" align=""left"" style=""white-space:nowrap"">" & vbcrlf
  response.write "                      <input type=""text"" name=""sc_toDate"" id=""sc_toDate"" value=""" & lcl_sc_toDate & """ size=""10"" maxlength=""10"" onchange=""clearMsg('sc_toDateCalPop');"" />" & vbcrlf
  response.write "                      <img src=""../images/calendar.gif"" id=""sc_toDateCalPop"" border=""0"" onclick=""clearMsg('sc_toDateCalPop');"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td class=""labelcolumn"">Rental Name:</td>" & vbcrlf
  response.write "                  <td align=""left"">" & vbcrlf
  response.write "                      <input type=""text"" name=""sc_rentalname"" id=""sc_rentalname"" value=""" & lcl_sc_rentalname & """ size=""30"" maxlength=""100"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td class=""labelcolumn"">Renter:</td>" & vbcrlf
  response.write "                  <td align=""left"">" & vbcrlf
  response.write "                      <input type=""text"" name=""sc_renter"" id=""sc_renter"" value=""" & lcl_sc_renter & """ size=""30"" maxlength=""100"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td class=""labelcolumn"">Location:</td>" & vbcrlf
  response.write "                  <td align=""left"">" & vbcrlf
  response.write "                      <input type=""text"" name=""sc_location"" id=""sc_location"" value=""" & lcl_sc_location & """ size=""30"" maxlength=""100"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td class=""labelcolumn"">Reservation Types:</td>" & vbcrlf
  response.write "                  <td width=""100%"" style=""text-align:left; white-space:nowrap"">" & vbcrlf
                                        'buildReservationTypeOptions session("orgid"), _
                                        '                            iReservationTypeID

                                        buildReservationTypeCheckboxes session("orgid"), _
                                                                       iReservationTypeID
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "            </p>" & vbcrlf
  response.write "            <p>" & vbcrlf
  response.write "              <input type=""button"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" />" & vbcrlf
  response.write "            </p>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
  response.write "          <div align=""right"">" & vbcrlf
  response.write "            <input type=""button"" name=""exportButton"" id=""exportButton"" value=""Download to Excel"" class=""button"" />" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
                            displaySummary session("orgid"), _
                                           iReservationTypeId, _
                                           lcl_sc_fromDate, _
                                           lcl_sc_toDate, _
                                           lcl_sc_rentalname, _
                                           lcl_sc_renter, _
                                           lcl_sc_location
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displaySummary(iOrgID, iReservationTypeID, iSCFromDate, iSCToDate, iSCRentalName, iSCRenter, iSCLocation)

  dim sOrgID, sReservationTypeID, lcl_bgcolor, lcl_linecount
  dim lcl_previous_rentalname, lcl_previous_numberofdays, lcl_previous_totalhours, lcl_previous_numberattending
  dim lcl_subtotal_numberofdays, lcl_subtotal_totalhours, lcl_subtotal_numberattending
  dim lcl_total_numberofdays, lcl_total_totalhours, lcl_total_numberattending
  dim sSCFromDate, sSCToDate, sSCRentalName, sSCRenter, sSCLocation

  lcl_bgcolor                  = "#eeeeee"
  lcl_linecount                = 0
  lcl_previous_rentalname      = ""
  lcl_previous_numberofdays    = clng(0)
  lcl_previous_totalhours      = clng(0)
  lcl_previous_numberattending = clng(0)

  sOrgID                       = 0
  sReservationTypeID           = ""
  lcl_subtotal_numberofdays    = 0
  lcl_subtotal_totalhours      = clng(0)
  lcl_subtotal_numberattending = clng(0)
  lcl_total_numberofdays       = clng(0)
  lcl_total_totalhours         = clng(0)
  lcl_total_numberattending    = clng(0)
  sSCFromDate                  = ""
  sSCToDate                    = ""
  sSCRentalName                = ""
  sSCRenter                    = ""
  sSCLocation                  = ""
  session("RentalSummaryList") = ""

  if iOrgID <> "" then
     if not containsApostrophe(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iSCFromDate <> "" then
     if not containsApostrophe(iSCFromDate) then
        sSCFromDate = cdate(iSCFromDate)
        sSCFromDate = dbsafe(sSCFromDate)
        sSCFromDate = "'" & sSCFromDate & "'"
     end if
  end if

  if iSCToDate <> "" then
     if not containsApostrophe(iSCToDate) then
        sSCToDate = cdate(iSCToDate)
        sSCToDate = dbsafe(sSCToDate)
        sSCToDate = "'" & sSCToDate & "'"
     end if
  end if

  if iReservationTypeID <> "" then
     if not containsApostrophe(iReservationTypeID) then
        sReservationTypeID = iReservationTypeID
     end if
  end if

  if iSCRentalName <> "" then
     sSCRentalName = ucase(iSCRentalName)
     sSCRentalName = dbsafe(sSCRentalName)
     sSCRentalName = "'%" & sSCRentalName & "%'"
  end if

  if iSCRenter <> "" then
     sSCRenter = ucase(iSCRenter)
     sSCRenter = dbsafe(sSCRenter)
     sSCRenter = "'%" & sSCRenter & "%'"
  end if

  if iSCLocation <> "" then
     sSCLocation = ucase(iSCLocation)
     sSCLocation = dbsafe(sSCLocation)
     sSCLocation = "'%" & sSCLocation & "%'"
  end if

  sSQL = "SELECT r.rentalname, "
  'sSQL = sSQL & " u1.userfname + ' ' + u1.userlname as Renter, "
  sSQL = sSQL & " u1.userfname + ' ' + u1.userlname as rentername, rest.description as restype, "
  sSQL = sSQL & " (SELECT ISNULL(SUM(p.paymenttotal),0) FROM egov_class_payment P INNER JOIN egov_journal_entry_types J ON P.journalentrytypeid = J.journalentrytypeid WHERE j.journalentrytype = 'rentalpayment' AND p.reservationid = rd.reservationid) as TotalPaid, "
  sSQL = sSQL & " u1.usercity, u1.userstate, u1.userzip, "
  sSQL = sSQL & " rr.rentaluserid, "
  sSQL = sSQL & " rd.reservationid, "
  sSQL = sSQL & " cl.name as location, "
  sSQL = sSQL & " rt.reservationtype, "
  sSQL = sSQL & " (select cs.seasonname "
  sSQL = sSQL &  " from egov_class_seasons cs "
  sSQL = sSQL &  " where cs.classseasonid IN (select c.classseasonid "
  sSQL = sSQL &                             " from egov_class c "
  sSQL = sSQL &                                  " inner join egov_class_time ct on c.classid = ct.classid "
  sSQL = sSQL &                             " where c.orgid = " & sOrgID
  sSQL = sSQL &                             " and ct.rentalid = rd.rentalid "
  sSQL = sSQL &                             " and ct.reservationid = rd.reservationid)) as season, "
  'sSQL = sSQL & " (select cast((dbo.getTotalHoursByReservationID(rd.reservationid, " & sSCFromDate & ", " & sSCToDate & ")/24) as decimal(6,2))) as number_of_days, "
  sSQL = sSQL & " dbo.countDaysByReservationID(rd.reservationid, " & sSCFromDate & ", " & sSCToDate & ") as total_days, "
  sSQL = sSQL & " dbo.getTotalHoursByReservationID(rd.reservationid, " & sSCFromDate & ", " & sSCToDate & ") as total_hours, "
  sSQL = sSQL & " dbo.getTotalNumberAttendingByReservationID(rd.reservationid) as number_attending "
  sSQL = sSQL & " FROM egov_rentalreservationdates rd "
  sSQL = sSQL &      " left outer join egov_rentals r on rd.rentalid = r.rentalid "
  sSQL = sSQL &      " left outer join egov_rentalreservations rr on rd.reservationid = rr.reservationid "
  'sSQL = sSQL &      " left outer join egov_users u1 on rr.rentaluserid = u1.userid "
  sSQL = sSQL &      " left outer join egov_rentalreservationtypes rt on rr.reservationtypeid = rt.reservationtypeid "
  sSQL = sSQL &      " left outer join egov_class_location cl on r.locationid = cl.locationid "
  sSQL = sSQL & " left outer join egov_users u1 on rr.rentaluserid = u1.userid  "
  sSQL = sSQL & " LEFT JOIN egov_poolpassresidenttypes rest ON rt.orgid = rd.orgid AND u1.residenttype = rest.resident_type "
  sSQL = sSQL & " WHERE rd.orgid = " & sOrgID
  sSQL = sSQL & " AND rd.statusid NOT IN (select rs.reservationstatusid "
  sSQL = sSQL &                         " from egov_rentalreservationstatuses rs "
  sSQL = sSQL &                         " where rs.isCancelled = 1) "

  if sSCFromDate <> "" AND sSCToDate <> "" then
     sSQL = sSQL & " AND reservationstarttime >= " & sSCFromDate
     sSQL = sSQL & " AND reservationendtime <= "   & sSCToDate
  end if

  if sReservationTypeID <> "" then
     sSQL = sSQL & " AND rt.reservationtypeid IN (" & sReservationTypeID & ") "
  end if

  if sSCRentalName <> "" then
     sSQL = sSQL & " AND upper(r.rentalname) LIKE (" & sSCRentalName & ") "
  end if

  if sSCRenter <> "" then
     sSQL = sSQL & " AND upper(u1.userfname + ' ' + u1.userlname) LIKE (" & sSCRenter & ") "
  end if

  if sSCLocation <> "" then
     sSQL = sSQL & " AND upper(cl.name) LIKE (" & sSCLocation & ") "
  end if

  'sSQL = sSQL & " AND u1.FirstName + u1.LastName <> '' "
  sSQL = sSQL & " GROUP BY r.rentalname, "
  sSQL = sSQL &          " rr.rentaluserid, "
  'sSQL = sSQL &          " u1.userfname + ' ' + u1.userlname, "
  sSQL = sSQL & " u1.userfname + ' ' + u1.userlname, rest.description, "
  sSQL = sSQL & " u1.usercity, u1.userstate, u1.userzip, "
  sSQL = sSQL &          " cl.name, "
  sSQL = sSQL &          " rt.reservationtype, "
  sSQL = sSQL &          " rd.reservationid, "
  sSQL = sSQL &          " rd.rentalid "
  sSQL = sSQL & " ORDER BY  r.rentalname, "
  'sSQL = sSQL &          " u1.userfname + ' ' + u1.userlname, "
  sSQL = sSQL &          " cl.name, "
  sSQL = sSQL &          " rt.reservationtype "
  'response.write sSQL & "<p>&nbsp;</p>" & vbcrlf

  session("RentalSummaryList") = sSQL

 	set oRentalSummaryList = Server.CreateObject("ADODB.Recordset")
	 oRentalSummaryList.Open sSQL, Application("DSN"), 3, 1

  if not oRentalSummaryList.eof then
     response.write "<table id=""rentalSummaryTable"" cellpadding=""2"" cellspacing=""0"" border=""0"">" & vbcrlf
     response.write "  <tr align=""left"" valign=""bottom"">" & vbcrlf
     response.write "      <th>Rental Name</th>" & vbcrlf
     response.write "      <th>Renter</th>" & vbcrlf
     response.write "      <th>Resident Type</th>" & vbcrlf
     response.write "      <th>Total Paid</th>" & vbcrlf
     response.write "      <th align=""center"">Reservation<br />ID</th>" & vbcrlf
     response.write "      <th>Location</th>" & vbcrlf
     response.write "      <th>Reservation Type</th>" & vbcrlf
     response.write "      <th>Season</th>" & vbcrlf
     response.write "      <th align=""center""># of<br />Days</th>" & vbcrlf
     response.write "      <th align=""center"">Total<br />Hours</th>" & vbcrlf
     response.write "      <th align=""center"">Number<br />Attending</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     do while not oRentalSummaryList.eof

        lcl_linecount        = lcl_linecount + 1
        lcl_bgcolor          = changeBGColor(lcl_bgcolor, "#eeeeee", "#ffffff")
        lcl_style_row        = ""
        lcl_rentalname       = "&nbsp;"
        lcl_rentaluserid     = 0
        lcl_renter           = "&nbsp;"
        lcl_reservationid    = "&nbsp;"
        lcl_location         = "&nbsp;"
        lcl_reservationtype  = "&nbsp;"
        lcl_season           = "&nbsp;"
        lcl_total_days       = 0
        lcl_total_hours      = 0
        lcl_number_attending = 0

        if oRentalSummaryList("rentalname") <> "" then
           lcl_rentalname = oRentalSummaryList("rentalname")
        end if

        'if oRentalSummaryList("renter") <> "" then
        '   lcl_renter = oRentalSummaryList("renter")
        'end if

        if oRentalSummaryList("rentaluserid") <> "" then
           lcl_rentaluserid = oRentalSummaryList("rentaluserid")
        end if

        if oRentalSummaryList("reservationid") <> "" then
           lcl_reservationid = oRentalSummaryList("reservationid")
        end if

        if oRentalSummaryList("location") <> "" then
           lcl_location = oRentalSummaryList("location")
        end if

        if oRentalSummaryList("reservationtype") <> "" then
           lcl_reservationtype = oRentalSummaryList("reservationtype")
        end if

        if oRentalSummaryList("season") <> "" then
           lcl_season = oRentalSummaryList("season")
        end if

        if oRentalSummaryList("total_days") <> "" then
           lcl_total_days = oRentalSummaryList("total_days")
        end if

        if oRentalSummaryList("total_hours") <> "" then
           lcl_total_hours = oRentalSummaryList("total_hours")
        end if

        'if lcl_total_hours > 0 then
        '   lcl_total_days = lcl_total_hours / 24

        '   lcl_number_of_days = FormatNumber(lcl_total_days,,-1)
        'end if

        if oRentalSummaryList("number_attending") <> "" then
           lcl_number_attending = oRentalSummaryList("number_attending")
        end if

       'Calculate the sub-totals
        lcl_subtotal_numberofdays    = lcl_subtotal_numberofdays    + lcl_previous_numberofdays
        lcl_subtotal_totalhours      = lcl_subtotal_totalhours      + lcl_previous_totalhours
        lcl_subtotal_numberattending = lcl_subtotal_numberattending + lcl_previous_numberattending

       'Get the correct Renter Name
        if lcl_reservationtype <> "" then
           'lcl_renter = getRenterName(lcl_reservationtype, lcl_rentaluserid)
	   lcl_renter = oRentalSummaryList("rentername")
        end if

       'Determine if a new "rental name" (row) is to be started.
        if lcl_rentalname = lcl_previous_rentalname then
           lcl_rentalname = "&nbsp;"
        else
           lcl_rentalname = "<strong>" & lcl_rentalname & "</strong>"

           if lcl_linecount > 1 then
              lcl_style_row = " class=""rentalSummary_newRow"""

              response.write "  <tr class=""rentalSummary_subTotalRow"">" & vbcrlf
              response.write "      <td colspan=""8"" style=""text-align:right"">" & lcl_previous_rentalname & " Sub-Totals:&nbsp;</td>" & vbcrlf
              response.write "      <td>" & lcl_subtotal_numberofdays    & "</td>" & vbcrlf
              response.write "      <td>" & lcl_subtotal_totalhours      & "</td>" & vbcrlf
              response.write "      <td>" & lcl_subtotal_numberattending & "</td>" & vbcrlf
              response.write "  </tr>" & vbcrlf
           end if

          'Calculate the totals and reset the sub-totals
           lcl_total_numberofdays    = lcl_total_numberofdays    + lcl_subtotal_numberofdays
           lcl_total_totalhours      = lcl_total_totalhours      + lcl_subtotal_totalhours
           lcl_total_numberattending = lcl_total_numberattending + lcl_subtotal_numberattending

           lcl_subtotal_numberofdays    = 0
           lcl_subtotal_totalhours      = 0
           lcl_subtotal_numberattending = 0

        end if

       'Build the reservationid URL
        lcl_reservationid_url = lcl_reservationid

        if lcl_orghasfeature_edit_reservations then
           lcl_reservationid_url = "<a href=""reservationedit.asp?reservationid=" & lcl_reservationid_url & """>" & lcl_reservationid & "</a>"
        end if

        response.write "  <tr align=""left"" valign=""top"" bgcolor=""" & lcl_bgcolor & """" & lcl_style_row & ">" & vbcrlf
        response.write "      <td class=""td_nowrap"">" & lcl_rentalname        & "</td>" & vbcrlf
        response.write "      <td class=""td_nowrap"">" & lcl_renter            & "</td>" & vbcrlf
        response.write "      <td>"                     & oRentalSummaryList("restype")          & "</td>" & vbcrlf
        response.write "      <td>"                     & FormatCurrency(oRentalSummaryList("TotalPaid"))          & "</td>" & vbcrlf
        response.write "      <td align=""center"">"    & lcl_reservationid_url & "</td>" & vbcrlf
        response.write "      <td>"                     & lcl_location          & "</td>" & vbcrlf
        response.write "      <td class=""td_nowrap"">" & lcl_reservationtype   & "</td>" & vbcrlf
        response.write "      <td>"                     & lcl_season            & "</td>" & vbcrlf
        response.write "      <td align=""center"">"    & lcl_total_days        & "</td>" & vbcrlf
        response.write "      <td align=""center"">"    & lcl_total_hours       & "</td>" & vbcrlf
        response.write "      <td align=""center"">"    & lcl_number_attending  & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        lcl_previous_rentalname      = oRentalSummaryList("rentalname")
        lcl_previous_numberofdays    = clng(oRentalSummaryList("total_days"))
        lcl_previous_totalhours      = clng(oRentalSummaryList("total_hours"))
        lcl_previous_numberattending = clng(oRentalSummaryList("number_attending"))

        oRentalSummaryList.movenext
     loop

     if lcl_linecount > 0 then

       'Calculate the sub-totals and totals
        lcl_subtotal_numberofdays    = lcl_subtotal_numberofdays    + lcl_previous_numberofdays
        lcl_subtotal_totalhours      = lcl_subtotal_totalhours      + lcl_previous_totalhours
        lcl_subtotal_numberattending = lcl_subtotal_numberattending + lcl_previous_numberattending

        lcl_total_numberofdays    = lcl_total_numberofdays    + lcl_subtotal_numberofdays
        lcl_total_totalhours      = lcl_total_totalhours      + lcl_subtotal_totalhours
        lcl_total_numberattending = lcl_total_numberattending + lcl_subtotal_numberattending

        response.write "  <tr class=""rentalSummary_subTotalRow"">" & vbcrlf
        response.write "      <td colspan=""8"" style=""text-align:right"">" & lcl_previous_rentalname & " Sub-Totals:&nbsp;</td>" & vbcrlf
        response.write "      <td>" & lcl_subtotal_numberofdays    & "</td>" & vbcrlf
        response.write "      <td>" & lcl_subtotal_totalhours      & "</td>" & vbcrlf
        response.write "      <td>" & lcl_subtotal_numberattending & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "  <tr class=""rentalSummary_subTotalRow"">" & vbcrlf
        response.write "      <td colspan=""8"" style=""text-align:right"">TOTALS:&nbsp;</td>" & vbcrlf
        response.write "      <td>" & lcl_total_numberofdays    & "</td>" & vbcrlf
        response.write "      <td>" & lcl_total_totalhours      & "</td>" & vbcrlf
        response.write "      <td>" & lcl_total_numberattending & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
     end if

     response.write "</table>" & vbcrlf

  end if

  oRentalSummaryList.close
  set oRentalSummaryList = nothing

end sub

'------------------------------------------------------------------------------
sub ShowFloatingYearPicks(iSelectedYear, sPickName)

 	dim iYear, lcl_selected_year

 	response.write "<select name=""" & sPickName & """ id=""" & sPickName & """>"

 	for iYear=Year(Date())-5 to Year(Date())+5
     lcl_selected_year = ""

     if clng(iSelectedYear) = clng(iYear) then
        lcl_selected_year = " selected=""selected"""
     end if

     response.write "  <option value=""" & iYear & """" & lcl_selected_year & ">" & iYear & "</option>" & vbcrlf
  next

	 response.write "</select>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub ShowMonthPicks(iSelectedMonth, sPickName)

  dim iMonth, lcl_selected_month

  response.write "<select name=""" & sPickName & """ id=""" & sPickName & """>" & vbcrlf

  for iMonth = 1 to 12
     lcl_selected_month = ""

     if clng(iSelectedMonth) = clng(iMonth) then
        lcl_selected_month = " selected=""selected"""
     end if

   		response.write "  <option value=""" & iMonth & """" & lcl_selected_month & ">" & MonthName(iMonth) & "</option>" & vbcrlf
  next

  response.write "</select>" & vbcrlf

end sub

'--------------------------------------------------------------------------------------------------
sub buildReservationTypeCheckboxes(iOrgID, iReservationTypeID)

  dim sSQL, sOrgID, sReservationTypeID, sLineCount

  sLineCount                 = 0
  sOrgID                     = 0
  sReservationTypeID         = ""
  lcl_checkedReservationType = ""

  if iOrgID <> "" then
     if isnumeric(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iReservationTypeID <> "" then
     if not containsApostrophe(iReservationTypeID) then
        sReservationTypeID = iReservationTypeID
     end if
  end if

 	sSQL = "SELECT reservationtypeid, "
  sSQL = sSQL & " reservationtype "
  sSQL = sSQL & " FROM egov_rentalreservationtypes "
  sSQL = sSQL & " WHERE orgid = " & sOrgID
	 sSQL = sSQL & " ORDER BY displayorder "

	 set oBuildReservationTypes = Server.CreateObject("ADODB.Recordset")
 	oBuildReservationTypes.Open sSQL, Application("DSN"), 0, 1

 	'response.write "<select name=""reservationtypeid"" id=""reservationtypeid"">" & vbcrlf
 	'response.write "  <option value=""0"">All Reservation Types</option>" & vbcrlf

 	if not oBuildReservationTypes.eof then
     do while not oBuildReservationTypes.eof
        sLineCount = sLineCount + 1

        if sReservationTypeID <> "" then
           lcl_checked_reservationtype = isCheckedReservationType(iOrgID, oBuildReservationTypes("reservationtypeid"), sReservationTypeID)
        end if

        'if clng(oBuildReservationTypes("reservationtypeid")) = sReservationTypeID then
        '   lcl_selected_reservationtype = " selected=""selected"""
        'end if

        'response.write "  <option value=""" & oBuildReservationTypes("reservationtypeid") & """" & lcl_selected_reservationtype & ">" & oBuildReservationTypes("reservationtype") & "</option>" & vbcrlf
        response.write "<input type=""checkbox"" name=""reservationtypeid"" id=""reservationtypeid" & sLineCount & """ value=""" & oBuildReservationTypes("reservationtypeid") & """" & lcl_checked_reservationtype & "> " & oBuildReservationTypes("reservationtype") & "&nbsp&nbsp;" & vbcrlf

     			oBuildReservationTypes.movenext
     loop
  end if

 	response.write "</select>" & vbcrlf

	 oBuildReservationTypes.close
	 set oBuildReservationTypes = nothing

end sub

'------------------------------------------------------------------------------
function isCheckedReservationType(iOrgID, iCurrentReservationTypeID, iReservationTypeIDs)
  dim lcl_return, sSQL, sOrgID, sCurrentReservationTypeID, sReservationTypeIDs

  sOrgID             = 0
  sReservationTypeID = ""
  lcl_return         = ""

  if iOrgID <> "" then
     if isnumeric(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iCurrentReservationTypeID <> "" then
     if not containsApostrophe(iCurrentReservationTypeID) then
        sCurrentReservationTypeID = iCurrentReservationTypeID
     end if
  end if

  if iReservationTypeIDs <> "" then
     if not containsApostrophe(iReservationTypeIDs) then
        sReservationTypeIDs = iReservationTypeIDs
     end if
  end if

 	sSQL = "SELECT distinct reservationtypeid "
  sSQL = sSQL & " FROM egov_rentalreservationtypes "
  sSQL = sSQL & " WHERE orgid = " & sOrgID
  sSQL = sSQL & " AND reservationtypeid IN (" & sReservationTypeIDs & ") "
  sSQL = sSQL & " aND reservationtypeid = "   & sCurrentReservationTypeID

	 set oIsCheckedReservationType = Server.CreateObject("ADODB.Recordset")
 	oIsCheckedReservationType.Open sSQL, Application("DSN"), 0, 1

 	if not oIsCheckedReservationType.eof then
     if oIsCheckedReservationType("reservationtypeid") <> "" then
        lcl_return = " checked=""checked"""
     end if
  end if

  oIsCheckedReservationType.close
  set oIsCheckedReservationType = nothing

  isCheckedReservationType = lcl_return

end function

'------------------------------------------------------------------------------
function getRenterName(iReservationType, iRentalUserID)

  dim lcl_return, lcl_db_column, lcl_db_table, sSQL
  dim sReservationType, sRentalUserID

  lcl_return       = ""
  lcl_db_column    = "firstname + ' ' + lastname"
  lcl_db_table     = "users"
  sReservationType = ""
  sRentalUserID    = 0
  sSQL             = ""

  if iReservationType <> "" then
     sReservationType = ucase(iReservationType)
  end if

  if iRentalUserID <> "" then
     sRentalUserID = clng(iRentalUserID)
  end if

  if sReservationType = "PUBLIC RESERVATION" then
     lcl_db_column = "userfname + ' ' + userlname"
     lcl_db_table  = "egov_users"
  end if

  sSQL = "SELECT " & lcl_db_column & " as renter "
  sSQL = sSQL & " FROM " & lcl_db_table
  sSQL = sSQL & " WHERE userid = " & sRentalUserID

	 set oGetRenterName = Server.CreateObject("ADODB.Recordset")
 	oGetRenterName.Open sSQL, Application("DSN"), 0, 1

  if not oGetRenterName.eof then
     lcl_return = oGetRenterName("renter")
  end if

  oGetRenterName.close
  set oGetRenterName = nothing

  getRenterName = lcl_return

end function

'------------------------------------------------------------------------------
'sub buildReservationTypeOptions(iOrgID, iReservationTypeID)

'  dim sSQL, sOrgID, sReservationTypeID

'  if iOrgID <> "" then
'     if isnumeric(iOrgID) then
'        sOrgID = clng(iOrgID)
'     end if
'  end if

'  if iReservationTypeID <> "" then
'     if isnumeric(iReservationTypeID) then
'        sReservationTypeID = clng(iReservationTypeID)
'     end if
'  end if

' 	sSQL = "SELECT reservationtypeid, "
'  sSQL = sSQL & " reservationtype "
'  sSQL = sSQL & " FROM egov_rentalreservationtypes "
'  sSQL = sSQL & " WHERE orgid = " & iOrgID
'	 sSQL = sSQL & " ORDER BY displayorder "

'	 set oBuildReservationTypes = Server.CreateObject("ADODB.Recordset")
' 	oBuildReservationTypes.Open sSQL, Application("DSN"), 0, 1

' 	response.write "<select name=""reservationtypeid"" id=""reservationtypeid"">" & vbcrlf
' 	response.write "  <option value=""0"">All Reservation Types</option>" & vbcrlf

' 	if not oBuildReservationTypes.eof then
'     do while not oBuildReservationTypes.eof
'        lcl_selected_reservationtype = ""

'        if clng(oBuildReservationTypes("reservationtypeid")) = sReservationTypeID then
'           lcl_selected_reservationtype = " selected=""selected"""
'        end if

'        response.write "  <option value=""" & oBuildReservationTypes("reservationtypeid") & """" & lcl_selected_reservationtype & ">" & oBuildReservationTypes("reservationtype") & "</option>" & vbcrlf

'     			oBuildReservationTypes.movenext
'     loop
'  end if

' 	response.write "</select>" & vbcrlf

'	 oBuildReservationTypes.close
'	 set oBuildReservationTypes = nothing

'end sub
%>
