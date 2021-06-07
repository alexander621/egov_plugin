<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: rentalsummaryreport_export.asp
' AUTHOR: David Boyer
' CREATED: 05/31/2012
' COPYRIGHT: Copyright 2012 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Export of the Rental Summary Report
'
' MODIFICATION HISTORY
' 1.0 05/31/2012	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
	 dim sSQL, oBuildRentalExport, sDate, sSearch, bIsArchive

 'Set up page options
	 sDate                = right("0" & month(date()),2) & right("0" & day(date()),2) & year(date())
	 server.scripttimeout = 9000
	 response.ContentType = "application/vnd.ms-excel"
	 response.AddHeader "Content-Disposition", "attachment;filename=RentalSummary_export_" & sDate & ".xls"

	 sSQL = session("RentalSummaryList")

 	set oBuildRentalExport = Server.CreateObject("ADODB.Recordset")
	 oBuildRentalExport.Open sSQL, Application("DSN"), 3, 1

  if not oBuildRentalExport.eof then

     response.write "<html>" & vbcrlf
     response.write "<body>" & vbcrlf
     response.write "<table id=""rentalSummaryTable"" cellpadding=""2"" cellspacing=""0"" border=""0"">" & vbcrlf
     response.write "  <tr align=""left"" valign=""bottom"">" & vbcrlf
     response.write "      <th>Rental Name</th>" & vbcrlf
     response.write "      <th>Renter</th>" & vbcrlf
     response.write "      <th>Resident Type</th>" & vbcrlf
     response.write "      <th>City</th>" & vbcrlf
     response.write "      <th>State</th>" & vbcrlf
     response.write "      <th>Zip</th>" & vbcrlf
     response.write "      <th>Total Paid</th>" & vbcrlf
     response.write "      <th align=""center"">Reservation<br />ID</th>" & vbcrlf
     response.write "      <th>Location</th>" & vbcrlf
     response.write "      <th>Reservation Type</th>" & vbcrlf
     response.write "      <th>Season</th>" & vbcrlf
     response.write "      <th align=""center""># of<br />Days</th>" & vbcrlf
     response.write "      <th align=""center"">Total<br />Hours</th>" & vbcrlf
     response.write "      <th align=""center"">Number<br />Attending</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.flush

     do while not oBuildRentalExport.eof

        lcl_linecount        = lcl_linecount + 1
        lcl_bgcolor          = changeBGColor(lcl_bgcolor, "#eeeeee", "#ffffff")
        lcl_style_row        = ""
        lcl_rentalname       = "&nbsp;"
        lcl_renter           = "&nbsp;"
        lcl_rentaluserid     = 0
        lcl_reservationid    = "&nbsp;"
        lcl_location         = "&nbsp;"
        lcl_reservationtype  = "&nbsp;"
        lcl_season           = "&nbsp;"
        lcl_total_days       = 0
        lcl_total_hours      = 0
        lcl_number_attending = 0

        if oBuildRentalExport("rentalname") <> "" then
           lcl_rentalname = oBuildRentalExport("rentalname")
        end if

        'if oBuildRentalExport("renter") <> "" then
        '   lcl_renter = oBuildRentalExport("renter")
        'end if

        if oBuildRentalExport("rentaluserid") <> "" then
           lcl_rentaluserid = oBuildRentalExport("rentaluserid")
        end if

        if oBuildRentalExport("reservationid") <> "" then
           lcl_reservationid = oBuildRentalExport("reservationid")
        end if

        if oBuildRentalExport("location") <> "" then
           lcl_location = oBuildRentalExport("location")
        end if

        if oBuildRentalExport("reservationtype") <> "" then
           lcl_reservationtype = oBuildRentalExport("reservationtype")
        end if

        if oBuildRentalExport("season") <> "" then
           lcl_season = oBuildRentalExport("season")
        end if

        if oBuildRentalExport("total_days") <> "" then
           lcl_total_days = oBuildRentalExport("total_days")
        end if

        if oBuildRentalExport("total_hours") <> "" then
           lcl_total_hours = oBuildRentalExport("total_hours")
        end if

        'if lcl_total_hours > 0 then
        '   lcl_total_days = lcl_total_hours / 24

        '   lcl_number_of_days = FormatNumber(lcl_total_days,,-1)
        'end if

        if oBuildRentalExport("number_attending") <> "" then
           lcl_number_attending = oBuildRentalExport("number_attending")
        end if

       'Calculate the sub-totals
        lcl_subtotal_numberofdays    = lcl_subtotal_numberofdays    + lcl_previous_numberofdays
        lcl_subtotal_totalhours      = lcl_subtotal_totalhours      + lcl_previous_totalhours
        lcl_subtotal_numberattending = lcl_subtotal_numberattending + lcl_previous_numberattending

       'Get the correct Renter Name
        if lcl_reservationtype <> "" then
           'lcl_renter = getRenterName(lcl_reservationtype, lcl_rentaluserid)
	   lcl_renter = oBuildRentalExport("rentername")
        end if

       'Determine if a new "rental name" (row) is to be started.
        if lcl_rentalname = lcl_previous_rentalname then
           lcl_rentalname = "&nbsp;"
        else
           lcl_rentalname = "<strong>" & lcl_rentalname & "</strong>"

           if lcl_linecount > 1 then
              lcl_style_row = " class=""rentalSummary_newRow"""

              response.write "  <tr class=""rentalSummary_subTotalRow"">"
              response.write "      <td colspan=""11"" style=""text-align:right"">" & lcl_previous_rentalname & " Sub-Totals:&nbsp;</td>"
              response.write "      <td>" & lcl_subtotal_numberofdays    & "</td>"
              response.write "      <td>" & lcl_subtotal_totalhours      & "</td>"
              response.write "      <td>" & lcl_subtotal_numberattending & "</td>"
              response.write "  </tr>"
	      response.flush
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

        response.write "  <tr align=""left"" valign=""top"" bgcolor=""" & lcl_bgcolor & """" & lcl_style_row & ">"
        response.write "      <td class=""td_nowrap"">" & lcl_rentalname        & "</td>"
        response.write "      <td class=""td_nowrap"">" & lcl_renter            & "</td>"
        response.write "      <td>"                     & oBuildRentalExport("restype")          & "</td>" & vbcrlf
        response.write "      <td>"                     & oBuildRentalExport("usercity")          & "</td>" & vbcrlf
        response.write "      <td>"                     & oBuildRentalExport("userstate")          & "</td>" & vbcrlf
        response.write "      <td>"                     & oBuildRentalExport("userzip")          & "</td>" & vbcrlf
        response.write "      <td>"                     & FormatCurrency(oBuildRentalExport("TotalPaid"))          & "</td>" & vbcrlf
        response.write "      <td align=""center"">"    & lcl_reservationid_url & "</td>"
        response.write "      <td>"                     & lcl_location          & "</td>"
        response.write "      <td class=""td_nowrap"">" & lcl_reservationtype   & "</td>"
        response.write "      <td>"                     & lcl_season            & "</td>"
        response.write "      <td align=""center"">"    & lcl_total_days        & "</td>"
        response.write "      <td align=""center"">"    & lcl_total_hours       & "</td>"
        response.write "      <td align=""center"">"    & lcl_number_attending  & "</td>"
        response.write "  </tr>"
	response.flush

        lcl_previous_rentalname      = oBuildRentalExport("rentalname")
        lcl_previous_numberofdays    = clng(oBuildRentalExport("total_days"))
        lcl_previous_totalhours      = clng(oBuildRentalExport("total_hours"))
        lcl_previous_numberattending = clng(oBuildRentalExport("number_attending"))

        oBuildRentalExport.movenext
     loop

     if lcl_linecount > 0 then

       'Calculate the sub-totals and totals
        lcl_subtotal_numberofdays    = lcl_subtotal_numberofdays    + lcl_previous_numberofdays
        lcl_subtotal_totalhours      = lcl_subtotal_totalhours      + lcl_previous_totalhours
        lcl_subtotal_numberattending = lcl_subtotal_numberattending + lcl_previous_numberattending

        lcl_total_numberofdays    = lcl_total_numberofdays    + lcl_subtotal_numberofdays
        lcl_total_totalhours      = lcl_total_totalhours      + lcl_subtotal_totalhours
        lcl_total_numberattending = lcl_total_numberattending + lcl_subtotal_numberattending

        response.write "  <tr class=""rentalSummary_subTotalRow"">"
        response.write "      <td colspan=""11"" style=""text-align:right"">" & lcl_previous_rentalname & " Sub-Totals:&nbsp;</td>"
        response.write "      <td>" & lcl_subtotal_numberofdays    & "</td>"
        response.write "      <td>" & lcl_subtotal_totalhours      & "</td>"
        response.write "      <td>" & lcl_subtotal_numberattending & "</td>"
        response.write "  </tr>"
        response.write "  <tr class=""rentalSummary_subTotalRow"">"
        response.write "      <td colspan=""11"" style=""text-align:right"">TOTALS:&nbsp;</td>"
        response.write "      <td>" & lcl_total_numberofdays    & "</td>"
        response.write "      <td>" & lcl_total_totalhours      & "</td>"
        response.write "      <td>" & lcl_total_numberattending & "</td>"
        response.write "  </tr>"
	response.flush
     end if

     response.write "</table>"
     response.write "</body>"
     response.write "</html>"

  end if

  oBuildRentalExport.close
  set oBuildRentalExport = nothing

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
%>
