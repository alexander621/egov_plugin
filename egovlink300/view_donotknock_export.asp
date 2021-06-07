<%
'SET UP PAGE OPTIONS
' sDate = Month(Date()) & Day(Date()) & Year(Date())
 sDate = year(date()) & month(date()) & day(date())
 sTime = hour(time()) & minute(time()) & second(time())
 server.scripttimeout = 4800
 response.ContentType = "application/msexcel"
 response.AddHeader "Content-Disposition", "attachment;filename=DoNotKnock_Export_" & sDate & "_" & sTime & ".XLS"

 lcl_canViewPeddlers   = request("vp")
 lcl_canViewSolicitors = request("vs")

 buildExport lcl_canViewPeddlers, lcl_canViewSolicitors

'------------------------------------------------------------------------------
sub buildExport(iCanViewPeddlers, iCanViewSolicitors)

  set oDoNotKnockExport = Server.CreateObject("ADODB.Recordset")

  sSQL = session("DONOTKNOCK_QUERY")
  oDoNotKnockExport.Open sSQL, Application("DSN"), 3, 1

  if not oDoNotKnockExport.eof then

     if iCanViewPeddlers then
        if iCanViewSolicitors then
           lcl_title = "Peddlers/Solicitors ""Do Not Knock"" Lists"
        else
           lcl_title = "Peddlers ""Do Not Knock"" List"
        end if
     else
        if iCanViewSolicitors then
           lcl_title = "Solicitors ""Do Not Knock"" List"
        end if
     end if

     response.write lcl_title & "<br /><br />" & vbcrlf

     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
     response.write "  <tr align=""center"">" & vbcrlf
     response.write "      <td nowrap=""nowrap"">Street Number</td>" & vbcrlf
     response.write "      <td nowrap=""nowrap"" align=""left"">Street Address</td>" & vbcrlf
     response.write "      <td>Resident Unit</td>" & vbcrlf

     if iCanViewPeddlers then
        response.write "      <td>On Peddlers List</td>" & vbcrlf
     end if

     if iCanViewSolicitors then
        response.write "      <td>On Solicitors List</td>" & vbcrlf
     end if

     response.write "  </tr>" & vbcrlf
     response.flush

     do while not oDoNotKnockExport.eof

        response.write "  <tr align=""center"">" & vbcrlf
        response.write "      <td>" & oDoNotKnockExport("db_streetnumber") & "</td>" & vbcrlf
        response.write "      <td align=""left"">" & oDoNotKnockExport("db_streetname") & "</td>" & vbcrlf
        response.write "      <td>" & oDoNotKnockExport("userunit")                     & "</td>" & vbcrlf

        if iCanViewPeddlers then
           if oDoNotKnockExport("isOnDoNotKnockList_peddlers") then
              lcl_display_peddlers = "YES"
           else
              lcl_display_peddlers = "NO"
           end if

           response.write "      <td>" & lcl_display_peddlers & "</td>" & vbcrlf
        end if

        if iCanViewSolicitors then
           if oDoNotKnockExport("isOnDoNotKnockList_solicitors") then
              lcl_display_solicitors = "YES"
           else
              lcl_display_solicitors = "NO"
           end if

           response.write "      <td>" & lcl_display_solicitors & "</td>" & vbcrlf
        end if

        response.write "  </tr>" & vbcrlf
	response.flush

        oDoNotKnockExport.movenext
     loop

     oDoNotKnockExport.close
     set oDoNotKnockExport = nothing

     response.write "</table>"
     response.flush

  end if
end sub
%>
