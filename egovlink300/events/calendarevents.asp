<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: calendarevents.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Calendar event display.
'
' MODIFICATION HISTORY
'	1.?	???      ??? - INITIAL VERSION.
' 1.2 08/21/08 David Boyer - Added Custom Calendar
' 1.3  11/19/13  Terry Foster - CLng Bug Fix
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim iSectionID, sDocumentTitle, dDate, oCmd, oRst, sEvents, bShowiCal, lcl_colspan_datetime

 session("RedirectPage") = request.servervariables("SCRIPT_NAME") & "?" & request.queryString()
 session("RedirectLang") = "Return to Calendar"

'Check to see if this is a custom calendar
 if trim(request("cal")) <> "" and isnumeric(replace(trim(request("cal")),"'","")) then
    'lcl_calendarfeature        = trim(request("cal"))
    'lcl_calendarfeature_url    = "&cal=" & lcl_calendarfeature
    lcl_calendarfeatureid      = CLng(replace(trim(request("cal")),"'",""))
    lcl_calendarfeature        = getFeatureByID(iorgid, lcl_calendarfeatureid)
    lcl_calendarfeature_url    = "&cal=" & lcl_calendarfeatureid
    lcl_calendarfeature_name   = " [" & getFeatureName(lcl_calendarfeature) & "]"
    lcl_displayHistory_feature = "displayhistoryinfo_customcalendars"
 else
    lcl_calendarfeature        = ""
    lcl_calendarfeature_url    = ""
    lcl_calendarfeature_name   = ""
    lcl_displayHistory_feature = "displayhistoryinfo"
 end if

 If Request("date") <> "" Then 
   	dDate = Request("date")
    dDate = Replace(dDate, "/", "")
 Else
   	dDate = ""
 End If 

 If IsDate(dDate) Then
    dDate = CDate(dDate)
 Else
    If IsDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year")) Then
       dDate = CDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year"))
    Else
       dDate = Date()
    End If
 End If
 'response.write dDate
 'response.end

 'Set the flag to show the iCal export
 bShowiCal = OrgHasFeature( iOrgId, "public ical export" )

'Check for org features
 lcl_orghasfeature_displayHistoryInfo = orghasfeature(iOrgID, lcl_displayHistory_feature)

'Build title
 if iorgid = 7 then
    lcl_title = sOrgName
 else
    lcl_title = "E-Gov Services " & sOrgName
 end if

 
%>
<html>
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	<meta name="viewport" content="width=device-width, initial-scale=1" />

<%
	response.write "<title>" & lcl_title & "</title>" & vbcrlf
%>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="calendar.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="javascript" src="../scripts/modules.js"></script>

	<script type="text/javascript">var addthis_config = {"data_track_clickback":true};</script>
	<script type="text/javascript" src="https://s7.addthis.com/js/250/addthis_widget.js#pubid=egovlink"></script>

 <!-- <script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script> -->
 <!-- <script type="text/javascript">var addthis_pub="cschappacher";</script> -->

	<script language="javascript">
	<!--

		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}

	//-->
	</script>

<style type="text/css">
   table#historyInfo,
   table#historyInfo td {
      border:        0pt solid #000000;
      margin-top:    5px;
   }
</style>
</head>

<!--#include file="../include_top.asp"-->
<%

if bShowiCal then
	lcl_colspan_datetime = " colspan=""2"""
else
	lcl_colspan_datetime = ""
end If

response.write "<p>" & vbcrlf
response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""90%"">" & vbcrlf
response.write "    <tr valign=""top"">" & vbcrlf
response.write "        <td>" & vbcrlf
response.write "            <font class=""pagetitle"">Calendar of Events" & lcl_calendarfeature_name & "</font>" & vbcrlf
						  checkForRSSFeed iorgid, "", "", "COMMUNITYCALENDAR", sEgovWebsiteURL
response.write "            <br />" & vbcrlf
						  RegisteredUserDisplay "../"
response.write "        </td>" & vbcrlf
response.write "        <td align=""right"">" & vbcrlf
						  displayAddThisButtonNew iorgid
response.write "        </td>" & vbcrlf
response.write "    </tr>" & vbcrlf
response.write "  </table>" & vbcrlf
response.write "</p>" & vbcrlf

response.write "<p><img src=""/" & sorgVirtualSiteName & "/images/arrow_2back.gif"" align=""absmiddle"">&nbsp;" & vbcrlf
response.write "  <a href=""calendar.asp?day=" & day(ddate) & "&month=" & month(ddate) & "&year=" & year(ddate) & lcl_calendarfeature_url & """>" & langBackToCalendar & "</a></p>" & vbcrlf

response.write "<div class=""calendartitle"">"
on error resume next
response.write    langEvents & ": " & FormatDateTime(dDate, vbLongDate) 
if err.number <> 0 then
       dDate = Date()
	response.write    langEvents & ": " & FormatDateTime(dDate, vbLongDate) 
end if
on error goto 0
response.write "</div>" & vbcrlf

response.write "<p>" & vbcrlf
response.write "  <table id=""eventlist"" border=""0"" cellpadding=""4"" cellspacing=""0"">" & vbcrlf
response.write "    <tr>" & vbcrlf
response.write "      <th>" & langDateTime & "</th>" & vbcrlf
response.write "      <th" & lcl_colspan_datetime & ">" & langEvent    & "</th>" & vbcrlf
response.write "    </tr>" & vbcrlf

'Retrieve the details about the event(s) for the date selected.
lcl_createdbyid            = ""
lcl_createdbyname          = ""
lcl_createddate            = ""
lcl_lastupdatedbyid        = ""
lcl_lastupdatedbyname      = ""
lcl_lastupdateddate        = ""
lcl_displayHistoryToPublic = 0
lcl_displayHistoryOption   = ""

sSQL = "SELECT e.EventID, "
sSQL = sSQL & " e.EventDate, "
sSQL = sSQL & " e.EventDuration, "
sSQL = sSQL & " t.TZAbbreviation, "
sSQL = sSQL & " e.Subject, "
sSQL = sSQL & " e.Message, "
sSQL = sSQL & " e.CategoryID, "
sSQL = sSQL & " c.Color, "
sSQL = sSQL & " e.CreatorUserID, "
sSQL = sSQL & " e.CreateDate, "
sSQL = sSQL & " e.ModifierUserID, "
sSQL = sSQL & " e.ModifiedDate, "
sSQL = sSQL & " e.displayHistoryToPublic, "
sSQL = sSQL & " e.displayHistoryOption, "
sSQL = sSQL & " (select u.firstname + ' ' + u.lastname from users u where u.userid = e.CreatorUserID) as createdbyname, "
sSQL = sSQL & " (select u.firstname + ' ' + u.lastname from users u where u.userid = e.ModifierUserID) as lastupdatedbyname "
sSQL = sSQL & " FROM Events e "
sSQL = sSQL &      " LEFT OUTER JOIN EventCategories c ON e.CategoryID = c.CategoryID, TimeZones t "
sSQL = sSQL & " WHERE t.TimeZoneID = e.EventTimeZoneID "
sSQL = sSQL & " AND e.OrgID = " & iorgid
sSQL = sSQL & " AND DateDiff(dd, e.EventDate, '" & dDate & "') = 0 "
sSQL = sSQL & " AND DateDiff(mm, e.EventDate, '" & dDate & "') = 0 "
sSQL = sSQL & " AND DateDiff(yy, e.EventDate, '" & dDate & "') = 0 "

if lcl_calendarfeature <> "" then
 sSQL = sSQL & " AND e.calendarfeature = '" & lcl_calendarfeature & "'"
else
 sSQL = sSQL & " AND (e.calendarfeature IS NULL OR e.calendarfeature = '')"
end if

sSQL = sSQL & " ORDER BY e.eventdate "

set oRst = Server.CreateObject("ADODB.Recordset")
oRst.Open sSQL, Application("DSN"), 3, 1

if not oRst.eof then
     do while not oRst.eof
        lcl_createdbyid            = oRst("CreatorUserID")
        lcl_createdbyname          = oRst("createdbyname")
        lcl_createddate            = oRst("CreateDate")
        lcl_lastupdatedbyid        = oRst("ModifierUserID")
        lcl_lastupdatedbyname      = oRst("lastupdatedbyname")
        lcl_lastupdateddate        = oRst("ModifiedDate")
        lcl_displayHistoryToPublic = oRst("displayHistoryToPublic")
        lcl_displayHistoryOption   = oRst("displayHistoryOption")

       	if oRst("EventDuration") > 0 Then
			If CLng(oRst("EventDuration")) = CLng(1440) Then
				dEnd = ""
			Else
				dEnd = DateAdd("n",oRst("EventDuration"),oRst("EventDate"))

				if DateDiff("d",dEnd,oRst("EventDate")) = 0 then
					dEnd = FormatDateTime(dEnd,vbLongTime)
				end if

				dEnd = " - " & dEnd
			End If 
       	else
        	  dEnd = ""
        end if

       	if oRst("CategoryID") <> 0 then
          'Get the category name
           sCategory = ""
           sCategory = getCategoryName(oRst("CategoryID"))

           if sCategory <> "" then
              sCategory = "(" & sCategory & ")"
           end if
			     else
			        sCategory = ""
		      end if

   				'BEGIN: Trim Code ------------------------------------------------------
   				'Used to trim seconds from dates displayed on calendar pages.
   				'9/9/2005 Vincent Evans

    				sDate1 = cStr(oRst("EventDate"))
    				sDate2 = cStr(dEnd)

    				iTrimDate1 = clng(InStrRev(sDate1,":"))
    				iTrimDate2 = clng(InStrRev(sDate2,":"))

   				'Retrieves AM/PM, trims final :00 and builds string
   	 			if iTrimDate1 > 0 then
   		   			sTemp  = Right(sDate1, 2)
   		   			sDate1 = Left(sDate1,iTrimDate1 - 1) & " " & sTemp
   		   			sTemp  = ""
    				end if

    				if iTrimDate2 > 0 then
    		  			sTemp  = Right(sDate2, 2)
    		  			sDate2 = Left(sDate2,iTrimDate2 - 1) & " " & sTemp
   	 	  			sTemp  = ""
   		 		end if

  				 'END: Trim Code --------------------------------------------------------

		     'Changed width from 75px to 25% to fix multilined time.
        'sEvents = sEvents & "<tr><td width=""25%"" valign=top>" & sDate1  & " " & sDate2 & " " & oRst("TZAbbreviation") & "</td><td><em><font color=""" & oRst("Color") & """>" & sCategory & "</em><strong>" & oRst("Subject") & "</font></strong><br />" & oRst("Message") & "</td></tr>"
        response.write "    <tr>" & vbcrlf
        response.write "        <td width=""25%"" valign=""top"">" & sDate1  & " " & sDate2 & "</td>" & vbcrlf
        response.write "        <td>" & vbcrlf
        response.write "            <em><font color=""" & oRst("Color") & """>" & sCategory & "</font></em>&nbsp;" & vbcrlf
        response.write "            <strong>" & oRst("Subject") & "</strong><br />" & vbcrlf
        response.write              oRst("Message") & vbcrlf

       'History Info
        if lcl_orghasfeature_displayHistoryInfo AND lcl_displayHistoryToPublic then
			response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""2"" id=""historyInfo"">" & vbcrlf

			displayHistoryInfo lcl_displayHistoryOption, "CREATEDBY", lcl_createdbyid, lcl_createdbyname, lcl_createdbydate

			displayHistoryInfo lcl_displayHistoryOption, "LASTUPDATED", lcl_lastupdatedbyid, lcl_lastupdatedbyname, lcl_lastupdateddate

			response.write "            </table>" & vbcrlf
        end if

        response.write "        </td>" & vbcrlf

		if bShowiCal then
			response.write "                          <td nowrap=""nowrap"" valign=""top"">" & vbcrlf
			response.write "                              <a href=""icalevent.asp?e=" & oRst("eventid") & """><img src=""../images/add_event.gif"" border=""0"" /> Add to My Calendar</a>" & vbcrlf
			response.write "                          </td>" & vbcrlf
		end If
		
        response.write "    </tr>" & vbcrlf

        oRst.movenext
     loop
  else
  	response.write "  <tr>" & vbcrlf
     response.write "      <td colspan=""2"">" & vbcrlf
     response.write "          <p><strong>There are no events currently scheduled for this day.</strong></p>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

  oRst.close
  set oRst = nothing

  response.write "</table>" & vbcrlf
  response.write "<p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>" & vbcrlf

 'BEGIN: Visitor Tracking -----------------------------------------------------
 	iSectionID = 55
 	if request("date") <> "" then
   		sDocumentTitle = "CALENDAR DATE: " & dDate   ' CDATE(request("date"))
 	else
	   	sDocumentTitle = "UNSPECIFIED CALENDAR DATE VIEW"
 	end if

  sURL        = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
 	datDate     = Date()
	 datDateTime = Now()
 	sVisitorIP  = request.servervariables("REMOTE_ADDR")

	 Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
 'END: Visitor Tracking -------------------------------------------------------
%>
<!--#Include file="../include_bottom.asp"-->
