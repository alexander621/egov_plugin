<!DOCTYPE html>
<!-- <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd"> -->
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: calendar.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Main Calendar display.
'
' MODIFICATION HISTORY
'	1.?	11/30/2007	Steve Loar - Changed to set day to the first when month is picked. Handles Feb EOM problem.
' 1.2 08/08/2008 David Boyer - Added Custom Calendar
'	2.0	08/11/2009	Steve Loar - Overhauled the calendar to work in IE8 and Firefox and to speed up load
'	2.1	12/10/2010	Steve Loar - Adding iCal export to add an event to user's calendar.
' 2.2 06/24/2011 David Boyer - Added history info to "Calendar Event List".
'	2.3	08/26/2011	Steve Loar - Changed Next and Previous to include the category selected
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 if not OrgHasFeature(iorgid,"calendar") then response.redirect "../default.asp"

 if Request.ServerVariables("REMOTE_ADDR") = "207.154.31.10" then
	  'block the Northern Lights Bots from attacking this page.
	   response.redirect "calendar.asp"
 end if

 session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
 session("RedirectLang") = "Return to Calendar"

 dim oEventsOrg, iYearDiff, bShowiCal
 Set oEventsOrg = New classOrganization

'Check to see if this is a custom calendar
 if trim(request("cal")) <> "" then
    if not isnumeric(trim(request("cal"))) then
       response.redirect "calendar.asp"
    else
       'lcl_calendarfeature        = trim(request("cal"))
       'lcl_calendarfeature_url    = "&cal=" & lcl_calendarfeature
       lcl_calendarfeatureid      = 0
       on error resume next
       lcl_calendarfeatureid      = CLng(trim(request("cal")))
       on error goto 0
       lcl_calendarfeature        = getFeatureByID(iorgid, lcl_calendarfeatureid)
       lcl_calendarfeature_url    = "&cal=" & lcl_calendarfeatureid
       lcl_calendarfeature_name   = " [" & getFeatureName(lcl_calendarfeature) & "]"
       lcl_displayHistory_feature = "displayhistoryinfo_customcalendars"
    end if

 else
    lcl_calendarfeatureid      = ""
    lcl_calendarfeature        = ""
    lcl_calendarfeature_url    = ""
    lcl_calendarfeature_name   = ""
    lcl_displayHistory_feature = "displayhistoryinfo"
 end if

Function GetDaysInMonth(iMonth, iYear)
 	Dim dTemp
 	dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
 	GetDaysInMonth = Day(dTemp)
End Function

Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth)
 	Dim dTemp
 	dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
 	GetWeekdayMonthStartsOn = WeekDay(dTemp)
End Function

Function SubtractOneMonth(dDate)
	 SubtractOneMonth = DateAdd("m", -1, dDate)
End Function

Function AddOneMonth(dDate)
	 AddOneMonth = DateAdd("m", 1, dDate)
End Function

Dim dDate     ' Date we're displaying calendar for
Dim iDIM      ' Days In Month
Dim iDOW      ' Day Of Week that month starts on
Dim iCurrent  ' Variable we use to hold current day of month as we write table
Dim iPosition ' Variable we use to hold current position in table
Dim dToday    ' Today
Dim titleSpan

dToday = Date()

'Set variables for values passed in
 If isDate(request("date")) Then 
    lcl_date = request("date")
	' SQL Server cannot handle really old dates
	'	- commented out for date range check to catch these now. SJL
'	If CDate(lcl_date) < CDate("1/1/1753") Then
'		lcl_date = Date()
'	End If
'	If CDate(lcl_date) > CDate("11/30/9999") Then
'		lcl_date = Date()
'	End If
 Else 
    lcl_date = ""
 End If 

on error resume next
 If IsNumeric(request("month")) Then 
   	lcl_month = CLng(request("month"))
 Else
   	response.redirect "calendar.asp"
 End If 

 If IsNumeric(request("day")) Then 
   	lcl_day   = CLng(request("day"))
 Else
   	response.redirect "calendar.asp"
 End If

 If IsNumeric(request("year")) Then 
   	lcl_year  = CLng(request("year"))
 Else
   	response.redirect "calendar.asp"
 End If
 on error goto 0

'Get selected date
 If IsDate(lcl_date) Then
 	  dDate = lcl_date
 Else
	If IsDate(lcl_month & "-" & lcl_day & "-" & lcl_year) Then
		dDate = lcl_month & "-" & lcl_day & "-" & lcl_year
	Else
		dDate = Date()
	End If
 End If

 ' To handle the bots that input dates way out of range we will catch the dates 
 ' more than 5 years plus or minus from today
iYearDiff = Abs(DateDiff("yyyy",dDate, Date()))

If iYearDiff > 5 Then
	response.redirect "outofrangedate.asp"
End If 

 bFilter = 0
  iCategory = CLng(0)
  on error resume next
 If IsNumeric(request("Category")) Then 
   	iCategory = CLng(Request("Category"))
 End If
 on error goto 0

 If Not iCategory = "" Then 
    If iCategory > CLng(0) Then 
      	bFilter = 1
    Else 
       bFilter = 0
    End If 
 End If 

'Now we have got the date.  Now get Days in the choosen month and the day of the week it starts on.
 iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
 iDOW = GetWeekdayMonthStartsOn(dDate)

'Set the flag to show the iCal export
 bShowiCal = OrgHasFeature( iOrgId, "public ical export" )

'Check for org features
 lcl_orghasfeature_displayHistoryInfo = orghasfeature(iOrgID, lcl_displayHistory_feature)
%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title>E-Gov Services - <%=sOrgName%></title>
<!--	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" /> -->
	
	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="calendar.css" />
	<link rel="stylesheet" type="text/css" href="calendarprint.css" media="print" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script type="text/javascript" src="https://code.jquery.com/jquery-1.6.2.min.js"></script>
	<script type="text/javascript" src="../scripts/modules.js"></script>
	<script type="text/javascript">var addthis_config = {"data_track_clickback":true};</script>
	<script type="text/javascript" src="https://s7.addthis.com/js/250/addthis_widget.js#pubid=egovlink"></script>

<!--	<script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script> -->
<!--	<script type="text/javascript">var addthis_pub="cschappacher";</script> -->

	<script type="text/javascript">
	<!--

		function goNextPrevious( month, year, url )
		{
			location.href = "calendar.asp?day=01&month=" + month + "&year=" + year + "&Category=" + $("#category").val() + url;
		}

		function SubmitByMonth()
		{
			document.frmDate.day.value = "01";
			document.frmDate.submit();
		}

		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}

		if(document.all && !document.getElementById) 
		{
			document.getElementById = function(id) 
			{
				 return document.all[id];
			}
		}

		function PrintPreview()
		{
			window.print();
		}
    
    function createOneMonthiCal( )
    {
      // need to know start date and category
      location.href='icalmonthexport.asp?startdate=' + $('#startDate').val() + '&categoryid=' + $('#category').val();
      
      return false;
    }
    
	//-->
	</script>

 <style type="text/css">
   <!--
		.unused  {font-family:Arial,Tahoma,Verdana; font-size:13px; color:#cc0000;}
		OPTION.color0000FF  {font-family:Arial,Tahoma,Verdana; font-size:13px; color:#0000ff;}
		OPTION.color006600  {font-family:Arial,Tahoma,Verdana; font-size:13px; color:#006600;}
		OPTION.colorCC0066  {font-family:Arial,Tahoma,Verdana; font-size:13px; color:#CC0066;}
		OPTION.colorFF9900  {font-family:Arial,Tahoma,Verdana; font-size:13px; color:#FF9900;}
		OPTION.color9933CC  {font-family:Arial,Tahoma,Verdana; font-size:13px; color:#9933CC;}
		OPTION.colorCC0000  {font-family:Arial,Tahoma,Verdana; font-size:13px; color:#CC0000;}
		OPTION.color0099FF  {font-family:Arial,Tahoma,Verdana; font-size:13px; color:#0099FF;}
		OPTION.colorFF33CC  {font-family:Arial,Tahoma,Verdana; font-size:13px; color:#FF33CC;}
		OPTION.colorFF0000  {font-family:Arial,Tahoma,Verdana; font-size:13px; color:#FF0000;}

		table#historyInfo, table#historyInfo td 
		{
			border: 0px solid #000000;
			margin-top: 5px;
		}
   //-->
 </style>

</head>

<!--#Include file="../include_top.asp"-->
<%
  response.write "<form id=""frmDate"" name=""frmDate"" action=""calendar.asp"" method=""get"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""day"" id=""day"" value=""" & Day(Now) & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""cal"" id=""cal"" value=""" & request("cal") & """ />" & vbcrlf

  response.write "<table border=""0"" id=""calendarbody"" border=""0"" cellspacing=""0"" cellpadding=""0"" width=""90%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <p>" & vbcrlf

 'BEGIN: Monthly Calendar -----------------------------------------------------
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""90%"">" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td align=""right"">" & vbcrlf
  response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "                      <tr valign=""top"">" & vbcrlf
  response.write "                          <td>" & vbcrlf
                                                displayAddThisButtonNew iorgid
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                    </table>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td>" & vbcrlf

 'Build the welcome message
  lcl_org_name        = oEventsOrg.GetOrgName()
  lcl_org_state       = oEventsOrg.GetState()
  'lcl_org_featurename = "Calendar of Events" & lcl_calendarfeature_name
  lcl_featurename     = oEventsOrg.GetOrgFeatureName("calendar")
  lcl_org_featurename = lcl_featurename & lcl_calendarfeature_name

  oEventsOrg.buildWelcomeMessage iorgid, _
                                 lcl_orghasdisplay_action_page_title, _
                                 lcl_org_name, _
                                 lcl_org_state, _
                                 lcl_org_featurename
  checkForRSSFeed iorgid, "", "", "COMMUNITYCALENDAR", sEgovWebsiteURL

  response.write "                    <br />" & vbcrlf
                                      RegisteredUserDisplay( "../" )
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td align=""right"">" & vbcrlf

  if checkForCustomCalendars( iorgid ) = "Y" then
     response.write "<strong>Show Calendar: </strong>" & vbcrlf
     response.write "<select name=""calendarfeature"" onChange=""location.href='calendar.asp?cal='+this.value"">" & vbcrlf
     response.write "  <option value="""">" & lcl_featurename & "</option>" & vbcrlf
                       displayCustomCalendarOptions iorgid, lcl_calendarfeatureid
     response.write "</select>" & vbcrlf
  end if

  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "          </p>" & vbcrlf
  response.write "          <table border=""0"" cellpadding=""2"" cellspacing=""0"" style=""max-width:800px;"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td align=""center"">" & vbcrlf
  response.write "                    <div id=""itemshow"">" & vbcrlf
  response.write "                    <strong>" & vbcrlf

  if blnCalRequest then
     response.write "<a href=""../action.asp?actionid=" & iCalForm & lcl_calendarfeature_url & """>REQUEST to add Calendar Item</a> | " & vbcrlf
  end if

  response.write "                    <a href=""#listView"">View Calendar by Event List</a> |" & vbcrlf
  response.write "                    <a href=""searchevents.asp" & replace(lcl_calendarfeature_url,"&","?") & """>SEARCH for Calendar Items</a> |" & vbcrlf
  response.write "                    <a href=""#"" onclick=""PrintPreview();"">Print the Calendar</a>" & vbcrlf
  response.write "                    </strong>" & vbcrlf
  response.write "                    </div>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf

  if orghasdisplay( iOrgId, "calendar notice" ) then
     response.write "<p id=""calendarnotice"">" & vbcrlf
                       GetOrgDisplay iOrgId, "calendar notice"
     response.write "</p>" & vbcrlf
  end if

  response.write "          <table id=""calendar"" class=""respTable"" border=""0"" cellpadding=""2"" cellspacing=""0"">" & vbcrlf
 ' response.write "            <tr height=""2%"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <th align=""center"" colspan=""7"">" & vbcrlf
  response.write "                    <table id=""calendarheader"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
  response.write "                      <tr valign=""top"">" & vbcrlf
  'response.write "                          <td height=""30"" width=""15%"" style=""border:0pt solid #ff0000;"" nowrap=""nowrap"">" & vbcrlf
  response.write "                          <td width=""15%"" style=""border:0pt solid #ff0000;"" nowrap=""nowrap"">" & vbcrlf
  response.write "                              <span class=""noprint"">" & vbcrlf
  													dNextPreviousDate = SubtractOneMonth(dDate)
  response.write "									<a href=""#"" onclick=""goNextPrevious( '" & Month(dNextPreviousDate) & "', '" & Year(dNextPreviousDate) & "', '" & lcl_calendarfeature_url & "' )""><img src=""../images/arrow_back.gif"" align=""absmiddle"" border=""0"" /></a>" & vbcrlf
  response.write "									<a href=""#"" onclick=""goNextPrevious( '" & Month(dNextPreviousDate) & "', '" & Year(dNextPreviousDate) & "', '" & lcl_calendarfeature_url & "' )""><span class=""calendarheadertext"">" & langPreviousMonth & "</span></a>" & vbcrlf
'  response.write "                                &nbsp;<a href=""calendar.asp?date=" & SubtractOneMonth(dDate) & lcl_calendarfeature_url & """><img src=""../images/arrow_back.gif"" align=""absmiddle"" border=""0"" /></a>&nbsp;" & vbcrlf
'  response.write "                                <a href=""calendar.asp?date=" & SubtractOneMonth(dDate) & lcl_calendarfeature_url & """>" & langPreviousMonth & "</a>" & vbcrlf
  response.write "                              </span>" & vbcrlf
  response.write "                              &nbsp;&nbsp;" & vbcrlf

 'These selected statements were added for ADA compliance
  response.write "                              <select id=""month"" name=""month"" onChange=""SubmitByMonth();"">" & vbcrlf
                                                  buildOption "MONTH", 1, month(dDate)
                                                  buildOption "MONTH", 2, month(dDate)
                                                  buildOption "MONTH", 3, month(dDate)
                                                  buildOption "MONTH", 4, month(dDate)
                                                  buildOption "MONTH", 5, month(dDate)
                                                  buildOption "MONTH", 6, month(dDate)
                                                  buildOption "MONTH", 7, month(dDate)
                                                  buildOption "MONTH", 8, month(dDate)
                                                  buildOption "MONTH", 9, month(dDate)
                                                  buildOption "MONTH", 10, month(dDate)
                                                  buildOption "MONTH", 11, month(dDate)
                                                  buildOption "MONTH", 12, month(dDate)

  response.write "                              </select>" & vbcrlf
  response.write "                              <select id=""year"" name=""year"" onChange=""document.getElementById('frmDate').submit();"">" & vbcrlf

                                                for x = (Year(dDate) - 5) To (Year(dDate) + 5)
                                                    lcl_selected_year = ""

                                                    if Year(dDate) = x then 
                                                    			lcl_selected_year = " selected=""selected"""
                               			                  end if

                                                    response.write "  <option value=""" & x & """" & lcl_selected_year & ">" & x & "</option>" & vbcrlf
                                                next

  response.write "                              </select>" & vbcrlf
  response.write "                              <span class=""noprint"">" & vbcrlf
  response.write "                                &nbsp;&nbsp;&nbsp;" & vbcrlf
'  response.write "                                <a href=""calendar.asp?date=" & AddOneMonth(dDate) & lcl_calendarfeature_url & """>" & langNextMonth & "</a>" & vbcrlf
'  response.write "                                <a href=""calendar.asp?date=" & AddOneMonth(dDate) & lcl_calendarfeature_url & """><img src=""../images/arrow_forward.gif"" align=""absmiddle"" border=""0"" /></a>" & vbcrlf
													dNextPreviousDate = AddOneMonth(dDate)
  response.write "									<a href=""#"" onclick=""goNextPrevious( '" & Month(dNextPreviousDate) & "', '" & Year(dNextPreviousDate) & "', '" & lcl_calendarfeature_url & "' )""><span class=""calendarheadertext"">" & langNextMonth & "</span></a>" & vbcrlf
  response.write "									<a href=""#"" onclick=""goNextPrevious( '" & Month(dNextPreviousDate) & "', '" & Year(dNextPreviousDate) & "', '" & lcl_calendarfeature_url & "' )""><img src=""../images/arrow_forward.gif"" align=""absmiddle"" border=""0"" /></a>" & vbcrlf
'  response.write "                                <img src=""../images/spacer.gif"" width=""1"" height=""30"" border=""1"" align=""absmiddle"" />" & vbcrlf
  response.write "                              </span>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                          <td align=""right"" width=""50%"" nowrap=""nowrap"">" & vbcrlf
  response.write "                              <div class=""noprint"">" & vbcrlf
  response.write "                                <span class=""view_text_label calendarheadertext"">View:</span>" & vbcrlf
  response.write "                                <select id=""category"" name=""Category"" class=""time"" onchange=""document.getElementById('frmDate').submit();"">" & vbcrlf
  response.write "                                  <option value=""0"">All Categories</option>" & vbcrlf
                                                    getEventCategoryOptions iorgid, lcl_calendarfeature, iCategory
  response.write "                                </select>" & vbcrlf
  response.write "                              </div>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                    </table>" & vbcrlf
  response.write "                </th>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  'response.write "            <tr id=""calendardayrow"" height=""10"">" & vbcrlf
  'response.write "                <td width=""13%"" height=""10"">Sun</td>" & vbcrlf
  'response.write "                <td width=""13%"" height=""10"">Mon</td>" & vbcrlf
  'response.write "                <td width=""13%"" height=""10"">Tues</td>" & vbcrlf
  'response.write "                <td width=""13%"" height=""10"">Wed</td>" & vbcrlf
  'response.write "                <td width=""13%"" height=""10"">Thurs</td>" & vbcrlf
  'response.write "                <td width=""13%"" height=""10"">Fri</td>" & vbcrlf
  'response.write "                <td width=""13%"" height=""10"">Sat</td>" & vbcrlf
  'response.write "            </tr>" & vbcrlf
'  response.write "            <tr id=""calendardayrow"">" & vbcrlf
'  response.write "                <td width=""13%"">Sun</td>" & vbcrlf
'  response.write "                <td width=""13%"">Mon</td>" & vbcrlf
'  response.write "                <td width=""13%"">Tues</td>" & vbcrlf
'  response.write "                <td width=""13%"">Wed</td>" & vbcrlf
'  response.write "                <td width=""13%"">Thurs</td>" & vbcrlf
'  response.write "                <td width=""13%"">Fri</td>" & vbcrlf
'  response.write "                <td width=""13%"">Sat</td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf
  response.write "            <tr id=""calendardayrow"" class=""respHide"">" & vbcrlf
  response.write "                <td>Sun</td>" & vbcrlf
  response.write "                <td>Mon</td>" & vbcrlf
  response.write "                <td>Tues</td>" & vbcrlf
  response.write "                <td>Wed</td>" & vbcrlf
  response.write "                <td>Thurs</td>" & vbcrlf
  response.write "                <td>Fri</td>" & vbcrlf
  response.write "                <td>Sat</td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Write spacer cells at beginning of first row if month doesn't start on a Sunday.
  if iDOW <> 1 then
     response.write "            <tr>" & vbcrlf

     iPosition = 1
     do while iPosition < iDOW
        response.write "                <td class=""respHide"">&nbsp;</td>"

        iPosition = iPosition + 1
     loop
  end if

	'Write days of month in proper day slots
		iCurrent  = 1
		iPosition = iDOW

		do while iCurrent <= iDIM

    'If we're at the begginning of a row then write TR
    	if iPosition = 1 then
				   	response.write "            <tr>" & vbcrlf
   		end if

   		'response.write "                <td height=""55"" id=""day_" & iCurrent & """ valign=""top"">" & vbcrlf
   		response.write "                <td id=""day_" & iCurrent & """ valign=""top"">" & vbcrlf

   	'If the day we're writing is the selected day then highlight it.
   		if iCurrent = Day(Date()) and Month(dDate) = Month(Date()) and Year(dDate) = Year(Date()) then
			   		response.write "<strong>" & iCurrent & " - Today</strong>" & vbcrlf
   		else
			   		response.write "<strong>" & iCurrent & "</strong>" & vbcrlf
   		end if

   		ShowEvents Month(dDate) & "/" & iCurrent & "/" & Year(dDate), lcl_calendarfeature, iCategory, lcl_calendarfeature_url

   		response.write "                </td>" & vbcrlf

   	'If we're at the endof a row then write /TR
   		if iPosition = 7 then
			   		response.write "            </tr>" & vbcrlf
   					iPosition = 0
   		end if

   	'Increment variables
   		iCurrent  = iCurrent  + 1
   		iPosition = iPosition + 1
		loop

	'Write spacer cells at end of last row if month doesn't end on a Saturday.
		if iPosition <> 1 then
				 do while iPosition <= 7
					   response.write "                <td class=""respHide"">&nbsp;</td>"

					   iPosition = iPosition + 1
 				loop

  			response.write "            </tr>" & vbcrlf
  end if

  response.write "          </table>" & vbcrlf
  response.write "          </form>" & vbcrlf
 'END: Monthly Calendar -------------------------------------------------------

 'BEGIN: Calendar Event List --------------------------------------------------
	 lcl_colspan_datetime = ""

  if bShowiCal then
	  lcl_colspan_datetime = " colspan=""2"""
    titleSpan = " colspan=""2"""
  else
    titleSpan = " colspan=""3"""
  end if

sStartOfMonth              = CDate(Month(dDate) & "/1/" & Year(dDate))
sCalendarDate              = CStr(sStartOfMonth)

'response.write "          <table border=""0"" cellpadding=""2"" cellspacing=""0"" style=""width:800px;"">" & vbcrlf
response.write "          <table id=""calendarEventList"" border=""0"" cellpadding=""2"" cellspacing=""0"">" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td align=""center"">" & vbcrlf
response.write "                    <a name=""listView""></a>" & vbcrlf
response.write "                    <p>" & vbcrlf

								 'formerly the table was class="cal"
response.write "                    <table id=""eventlist"" border=""0"" cellpadding=""4"" cellspacing=""0"" width=""90%"">" & vbcrlf
response.write "                      <tr>" & vbcrlf
response.write "                          <td" & titleSpan & " align=""center"" class=""calendartitle"">Calendar Event List</td>" & vbcrlf
if bShowiCal then
  response.write "<td nowrap>"
  response.write "<input id=""startDate"" type=""hidden"" value=""" & sCalendarDate & """ />"
  response.write "<a href=""javascript:void(0);"" onclick=""createOneMonthiCal( );"" id=""iCalMonthLink""><img src=""../images/add_event.gif"" border=""0"" /><span class=""respHide""> Add All to My Calendar</span></a></td>" & vbcrlf
end if 
response.write "                      </tr>" & vbcrlf
response.write "                      <tr>" & vbcrlf
response.write "                          <th>" & langDateTime & "</th>" & vbcrlf
response.write "                          <th" & lcl_colspan_datetime & ">" & langEvent & "</th>" & vbcrlf
response.write "                      </tr>" & vbcrlf

'Setup the query to pull all of the details for the events for a month
'sStartOfMonth              = CDate(Month(dDate) & "/1/" & Year(dDate)) ' moved to above table
'sCalendarDate              = CStr(sStartOfMonth) ' moved to above table
dNextDay                   = DateAdd("m", 1, sStartOfMonth)
sNextDay                   = CStr(dNextDay) & " 00:00:00 AM"
lcl_createdbyid            = ""
lcl_createdbyname          = ""
lcl_createddate            = ""
lcl_lastupdatedbyid        = ""
lcl_lastupdatedbyname      = ""
lcl_lastupdateddate        = ""
lcl_displayHistoryToPublic = 0
lcl_displayHistoryOption   = ""

sSql = "SELECT e.eventid, "
sSql = sSql & " e.eventdate, "
sSql = sSql & " e.eventduration, "
sSql = sSql & " e.subject, "
sSql = sSql & " e.message, "
sSql = sSql & " e.categoryid, "
sSql = sSql & " c.color, "
sSql = sSql & " e.CreatorUserID, "
sSql = sSql & " e.CreateDate, "
sSql = sSql & " e.ModifierUserID, "
sSql = sSql & " e.ModifiedDate, "
sSql = sSql & " e.displayHistoryToPublic, "
sSql = sSql & " e.displayHistoryOption, "
sSql = sSql & " (select u.firstname + ' ' + u.lastname from users u where u.userid = e.CreatorUserID) as createdbyname, "
sSql = sSql & " (select u.firstname + ' ' + u.lastname from users u where u.userid = e.ModifierUserID) as lastupdatedbyname "
sSql = sSql & " FROM events e "
sSql = sSql &      " LEFT JOIN eventcategories c ON e.categoryid = c.categoryid "
sSql = sSql & " WHERE e.orgid = " & iorgid
sSql = sSql & " AND E.eventdate >= '" & sCalendarDate & "' "
sSql = sSql & " AND E.eventdate < '" & sNextDay & "' "

if lcl_calendarfeature <> "" then
	sSql = sSql & " AND UPPER(e.calendarfeature) = '" & UCase(lcl_calendarfeature) & "' "
else
	sSql = sSql & " AND (e.calendarfeature = '' OR e.calendarfeature IS NULL)"
end if

if bFilter then
	sSql = sSql & " AND e.categoryid = " & iCategory
end if

'Setup the order by
sSql = sSql & " ORDER BY e.eventdate, eventid "
'response.write sSql & "<br /><br />"

set oRst = Server.CreateObject("ADODB.Recordset")
oRst.Open sSql, Application("DSN"), 3, 1

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
				' calculate the end date and time using the duration in minutes
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
			sCategory = "(" & getCategoryName(oRst("categoryid")) & ") "
		else
			sCategory = ""
		end if

		'Format the event date/time (remove the seconds from the date(s))
		lcl_event_date = oRst("eventdate")

		'This displays the "12:00:00 AM" IF a duration exists
		if left(FormatDateTime(oRst("eventdate"),vbLongTime),11) = "12:00:00 AM" and oRst("eventduration") > 0 Then
			If CLng(oRst("EventDuration")) <> CLng(1440) Then
				lcl_event_date = oRst("eventdate") & " " & formatdatetime(oRst("eventdate"),vbLongTime)
			End If 
		end if

		if oRst("eventduration") > 0 Then
			If CLng(oRst("EventDuration")) <> CLng(1440) Then
				if Left(FormatDateTime(Replace(dEnd," - ", ""),vbLongTime),11) = "12:00:00 AM" then
					lcl_end_time = replace(dEnd," - ", "")
					lcl_end_time = lcl_end_time & " " & formatdatetime(lcl_end_time,vbLongTime)
					lcl_end_time = " - " & lcl_end_time
					dEnd         = lcl_end_time
				end If
			End If 
		end if

		' formatEventDateTime is in common.asp
		formatEventDateTime lcl_event_date, dEnd, sDate1, sDate2

        response.write "                      <tr align=""left"">" & vbcrlf
		response.write "                          <td width=""25%"" valign=""top"">" & vbcrlf
        response.write                                sDate1 & " " & sDate2 & vbcrlf
        response.write "                          </td>" & vbcrlf
		response.write "                          <td>" & vbcrlf
        response.write "                              <font color=""" & oRst("Color") & """>" & vbcrlf
        response.write "                                <em>" & sCategory & "</em>" & vbcrlf
        response.write "                                <strong>" & oRst("Subject") & "</strong>" & vbcrlf
        response.write "                              </font><br />" & vbcrlf
        response.write                                oRst("Message") & vbcrlf

       'History Info
        if lcl_orghasfeature_displayHistoryInfo AND lcl_displayHistoryToPublic then
			response.write "                                 <table border=""0"" cellspacing=""0"" cellpadding=""2"" id=""historyInfo"">" & vbcrlf
			displayHistoryInfo lcl_displayHistoryOption, "CREATEDBY", lcl_createdbyid, lcl_createdbyname, lcl_createdbydate
			displayHistoryInfo lcl_displayHistoryOption, "LASTUPDATED", lcl_lastupdatedbyid, lcl_lastupdatedbyname, lcl_lastupdateddate
			response.write "                                 </table>" & vbcrlf
        end if

        response.write "                          </td>" & vbcrlf

		if bShowiCal then
			response.write "                          <td nowrap=""nowrap"" valign=""top"">" & vbcrlf
			response.write "                              <a href=""icalevent.asp?e=" & oRst("eventid") & """><img src=""../images/add_event.gif"" border=""0"" /><span class=""respHide""> Add to My Calendar</span></a>" & vbcrlf
			response.write "                          </td>" & vbcrlf
		end if

   					response.write "                      </tr>" & vbcrlf

     			oRst.movenext
				 loop
  else
     if bShowiCal then
        lcl_colspan_datetime = " colspan=""2"""
     else
        lcl_colspan_datetime = ""
     end if

					response.write "                      <tr>"
 				response.write "                          <td" & lcl_colspan_datetime & "><p><strong>There are no events currently scheduled for this day.</strong></p></td>" & vbcrlf
 				response.write "                      </tr>"
  end if

		oRst.close
		set oRst = nothing 

  response.write "                    </table>" & vbcrlf
  response.write "                    </p>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
 'END: Calendar Event List ----------------------------------------------------

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "<p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>" & vbcrlf

'BEGIN: Visitor Tracking ------------------------------------------------------
	iSectionID     = 5
	sDocumentTitle = "MAIN"
	sURL           = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
	datDate        = Date()
	datDateTime    = Now()
	sVisitorIP     = request.servervariables("REMOTE_ADDR")
	Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
'END: Visitor Tracking --------------------------------------------------------
%>
<!--#Include file="../include_bottom.asp"-->
<%

'------------------------------------------------------------------------------
Function escDblQuote( strDB )
 	If VarType( strDB ) <> vbString Then 
	   	escDblQuote = strDB 
 	Else 
	   	escDblQuote = Replace( strDB, Chr(34), "\" & Chr(34) )
 	End If 
End Function

'------------------------------------------------------------------------------
Function GetFeatureName( sFeature )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(FO.featurename,F.featurename) AS featurename "
	sSql = sSql & " FROM egov_organizations_to_features FO, egov_organization_features F "
	sSql = sSql & " WHERE FO.featureid = F.featureid "
	sSql = sSql & " AND FO.orgid = " & iorgid
	sSql = sSql & " AND feature = '" & sFeature & "'" 

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		  GetFeatureName = oRs("featurename")
	Else
		  GetFeatureName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function


'------------------------------------------------------------------------------
Sub ShowEvents( ByVal sCalendarDate, ByVal sFeature, ByVal iCategory, ByVal lcl_calendarfeature_url )
	Dim sSql, oRs, iPerday, dNextDay, sNextDay, dCalendarDate, sNewDate

	iPerday = 0
	dCalendarDate = CDate(sCalendarDate)
	sNewDate = CStr(dCalendarDate) 
	dNextDay = DateAdd("D", 1, CDate(sCalendarDate))
	sNextDay = CStr(dNextDay)
	lcl_numItemsPerDay = getItemsPerDay(iorgid)

	sSql = "SELECT TOP " & clng(lcl_numItemsPerDay) + 1  & " E.eventdate, "
	sSql = sSql & " E.subject,  ISNULL(C.color,'black') AS color "
	sSql = sSql & " FROM events E "
	sSql = sSql & " LEFT JOIN eventcategories C ON E.categoryid = C.categoryid "
	sSql = sSql & " WHERE E.orgid = " & iorgid
	sSql = sSql & " AND E.eventdate >= '" & sNewDate & "' "
	sSql = sSql & " AND E.eventdate < '"  & sNextDay & "' "

	If CLng(iCategory) > CLng(0) Then 
		sSql = sSql & " AND E.categoryid = " & iCategory
	End If 

	If lcl_calendarfeature <> "" Then 
		sSql = sSql & " AND UPPER(E.calendarfeature) = '" & UCASE(sFeature) & "' "
	Else 
		sSql = sSql & " AND (E.calendarfeature = '' OR E.calendarfeature IS NULL)"
	End If 

	'Setup the ORDER BY
	sSql = sSql & " ORDER BY E.eventdate, eventid"

	'response.write sSql & "<br /><br />"

	session("calendarSQL") = "sCalendarDate:" & sCalendarDate & " dNextDay:" & dNextDay & " sSql:" & sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	session("calendarSQL") = ""

 	If Not oRs.EOF Then 
		response.write "<ul class=""eventlist"">" & vbcrlf

		Do While Not oRs.EOF
			response.write "<li>" & vbcrlf

			iPerday = iPerday + 1

			If iPerday > clng(lcl_numItemsPerDay) Then 
				response.write "<a href='calendarevents.asp?date="& Month(sCalendarDate) & "-" & Day(CDate(sCalendarDate)) & "-" & Year(sCalendarDate) & lcl_calendarfeature_url & "'>More...</a>" & vbcrlf
				Exit Do 
			Else 
				response.write "<span style='cursor:pointer; cursor:hand;' onclick='document.location.href=""calendarevents.asp?date="& Month(CDate(sCalendarDate)) & "-" & Day(CDate(sCalendarDate)) & "-" & Year(CDate(sCalendarDate)) & lcl_calendarfeature_url & """'><font size=1 color='" & oRs("Color") & "'>" & oRs("subject") & "</font></span>" & vbcrlf
			End If 

			response.write "</li>" & vbcrlf

			oRs.MoveNext
		Loop 
		response.write "</ul>" & vbcrlf
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
sub oldDisplay()
'Determine what days we have events on and mark in calendar
 lcl_numItemsPerDay = getItemsPerDay(iorgid)

'Setup the query to pull all of the events for the month
	sSql = "SELECT e.eventdate, e.subject, c.color "
	sSql = sSql & " FROM events as e "
	sSql = sSql & " LEFT JOIN eventcategories as c ON e.categoryid = c.categoryid "
	sSql = sSql & " WHERE e.orgid = " & iorgid
	sSql = sSql & " AND datediff(mm, e.eventdate, '" & dDate & "') = 0 "
	sSql = sSql & " AND datediff(yy, e.eventdate, '" & dDate & "') = 0 "
	
	if lcl_calendarfeature <> "" then
  		sSql = sSql & " AND UPPER(e.calendarfeature) = '" & UCASE(lcl_calendarfeature) & "' "
	else
  		sSql = sSql & " AND (e.calendarfeature = '' OR e.calendarfeature IS NULL)"
	end if

'Setup the ORDER BY
	sSql = sSql & " ORDER BY e.eventdate "

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	sScript = ""

	if not oRs.eof then
  		sScript = "<script type=""text/javascript"">" & vbcrlf

   	do while not oRs.eof

    			perday  = 0
     		sEvents = ""
    			iDay    = Day(oRs("EventDate"))

   			'Determine if the events returned for the month are filtered or not.
    			if bFilter then  	'FILTERED
      				sSql2 = sSql & " AND e.categoryid = " & iCategory
      				sSql2 = sSql2 & lcl_orderby
    			else          				'NOT FILTERED
      				sSql2 = sSql & lcl_orderby
       end if

    			set oRs2 = Server.CreateObject("ADODB.Recordset")
    			oRs2.Open sSql2, Application("DSN"), 3, 1

    			if not oRs2.eof then
      				do while not oRs2.eof
        					if day(oRs2("EventDate")) = iDay then
          						perday = perday + 1

          						if perday > clng(lcl_numItemsPerDay) then
            							sEvents = sEvents & "<li><a href='calendarevents.asp?date="& Month(dDate) & "-" & iDay & "-" & Year(dDate) & lcl_calendarfeature_url & "'>More...</a></li>"
            							exit do
                else
             						truncMessage = oRs2("Subject")

            							'dcajacob 6/15/05 removed at request of Peter,John
            							'If Len(truncMessage) > 22 Then
              					'   truncMessage=Left(truncMessage,19) & "..."
            							'End If
            							sEvents = sEvents & "<li><span style='cursor:pointer; cursor:hand;' onclick='document.location.href=\""calendarevents.asp?date="& Month(dDate) & "-" & iDay & "-" & Year(dDate) & lcl_calendarfeature_url & "\""'><font size=1 style='text-transform:uppercase' color='" & oRs2("Color") & "'>" & escDblQuote(truncMessage) & "</font></span></li>"
          						end if
             end if

        					oRs2.MoveNext
          loop
       end if

    			oRs2.Close
    			set oRs2 = nothing 

    			sTemp = "document.getElementById('day_" & iDay & "').innerHTML = ""<strong>" & iDay & "</strong><br />" & sEvents & """;" & vbcrlf
    			'sScript = sScript & sTemp & "document.all.day_" & iDay & ".style.backgroundColor = '#99ccff';" &  vbCrLf
    			sScript = sScript & sTemp & vbcrlf

    			oRs.MoveNext
    loop 

  		sScript = sScript & "</script>" & vbcrlf

 end if

	oRs.close
	set oRs = nothing 

	response.write sScript

end sub

'------------------------------------------------------------------------------
sub buildOption(iOptionType, iValue, iCurrentValue)

  lcl_isSelected   = ""
  lcl_optiontype   = ""
  lcl_value        = ""
  lcl_currentvalue = ""
  lcl_displayvalue = ""

  if iOptionType <> "" then
     lcl_optiontype = ucase(iOptionType)
  end if

  if iValue <> "" then
     lcl_value = iValue
  end if

  if iCurrentValue <> "" then
     lcl_currentvalue = iCurrentValue
  end if

 'Determine which option is select for the option type being built
  if lcl_value = lcl_currentvalue then
     lcl_isSelected   = " selected=""selected"""
  end if

  if lcl_optiontype = "MONTH" then
     if lcl_value = 1 then
        lcl_displayvalue = langMonth01
     elseif lcl_value = 2 then
        lcl_displayvalue = langMonth02
     elseif lcl_value = 3 then
        lcl_displayvalue = langMonth03
     elseif lcl_value = 4 then
        lcl_displayvalue = langMonth04
     elseif lcl_value = 5 then
        lcl_displayvalue = langMonth05
     elseif lcl_value = 6 then
        lcl_displayvalue = langMonth06
     elseif lcl_value = 7 then
        lcl_displayvalue = langMonth07
     elseif lcl_value = 8 then
        lcl_displayvalue = langMonth08
     elseif lcl_value = 9 then
        lcl_displayvalue = langMonth09
     elseif lcl_value = 10 then
        lcl_displayvalue = langMonth10
     elseif lcl_value = 11 then
        lcl_displayvalue = langMonth11
     elseif lcl_value = 12 then
        lcl_displayvalue = langMonth12
     end if
  else
     lcl_displayvalue = lcl_value
  end if

 'Build the option
  response.write "<option value=""" & lcl_value & """" & lcl_isSelected & ">" & lcl_displayvalue & "</option>" & vbcrlf

end sub
%>
