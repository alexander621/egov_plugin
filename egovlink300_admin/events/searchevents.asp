<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: searchevents.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Calendar event display.
'
' MODIFICATION HISTORY
'	1.?	???	     ??? - INITIAL VERSION
' 1.2 08/20/08 David Boyer - Added Custom Calendar
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("internal_calendars") = "Y" OR isFeatureOffline("custom_calendars") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../" ' Override of value from common.asp

'Check to see if any Custom Calendars exist
 lcl_hasCustomCalendars = checkForCustomCalendars(session("orgid"))

'Check to see if this is a Custom Calendar
' if trim(request("cal")) <> "" then
'    lcl_calendarfeature = trim(request("cal"))
' else
'    if trim(request("calendarfeature")) <> "" then
'       lcl_calendarfeature = trim(request("calendarfeature"))
'    else
'       lcl_calendarfeature = getFirstCalendarInList(session("orgid"))
'    end if
' end if

' lcl_calendarfeature_url  = "&cal=" & lcl_calendarfeature

' if lcl_calendarfeature <> "" then
'    lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"
' else
'    lcl_calendarfeature_name = ""
' end if

'Allow the user to access the Internal Calendars if any/all of the following:
 '1. The user has the "edit events" permission assigned
 '2. The user has a specific Custom Calendar feature assigned: [session("calendarfeature") <> ""]
' if trim(request("cal")) <> "" then
'    if OrgHasFeature(trim(request("cal"))) AND UserHasPermission(session("userid"), trim(request("cal"))) then
'       session("calendarfeature") = trim(request("cal"))
       'lcl_calendarfeature = trim(request("cal"))
'       lcl_calendarfeature_url  = "&cal=" & session("calendarfeature")
'       lcl_calendarfeature_name = " [" & GetFeatureName(session("calendarfeature")) & "]"
'    else
'      	response.redirect sLevel & "permissiondenied.asp"
'    end if
' else
'    if NOT UserHasPermission( Session("UserId"), "edit events" ) then
' 	     response.redirect sLevel & "permissiondenied.asp"
'    end if

'    session("calendarfeature") = ""
' end if

 if trim(request("cal")) <> "" then
    if not isnumeric(trim(request("cal"))) then
       response.redirect sLevel & "permissiondenied.asp"
    else
       lcl_calendarfeatureid = CLng(trim(request("cal")))
    end if
 else
    lcl_calendarfeatureid = getFirstCalendarInList(session("orgid"))
 end if

 if lcl_calendarfeatureid <> "" then
    lcl_calendarfeature = getFeatureByID(session("orgid"), lcl_calendarfeatureid)

    if OrgHasFeature(lcl_calendarfeature) AND UserHasPermission(session("userid"), lcl_calendarfeature) then
       lcl_calendarfeature_url  = "&cal=" & lcl_calendarfeatureid
       lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"
    else
       response.redirect sLevel & "permissiondenied.asp"
    end if
 else
    lcl_calendarfeatureid    = ""
    lcl_calendarfeature      = ""
    lcl_calendarfeature_url  = ""
    lcl_calendarfeature_name = ""

    if NOT UserHasPermission( Session("UserId"), "edit events" ) then
 	     response.redirect sLevel & "permissiondenied.asp"
    end if

 end if


 lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=Hide
%>
<html>
<head>
	<title><%=langBSEvents%><%=lcl_calendar_name%></title>
 <%
   if session("orgid") = 7 then
      lcl_title = sOrgName
   else
      lcl_title = "E-Gov Services " & sOrgName
   end if

   response.write "<title>" & lcl_title & "</title>" & vbcrlf
 %>
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script src="../scripts/selectAll.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

	<script language=javascript>
	<!--

		function openWin2(url, name) 
		{
			popupWin = window.open(url, name,"resizable,width=500,height=450");
		}

		function fnCheckNew() 
		{
			if ((document.frmSearch.Keyword.value != '') && ((document.frmSearch.Subject.checked == true) || (document.frmSearch.Descrip.checked == true)) ) 
			{
				return true;
			}
			else 
			{
				return false;
			}
		}


		function doCalendar( sField ) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function storeCaret (textEl) 
		{
			if (textEl.createTextRange)
				textEl.caretPos = document.selection.createRange().duplicate();
		}

		function insertAtCaret (textEl, text) 
		{
			if (textEl.createTextRange && textEl.caretPos) 
			{
				var caretPos = textEl.caretPos;
			caretPos.text =
				caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
				text + ' ' : text;
			}
			else
				textEl.value = textEl.value + text;
		}

		function doPicker(sFormField) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("../picker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

	//-->
	</script>
</head>

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<body topmargin="0" leftmargin="0" bottommargin="0" rightmargin="0" marginwidth="0" marginheight="0">

<div id="content">
	 <div id="centercontent">

<p>
<h3>Calendar Search<%=lcl_calendarfeature_name%></h3><p>
<%
  Dim dDate
  'Variables for searching by date
  dim dDateSearch
  Dim dDateStart
  Dim dDateEnd

  if IsDate(Request.QueryString("date")) then
     dDate = CDate(Request.QueryString("date"))
  else
     if IsDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year")) then
        dDate = CDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year"))
     else
        dDate = Date()
     end if
  end if

  bResults = 0

  if request.form("_task") = "search" then

     Dim oCmd, oRst, sEvents
     lcl_where_clause = ""

     if (UCASE(Request.Form("Subject")) = "ON" AND UCase(Request.Form("Descrip")) = "ON") then
  	      'sProc    = "SearchMonthEventsBySubjectDescrip"
         lcl_where_clause = " AND (UPPER(subject) LIKE ('%" & UCASE(request("keyword")) & "%') OR UPPER(message) LIKE ('%" & UCASE(request("keyword")) & "%')) "
        	bSearch  = 1
     	  	sSub     = "CHECKED"
     	  	sDescrip = "CHECKED"
  	  elseif UCASE(Request.Form("Subject")) = "ON" then
      	  'sProc    = "SearchMonthEventsBySubject"
         lcl_where_clause = " AND UPPER(subject) LIKE ('%" & UCASE(request("keyword")) & "%') "
     	  	bSearch  = 1
     	  	sSub     = "CHECKED"
  	  elseif UCASE(Request.Form("Descrip")) = "ON" then
     	  	'sProc    = "SearchMonthEventsByDescrip"
         lcl_where_clause = " AND UPPER(message) LIKE ('%" & UCASE(request("keyword")) & "%') "
     	  	bSearch  = 1
     	  	sDescrip = "CHECKED"
     else
     	   sSub     = ""
     	   sDescrip = ""
     end if

    'Section used to determine which stored procedure to call and pass correct variables.
     lcl_execute_query = "Y"

    'Build the query
     sSql = "SELECT e.EventID, e.EventDate, e.EventDuration, t.TZAbbreviation, e.Subject, e.Message, e.CategoryID, c.Color "
     sSql = sSql & " FROM Events e "
     sSql = sSql &   " INNER JOIN TimeZones t ON t.timezoneid = e.eventtimezoneid "
     sSql = sSql &   " LEFT JOIN EventCategories c ON e.categoryid = c.categoryid "
     sSql = sSql & " WHERE e.orgid = " & session("orgid")

     if lcl_calendarfeature <> "" then
        sSql = sSql & " AND e.calendarfeature = '" & dbsafe(lcl_calendarfeature) & "' "
     else
        sSql = sSql & " AND (e.calendarfeature IS NULL OR e.calendarfeature <> '') "
     end if

     sSql = sSql & lcl_where_clause

     if UCASE(request.form("DateSearch")) = 1 then
        'nothing to add to query

    'Before Date
     elseif UCASE(Request.Form("DateSearch")) = 2 AND Request.Form("DatePickerBefore") <> "" then
   	    if IsDate(Request.Form("DatePickerBefore")) then
     		    dDateSearch = CDate(Request.Form("DatePickerBefore"))
       	else
		         dDateSearch = Date()
        end if

        lcl_where_clause_before = " AND e.eventdate <= '" & dbsafe(dDateSearch) & "' "

    'After Date
     elseif UCASE(Request.Form("DateSearch")) = 3 AND Request.Form("DatePickerAfter") <> "" then
   	    if IsDate(Request.Form("DatePickerAfter")) then
     		    dDateSearch = CDate(Request.Form("DatePickerAfter"))
   	    else
     		    dDateSearch = Date()
       	end if

        lcl_where_clause_after = " AND e.eventdate >= '" & dbsafe(dDateSearch) & "' "

    'Between dates (Range)
     elseif UCASE(Request.Form("DateSearch")) = 4 AND Request.Form("DatePickerStart") <> "" AND Request.Form("DatePickerEnd") <> "" then
   	    if IsDate(Request.Form("DatePickerStart")) AND IsDate(Request.Form("DatePickerEnd")) then
     		    dDateStart = CDate(Request.Form("DatePickerStart"))
     		    dDateEnd = CDate(Request.Form("DatePickerEnd"))
   	    else
      	    if NOT IsDate(Request.Form("DatePickerStart")) OR NOT IsDate(Request.Form("DatePickerEnd")) then
       			    if NOT IsDate(Request.Form("DatePickerStart")) then
         				    dDateStart = Date()
           			end if
           			if NOT IsDate(Request.Form("DatePickerEnd")) then
             				dDateEnd = Date()
           			end if
          	end if
        end if

        if CDate(dDateStart) > CDate(dDateEnd) then
	          lcl_start_date = dDateEnd
           lcl_end_date   = dDateStart
        else
           lcl_start_date = dDateStart
           lcl_end_date   = dDateEnd
        end if

        lcl_where_clause_between = " AND (e.eventdate BETWEEN '" & dbsafe(lcl_start_date) & "' AND '" & dbsafe(lcl_end_date) & "') "

     else
        lcl_execute_query = "N"
     end if

     if lcl_execute_query = "Y" then

       'Append BEFORE where clause
        if lcl_where_clause_before <> "" then
           sSql = sSql & lcl_where_clause_before
        end if

       'Append AFTER where clause
        if lcl_where_clause_after <> "" then
           sSql = sSql & lcl_where_clause_after
        end if

       'Append BETWEEN where clause
        if lcl_where_clause_between <> "" then
           sSql = sSql & lcl_where_clause_between
        end if

        sSql = sSql & " ORDER BY e.eventdate "
        set oRst = Server.CreateObject("ADODB.Recordset")
        oRst.Open sSql, Application("DSN"), 3, 1

        if not oRst.eof then
          	bResults = 1
           do while NOT oRst.eof
             	if oRst("EventDuration") > 0 then
              	  dEnd = DateAdd("n",oRst("EventDuration"),oRst("EventDate"))
              	  if DateDiff("d",dEnd,oRst("EventDate")) = 0 then
      	             dEnd = FormatDateTime(dEnd,vbLongTime)
              	  end if
              	  dEnd = " - " & dEnd
             	else
              	  dEnd = ""
              end if

             	if oRst("CategoryID") <> 0 then
                 sCategory = "(" & getCategoryName(oRst("CategoryID")) & ")"
           		 else
          			    sCategory = ""
             	end if

           		'-------------------------------------------------------------
           		'Used to trim seconds from dates displayed on calendar pages.
           		'9/9/2005 Vincent Evans
         				'Start trim code
         				'-------------------------------------------------------------
          				sDate1 = cStr(oRst("EventDate"))
      		    		sDate2 = cStr(dEnd)

         		 		iTrimDate1 = clng(InStrRev(sDate1,":"))
         		 		iTrimDate2 = clng(InStrRev(sDate2,":"))

         				'Retrieves AM/PM, trims final :00 and builds string
         		 		if iTrimDate1 > 0 then
         		   			sTemp = Right(sDate1, 2)
         		   			sDate1 = Left(sDate1,iTrimDate1 - 1) & " " & sTemp
         		   			sTemp = ""
         		  	end if

         		 		if iTrimDate2 > 0 then
         		   			sTemp = Right(sDate2, 2)
         		   			sDate2 = Left(sDate2,iTrimDate2 - 1) & " " & sTemp
         		   			sTemp = ""
      		    		end if

         				'-------------------------------------------------------------
         				'End trim code
         				'-------------------------------------------------------------

         				'Changed width from 75px to 25% to fix multilined time.
         		   sEvents = sEvents & "<tr>" & vbcrlf
              sEvents = sEvents & "    <td width=""25%"" valign=""top"" nowrap>" & sDate1 & " " & sDate2 & " " & oRst("TZAbbreviation") & "</td>" & vbcrlf
              sEvents = sEvents & "    <td><font color=""" & oRst("Color") & """ style=""font-family: Tahoma,Arial;font-size: 11px;""><i>" & sCategory & "</i>&nbsp;<b>" & oRst("Subject") & "</font></b><br />" & oRst("Message") & "</td>" & vbcrlf
              sEvents = sEvents & "</tr>" & vbcrlf

              oRst.MoveNext
           loop
        end If
        oRst.Close 
        set oRst = Nothing 
     end if
  Else 

   	'DEFAULT FORM VALUES
    	sSub     = "CHECKED"
    	sDescrip = "CHECKED"

  end if

 'Determine which "Search By Date" has been selected:
  if request("DateSearch") <> "" then
     lcl_datesearch = request("DateSearch")
  else
     lcl_datesearch = ""
  end if

 'Determine which date search option to "check"
  lcl_checked_all     = ""
  lcl_checked_before  = ""
  lcl_checked_after   = ""
  lcl_checked_between = ""

  if lcl_datesearch = "1" then
     lcl_checked_all     = " checked"
  elseif lcl_datesearch  = "2" then
     lcl_checked_before  = " checked"
  elseif lcl_datesearch  = "3" then
     lcl_checked_after   = " checked"
  elseif lcl_datesearch  = "4" then
     lcl_checked_between = " checked"
  else
     lcl_checked_all     = " checked"
  end if

  lcl_before_date = request("datepickerbefore")
  lcl_after_date  = request("datepickerafter")
  lcl_start_date  = request("datepickerstart")
  lcl_end_date    = request("datepickerend")
%>

<p class="title">
<img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="calendar.asp<%=replace(lcl_calendarfeature_url,"&","?")%>"><%=langBackToCalendar%></a>

<div style="padding:10px;">
<form name="frmSearch" method="post" action="searchevents.asp">
  <input type="<%=lcl_hidden%>" name="_task" value="search" />
  <input type="<%=lcl_hidden%>" name="cal" value="<%=lcl_calendarfeatureid%>" />

<div class="shadow" style="width: 500px;">
<table border="0" cellspacing="0" cellpadding="2" class="tablelist" style="width: 500px;">
  <tr>
      <th colspan="2" align="left">Search For an Event</th>
  </tr>
  <tr>
	     <td width="25%">Keyword or Phrase: </td>
	     <td>
	         <input type="text" name="Keyword" style="width:200px;" maxlength="50" value="<%=Request.Form("Keyword")%>" />
	         <input class="button" type="button" value="Search" onclick="javascript:if(fnCheckNew()) {document.frmSearch.submit();} else {alert('Please be sure to enter a keyword and\/or make sure to check a field to search!');}" />
	     </td>
	 </tr>
	 <tr>
	     <td >Search In:</td>
	     <td>
	         Subject: <input type="checkbox" name="Subject" <%=sSub%> value="ON">&nbsp;&nbsp;&nbsp;&nbsp;
	         Description: <input type="checkbox" name="Descrip" <%=sDescrip%> value="ON">
      </td>
	 </tr>
	 <tr>
	     <td valign="top">Search By Date:</td>
	     <td>
					     <table border="0" cellspacing="0" cellpadding="0" width="100%">
						      <tr>
							         <td>All: </td>
	              	<td colspan="3"><input type="radio" name="DateSearch" value="1"<%=lcl_checked_all%>></td>
      						</tr>
						      <tr>
         							<td>Before: </td>
         							<td><input type="radio" name="DateSearch" value="2"<%=lcl_checked_before%>></td>
          						<td>Date:</td>
         							<td>
                    <input type="text" name="DatePickerBefore" value="<%=lcl_before_date%>" style="width:133px;" maxlength="50" >
            								&nbsp;<a href="javascript:void doCalendar('DatePickerBefore');"><img src="../images/calendar.gif" border="0" /></a>
         							</td>
      						</tr>
      						<tr>
						         	<td>After:</td>
         							<td><input type="radio" name="DateSearch" value="3"<%=lcl_checked_after%>></td>
         							<td>Date:</td>	
         							<td>
                    <input type="text" name="DatePickerAfter" value="<%=lcl_after_date%>" style="width:133px;" maxlength="50" >
            								&nbsp;<a href="javascript:void doCalendar('DatePickerAfter');"><img src="../images/calendar.gif" border="0" /></a>
         							</td>					
      						</tr>
      						<tr>
         							<td>Between:</td>
         							<td><input type="radio" name="DateSearch" value="4"<%=lcl_checked_between%>></td>
         							<td>Start:</td>						
         							<td>
                    <input type="text" name="DatePickerStart" value="<%=lcl_start_date%>" style="width:133px;" maxlength="50" >
            								&nbsp;<a href="javascript:void doCalendar('DatePickerStart');"><img src="../images/calendar.gif" border="0" /></a>
	        							</td>
      						</tr>
      						<tr>
						         	<td colspan="2">&nbsp;</td>
         							<td>End:</td>
         							<td>
            								<input type="text" name="DatePickerEnd" value="<%=lcl_end_date%>" style="width:133px;" maxlength="50" >
            								&nbsp;<a href="javascript:void doCalendar('DatePickerEnd')"><img src="../images/calendar.gif" border="0" /></a>
         							</td>
       					</tr>
     					</table>
				  </td>
  </tr>            
</table>
</div>
</form>

<%
  if bResults then

     response.write "<div class=""shadow"">" & vbcrlf
     response.write "<table border=""0""  cellspacing=""0"" cellpadding=""2"" class=""tablelist"">" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <th align=""left"">" & langDateTime & "</th>" & vbcrlf
     response.write "      <th align=""left"">" & langEvent    & "</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write sEvents
     response.write "</table>" & vbcrlf
     response.write "</div>" & vbcrlf

  elseif bSearch then

     response.write "<div class=""box_header4"">Search Results</div>" & vbcrlf
  			response.write "<div class=""groupSmall4""><b>Your search yielded no results.</b></div>" & vbcrlf

  else
  end if
%>

  </div>
</div>

<!--#Include file="../admin_footer.asp"-->

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' BEGIN: VISITOR TRACKING
'--------------------------------------------------------------------------------------------------
'	iSectionID = 55
'	If request("date") <> "" Then
'		sDocumentTitle = "CALENDAR DATE: " & CDATE(request("date"))
'	Else
'		sDocumentTitle = "UNSPECIFIED CALENDAR DATE VIEW"
'	End If
'	sURL = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
'	datDate = Date()
'	datDateTime = Now()
'	sVisitorIP = request.servervariables("REMOTE_ADDR")
'	Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,session("orgid"))
'--------------------------------------------------------------------------------------------------
' END: VISITOR TRACKING
'--------------------------------------------------------------------------------------------------

function dbsafe(p_value)
  lcl_return = ""

  lcl_return = replace(p_value,"'","''")

  dbsafe = lcl_return

end function
%>
