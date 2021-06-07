<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
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
' 1.3  11/19/13  Terry Foster - CLng Bug Fix
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if this is a custom calendar
 if trim(request("cal")) <> "" and isnumeric(replace(trim(request("cal")),"'","")) then
    'lcl_calendarfeature      = trim(request("cal"))
    'lcl_calendarfeature_url  = "&cal=" & lcl_calendarfeature
    lcl_calendarfeatureid    = CLng(replace(trim(request("cal")),"'",""))
    lcl_calendarfeature      = getFeatureByID(iorgid, lcl_calendarfeatureid)
    lcl_calendarfeature_url  = "&cal=" & lcl_calendarfeatureid
    lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"
 else
    lcl_calendarfeatureid    = ""
    lcl_calendarfeature      = ""
    lcl_calendarfeature_url  = ""
    lcl_calendarfeature_name = ""
 end if

 lcl_title = sOrgName

 if iorgid <> 7 then
    lcl_title = "E-Gov Services " & sOrgName
 end if

 'Variables for searching by date
  dim dDate, dDateSearch, dDateStart, dDateEnd

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

  lcl_task = request("_task")
  lcl_task = UCASE(lcl_task)

  if lcl_task = "SEARCH" then

     Dim oCmd, oRst, sEvents
     lcl_where_clause = ""

    'Get all of the form variables and validate them
     lcl_keyword     = request("keyword")
     lcl_keyword     = UCASE(lcl_keyword)
     lcl_keyword     = dbsafe(lcl_keyword)

     lcl_subject     = request("subject")
     lcl_subject     = UCASE(lcl_subject)

     lcl_description = request("descrip")
     lcl_description = UCASE(lcl_description)

     lcl_dateSearch  = request("datesearch")
     lcl_dateSearch  = ucase(lcl_dateSearch)
	 if not isnumeric(lcl_dateSearch) then lcl_dateSearch = "0"

    '*** These dates are validated in the code below ***
     lcl_datePickerBefore = request("datePickerBefore")
     lcl_datePickerAfter  = request("datePickerAfter")
     lcl_datePickerStart  = request("datePickerStart")
     lcl_datePickerEnd    = request("datePickerEnd")

     if lcl_subject = "ON" AND lcl_description = "ON" then
  	      'sProc    = "SearchMonthEventsBySubjectDescrip"
         lcl_where_clause = " AND (UPPER(subject) LIKE ('%" & lcl_keyword & "%') "
         lcl_where_clause = lcl_where_clause & " OR UPPER(message) LIKE ('%" & lcl_keyword & "%')) "
        	bSearch  = 1
     	  	sSub     = " checked=""checked"""
     	  	sDescrip = " checked=""checked"""
  	  elseif lcl_subject = "ON" then
      	  'sProc    = "SearchMonthEventsBySubject"
         lcl_where_clause = " AND UPPER(subject) LIKE ('%" & lcl_keyword & "%') "
     	  	bSearch  = 1
     	  	sSub     = " checked=""checked"""
  	  elseif lcl_description = "ON" then
     	  	'sProc    = "SearchMonthEventsByDescrip"
         lcl_where_clause = " AND UPPER(message) LIKE ('%" & lcl_keyword & "%') "
     	  	bSearch  = 1
     	  	sDescrip = " checked=""checked"""
     else
     	   sSub     = ""
     	   sDescrip = ""
     end if

    'Section used to determine which stored procedure to call and pass correct variables.
     lcl_execute_query = "Y"

    'Build the query
     sSQL = "SELECT e.EventID, "
     sSQL = sSQL & " e.EventDate, "
     sSQL = sSQL & " e.EventDuration, "
     sSQL = sSQL & " t.TZAbbreviation, "
     sSQL = sSQL & " e.Subject, "
     sSQL = sSQL & " e.Message, "
     sSQL = sSQL & " e.CategoryID, "
     sSQL = sSQL & " c.Color "
     sSQL = sSQL & " FROM Events e "
     sSQL = sSQL &   " INNER JOIN TimeZones t ON t.timezoneid = e.eventtimezoneid "
     sSQL = sSQL &   " LEFT JOIN EventCategories c ON e.categoryid = c.categoryid "
     sSQL = sSQL & " WHERE e.orgid = " & iorgid

     if lcl_calendarfeature <> "" then
        sSQL = sSQL & " AND e.calendarfeature = '" & dbsafe(lcl_calendarfeature) & "' "
     else
        sSQL = sSQL & " AND (e.calendarfeature IS NULL OR e.calendarfeature = '') "
     end if

     sSQL = sSQL & lcl_where_clause

     if lcl_dateSearch = 1 then
        'nothing to add to query

    'Before Date
     elseif lcl_dateSearch = 2 AND lcl_datePickerBefore <> "" then
   	    if IsDate(lcl_datePickerBefore) then
     		    dDateSearch = CDate(lcl_datePickerBefore)
       	else
		         dDateSearch = Date()
        end if

        lcl_where_clause_before = " AND e.eventdate <= '" & dbsafe(dDateSearch) & "' "

    'After Date
     elseif lcl_dateSearch = 3 AND lcl_datePickerAfter <> "" then
   	    if IsDate(lcl_datePickerAfter) then
     		    dDateSearch = CDate(lcl_datePickerAfter)
   	    else
     		    dDateSearch = Date()
       	end if

        lcl_where_clause_after = " AND e.eventdate >= '" & dbsafe(dDateSearch) & "' "

    'Between dates (Range)
     elseif lcl_dateSearch = 4 AND lcl_datePickerStart <> "" AND lcl_datePickerEnd <> "" then
   	    if IsDate(lcl_datePickerStart) AND IsDate(lcl_datePickerEnd) then
     		    dDateStart = CDate(lcl_datePickerStart)
     		    dDateEnd = CDate(lcl_datePickerEnd)
   	    else
      	    if NOT IsDate(lcl_datePickerStart) OR NOT IsDate(lcl_datePickerEnd) then
       			    if NOT IsDate(lcl_datePickerStart) then
         				    dDateStart = Date()
           			end if
           			if NOT IsDate(lcl_datePickerEnd) then
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

        lcl_where_clause_between = " AND (e.eventdate BETWEEN '" & dbsafe(lcl_start_date) & "' "
        lcl_where_clause_between = lcl_where_clause_between & " AND '" & dbsafe(lcl_end_date) & "') "

     else
        lcl_execute_query = "N"
     end if

     if lcl_execute_query = "Y" then

       'Append BEFORE where clause
        if lcl_where_clause_before <> "" then
           sSQL = sSQL & lcl_where_clause_before
        end if

       'Append AFTER where clause
        if lcl_where_clause_after <> "" then
           sSQL = sSQL & lcl_where_clause_after
        end if

       'Append BETWEEN where clause
        if lcl_where_clause_between <> "" then
           sSQL = sSQL & lcl_where_clause_between
        end if

        sSQL = sSQL & " ORDER BY e.eventdate "
        set oRst = Server.CreateObject("ADODB.Recordset")
        oRst.Open sSQL, Application("DSN"), 3, 1

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
              sEvents = sEvents & "    <td width=""25%"" valign=""top"" nowrap>" & vbcrlf
              sEvents = sEvents &          sDate1 & " " & sDate2 & " " & oRst("TZAbbreviation") & "</td>" & vbcrlf
              sEvents = sEvents & "    <td><i><font color=""" & oRst("Color") & """>" & sCategory & "</i>&nbsp;<b>" & oRst("Subject") & "</font></b><br />" & oRst("Message") & "</td>" & vbcrlf
              sEvents = sEvents & "</tr>" & vbcrlf

              oRst.MoveNext
           loop
        end if
        set oRst = nothing
     end if
  else

   	'DEFAULT FORM VALUES
    	sSub     = " checked=""checked"""
    	sDescrip = " checked=""checked"""

  end if

 'Determine which "Search By Date" has been selected:
  if request("DateSearch") <> "" then
     lcl_dateSearch = request("DateSearch")
     lcl_dateSearch = UCASE(lcl_dateSearch)
  else
     lcl_dateSearch = ""
  end if

 'Determine which date search option to "check"
  lcl_checked_all     = ""
  lcl_checked_before  = ""
  lcl_checked_after   = ""
  lcl_checked_between = ""

  if lcl_dateSearch = "1" then
     lcl_checked_all     = " checked=""checked"""
  elseif lcl_dateSearch  = "2" then
     lcl_checked_before  = " checked=""checked"""
  elseif lcl_dateSearch  = "3" then
     lcl_checked_after   = " checked=""checked"""
  elseif lcl_dateSearch  = "4" then
     lcl_checked_between = " checked=""checked"""
  else
     lcl_checked_all     = " checked=""checked"""
  end if

  lcl_before_date = request("datepickerbefore")
  lcl_after_date  = request("datepickerafter")
  lcl_start_date  = request("datepickerstart")
  lcl_end_date    = request("datepickerend")
%>
<html>
<head>
  <title><%=lcl_title%></title>

	 <link rel="stylesheet" type="text/css" href="../css/styles.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
	 <link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	 <script type="text/javascript" src="../scripts/modules.js"></script>
	 <script type="text/javascript" src="../scripts/easyform.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>

	<script type="text/javascript">

 $(document).ready(function() {
    $('#returnButton').click(function() {
       location.href = 'calendar.asp'<%=replace(lcl_calendarfeature_url,"&","?")%>;
    });

    $('#searchButton').click(function() {
       if(fnCheckNew()) {
          $('#searchForm').submit();
       } else {
          alert('Please be sure to enter a keyword and/or make sure to check a field to search.');
       }
    });
 });

	function openWin2(url, name) 
	{
	  popupWin = window.open(url, name,"resizable,width=500,height=450");
	}

	function fnCheckNew() {
   if (($('#Keyword').val() != '') && ($('#Subject').prop('checked') || $('#Descrip').prop('checked'))) {
			     return true;
 		} else {
     			return false;
 		}
	}

	function doDate(returnfield, num) {
	  w = (screen.width - 350)/2;
	  h = (screen.height - 350)/2;
	  eval('DatePickerWin=window.open("calendarpicker.asp?r=" + returnfield + "&n=" + num, "_calendar", "width=350,height=250,toolbar=0,status=yes,scrollbars=0,menubar=0,left=' + w + ',top=' + h + '")');
 }
	</script>

<style type="text/css">
   #content {
      padding-top: 10px;
      padding-left: 5px;
   }

   .purchasereport,
   .groupSmall4 {
   	  border-bottom-left-radius: 5px;
      border-bottom-right-radius: 5px;
	     -moz-border-radius-bottomleft: 5px;
	     -moz-border-radius-bottomright: 5px;
   	  -webkit-border-bottom-left-radius: 5px;
   	  -webkit-border-bottom-right-radius: 5px;
   }
</style>
</head>

<!--#Include file="../include_top.asp"-->
<%
  response.write "<p>" & vbcrlf
  response.write "<font class=""pagetitle"">Calendar Search" & lcl_calendarfeature_name & "</font><br />" & vbrlf

  RegisteredUserDisplay( "../" )

  response.write "<p class=""title"">" & vbcrlf
  response.write "<input type=""button"" name=""returnButton"" id=""returnButton"" value=""Back to Calendar"" class=""button"" />" & vbcrlf

  response.write "<form name=""searchForm"" id=""searchForm"" method=""post"" action=""searchevents.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""_task"" id=""_task"" value=""search"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""cal"" id=""cal"" value=""" & lcl_calendarfeatureid & """ />" & vbcrlf
  response.write "<div class=""box_header4"">Search For an Event</div>" & vbcrlf
  response.write "  <div class=""groupSmall4"">" & vbcrlf

  response.write "<table cellspacing=""0"" border=""0"">" & vbrlf
  response.write "  <tr>" & vbrlf
  response.write "	     <td valign=""top"" width=""20%"">Keyword or Phrase: </td>" & vbrlf
  response.write "	     <td>" & vbrlf
  response.write "	         <input type=""text"" name=""Keyword"" style=""width:200px;"" maxlength=""50"" value=""" & lcl_keyword & """ />" & vbrlf
  response.write "	         <input type=""button"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" />" & vbrlf
  response.write "	     </td>" & vbrlf
  response.write "	 </tr>" & vbrlf
  response.write "	 <tr>" & vbrlf
  response.write "	     <td valign=""top"">Search In:</td>" & vbrlf
  response.write "	     <td>" & vbrlf
  response.write "	         Subject: <input type=""checkbox"" name=""Subject"" id=""Subject"" value=""ON""" & sSub & " />&nbsp;&nbsp;&nbsp;&nbsp;" & vbrlf
  response.write "	         Description: <input type=""checkbox"" name=""Descrip"" id=""Descrip"" value=""ON""" & sDescrip & " />" & vbrlf
  response.write "      </td>" & vbrlf
  response.write "	 </tr>" & vbrlf
  response.write "	 <tr>" & vbrlf
  response.write "	     <td valign=""top"">Search By Date:</td>" & vbrlf
  response.write "	     <td>" & vbrlf
  response.write "					     <table>" & vbrlf
  response.write "						      <tr>" & vbrlf
  response.write "							         <td>All: </td>" & vbrlf
  response.write "	              	<td><input type=""radio"" name=""DateSearch"" id=""DateSearch1"" value=""1""" & lcl_checked_all & " /></td>" & vbrlf
  response.write "      						</tr>" & vbrlf
  response.write "						      <tr>" & vbrlf
  response.write "         							<td>Before: </td>" & vbrlf
  response.write "         							<td><input type=""radio"" name=""DateSearch"" id=""DateSearch2"" value=""2""" & lcl_checked_before & " /></td>" & vbrlf
  response.write "          						<td>Date:</td>" & vbrlf
  response.write "         							<td>" & vbrlf
  response.write "                    <input type=""text"" name=""DatePickerBefore"" id=""DatePickerBefore"" value=""" & lcl_before_date & """ style=""width:133px;"" maxlength=""50"" />" & vbrlf
  response.write "            								&nbsp;<a href=""javascript:void doDate('DatePickerBefore',1);""><img src=""../images/calendar.gif"" border=""0"" /></a>" & vbrlf
  response.write "         							</td>" & vbrlf
  response.write "      						</tr>" & vbrlf
  response.write "      						<tr>" & vbrlf
  response.write "						         	<td>After:</td>" & vbrlf
  response.write "         							<td><input type=""radio"" name=""DateSearch"" id=""DateSearch3"" value=""3""" & lcl_checked_after & " /></td>" & vbrlf
  response.write "         							<td>Date:</td>" & vbrlf
  response.write "         							<td>" & vbrlf
  response.write "                    <input type=""text"" name=""DatePickerAfter"" id=""DatePickerAfter"" value=""" & lcl_after_date & """ style=""width:133px;"" maxlength=""50"" />" & vbrlf
  response.write "            								&nbsp;<a href=""javascript:void doDate('DatePickerAfter',1);""><img src=""../images/calendar.gif"" border=""0"" /></a>" & vbrlf
  response.write "         							</td>" & vbrlf
  response.write "      						</tr>" & vbrlf
  response.write "      						<tr>" & vbrlf
  response.write "         							<td>Between:</td>" & vbrlf
  response.write "         							<td><input type=""radio"" name=""DateSearch"" id=""DateSearch4"" value=""4""" & lcl_checked_between & """ /></td>" & vbrlf
  response.write "         							<td>Start:</td>" & vbrlf
  response.write "         							<td>" & vbrlf
  response.write "                    <input type=""text"" name=""DatePickerStart"" id=""DatePickerStart"" value=""" & lcl_start_date & """ style=""width:133px;"" maxlength=""50"" />" & vbrlf
  response.write "            								&nbsp;<a href=""javascript:void doDate('DatePickerStart',1);""><img src=""../images/calendar.gif"" border=""0"" /></a>" & vbrlf
  response.write "	        							</td>" & vbrlf
  response.write "      						</tr>" & vbrlf
  response.write "      						<tr>" & vbrlf
  response.write "						         	<td colspan=""2"">&nbsp;</td>" & vbrlf
  response.write "         							<td>End:</td>" & vbrlf
  response.write "         							<td>" & vbrlf
  response.write "            								<input type=""text"" name=""DatePickerEnd"" id=""DatePickerEnd"" value=""" & lcl_end_date & """ style=""width:133px;"" maxlength=""50"" />" & vbrlf
  response.write "            								&nbsp;<a href=""javascript:void doDate('DatePickerEnd',1)""><img src=""../images/calendar.gif"" border=""0"" /></a>" & vbrlf
  response.write "         							</td>" & vbrlf
  response.write "       					</tr>" & vbrlf
  response.write "     					</table>" & vbrlf
  response.write "				  </td>" & vbrlf
  response.write "  </tr>" & vbrlf
  response.write "</table>" & vbrlf
  response.write "  </div>" & vbrlf
  response.write "</form>" & vbrlf
  response.write "<div id=""content"">" & vbcrlf

  if bResults then
     response.write "<table border=""0""  cellspacing=""0"" cellpadding=""2"" class=""purchasereport"" style=""width:600px;"">" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <th align=""left"">" & langDateTime & "</th>" & vbcrlf
     response.write "      <th align=""left"">" & langEvent & "</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write sEvents
     response.write "</table>" & vbcrlf

  elseif bSearch then
     response.write "<div class=""box_header4"">Search Results</div>" & vbcrlf
  			response.write "<div class=""groupSmall4""><strong>Your search yielded no results.</strong></div>" & vbcrlf
  end if

  response.write "</div>" & vbcrlf
  response.write "<p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>" & vbcrlf

'--------------------------------------------------------------------------------------------------
' BEGIN: VISITOR TRACKING
'--------------------------------------------------------------------------------------------------
	iSectionID = 55

	If request("date") <> "" Then
		sDocumentTitle = "CALENDAR DATE: " & CDATE(request("date"))
	Else
		sDocumentTitle = "UNSPECIFIED CALENDAR DATE VIEW"
	End If
	sURL = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
	datDate = Date()
	datDateTime = Now()
	sVisitorIP = request.servervariables("REMOTE_ADDR")
	Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
'--------------------------------------------------------------------------------------------------
' END: VISITOR TRACKING
'--------------------------------------------------------------------------------------------------

function dbsafe(p_value)
  lcl_return = ""

  lcl_return = replace(p_value,"'","''")

  dbsafe = lcl_return

end function
%>

<!--#Include file="../include_bottom.asp"-->

