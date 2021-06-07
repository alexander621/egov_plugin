<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
Response.Buffer = True
%>

<%
' CAPTURE CURRENT PATH
Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
Session("RedirectLang") = "Return to Action Line"
%>


<%If iorgid = 7 Then %>
	<title><%=sOrgName%></title>
<%Else%>
	<title>E-Gov Services <%=sOrgName%></title>
<%End If%>

<link rel="stylesheet" href="../css/styles.css" type="text/css">

	<link href="../global.css" rel="stylesheet" type="text/css">
	<script language="Javascript" src="../scripts/modules.js"></script>
<link rel="stylesheet" href="../css/styles.css" type="text/css">
<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" type="text/css">
<script language="Javascript" src="scripts/easyform.js"></script>

<script language=javascript>
function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}

function fnCheckNew(){

		if ((document.all.Search.Keyword.value != '') && ((document.all.Search.Subject.checked == true) || (document.all.Search.Descrip.checked == true)) ) {
			return true;
		}
		else
		{
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
  <!--
    body {scrollbar-base-color:#6699cc; scrollbar-highlight-color:#ffffff; scrollbar-arrow-color:#99ccff; font-family:Verdana,Tahoma,Arial; font-size:11px;}
    .cal {border-left:1px solid #0099ff; border-top:1px solid #0099ff; border-right:1px solid #0099ff;}
    .cal th {border-right:1px solid #0099ff; border-bottom:1px solid #0099ff; font-family:Tahoma,Arial; font-size:11px; color:#ffffff; text-align:left;}
    .cal td {border-bottom:1px solid #0099ff; font-family:Tahoma,Arial; font-size:11px;}
    select {font-family:Arial,Tahoma,Verdana; font-size:13px;}
  //-->
  </style>

</head>

<!--#Include file="../include_top.asp"-->

<!--BODY CONTENT-->

  <%
  Dim dDate
  'Variables for searching by date
  dim dDateSearch
  Dim dDateStart
  Dim dDateEnd


  If IsDate(Request.QueryString("date")) Then '1
    dDate = CDate(Request.QueryString("date"))
  Else
    If IsDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year")) Then '2
      dDate = CDate(Request.QueryString("month") & "-" & Request.QueryString("day") & "-" & Request.QueryString("year"))
    Else
      dDate = Date()
    End If '1
  End If '2

  bResults = 0

If Request.Form("_task") = "search" Then '3

  Dim oCmd, oRst, sEvents

  If (UCase(Request.Form("Subject")) = "ON" AND UCase(Request.Form("Descrip")) = "ON") Then	'4
	sProc = "SearchMonthEventsBySubjectDescrip"
  	bSearch = 1
  	sSub = "CHECKED"
  	sDescrip = "CHECKED"
  ElseIf UCase(Request.Form("Subject")) = "ON" Then
  	sProc = "SearchMonthEventsBySubject"
  	bSearch = 1
  	sSub = "CHECKED"
  ElseIf UCase(Request.Form("Descrip")) = "ON" Then
  	sProc = "SearchMonthEventsByDescrip"
  	bSearch = 1
  	sDescrip = "CHECKED"
  Else
    sSub = ""
    sDescrip = ""
  End If '3
  'Section used to determine which stored procedure to call and pass correct variables.
  
  If UCase(Request.Form("DateSearch")) = 1 Then '5
    Set oCmd = Server.CreateObject("ADODB.Command")
    With oCmd
	  'Orignal Code for calendar search:
      .ActiveConnection = Application("DSN")
      .CommandText = sProc
      .CommandType = adCmdStoredProc
      .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, iorgid)
      .Parameters.Append oCmd.CreateParameter("Date", adDateTime, adParamInput, 4, dDate)
      If bSearch Then .Parameters.Append oCmd.CreateParameter("Keyword", adVarChar, adParamInput, 50, Request.Form("Keyword"))
    End With

  'Before Date
  ElseIf UCase(Request.Form("DateSearch")) = 2 AND Request.Form("DatePickerBefore") <> "" Then
  	If  IsDate(Request.Form("DatePickerBefore")) Then '6
		dDateSearch = CDate(Request.Form("DatePickerBefore"))
	Else
		dDateSearch = Date()
	End If '4
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
		  'Altered to allow for before/after date search. 0 bit = before 1 bit = after date
		  .ActiveConnection = Application("DSN")
		  .CommandText = sProc & "BefAft"
		  .CommandType = adCmdStoredProc
		  .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, iorgid)
		  .Parameters.Append oCmd.CreateParameter("Date", adDateTime, adParamInput, 4, dDateSearch)
		  If bSearch Then .Parameters.Append oCmd.CreateParameter("Keyword", adVarChar, adParamInput, 50, Request.Form("Keyword"))
		  .Parameters.Append oCmd.CreateParameter("BefAft", adBit, adParamInput, 1, 0)
		End With	

  'After Date
  ElseIf UCase(Request.Form("DateSearch")) = 3 AND Request.Form("DatePickerAfter") <> "" Then
	If IsDate(Request.Form("DatePickerAfter")) Then '7
		dDateSearch = CDate(Request.Form("DatePickerAfter"))
	Else
		dDateSearch = Date()
	End If '6
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
		  .ActiveConnection = Application("DSN")
		  .CommandText = sProc & "BefAft"
		  .CommandType = adCmdStoredProc
		  .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, iorgid)
		  .Parameters.Append oCmd.CreateParameter("Date", adDateTime, adParamInput, 4, dDateSearch)
		  If bSearch Then .Parameters.Append oCmd.CreateParameter("Keyword", adVarChar, adParamInput, 50, Request.Form("Keyword"))
		  .Parameters.Append oCmd.CreateParameter("BefAft", adBit, adParamInput, 1, 1)
		End With
    
  'Between dates (Range)
  ElseIf UCase(Request.Form("DateSearch")) = 4 AND Request.Form("DatePickerStart") <> "" AND Request.Form("DatePickerEnd") <> "" Then
	If IsDate(Request.Form("DatePickerStart")) And IsDate(Request.Form("DatePickerEnd")) Then '8
		dDateStart = CDate(Request.Form("DatePickerStart"))
		dDateEnd = CDate(Request.Form("DatePickerEnd"))
	Else
		If Not IsDate(Request.Form("DatePickerStart")) Or Not IsDate(Request.Form("DatePickerEnd")) Then
			If Not IsDate(Request.Form("DatePickerStart")) Then 
				dDateStart = Date()
			End If
			If Not IsDate(Request.Form("DatePickerEnd")) Then 
				dDateEnd = Date()
			End If
		End If
	End If '7
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
		  .ActiveConnection = Application("DSN")
		  .CommandText = sProc & "Between"
		  .CommandType = adCmdStoredProc
		  .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, iorgid)
		  .Parameters.Append oCmd.CreateParameter("DateStart", adDateTime, adParamInput, 4, dDateStart)
		  .Parameters.Append oCmd.CreateParameter("DateEnd", adDateTime, adParamInput, 4, dDateEnd)
		  If bSearch Then .Parameters.Append oCmd.CreateParameter("Keyword", adVarChar, adParamInput, 50, Request.Form("Keyword"))
		End With  
  End If '8




	'response.write sProc
	'response.end

    Set oRst = Server.CreateObject("ADODB.Recordset")
    With oRst
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
      .Open oCmd
    End With
    Set oCmd = Nothing

    If Not oRst.EOF Then '9
    	bResults = 1
      Do While Not oRst.EOF
      	If oRst("EventDuration") > 0 Then '10
      	  dEnd = DateAdd("n",oRst("EventDuration"),oRst("EventDate"))
      	  If DateDiff("d",dEnd,oRst("EventDate")) = 0 Then '11
      	    dEnd = FormatDateTime(dEnd,vbLongTime)
      	  End If '9
      	  dEnd = " - " & dEnd
      	Else
      	  dEnd = ""
        End If '10

      	If oRst("CategoryID") <> 0 Then '12
      		Set oCmd2 = Server.CreateObject("ADODB.Command")
			    With oCmd2
			      .ActiveConnection = Application("DSN")
			      .CommandText = "GetCategoryName"
			      .CommandType = adCmdStoredProc
			      .Parameters.Append oCmd2.CreateParameter("CategoryID", adInteger, adParamInput, 4, oRst("CategoryID"))
			    End With

			    Set oRst2 = Server.CreateObject("ADODB.Recordset")
			    With oRst2
			      .CursorLocation = adUseClient
			      .CursorType = adOpenStatic
			      .LockType = adLockReadOnly
			      .Open oCmd2
			    End With
			    Set oCmd2 = Nothing

			    sCategory = "(" & oRst2("CategoryName") & ") "
		Else
			    sCategory = ""
		End If '11

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
		If iTrimDate1 > 0 Then '13
			sTemp = Right(sDate1, 2)
			sDate1 = Left(sDate1,iTrimDate1 - 1) & " " & sTemp
			sTemp = ""
		End If '12

		If iTrimDate2 > 0 Then '14
			sTemp = Right(sDate2, 2)
			sDate2 = Left(sDate2,iTrimDate2 - 1) & " " & sTemp
			sTemp = ""
		End If '13

		'-------------------------------------------------------------
		'End trim code
		'-------------------------------------------------------------

		'Changed width from 75px to 25% to fix multilined time.

        sEvents = sEvents & "<tr><td width=""25%"" valign=top>" & sDate1 & " " & sDate2 & " " & oRst("TZAbbreviation") & "</td><td><i><font color=""" & oRst("Color") & """>" & sCategory & "</i><b>" & oRst("Subject") & "</font></b><br>" & oRst("Message") & "</td></tr>"
        oRst.MoveNext
      Loop
    End If '14
    Set oRst = Nothing

Else

	' DEFAULT FORM VALUES
 	sSub = "CHECKED"
  	sDescrip = "CHECKED"

End If '15
  %>

<div style="width:90%;margin-left:20px;" >
  <p class=title><%=langEvents%>: <%= FormatDateTime(dDate, vbLongDate) %> <br>
  <img src="/<%=sorgVirtualSiteName%>/images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="calendar.asp"><%=langBackToCalendar%></a>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
            <!-- START: NEW CATEGORY -->
      <td colspan="2" valign="top">
        <form name="Search" method=post action="searchevents.asp" method="post">
          <input type="hidden" name="_task" value="search">

          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:if(fnCheckNew()) {document.all.Search.submit();} else {alert('Please be sure to enter a keyword and\/or make sure to check a field to search!');}">Search</a></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
            <tr>
              <th align="left" colspan="2">Search For an Event</th>
            </tr>
            <tr>
              <td valign="top" width="20%">Keyword or Phrase:</td>
              <td>
                <input type="text" name="Keyword" style="width:200px;" maxlength="50" value="<%=Request.Form("Keyword")%>">
              </td>
            </tr>
            <tr>
              <td valign="top">Search In:</td>
              <td>
              	Subject: <input type="checkbox" name="Subject" <%=sSub%> value="ON">&nbsp;&nbsp;&nbsp;&nbsp;
              	Description: <input  type="checkbox" name="Descrip" <%=sDescrip%> value="ON">
            </tr>
            <tr>
              <td valign="top">Search By Date:</td>
              <td>
				<table>
					<tr>
						<td>All: </td>
              			<td><INPUT type="radio" name=DateSearch checked=1 value=1></td>
					</tr>
					<tr>
						<td>Before: </td>
						<td><INPUT type="radio" name=DateSearch value=2></td>
						<td>Date:</td>
						<td><input type="text" name="DatePickerBefore" <%=sDateBefore%> style="width:133px;" maxlength="50" >
							&nbsp;<a href="javascript:void doDate('DatePickerBefore',1);"><%=langChoose%></a>
						</td>
					</tr>
					<tr>
						<td>After:</td>
						<td><INPUT type="radio" name=DateSearch value=3></td>
						<td>Date:</td>	
						<td><input type="text" name="DatePickerAfter" <%=sDateAfter%> style="width:133px;" maxlength="50" >
							&nbsp;<a href="javascript:void doDate('DatePickerAfter',1);"><%=langChoose%></a>
						</td>					
					</tr>
					<tr>
						<td>Between:</td>
						<td><INPUT type="radio" name=DateSearch value=4></td>
						<td>Start:</td>						
						<td><input type="text" name="DatePickerStart" <%=sDateStart%> style="width:133px;" maxlength="50" >
							&nbsp;<a href="javascript:void doDate('DatePickerStart',1);"><%=langChoose%></a>

						</td>
					</tr>
					<tr>
						<td colspan=2>&nbsp;</td>
						<td>
							End:
						</td>
						<td>
							<input type="text" name="DatePickerEnd" <%=sDateEnd%> style="width:133px;" maxlength="50" >
							&nbsp;<a href="javascript:void doDate('DatePickerEnd',1)"><%=langChoose%></a>
						</td>
					</tr>
				</table>
			  </td>
            </tr>            
          </table>
          <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:if(fnCheckNew()) {document.all.Search.submit();} else {alert('Please be sure to enter a keyword and\/or make sure to check a field to search!');}">Search</a>


		  </div>
		</form>
      </td>
        <!-- END: NEW CATEGORY -->
    </tr>
  </table>

  <p>
  <table border="0" cellpadding="4" cellspacing="0" class="cal" width="100%">
    <% If bResults Then
    %><tr bgcolor="#1c4aab"><th><%=langDateTime%></th><th><%=langEvent%></th></tr><%
			response.write sEvents
		ElseIf bSearch Then
	  %><tr bgcolor="#1c4aab"><th colspan="2">Search Results</th></tr><%
			response.write "<tr><TD colspan=2><P><B>Your search yielded no results.</b></P></td></tr>"
	  Else
		End if
	%>
  </table>
</div>


<!--SPACING CODE-->
<p>&nbsp;<bR>&nbsp;<bR>&nbsp;<bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="../include_bottom.asp"-->

<%
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
%>
