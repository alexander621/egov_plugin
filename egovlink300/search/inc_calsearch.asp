
<!--BODY CONTENT-->
<style type="text/css">
  <!--
    body {scrollbar-base-color:#6699cc; scrollbar-highlight-color:#ffffff; scrollbar-arrow-color:#99ccff; font-family:Verdana,Tahoma,Arial; font-size:11px;}
    .cal {border-left:1px solid #0099ff; border-top:1px solid #0099ff; border-right:1px solid #0099ff;}
    .cal th {border-right:1px solid #0099ff; border-bottom:1px solid #0099ff; font-family:Tahoma,Arial; font-size:11px; color:#ffffff; text-align:left;}
    .cal td {border-bottom:1px solid #0099ff; font-family:Tahoma,Arial; font-size:11px;}
    select {font-family:Arial,Tahoma,Verdana; font-size:13px;}
  //-->
  </style>

  <%
  Dim dDate
  'Variables for searching by date
  dim dDateSearch
  Dim dDateStart
  Dim dDateEnd


  bResults = 0

  Dim oCmd, sEvents

  'Section used to determine which stored procedure to call and pass correct variables.
  
    sSQL = "SELECT e.EventID, e.EventDate, e.EventDuration, t.TZAbbreviation, e.Subject, e.Message, e.CategoryID, c.Color "
	sSql = sSql & " FROM Events e LEFT JOIN TimeZones t ON t.TimeZoneID = e.EventTimeZoneID "
	sSql = sSql & " LEFT JOIN EventCategories c ON e.CategoryID = c.CategoryID WHERE e.OrgID = '" & iOrgID
	sSql = sSql & "' AND (Subject LIKE '%" & DBsafe(request("SearchString")) & "%' OR Message LIKE '%" & DBsafe(request("SearchString")) & "%') ORDER BY e.EventDate"
    
	Set oRst = Server.CreateObject("ADODB.Recordset")
	oRst.Open sSQL, Application("DSN"), 3, 1

    RecordCount = oRst.RecordCount

    If Not oRst.EOF Then '9
    	bResults = 1
		i=0
		Do While Not oRst.EOF and i < 5
      
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
		i = i + 1
      Loop
    End If '14
    Set oRst = Nothing

  %>

<div style="width:90%;margin-left:20px;" >
  <p>
  <form name="eventsearch" action="../events/searchevents.asp" Method="post">
     <input type="hidden" name="_task" value="search" />
	 <input type="hidden" name="Keyword" value="<%=request("searchstring")%>" />
	 <input type="hidden" name="Subject" value="ON" />
	 <input type="hidden" name="Descrip" value="ON" />
	 <input type="hidden" name="DateSearch" value="1" />
  </form>
    <% If bResults Then%>
	<font class=label>Events 1 to <% if RecordCount < 5 then%><%=RecordCount%><%else%>5<%end if%> of <%=RecordCount%> matching the query "<i><%=request("searchstring")%></i>".</font>
    <br>
<a href="javascript:document.eventsearch.submit();">VIEW ALL RESULTS</a>

	<% end if %>
  <table border="0" cellpadding="4" cellspacing="0" class="tablelist" width="100%">
    <% If bResults Then
    %><tr style="height:26px;"><th class="subheading"><%=langDateTime%></th><th class="subheading"><%=langEvent%></th></tr><%
			response.write sEvents
		ElseIf bSearch Then
	  %><tr bgcolor="#1c4aab"><th colspan="2">Search Results</th></tr><%
			response.write "<tr><td colspan=2><p><b>Your search yielded no results.</b></p></td></tr>"
	  Else
	  %><tr bgcolor="#1c4aab"><th colspan="2">Search Results</th></tr><%
			response.write "<tr><td colspan=2><p><b>Your search yielded no results.</b></p></td></tr>"
		End if
	%>
  </table>
    <% If bResults Then%>
<a href="javascript:document.eventsearch.submit();">VIEW ALL RESULTS</a>
	<% end if %>
</div>



