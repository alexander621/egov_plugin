<!-- #include file="../includes/common.asp" //-->

<%
Dim oCmd, oRst, dDate, iDuration, sTimeZones, sLinks, bShown

If Not HasPermission("CanEditEvents") Then Response.Redirect "../default.asp"

If Request.Form("_task") = "newevent" Then

  dDate = CDate(Request.Form("DatePicker") & " " & Request.Form("Hour") & ":" & Request.Form("Minute") & " " & Request.Form("AMPM"))
  
  iDuration = Request.Form("Duration")
  If iDuration & "" <> "" Then
    iDuration = CLng(iDuration) * clng(Request.Form("DurationInterval"))
  Else
    iDuration = -1
  End If
  
  If Request.Form("CustomCategory") <> "" Then
  ' Create a New Category for this Organization
    Set oCmd = Server.CreateObject("ADODB.Command")
    With oCmd
	    .ActiveConnection = Application("DSN")
	    .CommandText = "NewCategory"
	    .CommandType = adCmdStoredProc
	    .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
	    .Parameters.Append oCmd.CreateParameter("CategoryName", adVarChar, adParamInput, 50, Request.Form("CustomCategory"))
	    .Parameters.Append oCmd.CreateParameter("Color", adVarChar, adParamInput, 7, "#000000")
	    .Execute
	  End With
	  Set oCmd = Nothing
	  
	  Set oCmd = Server.CreateObject("ADODB.Command")
	  With oCmd
	  	.ActiveConnection = Application("DSN")
	    .CommandText = "GetCategoryID"
	    .CommandType = adCmdStoredProc
	    .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
	    .Parameters.Append oCmd.CreateParameter("CategoryName", adVarChar, adParamInput, 50, Request.Form("CustomCategory"))
	    .Execute
    End With
    
    Set oRst = Server.CreateObject("ADODB.Recordset")
	  With oRst
	    .CursorLocation = adUseClient
	    .CursorType = adOpenStatic
	    .LockType = adLockReadOnly
	    .Open oCmd
	  End With
	  Set oCmd = Nothing
	  
	  iCategoryID = oRst("CategoryID")
	  
	  Set oRst = Nothing
	  
	Else
	  iCategoryID = Request.Form("Category")
  End if

  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "NewEvent"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
    .Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
    .Parameters.Append oCmd.CreateParameter("EventDate", adDateTime, adParamInput, 4, dDate)
    .Parameters.Append oCmd.CreateParameter("TimeZone", adInteger, adParamInput, 4, Request.Form("TimeZone"))
    .Parameters.Append oCmd.CreateParameter("Duration", adInteger, adParamInput, 4, iDuration)
    .Parameters.Append oCmd.CreateParameter("CategoryID", adInteger, adParamInput, 4, iCategoryID)
    .Parameters.Append oCmd.CreateParameter("Subject", adVarChar, adParamInput, 50, Request.Form("Subject"))
    .Parameters.Append oCmd.CreateParameter("Message", adVarChar, adParamInput, 5000, Request.Form("Message"))
    .Execute
  End With
  Set oCmd = Nothing

  Response.Redirect "../events"

ELSE
  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "ListTimeZones"
    .CommandType = adCmdStoredProc
    .Execute
  End With
  
  Set oRst = Server.CreateObject("ADODB.Recordset")
  With oRst
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open oCmd
  End With
  Set oCmd = Nothing
  
  Do While Not oRst.EOF
    sTimeZones=sTimeZones & "<option "
    if oRst("TimeZoneID") = 1 then sTimeZones=sTimeZones & "SELECTED"
    sTimeZones=sTimeZones & " value=""" & oRst("TimeZoneID") & """>" & oRst("TZName") & "</option>"
    oRst.movenext
  Loop
  
  if oRst.State=1 then oRst.Close
  set oRst=nothing
  

End If
%>

<html>
<head>
  <title><%=langBSEVents %></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script language="Javascript">
  <!--
    function doCalendar() {
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("calendarpicker.asp?p=1", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }
  //-->
  </script>
  <SCRIPT>
     function storeCaret (textEl) {
       if (textEl.createTextRange) 
         textEl.caretPos = document.selection.createRange().duplicate();
     }
     function insertAtCaret (textEl, text) {
       if (textEl.createTextRange && textEl.caretPos) {
         var caretPos = textEl.caretPos;
         caretPos.text =
           caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
             text + ' ' : text;
       }
       else
         textEl.value  = text;
     }
     </SCRIPT>


	     <script language="Javascript">
  <!--
    function doPicker(sFormField) {
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("../picker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }


	function fnCheckSubject(){
		
		if (document.NewEvent.Subject.value != '') {
			return true;
		}
		else
		{
			return false;
		}

		
	}
  //-->
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="document.all.DatePicker.focus();">
  <%DrawTabs tabHome,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_home.jpg"></td>
      <td><font size="+1"><b><%=langEvents%>: <%=langNewEvent%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langBackToEventList%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" nowrap>

        <!-- START: QUICK LINKS MODULE //-->
        
        <%
        sLinks = "<div style=""padding-bottom:8px;""><b>" & langEventLinks & "</b></div>"

        If HasPermission("CanEditEvent") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/calendar.gif"" align=""absmiddle"">&nbsp;<a href=""newevent.asp"">" & langNewEvent & "</a></div>"
          bShown = True
        End If
        
        If HasPermission("CanEditEvent") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/calendar.gif"" align=""absmiddle"">&nbsp;<a href=""../events"">" & langEditEvents & "</a></div>"
          bShown = True
        End If
        
        If HasPermission("CanEditEvent") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/calendar.gif"" align=""absmiddle"">&nbsp;<a href=""eventcategories.asp"">Manage Categories</a></div>"
          bShown = True
        End If
        
        If bShown Then
          Response.Write sLinks & "<br>"
        End If
        %>

        <% Call DrawQuicklinks("", 1) %>
        <!-- END: QUICK LINKS MODULE //-->

      </td>
        <!-- START: NEW EVENT -->
      <td colspan="2" valign="top">
        <form name="NewEvent" method=post action="newevent.asp" method="post">
          <input type="hidden" name="_task" value="newevent">

          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:if(fnCheckSubject()) {document.all.NewEvent.submit();} else {alert('Please enter a subject!');}"><%=langCreate%></a></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
            <tr>
              <th align="left" colspan="2"><%=langNewEvent%></th>
            </tr>
            <tr>
              <td valign="top"><%=langDate%>:</td>
              <td><input type="text" name="DatePicker" style="width:133px;" maxlength="50" value="<%= Date() %>">&nbsp;<a href="javascript:void doCalendar();"><%=langChoose%></a></td>
            </tr>
            <tr>
              <td valign="top" nowrap><%=langStartTime%>:</td>
              <td width="100%">
                <select name="Hour" class="time">
                  <option value="1">1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9" selected>9</option>
                  <option value="10">10</option>
                  <option value="11">11</option>
                  <option value="12">12</option>
                </select>
                :
                <select name="Minute" class="time">
                  <option value="00">00</option>
                  <option value="05">05</option>
                  <option value="10">10</option>
                  <option value="15">15</option>
                  <option value="20">20</option>
                  <option value="25">25</option>
                  <option value="30">30</option>
                  <option value="35">35</option>
                  <option value="40">40</option>
                  <option value="45">45</option>
                  <option value="50">50</option>
                  <option value="55">55</option>
                </select>
                <select name="AMPM" class="time">
                  <option value="AM">AM</option>
                  <option value="PM">PM</option>
                </select>
            </tr>
            <tr>
              <td valign="top" nowrap><%=langTimeZone%>:</td>
              <td>
                <select name="Timezone" class=time>
                  <%=sTimeZones%>
                </select>
              </td>
            <tr>
              <td valign="top"><%=langDuration%>:</td>
              <td>
                <input type="text" name="Duration" style="width:50px;" maxlength="5">
                <select name="DurationInterval" class="time" style="width:80px;">
                  <option value="1"><%=langMinutes%></option>
                  <option value="60"><%=langHours%></option>
                  <option value="1440"><%=langDays%></option>
                  <option value="10080"><%=langWeeks%></option>
                </select>
              </td>
            </tr>         
            <tr>
              <td valign="top">Category:</td>
              <td>
              	Choose:
                <select name="Category" class="time" style="width:80px;">
                  <option value="0">None</option>
<%
									Set oCmd = Server.CreateObject("ADODB.Command")
								  With oCmd
								    .ActiveConnection = Application("DSN")
								    .CommandText = "ListEventCategories"
								    .CommandType = adCmdStoredProc
								    .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
								    .Execute
								  End With
								  
								  Set oRst = Server.CreateObject("ADODB.Recordset")
								  With oRst
								    .CursorLocation = adUseClient
								    .CursorType = adOpenStatic
								    .LockType = adLockReadOnly
								    .Open oCmd
								  End With
									  
								   Do While Not oRst.EOF
								   	 sOption = "<option value=""" & oRst("CategoryID") & """>" & oRst("CategoryName") & "</option>" & vbCrLf
								     Response.write(sOption)
								     oRst.movenext
									 Loop
%>  
                </select>
                OR New Category:
                <input type="text" name="CustomCategory" style="width:133px;" maxlength="50">
              </td>
            </tr>
            <tr>
              <td valign="top"><%=langSubject%>:</td>
              <td><input type="text" name="Subject" style="width:400px;" maxlength="50"></td>
            </tr>
            <tr>
              <td valign="top"><%=langDetails%>:&nbsp;&nbsp;&nbsp;&nbsp;</td>
              <td><textarea name="Message" rows="5" style="width:400px;" ONSELECT="storeCaret(this);" ONCLICK="storeCaret(this);" ONKEYUP="storeCaret(this);" ONDBLCLICK="storeCaret(this);" ></textarea>
			  
	
				<input type=button value="Add Link" onClick="doPicker('NewEvent.Message');"> 

			  
			  </td>
            </tr>
          </table>
          <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:if(fnCheckSubject()) {document.all.NewEvent.submit();} else {alert('Please enter a subject!');}"><%=langCreate%></a>
		    
		  
		  </div>
		</form>
      </td>
        <!-- END: NEW EVENT -->
    </tr>
  </table>
</body>
</html>
