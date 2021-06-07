<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
  Dim sError

 'Set Timezone information into session
  session("iUserOffset") = request.cookies("tz")

 'Override of value from common.asp
  sLevel = ""

 'Check for org features
  lcl_orghasfeature_action_line                     = orghasfeature("action line")
  lcl_orghasfeature_admin_locations                 = orghasfeature("admin locations")
  lcl_orghasfeature_customreports_helpdocumentation = orghasfeature("customreports_helpdocumentation")

 'Check for user permissions
  lcl_userhaspermission_action_line                     = userhaspermission(session("userid"),"action line")
  lcl_userhaspermission_site_linker                     = userhaspermission(session("userid"),"site linker")
  lcl_userhaspermission_product_usage_rpt               = userhaspermission(session("userid"),"product usage rpt")
  lcl_userhaspermission_customreports_helpdocumentation = userhaspermission(session("userid"),"customreports_helpdocumentation")

 'Check to see if there is a userid
  if session("userid") = 0 OR session("userid") = "" then
     lcl_welcome_message = langWelcomeGuest
  else
     lcl_welcome_message = langWelcomeBack & " " & session("FullName")
  end if
%>
<html>
<head>
  <title><%=langBSHome%></title>

  <link rel="stylesheet" type="text/css" href="global.css" />
  <link rel="stylesheet" type="text/css" href="menu/menu_scripts/menu.css" />

  <script language="javascript" src="scripts/modules.js"></script>
  <script language="javascript" src="scripts/ajaxLib.js"></script>

<script language="javascript">
<!--

//Set timezone in cookie to retrieve later
var d=new Date();
if(d.getTimezoneOffset) {
		 var iMinutes = d.getTimezoneOffset();
		 document.cookie = "tz=" + iMinutes;
}

//function doPicker(sFormField) {
//		w = (screen.width - 350)/2;
//		h = (screen.height - 350)/2;
//		eval('window.open("sitelinker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=470,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
//}

function doPicker(sFormField) {
  w = 600;
  h = 500;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  lcl_showFolderStart = "";

  //lcl_showFolderStart = "&folderStart=published_documents";
  //lcl_showFolderStart = "&folderStart=unpublished_documents";
  lcl_showFolderStart = "&folderStart=CITY_ROOT";


  pickerURL  = "sitelinker/default.asp";
  pickerURL += "?name=" + sFormField;
  pickerURL += lcl_showFolderStart;

  eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function insertAtCaret (textEl, text) {
  if (textEl.createTextRange && textEl.caretPos) {
		    var caretPos = textEl.caretPos;
  			 caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? text + ' ' : text;
  } else {
   			textEl.value = textEl.value + text;
	 }
}

function openWin2(url, name) {
		popupWin = window.open(url, name, "resizable,width=500,height=450");
}

function SetLocation( iUserId ) {
		//alert(document.LocationForm.locationid.options[document.LocationForm.locationid.selectedIndex].value);
		//doAjax('../includes/setuserlocation.asp', 'locationid=' + document.LocationForm.locationid.options[document.LocationForm.locationid.selectedIndex].value, 'LocationSet', 'get', '0');
		doAjax('includes/setuserlocation.asp', 'locationid=' + document.LocationForm.locationid.options[document.LocationForm.locationid.selectedIndex].value, '', 'get', '0');
}

function LocationSet( iLocationId ) {
		// Nothing happens here
		//alert( iLocationId );
}

//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content" style="width:auto;">
 	<div id="centercontent">

<table id="bodytable" border="0" cellpadding="0" cellspacing="0" class="start">
  <tr valign="top">
    		<td>
          <div id="featurewelcome"><%=lcl_welcome_message%>!</div>
      			 <div id="featuredate">Today is <%=FormatDateTime(Date(), vbLongDate)%></div>
	     </td>
      <td align="center">
          <%
           'Check for delegate assignees/assignments
            if lcl_orghasfeature_action_line AND lcl_userhaspermission_action_line then
               checkForDelegateAssigned    session("userid")
               checkForDelegateAssignments session("userid")
            end if
          %>
      </td>
  </tr>
  <tr>
		    <td valign="top" colspan="2">
       			<!-- the following is the login in form -->
   			<%
        If session("UserID") = 0 or session("UserID") = "" Then 
    				   If sError <> "" Then Response.Write sError 
   			%>
        		<form method="post" action="login.asp">
          <table border="0" cellpadding="5"  cellspacing="0" class="messagehead" width="151" height="100">
       					<tr>
          						<td width="151" colspan="2" bgcolor="#93bee1" class="section_hdr" style="border-bottom:1px solid #336699;" ><%=langMember+" "+langLogIn%></td>
       					</tr>
       					<tr>
          						<td><font face="Tahoma,Arial,Verdana" size="1"><%=langUserName%>:</font></td>
          						<td><input name="Username" size="10" style="width:78px; height:19px;"></td>
       					</tr>
       					<tr>
          						<td><font face="Tahoma,Arial,Verdana" size="1"><%=langPassword%>:</font></td>
         	 					<td><input type="Password" name="password" size="10" style="font-size:10px; font-family:Tahoma; width:78px; height:19px;"></td>
       					</tr>
       					<tr>
          						<td colspan="2"><input type="checkbox" name="SaveLogin"><font face="Tahoma,Arial,Verdana" size="1">Log me in automatically</font></td>
       					</tr>
       					<tr>
          						<td>&nbsp;</td>
          						<td><input type="submit" class="button" value="Login" style="font-family:Tahoma,Arial,Verdana; font-size:10px; width:55px;"></td>
       					</tr>
         			<tr>
          						<td colspan="2"><a href="dirs/lookuppassword.asp"><FONT SIZE="1"><%=langForgotPass%></FONT></A></td>
        				</tr>
      				</table>
     	  				<input type="hidden" name="_task" value="login" />
      				</form>
    			 <% end if %>
         	<p id="pagenote">
        		Use the Drop Down Navigation above to access the features of the site.
         	</p>
		<%
		sSQL = "SELECT emailaddress " _
			& " FROM users lu " _
			& " INNER JOIN users u ON u.orgid = lu.orgid and u.isdeleted = 0 " _
			& " INNER JOIN emailsuppressionlist sl ON sl.emailaddress = u.Email " _
			& " WHERE lu.userid = " & session("userid")
		set oRs = Server.CreateObject("ADODB.RecordSet")
		oRs.Open sSQL, Application("DSN"), 3, 1 

		if not oRs.EOF then%>
         	<p id="pagenote">
			The following email addresses are suppressed (please contact support if there is an error or you can delete the associated user to remove this message):<br />
			<% do while not oRs.EOF %>
				<%=oRs("emailaddress")%><br />
				<% oRs.MoveNext
			loop%>

         	</p>
		<% end if 
		oRs.Close
		Set oRs = Nothing
		%>
       <%
         if lcl_orghasfeature_admin_locations then
            response.write "<div class=""featuretitles"">Your Location Today</div>" & vbcrlf
            response.write "<p class=""featurebody"">" & vbcrlf
            response.write "<form name=""LocationForm"" method=""post"" action=""default.asp"">" & vbcrlf
            response.write "Location: &nbsp; " & vbcrlf

            ShowUserLocations

            response.write "</form>" & vbcrlf
            response.write "</p>" & vbcrlf
            response.write "</div>" & vbcrlf
         end if

        'BEGIN: General Tools -------------------------------------------------
         response.write "<div class=""featuretitles"">General Tools</div>" & vbcrlf
         response.write "<p class=""featurebody"">" & vbcrlf

        'Your User Profile ----------------------------------------------------
         response.write "  <a href=""dirs/update_user.asp?userid=" & session("userid") & """><strong>Your User Profile</strong></a> " & vbcrlf
         response.write "  allows you to edit your user profile and change your password." & vbcrlf

        'E-Gov Site Linker ----------------------------------------------------
         if lcl_userhaspermission_site_linker then
            response.write "<br />" & vbcrlf
            'response.write "<a href=""javascript:doPicker('UpdateEvent.Message')""><strong>E-Gov Site Linker</strong></a>" & vbcrlf
            response.write "<a href=""javascript:doPicker('UpdateEvent.Message');""><strong>E-Gov Site Linker</strong></a>" & vbcrlf
            response.write "  allows staff to create links to other pages within the E-Gov Site." & vbcrlf
         end if
	 
	 'E-Gov Alexa Searches
         if session("orgid") = "18" or session("orgid") = "76" or session("orgid") = "5" or session("orgid") = "26" or session("orgid") = "220" or session("orgid") = "208" or session("orgid") = "223" or session("orgid") = "187" then
            response.write "<br />" & vbcrlf
            response.write "<a href=""alexa_searches.asp""><strong>Voice Assistant Search Log</strong></a>" & vbcrlf
         end if

         response.write "<br />" & vbcrlf

         if UserIsRootAdmin( session("UserID") ) then
           'Organization Features ---------------------------------------------
            response.write "<br />" & vbcrlf
            response.write "<a href=""emailsuppression.asp""><strong>Email Suppression List</strong></a> " & vbcrlf
            response.write "  allows the root admin user to see the suppressed email list." & vbcrlf
           'Organization Features ---------------------------------------------
            response.write "<br />" & vbcrlf
            response.write "<a href=""admin/featureselection.asp""><strong>Organization Features</strong></a> " & vbcrlf
            response.write "  allows the root admin user to set organization features." & vbcrlf

           'Organization Properties -------------------------------------------
            response.write "<br />" & vbcrlf
            response.write "<a href=""admin/edit_org.asp""><strong>Organization Properties</strong></a> " & vbcrlf
            response.write "  allows the root admin user to maintain this organization's properties." & vbcrlf

           'Organization 'Edit Displays' --------------------------------------
            response.write "<br />" & vbcrlf
            response.write "<a href=""admin/clientdisplaylist.asp""><strong>Organization Displays</strong></a> " & vbcrlf
            response.write "  allows the root admin user to maintain this organization's displays." & vbcrlf

           'Orgs-to-Feature Summary -------------------------------------------
            response.write "<br />" & vbcrlf
            response.write "<a href=""admin/orgs_to_feature_summary.asp""><strong>Orgs-to-Feature Summary</strong></a> " & vbcrlf
            response.write "  allows the root admin user to find all organizations associated to a specific feature." & vbcrlf

           'Organization 'Assign Feature to Orgs' -----------------------------
            'response.write "<br />" & vbcrlf
            'response.write "<a href=""admin/featureassign.asp?assigntype=ORG&loc=HOME""><strong>Organization 'Assign Feature to Orgs'</strong></a> " & vbcrlf
            'response.write "  allows the root admin user to assign a feature to orgs." & vbcrlf

           'Organization 'Assign Feature to Users' ----------------------------
            'response.write "<br />" & vbcrlf
            'response.write "<a href=""admin/featureassign.asp?assigntype=USER&loc=HOME""><strong>Organization 'Assign Feature to Users'</strong></a> " & vbcrlf
            'response.write "  allows the root admin user to assign a feature to users." & vbcrlf

 			        response.write "<br />" & vbcrlf

           'Outage Maintenance ------------------------------------------------
 			        response.write "<br />" & vbcrlf
            response.write "<a href=""admin/outage_maint.asp""><strong>Outage Maintenance</strong></a>" & vbcrlf
            response.write "  allows the root admin user to turn off features for all organizations for code promotions." & vbcrlf

           'Page Log Summary --------------------------------------------------
            response.write "<br />" & vbcrlf
            response.write "<a href=""admin/pagelogsummary.asp""><strong>Page Log Summary</strong></a> " & vbcrlf
            response.write "  allows the root admin user to view the page log summary." & vbcrlf

            response.write "<br />" & vbcrlf
            response.write "<!--a href=""getcoordinates.aspx?orgid=" & session("orgid") & """><strong>Populate Coordinates</strong></a> " & vbcrlf
            response.write "  allows the root admin user to populate geocoordinates.-->" & vbcrlf
         end if

        'E-Gov Product Usage Stats --------------------------------------------
         if lcl_userhaspermission_product_usage_rpt then
            response.write "<br />" & vbcrlf
            response.write "<a href=""admin/org_stats.asp""><strong>E-Gov Product Usage Stats</strong></a> " & vbcrlf
            response.write "  allows you to view organization features usage." & vbcrlf
         end if

        'E-Gov Help Documentation ---------------------------------------------
         if lcl_orghasfeature_customreports_helpdocumentation AND lcl_userhaspermission_customreports_helpdocumentation then
            response.write "<br />" & vbcrlf
            response.write "<a href=""admin/org_stats.asp""><strong>E-Gov Help Documentation</strong></a> " & vbcrlf
            response.write "  allows you to view all of the help documentation that has been uploaded." & vbcrlf
         end if
        'END: General Tools ---------------------------------------------------

        'BEGIN: Recent Updates ------------------------------------------------
         response.write "</p>" & vbcrlf
         response.write "<div class=""featuretitles"">Recent Updates</div>" & vbcrlf
         response.write "<p class=""featurebody"">" & vbcrlf

         ShowNewsItems

         response.write "</p>" & vbcrlf
        'END: Recent Updates --------------------------------------------------
       %>
      </td>
  </tr>
</table>

  </div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
Sub ShowNewsItems()
	Dim sSql, oItems

	sSQL = "SELECT itemdate, itemtitle, itemtext, ISNULL(itemlinkurl,'') AS itemlinkurl "
 sSQL = sSQL & " FROM egov_news_items I, organizations O "
	sSQL = sSQL & " WHERE O.isadminnewssource = 1 "
 sSQL = sSQL & " AND O.orgid = I.orgid "
 sSQL = sSQL & " AND I.itemdisplay = 1 "
	sSQL = sSQL & " AND (publicationstart is null OR publicationstart <= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
	sSQL = sSQL & " AND (publicationend is null OR publicationend >= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
	sSQL = sSQL & " ORDER BY itemorder"

	set oItems = Server.CreateObject("ADODB.Recordset")
	oItems.Open sSQL, Application("DSN"), 3, 1

	if not oItems.eof then
  		do while not oItems.eof
    			response.write "<strong>" & oItems("itemdate") &  " &mdash; " & oItems("itemtitle") &  "</strong><br />" & oItems("itemtext") & vbcrlf

    			if oItems("itemlinkurl") <> "" then
      				response.write "<br /><a href=""" & oItems("itemlinkurl") & """ target=""_topwin""><strong>More &gt;&gt;</strong></a>" & vbcrlf
       end if

    			response.write "<br /><br />" & vbcrlf

    			oItems.movenext
    loop
	else
  		response.write "There are no updates at the present time. Please check back again for future updates." & vbcrlf
	end if

	oItems.Close
	set oItems = nothing 

End Sub 

'------------------------------------------------------------------------------
sub checkForDelegateAssignments(iUserID)

  if iUserID <> "" then
     sSQL = "SELECT userid, firstname, lastname "
     sSQL = sSQL & " FROM users "
     sSQL = sSQL & " WHERE delegateid = " & iUserID
     sSQL = sSQL & " ORDER BY lastname, firstname "

    	set oDelegates = Server.CreateObject("ADODB.Recordset")
    	oDelegates.Open sSQL, Application("DSN"), 3, 1

     if not oDelegates.eof then
        iRowCount = 0

        response.write "<fieldset>" & vbcrlf
        response.write "  <legend style=""margin-bottom:10px;"">You have been assigned as a delegate for:&nbsp;</legend>" & vbcrlf

        do while not oDelegates.eof
           iRowCount = iRowCount + 1

           if iRowCount > 1 then
              response.write "<br />" & vbcrlf
           end if

           response.write oDelegates("firstname") & " " & oDelegates("lastname")

           oDelegates.movenext
        loop

        response.write "</fieldset><br />" & vbcrlf

     end if

     oDelegates.close
     set oDelegates = nothing

  end if

end sub

'------------------------------------------------------------------------------
sub checkForDelegateAssigned(iUserID)

  if iUserID <> "" then
     sSQL = "SELECT userid, firstname, lastname "
     sSQL = sSQL & " FROM users "
     sSQL = sSQL & " WHERE userid = (select delegateid "
     sSQL = sSQL &                 " from users "
     sSQL = sSQL &                 " where userid = " & iUserID & ") "
     sSQL = sSQL & " ORDER BY lastname, firstname "

    	set oDelegates = Server.CreateObject("ADODB.Recordset")
    	oDelegates.Open sSQL, Application("DSN"), 3, 1

     if not oDelegates.eof then
        iRowCount = 0

        response.write "<fieldset>" & vbcrlf
        response.write "  <legend style=""margin-bottom:10px;"">You have assigned the following as your delegate:&nbsp;</legend>" & vbcrlf

        do while not oDelegates.eof
           iRowCount = iRowCount + 1

           if iRowCount > 1 then
              response.write "<br />" & vbcrlf
           end if

           response.write oDelegates("firstname") & " " & oDelegates("lastname")

           oDelegates.movenext
        loop

        response.write "</fieldset><br />" & vbcrlf

     end if

     oDelegates.close
     set oDelegates = nothing

  end if

end sub
%>
