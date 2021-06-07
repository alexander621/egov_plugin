<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->

<% blnCanUpdateActionLine = HasPermission("CanEditActionRequests") %>

<html>
<head>
  <title><%=langBSActionLine%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabActionline,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><!--<img src="../images/icon_home.jpg">--></td>
      <td><font><b>Action Line Request Administration</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)"><%=langBackToStart%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" nowrap>

        <!-- START: QUICK LINKS MODULE //-->
        
        <%
        sLinks = "<div style=""padding-bottom:8px;""><b>Action Line Links</b></div>"

        If bCanEdit Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/calendar.gif"" align=""absmiddle"">&nbsp;<a href=""newevent.asp"">" & langNewEvent & "</a></div>"
          bShown = True
        End If
        
       ' If 1 Then
          'sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/edit.gif"" align=""absmiddle"">&nbsp;<a href=""actioncategories.asp"">Manage Form Categories</a></div>"
        '  bShown = True
       ' End If
        
        If bShown Then
          Response.Write sLinks & "<br />"
        End If
        %>

        <% 'Call DrawQuicklinks("", 1) %>
        <!-- END: QUICK LINKS MODULE //-->


		 <font size="1" face="Verdana,Tahoma,Arial"><b><%=langDesignby%></b><div class="logo"><A HREF="http://www.eclink.com"><img src="../images/poweredby.jpg" align="center" border="0"></A></div>

      </td>
        
      <td colspan="2" valign="top">
	  <!--BEGIN: ACTION LINE REQUEST LIST -->

	  <B>E-Gov Standard Tools</b>
      
	   <% If  blnCanUpdateActionLine Then %>

			<ul>
			 
			 <%'If session("orgid") = 15 Then%>
			<!--  <li><a href="action.asp">(E-Gov Forms Entry)</a> - Submit New Action Line Requests<br><br>-->
			  <%'end if%>
			  
			  <li><a href="action_line_list.asp">(E-Gov Request Manager)</a> - Manage Submitted Action Line Requests</li><br /><br />
			  <li><a href="manage_action_forms.asp">(E-Gov Alert Manager)</a> - Manage Your Action Line Request Forms</li><br /><br />
			  <li><a href="notification_report.asp">(E-Gov Notification Report)</a> - View/Manage Your Form Notificatins and Escalations</li><br /><br />
			  <li><a href="../admin/list_forms.asp">(E-Gov Forms Creator)</a> - Design/Modify Your Action Line Request Forms</li><br /><br />
			  <li><a href="actioncategories.asp">(E-Gov Form Categories Manager)</a> - Add/Edit/Delete Form Categories</li><br /><br />

			  <%'If session("orgid") = 26 or session("orgid") = 37 or session("orgid") = 8  Then%>
			  <% If OrgHasFeature( "action export" ) Then %>
					<li><a href="eval_export_form.asp">(E-Gov Action Line Data Export)</a> - Export submitted data to CSV file based on action form and date.<br /><br />
			  <% End If %>
			
			</ul>

		<%If session("orgquerytool") or session("OrgFormLetterOn") or session("orgfaq") Then %>
			<b>E-Gov Enterprise Tools</b>
			<ul>
		<%End If%>

			<!--FORM LETTER OPTION-->
			<%If session("OrgFormLetterOn") Then %>
				<li><a href="../formletters/list_letter.asp">(Form Letter Manager)</a> - Add/Modify Form Letters</li><br /><br />
			<%End If%>
			
			<!-- QUERY TOOL ACCESS -->
			<%If session("orgquerytool") Then %>
				<li><a href="../querytool/default.asp">(E-Gov Request Query/Reporting Tool)</a></li><br /><br />
			<%End If%>

			<!-- FAQ CREATOR -->
			<% if session("orgfaq") then %>	
				<li><a href="../faq/list_faq.asp">(FAQ Manager)</a> - Add/Modify Frequently Asked Questions</li><br /><br />
			<%End If%>
		
		<%If session("orgquerytool") or session("OrgFormLetterOn") or session("orgfaq") Then %>
			</ul><br /><br />
		<%End If%>
		
		<%Else%>
			<p>You do not have permission to access the <b>Action Line Tab</b>.  Please contact your E-Govlink administrator to inquire about gaining access to the <b>Action Line Tab</b>.</p>
		<%End If%>
	 
	  
	  <!-- END: ACTION LINE REQUEST LIST -->
      </td>
       
    </tr>
  </table>
</body>
</html>


