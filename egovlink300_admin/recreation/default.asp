<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CLIENT_TEMPLATE_PAGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06   JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

blnPoolUser = HasPermission("PoolUsers") %>



<!-- #include file="../includes/common.asp" //-->


<html>
<head>
	<title>E-Gov Administration Console</title>
	<link rel="stylesheet" type="text/css" href="querytool.css" />
	<link href="../global.css" rel="stylesheet" type="text/css">
</head>


<body>

 
<%DrawTabs tabRecreation,1%>


<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><!--<img src="../images/icon_home.jpg">--></td>
      <td><font><b>Recreation Module Management</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)"><%=langBackToStart%></a></td>
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
          Response.Write sLinks & "<br>"
        End If
        %>

        <% 'Call DrawQuicklinks("", 1) %>
        <!-- END: QUICK LINKS MODULE //-->


		 <font size="1" face="Verdana,Tahoma,Arial"><b><%=langDesignby%></b><div class="logo"><A HREF="http://www.eclink.com"><img src="../images/poweredby.jpg" align="center" border="0"></A></div>

      </td>
        
      <td colspan="2" valign="top">
<!--BEGIN PAGE CONTENT-->


<%If blnPoolUser Then%> 
	<div style="padding:20px;">
	<font style="FONT-SIZE: 14px;"><strong>Pool Passes</strong></font> 
		<p>
			<a href="../poolpass/poolpass_form.asp"><strong>Pool Pass Purchase</strong></a> allows authorized staff to perform pool pass purchases.<br />
			<a href="../poolpass/member_list.asp"><strong>Membership List</strong></a> view pool pass membership<br />
			<a href="../classes/roster_list.asp"><strong>Registration and Rosters</strong></a> allows authorized staff to register citizens for class\events and to view class\event rosters.<br />
		</p>
	</div>
<%Else%>
	<div style="padding:20px;">
	<% If OrgHasFeature( "facilities" ) Then %>
		<font style="FONT-SIZE: 14px;"><b>Facility Management</b></font> 
		<p>
			<a href="../recreation/facility_management.asp"><b>Manage Facilities</b></a> allows authorized staff to add\manage\delete reservable city facilities.<br>
			<a href="facility_calendar.asp"><b>Make/View Facility Reservations</b></a><br>
		<!--<a href="facility_reporting.asp?reportid=1"><b>Facility Totals</b></a><br>-->
		<!--<a href="facility_reporting.asp?reportid=2"><b>Facility Cancellations</b></a>-->
		</p>
		<br />
<% end if %>

<% If OrgHasFeature( "gifts" ) Then %>
		<font style="FONT-SIZE: 14px;"><b>Commemorative Gifts</b></font> 
		<p>
			<a href="../gifts/gift_form.asp"><b>New Commemorative Gift Purchase</b></a> make gift purchase.<br>
		 <!--<a href="../gifts/gift_payment_list.asp"><b>Commemorative Gift Reporting</b></a> view gifts purchased online.-->
		</p>
		<br />
<% end if %>

<% If OrgHasFeature( "memberships" ) Then %>
	<font style="FONT-SIZE: 14px;"><strong>Pool Passes</strong></font> 
    <p>
		<a href="../poolpass/poolpass_form.asp"><strong>Pool Membership Purchase</strong></a> allows authorized staff to perform pool membership purchases.<br />
		<a href="../poolpass/member_list.asp"><strong>Pool Membership</strong></a> view pool membership<br />
	    <a href="../poolpass/poolpass_rates.asp"><strong>Manage Pool Membership Rates</strong></a> allows authorized staff to manage pool pass rates.<br />
		<a href="../poolpass/poolpass_intro.asp"><strong>Manage Pool Membership Introductory Text</strong></a> edit the introductory text that appears on the public site<br />
	    <!--<a href="../poolpass/poolpass_list.asp"><strong>Pool Pass Report</strong></a> view pool pass purchases<br />-->
		<!--<a href="../poolpass/poolpass_type_report.asp"><strong>Pool Pass Counts By Type Report</strong></a> view the count of passes sold by type for a selected year-->
	</p>
	<br />
<% end if %>


<% If OrgHasFeature( "rec payments" ) Then %>
	<font style="FONT-SIZE: 14px;"><b>Recreation Payment Alerts</b></font> 
	<p>
		<a href="manage_recreation_alerts.asp"><b>Manage Recreation Payment Alerts</b></a><br />
		<a href="verisign_password_change.asp"><b>Manage Versign Password</b></a>
	</p>
	<br />
<% end if %>

<% If OrgHasFeature( "address update" ) Then %>
	<font style="FONT-SIZE: 14px;"><b>Citizen Registration</b></font> 
	<p>
		<a href="../manage_address_list.asp"><b>Manage Resident Address List</b></a>
	</p>
<br>
<% end if %>


<% If OrgHasFeature( "activities" ) Then %>
	<font style="FONT-SIZE: 14px;"><b>Events/Classes</b></font> 
	<p>
		<!--<a href="../classes/class_editongoing.asp"><strong>New Ongoing</strong></a> allows authorized staff to create a new ongoing class or event. <br />-->
		<a href="../classes/new_class.asp?classtypeid=3"><strong>Create a New Single Class/Event</strong></a> allows authorized staff to create a new single class or event.<br />
		<a href="../classes/new_class.asp?classtypeid=1"><strong>Create a New Series</strong></a> allows authorized staff to create a new series class or event.<br />
		<a href="../classes/class_list.asp"><strong>Manage Classes\Events</strong></a> allows authorized staff to add\manage\delete classes\events.<br />
		<a href="../classes/roster_list.asp"><strong>Registration and Rosters</strong></a> allows authorized staff to register citizens for class\events and to view class\event rosters.<br />
		<a href="../classes/instructor_mgmt.asp"><strong>Instructors</strong></a> allows authorized staff to add\manage\delete\assign instructors.<br />
		<a href="../classes/location_mgmt.asp"><strong>Locations</strong></a> allows authorized staff to add\manage\delete locations.<br />
		<a href="../classes/category_mgmt.asp"><strong>Categories</strong></a> allows authorized staff to add\manage\delete categories.<br />
		<a href="../classes/discount_mgmt.asp"><strong>Discounts</strong></a> allows authorized staff to add\manage\delete discounts.<br />
		<a href="../classes/poc_mgmt.asp"><strong>Point of Contact</strong></a> allows authorized staff to add\manage\delete\assign a Point of Contact.<br />
		<a href="../classes/dl_mgmt.asp"><strong>Manage Distribution Lists</strong></a> allows authorized staff to add\manage\delete mailing lists.<br />
		<a href="../classes/dl_sendmail.asp"><strong>Send Email Notifications</strong></a> allows authorized staff to send email to users on specified mailing lists.<br />
		<a href="../classes/class_waivers.asp"><strong>Waiver Management</strong></a> allows authorized staff to add\manage\delete waivers.<br />
		<!--<a href="../classes/class_statisticsreport.asp"><strong>Statistics Report</strong></a> View progam statistics.<br />-->
	</p>
	<br />
<% end if %>


<% If OrgHasFeature( "rec reports" ) Then %>
	<font style="FONT-SIZE: 14px;"><b>Recreation Reports</b></font> 
	<p>
		<a href="reports.asp"><strong>Recreation Reports Page</strong></a> view recreation reports.<br />
	</p>
	<br />
<% end if %>


</div>
<%End If%>
<!--END: PAGE CONTENT-->
      </td>
       
    </tr>
  </table>




<!--#Include file="../admin_footer.asp"-->  

</body>


</html>



<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------
%>


