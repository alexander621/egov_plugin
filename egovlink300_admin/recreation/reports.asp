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
      <td><font><b>E-GOV Reports</b></font><br>
	  <img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)"><%=langBackToStart%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" nowrap="nowrap">

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


		 <font size="1" face="Verdana,Tahoma,Arial"><b><%=langDesignby%></b><div class="logo"><a href="http://www.eclink.com"><img src="../images/poweredby.jpg" align="center" border="0"></a></div>

      </td>
        
      <td colspan="2" valign="top">
<!--BEGIN PAGE CONTENT-->


<div style="padding:20px;">
	<% If OrgHasFeature( "rec reports" ) Then %>
	<p>
		<font style="FONT-SIZE: 14px;"><b>Recreation Reports</b></font> 
	</p>
	<p>
	<a href="global_report.asp"><b>Recreation Financial Totals</b></a> view financial report for all recreation programs.<br />
	<a href="../purchases_report/purchases_list.asp"><b>Citizen Purchases</b></a> view recreation purchases for selected citizens.<br />
	
	<% If OrgHasFeature( "facilities" ) Then %>
		<a href="facility_reporting.asp?reportid=1"><b>Facility Totals</b></a> view monthly facility totals.<br />
		<a href="facility_reporting.asp?reportid=2"><b>Facility Cancellations</b></a><br />
		<a href="rpt_cleaning_crew.asp"><strong>Cleaning Crew Facility Report</strong></a> List of facility reservations with arrive/depart times.<br />
	<% End If %>
	
	<% If OrgHasFeature( "gifts" ) Then %>
		<a href="../gifts/gift_payment_list.asp"><b>Commemorative Gift Reporting</b></a> view gifts purchased online.<br />
	<% End If %>

	<% If OrgHasFeature( "memberships" ) Then %>
		<a href="../poolpass/poolpass_list.asp"><strong>Membership Report</strong></a> view Membership purchases<br />
		<a href="../poolpass/poolpass_type_report.asp"><strong>Membership Counts By Type Report</strong></a> view the count of Memberships sold by type for a selected year.<br />
	<% End If %>

	<% If OrgHasFeature( "activities" ) Then %>
		<a href="../classes/class_statisticsreport.asp"><strong>Classes and Events Statistics Report</strong></a> View progam statistics.<br />
	<% End If %>
	
	</p>
	<br />
<% End If %>
</div>

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


