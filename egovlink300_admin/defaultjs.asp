<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<% Dim sError

' Set Timezone information into session
Session("iUserOffset") = request.cookies("tz")

%>


<html>
<head>
  <title><%=langBSHome%></title>
  <link href="global.css" rel="stylesheet" type="text/css">
  <script language="Javascript" src="scripts/modules.js"></script>
    <script language="Javascript" > 
  //Set timezone in cookie to retrieve later
  var d=new Date()
  if (d.getTimezoneOffset){
  var iMinutes = d.getTimezoneOffset();
  document.cookie = "tz=" + iMinutes;
  }
  </script>
  <script language="Javascript">
  <!--
    function doPicker(sFormField) {
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("sitelinker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }

  //-->
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >
  <%DrawTabs tabHome,0%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><!--<img src="images/icon_home.jpg">--></td>
      <td colspan="2">
        <% if session("UserID")=0 or session("UserID")="" then	 %>
          <font style="font-size:14px;"><b><%=langWelcomeGuest%></b></font><br>
        <% else %>
          <font style="font-size:14px;"><b><%=langWelcomeBack & " " & Session("FullName") %>!</b></font><br>
        <% end if%>
        <div style="padding-top:2px;">Today is <%=FormatDateTime(Date(), vbLongDate)%></div>
	    </td>
    </tr>
    <tr>
      <td valign="top" width="151">

<!-- the following is the login in form -->
<% if session("UserID")=0 or session("UserID")="" then	
   If sError <> "" Then    Response.Write sError 
%>
<form method="POST" action="login.asp">

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
    <td></td>
    <td><input type=submit value="Login" style="font-family:Tahoma,Arial,Verdana; font-size:10px; width:55px;"></td>
  </tr>
  <tr>
    <td colspan="2"><a href="dirs/lookuppassword.asp"><FONT SIZE="1"><%=langForgotPass%></FONT></A></td>
  </tr>

</table>
    <input type="hidden" name="_task" value="login">
</form>
<% else %>
        <!--
        <div class="quicklink">&nbsp;&nbsp;<img src="images/newmail_small.jpg" align="absmiddle">&nbsp;<a href="messages">1 <%=langNewMessage%></a></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="images/newdisc.gif" align="absmiddle">&nbsp;0 <%=langNewDiscussions%></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="images/newpoll.gif" align="absmiddle">&nbsp;<a href="polls/">2 <%=langNewVotes%></a></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="images/newmeeting.gif" align="absmiddle">&nbsp;<a href="meetings.asp">1 <%=langNewMeeting%></a></div>
        //-->
  
		<% 'Call DrawQuicklinks("",0) %>


		<!--    
		<form action="docs/default.asp" method=post id=form1 name=frmSearch>
          <div style="padding-bottom:3px;"><%=langSearchDocuments%>:</div>
          <input type="text" name="SearchString" style="background-color:#eeeeee; border:1px solid #000000; width:144px;"><br>
          <div class="quicklink" align="right"><a href="#" onClick='document.frmSearch.submit()'><img src="images/go.gif" border="0"><%=langGo%></a>&nbsp;&nbsp;</div>
        </form> -->
<% end if%>
 
 <font size="1" face="Verdana,Tahoma,Arial"><b><%=langDesignby%></b><div class="logo"><A HREF="http://www.eclink.com"><img src="images/poweredby.jpg" align="center" border="0"></A></div>
        <!-- END: QUICK LINKS MODULE //-->

      </td>
      <td valign="top">

	<font style="FONT-SIZE: 14px;"><b>General Tools</b></font><br>
	<p><a href="javascript:doPicker('UpdateEvent.Message')"><b>E-Gov Site Linker</b></a> allows staff to create links to other pages within the E-Gov Site.</p><br>

	<font style="FONT-SIZE: 14px;"><b>Community Calendar (Calendar Tab)</b></font><br>
	<p><a href="events/"><b>E-Gov Calendar Manager</b></a> allows non-technical staff to maintain the calendar. Calendar items can be easily linked to web pages, documents (like meeting agendas), or request forms.</p><br>


	<font style="FONT-SIZE: 14px;"><B>Document Manager (Documents Tab)</b></font>
	<p><a href="docs/"><b>E-Gov Documents Manager</b></a> allows authorized staff to quickly and easily upload documents to your website from any PC-no web skills needed.  Documents and document folders can be accessed from your main website with direct links.
	</p><br>
	
	<font style="FONT-SIZE: 14px;"><B>Security Management (Directories Tab)</b></font>
	<p>
	<a href="dirs/"><b>E-Gov Security Manager</b></a>
	<UL>
	<li>Allows designated staff to control access to administrative functions of the website.
	<li>Access to each area (CRM,  Payments, Documents, and Calendar) can be separately controlled. 
	<li>CRM access can be limited to own requests, own department requests, or all requests.
	<li>Payment access can be limited to a particular payment type or all payments.
	<li>Document access can be limited to particular document folders or all documents
	<li>Only specific individuals will have access to security management.
	</uL>
	</p>
	<br>

	<font style="FONT-SIZE: 14px;"><B>Citizen Request Management - CRM (Action Line Tab)</B></font>
	<p>
	<UL>
	<li><B>E-Gov Forms Creator</B> let's you build a unique request form for each request type, so you always get the specific information you need to take action. No need to decipher phone messages or call back for more information. Forms can include check boxes, drop-down lists and other typical Web data entry formats. 
	<li><B>E-Gov Alert Manager</B> immediately routes the requests to the designated individuals in your organization for follow-up (and reminds them or their supervisors, if not resolved within a pre-specified time).
	<li><B>E-Gov Request Manager</B> tracks all actions taken with time stamped record of who entered update.  Requests can be reassigned (which generates alert email to newly responsible person).  
	<li><B>E-Gov Request Query/Reporting Tool - (Coming Soon!)</B> allows management to get up-to-the-minute results.  Examples:
		<uL>
			<li>Check status of any pending or completed request by requester name.
			<li>Determine age of open and completed requests by department.
			<li>Report on results for all requests for month.
		</ul>
	</ul>
	</p>
	<br>


		<font style="FONT-SIZE: 14px;"><B>Online Payments (Payments Tab)</b></font> 
		<p><a href="payments/"><b>E-Gov Payment Manager</b></a> tracks all payments received with time-stamped record.  Print list of payments received for entry into accounting systems.</p>
 
 <br>
 <% if session("orgId") = 15 then %>		
		<font style="FONT-SIZE: 14px;"><B>FAQ Manager</b></font> 
		<p><a href="faq/list_faq.asp"><b>FAQ Manager</b></a> allows authorized staff to quickly and easily add Frequently Asked Questions to your website from any PC-no web skills needed. FAQ answers can be easily linked to web pages, documents and document folders. </p>
<% end if %>

 <br>
 <% if session("orgId") = 15 then %>		
		<font style="FONT-SIZE: 14px;"><B>Form Letter Manager</b></font> 
		<p><a href="formletters/list_letter.asp"><b>Form Letter Manager</b></a> allows authorized staff to quickly and easily add Form Letters to your website from any PC-no web skills needed. Letters can be easily printed from a browser, copied into other word editing programs, or emailed. </p>
<% end if %>
<br>&nbsp;
 <div class="disclaimer">&nbsp;<%=langCopyright%></div>
      </td>

    </tr>
  </table>
</body>
</html>

<script language=javascript>
	function openWin2(url, name) {
  popupWin = window.open(url, name,
"resizable,width=500,height=450");
}
</script>
