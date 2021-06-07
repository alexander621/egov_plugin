<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!--#Include file="include_top_functions.asp"-->

<% Dim sError 


' CAPTURE CURRENT PATH
Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME")

%>

<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Park City - Online Services</title>
<link rel="stylesheet" href="css/parkcity.css" />
</head>
<body>
<div id="wrapper">
<div id="header">
	<!--<a href="http://www.montgomeryohio.org/design2/default.htm"><img src="custom/images/montgomery/logo_sm.gif" alt="Link to homepage" /></a>-->
	<h1>Park City</h1>
</div>
	
<div id="nav">
			
				<div id="subnav">
						<p><a href="http://www.egovlink.com/web">City Home</a></p>
		        <p><a href="http://www.egovlink.com/parkcity/">E-Gov Home</a></p>
	          <p><a href="http://www.egovlink.com/parkcity/action.asp">Action Line</a></p>
		        <p><a href="http://www.egovlink.com/parkcity/events/calendar.asp">Community Calendar</a></p>
	          <p><a href="http://www.egovlink.com/parkcity/docs/menu/home.asp">Documents</a></p>
		        <p><a href="http://www.egovlink.com/parkcity/recreation/facility_list.asp">Facility Reservations</a></p>
		        <p><a href="http://www.egovlink.com/parkcity/pool_pass/poolpass_form.asp">Pool Passes</a></p>
				<p><a href="http://www.egovlink.com/parkcity/classes/class_list.asp">Event/Class Registration</a></p>
	          <p><a href="http://www.egovlink.com/parkcity/gifts/gift_list.asp">Commemorative Gifts </a></p>
			</div>
		
		<div class="spacer">&nbsp;</div>
</div>

<div id="content">


<h1>Welcome to our Online Services</h1>
			<h2><a href="http://www.egovlink.com/parkcity/action.asp">Action Line</a> </h2>
			<p>Make suggestions, request information or request service at your convenience.</p>
			<h2><a href="http://www.egovlink.com/parkcity/events/calendar.asp">Community Calendar</a> </h2>
			<p>See what's going on, including schedules for all government meetings and City events. Use the handy search feature to find the specific events you are looking for.</p>	
			<h2><a href="http://www.egovlink.com/parkcity/docs/menu/home.asp">Documents</a> </h2>
			<p>Quickly find the documents you are looking for. All City documents are located in a single place with a handy search feature.</p>	
			<h2><a href="http://www.egovlink.com/parkcity/recreation/facility_list.asp">Facility Reservations</a> </h2>
			<p>Reserve one of the beautiful lodges for your next function.</p>	
			<h2><a href="http://www.egovlink.com/parkcity/pool_pass/poolpass_form.asp">Pool Passes</a> </h2>
			<p>Purchase your season pool pass online.</p>
			<h2><a href="http://www.egovlink.com/parkcity/classes/class_list.asp">Event &amp; Class Registrations</a> </h2>
			<p>Register for recreation classes and events. </p>
			<h2><a href="http://www.egovlink.com/parkcity/gifts/gift_list.asp">Commemorative Gifts</a> </h2>
			<p>Purchase a commemorative gift such as a tree, park bench or a brick paver to show your support or to honor a loved one.</p>

		<div class="spacer">&nbsp;</div>
</div>


<!--<div id="footer">
<p>Valid <a href="http://validator.w3.org/check?uri=http://www.realworldstyle.com/2col.html">XHTML</a> and <a href="http://jigsaw.w3.org/css-validator/validator?uri=http://www.realworldstyle.com/2col.css">CSS</a> &#8226; <a href="mailto:mark@realworldstyle.com">mark@realworldstyle.com</a></p>
</div>-->

<div id="footer">
		
			<p><a href="http://www.egovlink.com/web">City Home</a>
				| <a href="http://www.egovlink.com/parkcity/">E-Gov Home</a>
					| <a href="http://www.egovlink.com/parkcity/action.asp">Action Line</a>
						| <a href="http://www.egovlink.com/parkcity/events/calendar.asp">Community Calendar</a>
							| <a href="http://www.egovlink.com/parkcity/docs/menu/home.asp">Online Documents</a> <br />
							 <a href="http://www.egovlink.com/parkcity/recreation/facility_list.asp">Facility Reservations</a>
							| <a href="http://www.egovlink.com/parkcity/pool_pass/poolpass_form.asp">Pool Passes</a>
							| <a href="http://www.egovlink.com/parkcity/classes/class_list.asp">Event/Class Registration</a>
						| <a href="http://www.egovlink.com/parkcity/gifts/gift_list.asp">Commemorative Gifts</a>
			</p>

			<p>
			 <a href="user_login.asp">Login</a>
			| <a href="register.asp">Register</a>
	
	</p>
		
		<p>Copyright &copy; 2004-<script type="text/javascript"> 
		<!--
			var theDate=new Date();
			document.write(theDate.getFullYear());
		//-->
		</script> electronic commerce link, inc. dba egovlink</p>
		
	</div>

</div>
</body>
</html>
