<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!--#Include file="include_top_functions.asp"-->
<%
 Dim sError 

'CAPTURE CURRENT PATH
 Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME")
%>
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Montgomery, Ohio - Online Services</title>
<link rel="stylesheet" href="css/montgomery.css" />
</head>
<body>

<script type="text/javascript">
  var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
  document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));

  var pageTracker = _gat._getTracker("UA-367786-2");
  pageTracker._initData();
  pageTracker._trackPageview();
</script>

<div id="wrapper">
<div id="header"><a href="http://www.montgomeryohio.org/design2/default.htm"><img src="custom/images/montgomery/logo_sm.gif" alt="Link to homepage" /></a></div>
	
<div id="nav">
	 <div id="subnav">
	   <p><a href="http://www.montgomeryohio.org/">City Home</a></p>
	   <p><a href="http://www.egovlink.com/montgomery/">E-Gov Home</a></p>
  <%
    sSQL = "SELECT ISNULL(OTF.featurename, f.featurename) AS featurename, "
    sSQL = sSQL & "ISNULL(OTF.featuredescription, f.featuredescription) AS featuredescription, "
    sSQL = sSQL & "ISNULL(OTF.publicurl, f.publicurl) AS publicurl, "
    sSQL = sSQL & "ISNULL(OTF.publicdisplayorder, f.publicdisplayorder) AS publicdisplayorder, "
    sSQL = sSQL & "ISNULL(OTF.publicimageurl, f.publicimageurl) AS publicimageurl, "
    sSQL = sSQL & "ISNULL(OTF.publiccanview, f.haspublicview) AS publiccanviw, "
    sSQL = sSQL & "f.publicdisplayorder "
    sSQL = sSQL & "FROM egov_organizations_to_features OTF, egov_organization_features f "
    sSQL = sSQL & "WHERE OTF.featureid = f.featureid "
    sSQL = sSQL & "AND OTF.orgid = 26 "
    sSQL = sSQL & "AND OTF.publiccanview = 1"
    sSQL = sSQL & "ORDER BY 4 "

   	set oDropDown = Server.CreateObject("ADODB.Recordset")
   	oDropDown.Open sSQL, Application("DSN") , 3, 1

    if not oDropDown.eof then
       while not oDropDown.eof
          response.write "  <p><a href=""" & oDropDown("publicurl") & """>" & replace(oDropDown("featurename")," & ","/") & "</a></p>" & vbcrlf
          oDropDown.movenext
       wend

      'Set to first record so that we can display the features in the content area
       oDropDown.movefirst

    end if
  %>
  </div>
 	<div class="spacer">&nbsp;</div>
</div>

<div id="content">

<h1>Welcome to our Online Services</h1>
  <%
    if not oDropDown.eof then
       while not oDropDown.eof
          response.write "  <h2><a href=""" & oDropDown("publicurl") & """>" & oDropDown("featurename") & "</a> </h2>" & vbcrlf
          response.write "  <p>" & oDropDown("featuredescription") & "</p>" & vbcrlf
          oDropDown.movenext
       wend

      'Set to first record so that we can display the features in the footer
       oDropDown.movefirst

    end if
  %>
		<div class="spacer">&nbsp;</div>
</div>

<!--<div id="footer">
<p>Valid <a href="http://validator.w3.org/check?uri=http://www.realworldstyle.com/2col.html">XHTML</a> and <a href="http://jigsaw.w3.org/css-validator/validator?uri=http://www.realworldstyle.com/2col.css">CSS</a> &#8226; <a href="mailto:mark@realworldstyle.com">mark@realworldstyle.com</a></p>
</div>-->

<div id="footer">
  <p><a href="http://www.montgomeryohio.org/">City Home</a>
   | <a href="http://www.egovlink.com/montgomery/">E-Gov Home</a>
  <%
    if not oDropDown.eof then
       i = 2
       while not oDropDown.eof
          i = i + 1

          if i > 1 AND i < 6 then
             response.write "| <a href=""" & oDropDown("publicurl") & """>" & REPLACE(oDropDown("featurename")," & ","/") & "</a>" & vbcrlf
          elseif i = 6 then
             i = 1
             response.write "<br />" & vbcrlf
             response.write "<a href=""" & oDropDown("publicurl") & """>" & REPLACE(oDropDown("featurename")," & ","/") & "</a>" & vbcrlf
          end if
          oDropDown.movenext
       wend
    end if
  %>
  </p>

  <p><a href="user_login.asp">Login</a>
			| <a href="register.asp">Register</a>
		</p>

		<p>Copyright &copy; 2004-<%=Year(Date())%> electronic commerce link, inc. dba egovlink</p>
		
	</div>

</div>
</body>
</html>
