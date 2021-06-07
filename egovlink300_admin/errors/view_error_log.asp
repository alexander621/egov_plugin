<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title><%=Application("AppName") & " Version " & Application("AppVersion") & " - (VIEW ERROR LOG)"%></title>
	<META NAME="Generator" CONTENT="EditPlus">
	<META NAME="Author" CONTENT="ECLINK">
	<META NAME="Keywords" CONTENT="">
	<META NAME="Description" CONTENT="">
	<link rel="stylesheet" href="error_log.css" type="text/css" />
</head>
<body topmargin="0" leftmargin="0">
<%
 'BEGIN: Header ---------------------------------------------------------------
  response.write "<div class=""headerbox"">" & vbcrlf
  response.write "  <font class=""header"">" & Application("AppName") & " Version " & Application("AppVersion") & " Admin Console</font>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "<table bgcolor=""#225e82"" border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "	  <tr bgcolor=""#ffffff""><td height=""1"" colspan=""2""></td></tr>" & vbcrlf
  response.write "	  <tr bgcolor=""#666666""><td height=""1"" colspan=""2""></td></tr>" & vbcrlf
  response.write "</table>" & vbcrlf
 'END: Header -----------------------------------------------------------------

 'BEGIN: Content --------------------------------------------------------------
  response.write "<table cellpadding=""5"" cellspacing=""0"" class=""layout"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td class=""menu"" valign=""top"">" & vbcrlf
  response.write "          <br />" & vbcrlf
%>
 	<!--#include file="includes/menu.asp"-->
<%
  response.write "      </td>" & vbcrlf
  response.write "     	<td  valign=""top"">" & vbcrlf

		if request("errorid") = "" then
     if request("iwebappid") = "" then
  					'Display Summary by Application
        response.write "<font style=""font-size:18px;"">Error Log - Summary</font><br />" & vbcrlf
        response.write "<hr size=""1"" style=""width:500px; text-align:left; color:#000000;"">" & vbcrlf
        response.write "<div style=""width:600px;"">" & vbcrlf
        response.write "  <p>Click application name to review complete error list or click datetime to review the last error reported for the application.</p>" & vbcrlf
        response.write "		<p>" & vbcrlf
                        					ErrorLogSummary
        response.write "  </p>" & vbcrlf
        response.write "</div>" & vbcrlf
     else
       'Display List by Application
        response.write "<font style=""font-size:18px;"">Error Log - Web Application</font><br />" & vbcrlf
        response.write "<hr size=""1"" style=""width:500px; text-align:left; color:#000000;"">" & vbcrlf
        response.write "<div style=""width:600px;"">" & vbcrlf
        response.write "		<p>Click error number to review error details.</p>" & vbcrlf
        response.write "		<p>" & vbcrlf
                             WebAppLog request("iwebappid")
        response.write "		</p>" & vbcrlf
        response.write "</div>" & vbcrlf
     end if
  else
    'Display Individual Error
     response.write "<font style=""font-size:18px;"">Error Details - (" & request("errorid") & ")</font><br />" & vbcrlf
     response.write "<hr size=""1"" style=""width:500px; text-align:left; color:#000000;"">" & vbcrlf
     response.write "<div style=""width:600px;"">" & vbcrlf
     response.write "  <p>Review the error details and conditions below.</p>" & vbcrlf
     response.write "		<p>" & vbcrlf
                            ViewError
     response.write "		</p>" & vbcrlf
     response.write "</div>" & vbcrlf
  end if

  response.write "     	</td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
 'END: Content ----------------------------------------------------------------

 'BEGIN: Footer ---------------------------------------------------------------
  response.write "<table bgcolor=""#225e82"" border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "	  <tr bgcolor=""#666666""><td height=""1"" colspan=""2""></td></tr>" & vbcrlf
  response.write "	  <tr bgcolor=""#ffffff""><td height=""1"" colspan=""2""></td></tr>" & vbcrlf
  response.write "</table>" & vbcrlf

 'Copyright
  response.write "<div class=""footerbox"">" & vbcrlf
  response.write "  <font style=""color:#ffffff"">Copyright &copy;1996-" & year(now) & ". <em>Electronic Commerce</em> Link, Inc.  All Rights Reserved.<br />" & vbcrlf
  response.write "  <img class=""logo"" src=""images/eclink_logo.jpg"" />" & vbcrlf
  response.write "</div>" & vbcrlf
 'END: Footer -----------------------------------------------------------------

  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub ErrorLogSummary()
	
	sSQL = "SELECT * FROM vw_web_app_summary order by webappname"

	set oSummary = Server.CreateObject("ADODB.Recordset")
	oSummary.Open sSQL, Application("ErrorDSN") , 3, 1
	
	if not oSummary.eof then
    response.write "<table cellspacing=""0"">" & vbcrlf
    response.write "  <tr>" & vbcrlf
    response.write "      <td class=""excelheaderleft"">ID</td>" & vbcrlf
    response.write "      <td class=""excelheader"">Name</td>" & vbcrlf
    response.write "      <td class=""excelheader""># Errors reported</td>" & vbcrlf
    response.write "      <td class=""excelheader"">Date of Last Error</td>" & vbcrlf
    response.write "  </tr>" & vbcrlf

  		do while not oSummary.eof
       response.write "  <tr>" & vbcrlf
       response.write "      <td class=""exceldataleft"">" & oSummary("WebAppId") & "</td>" & vbcrlf
       response.write "      <td class=""exceldata""><a href=""view_error_log.asp?iwebappid=" & oSummary("WebAppId") & """>" & oSummary("WebAppName") & "</a></td>" & vbcrlf
       response.write "      <td class=""exceldata"" align=""center"">" & oSummary("totalerrors") & "</td>" & vbcrlf
       response.write "      <td class=""exceldata"">" & fnGetDateofLastError(oSummary("WebAppId")) & "</td>" & vbcrlf
       response.write "  </tr>" & vbcrlf

    			oSummary.movenext
    loop

  		response.write "</table>" & vbcrlf

 else
  		response.write "<p style=""padding-top:10px;""><center><font color=""#ff0000""><strong><em>No Web applications have been created!</em></strong></font></p>" & vbcrlf
 end if

	set oSummary = nothing 

end sub

'------------------------------------------------------------------------------
sub ViewError()

 'Get Error Information
 	sSQL = "SELECT * "
  sSQL = sSQL & " FROM errorlog "
  sSQL = sSQL & "      LEFT OUTER JOIN webapplications ON errorlog.webappid = webapplications.webappid "
  sSQL = sSQL & " WHERE rowid = " & request("errorid")

 	set oError = Server.CreateObject("ADODB.Recordset")
	 oError.Open sSQL, Application("ErrorDSN") , 3, 1
	
  if not oError.eof then
  		'BEGIN: Display Error Information -----------------------------------------

  		'Web Application Information / Datetime of Error
   		response.write "<strong><em>ec</em> link - ASP Script Error Log </strong><br />" & vbcrlf
     response.write "<div align=""left"">" & vbcrlf
     response.write "<font style=""font-family: arial,tahoma; font-size: 12px;"">" & vbcrlf
     response.write "<div class=""errorbox"">" & vbcrlf
     response.write "  <strong>Date Error Reported:</strong> "         & oError("errordatetime")& "<br />" & vbcrlf
     response.write "  <strong>Web Application: </strong>"             & oError("webappname")   & "<br />" & vbcrlf
     response.write "  <strong>Web Application Description: </strong>" & oError("webappdescription") & vbcrlf
     response.write "</div>" & vbcrlf
		
    'IIS ASP Error Object Information
   		response.write "<br />" & vbcrlf
     response.write "<strong>IIS ASP Error Object Information:</strong>" & vbcrlf
     response.write "<div class=""errorbox"">"  & vbcrlf
   		response.write "  <strong>File: </strong> "        & oError("file")        & "<br />" & vbcrlf
   		response.write "  <strong>Line: </strong> "        & oError("line")        & "<br />" & vbcrlf
  	 	response.write "  <strong>Number: </strong> "      & oError("number")      & "<br />" & vbcrlf
   		response.write "  <strong>Description: </strong> " & oError("description") & "<br />" & vbcrlf
   		response.write "  <strong>Category: </strong> "    & oError("Category")    & "<br />" & vbcrlf
   		response.write "  <strong>Column: </strong> "      & oError("Column")      & "<br />" & vbcrlf
   		'response.write "<strong>Source: </strong> " & oError("source") & "<br />" & vbcrlf
		   response.write "</div>" & vbcrlf
		
  		'Browser Information
   		response.write "<br /><strong>Client User Browser Information</strong><div class=""errorbox"">" & oError("browserinformation") & "</div>" & vbcrlf

  		'Application Collection
   		response.write "<br /><strong>ASP Application Collection</strong><div class=""errorbox"">" & oError("applicationcollection")& "</div>" & vbcrlf
		
  		'Request Form Collection
		   response.write "<br /><strong>Request Form Collection</strong><div class=""errorbox"">" & oError("requestformcollection") & "</div>" & vbcrlf
		
  		'Querystring Collection
		   response.write "<br /><strong>Querystring Collection</strong><div class=""errorbox"">" & oError("requestquerystringcollection") & "</div>" & vbcrlf

		  'Session Collection
		   response.write "<br /><strong>Session Collection</strong><div class=""errorbox"">" & oError("sessioncollection") & "</div>" & vbcrlf

		  'Cookies Collection
		   response.write "<br /><strong>Cookies Collection</strong><div class=""errorbox"">" & oError("cookiescollection") & "</div>" & vbcrlf

		  'Server Variables Collection
		   response.write "<br /><strong>Server Variable Collection</strong><div class=""errorbox"">" & oError("servervariablescollection") & "</div>" & vbcrlf

  else
     response.write "<div align=""left"">" & vbcrlf
     response.write "<font style=""font-family: arial,tahoma; font-size: 12px;"">" & vbcrlf
     response.write "<div style=""text-align:left; width:90%; border:1px solid #000000; font-family:arial,tahoma; font-size:12px; color:yellow; padding:5px; background-color:#ff0000;"">" & vbcrlf
     response.write "  -- NO ERROR FOUND MATCHING THAT ID --" & vbcrlf
     response.write "</div>" & vbcrlf
  end if

	 set oError = nothing

end sub

'------------------------------------------------------------------------------
function fnGetDateofLastError(iWebAppID)

	sReturnValue = "N/A"

	sSQL = "SELECT TOP 1 * FROM errorlog WHERE webappid='" & iWebAppID & "' ORDER BY errordatetime DESC"

	set oError = Server.CreateObject("ADODB.Recordset")
	oError.Open sSQL, Application("ErrorDSN") , 3, 1
	
	if not oError.eof then
  		sReturnValue = "<a title=""Click to view error"" alt=""Click to view error"" href=""view_error_log.asp?errorid=" & oError("rowid")& """>" & oError("errordatetime") & "</a>" & vbcrlf
 end if

	set oError = nothing

	fnGetDateofLastError = sReturnValue

end function

'------------------------------------------------------------------------------
sub WebAppLog( iwebapp )
	
	sSQL = "SELECT * "
 sSQL = sSQL & " FROM errorlog "
 sSQL = sSQL & "      LEFT OUTER JOIN webapplications ON errorlog.webappid = webapplications.webappid "
 sSQL = sSQL & " WHERE errorlog.webappid = '" & iwebapp & "' "
 sSQL = sSQL & " ORDER BY errordatetime desc"

	set oLog = Server.CreateObject("ADODB.Recordset")
	oLog.PageSize       = 25
	oLog.CacheSize      = 25
	oLog.CursorLocation = 3
	oLog.Open sSQL, Application("ErrorDSN") , 3, 1

 if not oLog.eof then
		 'Set Page to View
 		 if Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1 then
	    		oLog.AbsolutePage = 1
  	 else
		    	if clng(Request("pagenum")) <= oLog.PageCount then
       			oLog.AbsolutePage = Request("pagenum")
    			else
      				oLog.AbsolutePage = 1
    			end if
		  end if

 		'Display Recordset Statistics
  		abspage = oLog.AbsolutePage
  		pagecnt = oLog.PageCount
			
  		response.write "<table>" & vbcrlf
    response.write "  <tr>" & vbcrlf
    response.write "      <td valign=""top"">" & vbcrlf
    response.write "          <a href=""view_error_log.asp?pagenum=1&iwebappid=" & request("iwebappid") & """><img border=""0"" src=""images/nav_first.gif""></a>" & vbcrlf
    response.write "          <a href=""view_error_log.asp?pagenum="&abspage-1 &"&iwebappid=" &  request("iwebappid") & """><img border=""0"" src=""images/nav_back.gif""></a>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "          <strong>" & vbcrlf
    response.write "            Number of pages: <font style=""color:blue;""> " & oLog.PageCount & "</font> | " & vbcrlf
  		response.write "            Current page: <font style=""color:blue;"">" & abspage & "</font> | " & vbcrlf
		  response.write "            Number of Errors: <font style=""color:blue;"">" & oLog.RecordCount
  		response.write "          </strong>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "          <a href=""view_error_log.asp?pagenum="&abspage+1 &"&iwebappid=" &  request("iwebappid") & """><img border=""0"" src=""images/nav_forward.gif"" valign=""bottom""></a>" & vbcrlf
    response.write "          <a href=""view_error_log.asp?pagenum="&oLog.PageCount&"&iwebappid=" &  request("iwebappid") & """><img border=""0"" src=""images/nav_last.gif"" valign=""bottom""></a>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
    response.write "</table>" & vbcrlf
  		response.write "<table cellspacing=""0"" class=""excel"">" & vbcrlf
    response.write "  <tr>" & vbcrlf
    response.write "      <td class=""excelheaderleft"">Number</td>" & vbcrlf
    response.write "      <td class=""excelheader"">Description</td>" & vbcrlf
    response.write "      <td class=""excelheader"">File</td>" & vbcrlf
    response.write "      <td class=""excelheader"">Datetime</td>" & vbcrlf
    response.write "  </tr>" & vbcrlf

  		for intRec=1 To oLog.PageSize
       if not oLOG.eof then
      				response.write "  <tr>" & vbcrlf
          response.write "      <td class=""exceldataleft"">" & vbcrlf
          response.write "          <a title=""Click to view error"" alt=""Click to view error"" href=""view_error_log.asp?errorid=" & oLog("rowid") & """>" & oLog("number") & "</a>" & vbcrlf
          response.write "      </td>" & vbcrlf
          response.write "      <td class=""exceldata"" align=""center"">" & oLog("description") & "&nbsp;</td>" & vbcrlf
          response.write "      <td class=""exceldata"">" & oLog("file") & "</td>" & vbcrlf
          response.write "      <td class=""exceldata"">" & oLog("errordatetime") & "</td>" & vbcrlf
          response.write "  </tr>" & vbcrlf

      				oLog.MoveNext
       end if
    next

  		response.write "</table>" & vbcrlf
    response.write "<p>&nbsp;</p>" & vbcrlf

 else
  		response.write "<p><font style=""color:red;""><strong>No errors logged!</strong></font></p>" & vbcrlf
	end if

	set oLog = nothing 

end sub
%>
