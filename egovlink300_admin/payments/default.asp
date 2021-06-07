<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->

<% blnCanUpdatePayments = HasPermission("CanEditPaymentRequests") %>

<html>
<head>
  <title><%=langBSPayments%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabPayments,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><!--<img src="../images/icon_home.jpg">--></td>
      <td><font size="+1"><b>Payment Services Administration</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../default.asp"><%=langBackToStart%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" nowrap>
		<font size="1" face="Verdana,Tahoma,Arial"><b><%=langDesignby%></b><div class="logo"><A HREF="http://www.eclink.com"><img src="../images/poweredby.jpg" align="center" border="0"></A></div>
        <!-- START: QUICK LINKS MODULE //
        
        <%
        sLinks = "<div style=""padding-bottom:8px;""><b>" & langEventLinks & "</b></div>"

        If bCanEdit Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/calendar.gif"" align=""absmiddle"">&nbsp;<a href=""newevent.asp"">" & langNewEvent & "</a></div>"
          bShown = True
        End If
        
        If bShown Then
          Response.Write sLinks & "<br>"
        End If
        %>

        <% Call DrawQuicklinks("", 1) %>
        <!-- END: QUICK LINKS MODULE //-->

      </td>
        
      <td colspan="2" valign="top">
	  <!--BEGIN: ACTION LINE REQUEST LIST -->
      

	 <% If blnCanUpdatePayments Then %>
		
		<ul>
		  <li><a href="action_line_list.asp">(E-Gov Payment Receipt Manager)</a> - Manage Online Submitted Payments<br><br></li>
		  
		  <li><a href="manage_action_forms.asp">(E-Gov Payment Notification Manager)</a> - Manage Your Online Payment Forms<br><br></li>
		
<%		Dim sGatewayURL, sGatewayName
		sGatewayName = " "
		
		sGatewayURL = GetGatewayURL( session("payment_gateway"), sGatewayName )

		If sGatewayURL <> "" Then  %>
			<li><a href="<%=sGatewayURL%>" target="_MANAGER" >(<%=sGatewayName%> Manager)</a> - Reconcile Your Submitted Payments</li>
<%		End If %>
		  
		  <!--<li><a href="#">Online Payment Reports</a>-->

		 </ul>
	
	<%Else%>

		<p>You do not have permission to access the <b>Payments Tab</b>.  Please contact your E-Govlink administrator to inquire about gaining access to the <b>Payments Tab</b>.</p>

	<%End If%>
	  
	  <!-- END: ACTION LINE REQUEST LIST -->
      </td>
       
    </tr>
  </table>
</body>
</html>


<%


Function GetGatewayURL( iGatewayId, ByRef sGatewayName )
	Dim sSql, oURL

	sSql = "Select adminurl, admingatewayname from egov_payment_gateways where paymentgatewayid = "  & iGatewayId

	Set oURL = Server.CreateObject("ADODB.Recordset")
	oURL.Open sSQL, Application("DSN"), 0, 1

	If Not oURL.EOF Then 
		GetGatewayURL = oURL("adminurl")
		sGatewayName = oURL("admingatewayname")
	Else 
		GetGatewayURL = ""
		sGatewayName = ""
	End If 

	oURL.close
	Set oURL = Nothing

End Function 


%>