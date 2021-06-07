    </td>
   <td width=1 ><img src="../img/clearshim.gif" border="0" width=1 ></td>
  </tr>
</table>
<!--END OF PAGE CONTENT-->

<% If blnFooterOn Then %>
<!--BEGIN: FADE LINES-->
<table bgcolor="#d6d3ce" border="0" cellpadding="2" cellspacing="0" width="100%"  >
	  <tr bgcolor="#666666"><td height="1" colspan="2"></td></tr>
	  <tr bgcolor="#ffffff"><td height="1" colspan="2"></td></tr>
</table>
<!--END: FADE LINES-->


<!--BEGIN: BOTTOM MENU AND COPYRIGHT INFORMATION-->
<center>

<div class=footerbox  ><table width=100% cellspacing=0 cellpadding=0><tr><TD valign=top align=center >
<br>
<font class=footermenu >
<a class=afooter href="<%=sHomeWebsiteURL%>"><%=sHomeWebsiteTag%></a> |
<a class=afooter href="<%=sEgovWebsiteURL%>/"><%=sEgovWebsiteTag%></a>

<%If blnOrgAction Then%> |
<a class=afooter href="<%=sEgovWebsiteURL%>/action.asp"><%=sOrgActionName%></a>
<%End If%>

<%If blnOrgCalendar Then%> |
<a class=afooter href="<%=sEgovWebsiteURL%>/events/calendar.asp"><%=sOrgCalendarName%></a>
<%End If%>

<%If blnOrgDocument Then%> |
<a class=afooter href="<%=sEgovWebsiteURL%>/docs/menu/home.asp"><%=sOrgDocumentName%></a>
<%End If%>

<%If blnOrgPayment Then%> |
<a class=afooter href="<%=sEgovWebsiteURL%>/payment.asp"><%=sOrgPaymentName%></a>
<%End If%>


<%If blnOrgFaq Then%> |
<a class=afooter href="<%=sEgovWebsiteURL%>/faq.asp">FAQ<%'sOrgFaqName%></a>
<%End If%>

<%If sOrgRegistration Then%>
	<br><a class=afooter href="<%=sEgovWebsiteURL%>/user_login.asp">Login</a> | <a class=afooter href="<%=sEgovWebsiteURL%>/register.asp">Register</a>
<%End If%>


<br><bR>
<font class=footer>Copyright &copy;2004-2005. <i>electronic commerce</i> link, inc. dba <a href="http://www.egovlink.com" target="_NEW"><font class=footermenu>egovlink</font></a>.</font>


<!--BEGIN: DEMO CHECK TO ADD ADMIN LINK-->
<%If iorgid=5 OR iorgid=13 Then%>&nbsp;&nbsp;&nbsp;<a target="_new" href="<%=sEgovWebsiteURL%>/admin/" class=hidden><font color=white>Administrator</font></a><%End IF%>
<!--END: DEMO CHECK TO ADD ADMIN LINK-->


<br>&nbsp;
</font>
</td>
<!--<td width=50 background="<%=sEgovWebsiteURL%>/img/fade.gif" >&nbsp;</td>
<td align=right bgcolor=#FFFFFF>
<div class="logo"> <font  size="1" face="Verdana,Tahoma,Arial"><b>Powered by </b><br><A HREF="http://www.eclink.com"><img src="<%=sEgovWebsiteURL%>/img/poweredby.jpg" vspace=5 align="center" border="0"></A></div></td>-->
</tr></table>
</div>
<!--END: BOTTOM MENU AND COPYRIGHT INFORMATION-->
<%End If%>




</body>
</html>