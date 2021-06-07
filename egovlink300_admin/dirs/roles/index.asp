<link href="../../global.css" rel="stylesheet" type="text/css">
<!-- #include file="../../includes/common.asp" //-->
<!--#include file="../dir_constants.asp"-->
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="menu">
    <tr>
      <td background="../images/back_main.jpg">
          <%  DrawTabs tabCommittees,3  %>
      </td>
    </tr>

  </table>
  <table border="0" cellpadding="10" cellspacing="0" width="100%"  class="start" >
    <tr>
      <td valign="top" width='151'>
		 <center> <img src='../../images/icon_directory.jpg'></center>
	
	<TABLE border=0 cellspacing=0 cellpadding=0>
	  <TR><TD valign=top align=center height=20>&nbsp; </TD> </TR>
	  <TR><TD>
<!--#include file="../quicklink.asp"-->
	  </TD>	  </TR>	  </TABLE>

      </td>
      <td colspan="2" valign="top">

<table><tr><td><font size='+1'><b><%=langAdminTitle%></b></font><br><img src='../../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='display_committee.asp'><%=langGoBack%></td></tr></table><br><br>
<!--#include file="forum.asp"-->
<%
thisname=request.servervariables("script_name")
response.write "<a href="&thisname&"?iOfaction=" & ActNewPost &"&groupid="&request.querystring("groupid")&">"&langNewRecord&"</a>"
response.write "&nbsp;&nbsp;<a href="&thisname&"?iOfaction=" & ActDisplayRecords&">"&langRecordList&"</a><br>"

	%>
<%
SHowForum
%>


  </td>
  <td width='200'>&nbsp;</td>
    </tr>
 </table>
 <!--#include file='../footer.asp'-->
