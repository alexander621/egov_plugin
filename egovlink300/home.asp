<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<% Dim sError %>

<html>
<head>
  <title><%=langBSHome%></title>
  <link href="global.css" rel="stylesheet" type="text/css">
  <script language="Javascript" src="scripts/modules.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >
  <%DrawTabs tabHome,0%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="images/icon_home.jpg"></td>
      <td colspan="2">
        <% if session("UserID")=0 or session("UserID")="" then	 %>
          <font size="+1"><b><%=langWelcomeGuest%></b></font><br>
        <% else %>
          <font size="+1"><b><%=langWelcomeBack & " " & Session("FullName") %>!</b></font><br>
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
        <% Call DrawQuicklinks("",0) %>

		    <form action="docs/default.asp" method=post id=form1 name=frmSearch>
          <div style="padding-bottom:3px;"><%=langSearchDocuments%>:</div>
          <input type="text" name="SearchString" style="background-color:#eeeeee; border:1px solid #000000; width:144px;"><br>
          <div class="quicklink" align="right"><a href="#" onClick='document.frmSearch.submit()'><img src="images/go.gif" border="0"><%=langGo%></a>&nbsp;&nbsp;</div>
        </form>
<% end if%>
 <font size="1" face="Verdana,Tahoma,Arial"><b><%=langDesignby%></b><div class="logo"><A HREF="http://www.eclink.com"><img src="images/poweredby.jpg" align="center" border="0"></A></div>
        <!-- END: QUICK LINKS MODULE //-->

      </td>
      <td valign="top">

        <%= ShowAnnouncements() %>
        <br>
        <br>

        <!-- STOCK TICKER //-->
        <% If Session("UserID") > 0 And Session("ShowStockTicker") = 1 Then %>
        <table border="0" cellpadding="0" cellspacing="0" class="messagehead" width="100%">
          <tr>
            <th align="left">&nbsp;&nbsp;<%=langStockInformation%></th>
            <th align="right">
              <a href="admin/ToggleStockTicker.asp?redirect=../" onclick="if (!confirm('<%=langStockTickerOffNotice%>')) {return false;}">
              <img src="images/main_delete.jpg" border="0">&nbsp;</a>
            </th>
          </tr>
          <tr>
            <td colspan="2">
              <OBJECT id="IEXR2_WPQ_" type="application/x-oleobject" classid="clsid:52ADE293-85E8-11D2-BB22-00104B0EA281" Codebase="http://fdl.msn.com/public/investor/v7/ticker.cab#version=8,2000,0326,2" width="100%" height="34"> 
                <param name="ServerRoot" value="http://moneycentral.msn.com">
                <param NAME="AppRoot" value="">
                <param name="StockTarget" value="_stocks">
                <param name="NewsTarget" value="_news">
              </OBJECT>
            </td>
          </tr>
        </table>
        <br>
        <br>
        <% End If %>

        <%= ShowEvents() %>

        <div class="disclaimer">&nbsp;<%=langCopyright%></div>
      </td>
      <td width="200" valign="top"> 
	  <%  = ShowFavorites() %>
<% if session("UserID")=0 or session("UserID")="" then	
else
%>
	   
		<br>
		<br>
		<%= ShowPersonalFavorites() %>
<% end if %>
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