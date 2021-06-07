<%
PageIsRequiredByLogin = True
Dim sTimeout
sTimeout=Request.QueryString()
If sTimeout="timeout" then
  Session.Abandon
Else
  Session.Timeout=5
End If
%>
<!-- #include file="includes/common.asp" //-->

<html>
<head>
<title><%=langSessionTimeout%></title>
<style>
  input{
    width:60px;
    font-size:10pt;
    text-align:center;
    FONT-FAMILY: Verdana,Tahoma;
  }
  body{
    background-color:#93bee1;
  }
  div{
    font-size:10pt;
    text-align:center;
    FONT-FAMILY: Verdana,Tahoma
  }
</style>
<script>
  function clickedYes() {     
    window.opener.setTimeout("onTimeout()", 60000*15);
    window.document.location.assign("<%=RootPath%>resetTimeout.asp");
  }
  
  function clickedNo() {
    window.opener.location.assign("<%=RootPath%>signoff.asp");
    window.close();
  }
  
  function goAway() {
    window.document.location.replace("timeoutPopup.asp?timeout");
  }
</script>
</head>
<%If sTimeout="timeout" then%>
  <body leftmargin="0" topmargin="5" marginheight="0" marginwidth="0">
    <div><%=langBeenLoggedOff%><br><br>
    <input type="button" value="<%=langOK%>" onclick="clickedNo();" id=button3 name=button3></div>
</body>
<%Else%>
<body leftmargin="0" topmargin="5" marginheight="0" marginwidth="0" onload="window.setTimeout('goAway()',60000*5);">
    <div><%=langWillBeLoggedOff%><br><br>
    <input type="button" value="<%=langYes%>" onclick="clickedYes();" id=button1 name=button1>&nbsp;&nbsp;&nbsp;<input type=button value="<%=langNo%>" onclick="clickedNo();" id=button2 name=button2></div>
  </body>
<%End If%>
</html>