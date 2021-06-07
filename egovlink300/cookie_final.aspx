<%@ Page Language="C#" AutoEventWireup="true" CodeFile="cookie_final.aspx.cs" Inherits="cookie_final" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>


<!DOCTYPE html>
<html lang="en">
<head id="Head1" runat="server">
</head>
<body>
Steps 3 & 6: Read Cookies in ASP.NET
<br />
<%
  displayCookieData();
%>
<br />
<a href="cookie_test.aspx">Step 4: Write Cookie in ASP.NET</a>
<br />
<br />
<a href="cookie_test.asp">Try Again - Step 1: Write Cookie in Classic ASP</a>
</body>
</html>
