<!-- #include file="../includes/common.asp" //-->
<html>
<head>
  <title><%=langBSSignup%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
      <td><%DrawTabs 0,1%></td>
    </tr>
    <tr>
      <td height="100%">
        <table border="0" cellpadding="10" cellspacing="0" width="100%" height="100%">
          <tr>
            <td width="180" align="center" style="background-color:#93bee1; border-right:1px solid #336699; padding:10px;">
              <table border="0" cellpadding="4" cellspacing="0" class="signupmenu" width="100%" height="100%">
                <tr>
                  <th align="left">Signup</th>
                </tr>
                <tr>
                  <td valign="top" height="100%" style="padding:5px;">
                    <table border="0" cellpadding="2" cellspacing="0">
                      <tr>
                        <td><img src="../images/square_bullet.gif" align="absmiddle">&nbsp;<a href="../signup">Welcome</a></td>
                      </tr>
                      <tr>
                        <td><img src="../images/square_bullet.gif" align="absmiddle">&nbsp;<a href="chooseversion.asp"><b>Choose Version</b></a></td>
                      </tr>
                      <tr>
                        <td><img src="../images/square_bullet.gif" align="absmiddle">&nbsp;<a href="billinginfo.asp">Enter Billing Information</a></td>
                      </tr>
                      <tr>
                        <td><img src="../images/square_bullet.gif" align="absmiddle">&nbsp;<a href="summary.asp">Review & Purchase</a></td>
                      </tr>
                      <tr>
                        <td><img src="../images/square_bullet.gif" align="absmiddle">&nbsp;<a href="customize.asp">Customize</a></td>
                      </tr>
                      <tr>
                        <td><br><b>Help</b></td>
                      </tr>
                      <tr>
                        <td><img src="../images/square_bullet.gif" align="absmiddle">&nbsp;<a href="about.asp">About <%=Application("ProgramName")%></a></td>
                      </tr>
                      <tr>
                        <td><img src="../images/square_bullet.gif" align="absmiddle">&nbsp;<a href="http://home.eclink.com/ecteamlink" target="_demo">Demo Site</a></td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
            <td valign="top" style="font-family:Tahoma,Arial,Verdana;">
              <font size="3" color="#0000bb"><b>Choose your version of <%=Application("ProgramName")%></b></font><br>
              <br><br>
              <b>Titanium</b><br>
              $40,000 - Unlimited User<br>
              <br>
              If you have reached this page, you probably already know a lot about <%=Application("ProgramName")%> and how it can 
              help your company.  If you are not yet familiar with this great product, please click either of the Help links from the menu at the left.<br>
              <br><a href="#">Add To Cart</a>
              <br><hr size="1" color="#0000bb"><br>
              <b>Platinum</b><br>
              $25,000 - 5000 User<br>
              <br>
              If you have reached this page, you probably already know a lot about <%=Application("ProgramName")%> and how it can 
              help your company.  If you are not yet familiar with this great product, please click either of the Help links from the menu at the left.<br>
              <br><a href="#">Add To Cart</a>
              <br><hr size="1" color="#0000bb"><br>
              <b>Gold</b><br>
              $15,000 - 2000 User<br>
              <br>
              If you have reached this page, you probably already know a lot about <%=Application("ProgramName")%> and how it can 
              help your company.  If you are not yet familiar with this great product, please click either of the Help links from the menu at the left.<br>
              <br><a href="#">Add To Cart</a>
              <br><hr size="1" color="#0000bb"><br>
              <b>Standard</b><br>
              $5,000 - 500 User<br>
              <br>
              If you have reached this page, you probably already know a lot about <%=Application("ProgramName")%> and how it can 
              help your company.  If you are not yet familiar with this great product, please click either of the Help links from the menu at the left.<br>
              <br><a href="#">Add To Cart</a>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>