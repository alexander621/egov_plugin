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
                        <td><img src="../images/square_bullet.gif" align="absmiddle">&nbsp;<a href="chooseversion.asp">Choose Version</a></td>
                      </tr>
                      <tr>
                        <td><img src="../images/square_bullet.gif" align="absmiddle">&nbsp;<a href="billinginfo.asp"><b>Enter Billing Information</b></a></td>
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
              <font size="3" color="#0000bb"><b>Billing Information</b></font><br>
              <br>
              Please enter information on how you would like initial payment and monthly payments (if applicable) will be billed.<br><br>
              <form>
                <table border="0" cellpadding="4" cellspacing="0">
                  <tr><td>First name:</td><td><input type="text" size="50"></td></tr>
                  <tr><td>Last name:</td><td><input type="text" size="50"></td></tr>
                  <tr><td>Address:</td><td><input type="text" size="50"></td></tr>
                  <tr><td>City:</td><td><input type="text" size="50"></td></tr>
                  <tr><td>State:</td><td><input type="text" size="2"></td></tr>
                  <tr><td>Zip:</td><td><input type="text" size="10"></td></tr>
                  <tr><td colspan="2"><hr size="1" color="#0000bb"></td></tr>
                  <tr><td>Credit Card Number:</td><td><input type="text" size="50"></td></tr>
                  <tr><td>Expiration:</td><td><input type="text" size="50"></td></tr>
                  <tr><td>Name On Card:</td><td><input type="text" size="50"></td></tr>
                </table>
              </form>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>