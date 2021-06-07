<!-- #include file="loadfolder.inc" //-->
<html>
<head>
  <script language="JScript" src="menu.js" type="text/javascript"></SCRIPT>
  <link href="menu_ie4.css" type="text/css" rel="stylesheet">
  <script language="Javascript" src="../scripts/vbstring.js"></script>
  <script language="Javascript1.2">
  <!--
    menuOn = true;
    ie50 = false;

    function toggleMenu() {
      if (menuOn) {
        parent.parent.parent.fs_top.rows="0,*";
        parent.parent.fs_bottom.cols="0,*";
        parent.ecapture.cols="25,*";
        document.all.menu.style.display='none';
        document.all.menurestore.style.display='';
      }
      else {
        parent.parent.parent.fs_top.rows="52,*";
        parent.parent.fs_bottom.cols="169,*";
        parent.ecapture.cols="220,*";
        document.all.menu.style.display='';
        document.all.menurestore.style.display='none';
      }

      menuOn = !menuOn;
    }

    function writeStyleSheet() {
      document.write('<style type="text/css"><!--');

      if (contains(navigator.appVersion, "MSIE 5.0")) {
        ie50 = true;
        document.write('.button  {border:#003366 solid 1px; padding:1px; cursor:hand; height:1px;}');
        document.write('.buttona {border:#ffffff solid 1px; padding:1px; cursor:hand; height:1px; background-color:#336699;}');
      }
      else {
        document.write('.button  {border:#003366 solid 1px; padding:1px; cursor:hand;}');
        document.write('.buttona {border:#ffffff solid 1px; padding:1px; cursor:hand; background-color:#336699;}');
      }

      document.write('//--></style>');
    } writeStyleSheet();
  //-->
  </script>
</head>

<body text="#ffffff" bgcolor="#6699cc" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <div id="menu">
    <img src="images/spacer.gif" width="1" height="3" border="0"><br>
    <span class="button" onclick="toggleMenu()" onmouseover="this.className='buttona';status='Collapse Menu';" onmouseout="this.className='button';status='';"><img src="images/collapse.gif" width="18" height="18" border="0" alt="Collapse Menu"></span>
    <span class="button" onclick="document.location.reload()" onmouseover="this.className='buttona';status='Refresh Menu';" onmouseout="this.className='button';status='';"><img src="images/refresh.gif" width="18" height="18" border="0" alt="Refresh Menu"></span>
    <img src="images/v_divider.gif" height="18" width="2" border="0">
    <span class="button" onclick="parent.fraTopic.location.href='../load.asp?file=main.asp'" onmouseover="this.className='buttona';status='Home';" onmouseout="this.className='button';status='';"><img src="images/home.gif" width="18" height="18" border="0" alt="Home"></span>
    <img src="images/v_divider.gif" height="18" width="2" border="0">
    <span class="button" onclick="parent.fraTopic.location.href='../load.asp?file=addhelp.asp'" onmouseover="this.className='buttona';status='Add Help';" onmouseout="this.className='button';status='';"><img src="images/addHelp.gif" width="18" height="18" border="0" alt="Add Help"></span>
    <hr style="height:1px; color:#003366;">

    <nobr>
    <font color="#000033">&nbsp;<i>Online Help</i></font>
    <ul id=ulRoot style="display: block">
      <%= LoadFolder( Application("ECapture_ArticlesPath") & "/z.Help" ) %>
    </ul>
    </nobr>

    <iframe name="hiddenframe" src="" width=1 height=1></iframe>
  </div>
  <span id="menurestore" style="display:none;" align="center">
    <img src="images/spacer.gif" width="1" height="3" border="0"><br>
    <span class="button" onclick="toggleMenu()" onmouseover="this.className='buttona';status='Restore Menu';" onmouseout="this.className='button';status='';"><img src="images/restore.gif" width="18" height="18" border="0" alt="Restore Menu"></span><br>
    <span style="padding:2px;"><img src="images/h_divider.gif" height="2" width="18" border="0"></span><br>
    <img src="images/spacer.gif" height="5" width="1" border="0"><br>
    <span class="button" onclick="parent.fraTopic.location.href='../load.asp?file=main.asp'" onmouseover="this.className='buttona';status='Home';" onmouseout="this.className='button';status='';"><img src="images/home.gif" width="18" height="18" border="0" alt="Home"></span><br>
    <span style="padding:2px;"><img src="images/h_divider.gif" height="2" width="18" border="0"></span><br>
    <script language="Javascript">if (!ie50) { document.write("<br>"); }</script>
    <span class="button" onclick="parent.fraTopic.location.href='../load.asp?file=addhelp.asp'" onmouseover="this.className='buttona';status='Add Help';" onmouseout="this.className='button';status='';"><img src="images/addHelp.gif" width="18" height="18" border="0" alt="Add Help"></span>
  </span>
</body>
</html>