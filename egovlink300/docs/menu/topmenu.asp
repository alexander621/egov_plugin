<!-- #include file="../../includes/common.asp" //-->

<html>
<head>
  <title>top menu</title>
  <link href="../../global.css" rel="stylesheet" type="text/css">
  <script language="Javascript">
  <!--
    menuOn = true;
    ie50 = false;

    function toggleMenu() 
	{
		if (menuOn) 
		{
			if (parent.parent.parent.fs_top != null) 
			{
				parent.parent.parent.fs_top.rows="0,*";
				parent.parent.fs_bottom.cols="0,*";
			}
			parent.fstCols.cols="0,*";
			img_toggle.src = "images/restore.gif";
		}
		else 
		{
			if (parent.parent.parent.fs_top != null) 
			{
				parent.parent.parent.fs_top.rows="52,*";
				parent.parent.fs_bottom.cols="169,*";
			}
			parent.fstCols.cols="220,*";
			img_toggle.src = "images/collapse.gif";
		}
		menuOn = !menuOn;
    }

    function writeStyleSheet() 
	{
		document.write('<style type="text/css"><!--');

		if (navigator.appVersion.indexOf("MSIE 5.0") > 0) 
		{
			ie50 = true;
			document.write('.button  {border:#336699 solid 1px; padding:1px; cursor:pointer; height:1px;');
			document.write('.buttona {border:#ffffff solid 1px; padding:1px; cursor:pointer; height:1px; background-color:#336699;}');
		}
		else 
		{
			document.write('.button  {border:#336699 solid 1px; padding:1px; cursor:pointer;}');
			document.write('.buttona {border:#ffffff solid 1px; padding:1px; cursor:pointer; background-color:#6699cc;}');
		}

		document.write('//--></style>');
    } 
	
	writeStyleSheet();

  //-->
  </script>
</head>

<body bgcolor="#336699" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" style="border-bottom:1px solid #336699">
  <%DrawTabs tabDocuments, 2%>
</body>
</html>