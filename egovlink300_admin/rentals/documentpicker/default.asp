<%sLocationName = trim(GetVirtualName(Session("OrgID")))
sHostUrl = Request.ServerVariables("HTTP_HOST")
%>

<html>
<head>
  <title>Choose File...</title>
  <style type="text/css">
  <!--
    td, input, select {font-family:MS Sans Serif,Tahoma,Arial; font-size:11px;}
  //-->
  </style>
  <script language="Javascript">
  <!--
    function MakeActive(id) {
      document.all.exdoc.style.display = "none";
      document.all.newdoc.style.display = "none";
      document.all.newurl.style.display = "none";
	  document.all.newpay.style.display = "none";
      eval("document.all." + id + ".style.display = ''");
    }
    
    function saveSelection()
    {
      var objParent=window.opener;
      var path=document.frmFilePath.FilePath.value;
	  document.frmFilePath.FileLink.value = "<a href='" + document.all.currentpath.value + "/" + path + "'>" + path + "</a>";
	  objParent.addItem.itemID.value=document.all.currentpath.value + "/" + path;
	  objParent.addItem.link.value=path;
	  if(objParent.addItem.title.value=='')objParent.addItem.title.value=path;
      window.close();
    }

	function buildlink(sFormField)
	{
		var sLink = "";
		var path = document.frmFilePath.FilePath.value;

		if (document.frmFilePath.FilePath.value == "") 
		{
			alert("Please select an image.");
			return;
		}
		sLink = "https://<%=sHostUrl%>" + document.all.currentpath.value + "/" + path;
		sLink = sLink.replace("/custom/pub","");
		//alert( sLink );

	  var oFormField = 'window.opener.document.' + sFormField; 
	  
	  window.opener.insertAtURL(eval(oFormField), sLink);
	  window.close();
	}

	function buildactionlink(sFormField){
	  var sLocation = "<%=sLocationName%>";
	  var iFormID = document.frmAddArticle.iFormID.value;
	  var sLinkName = document.frmAddArticle.ALinkName.value;
	  if (sLinkName =='') {
		sLinkName = document.frmAddArticle.AFormName.value;
	  }
	 
	  sLink = "<a target='_EGOVLINK' href='http://www.egovlink.com/" + sLocation + "/action.asp?actionid=" + iFormID + "'>" + sLinkName + "</a>";

	  var oFormField = 'window.opener.document.' + sFormField; 
	  window.opener.insertAtCaret(eval(oFormField), sLink);
	  window.close();
	}

	function buildpaymentlink(sFormField){
	  var sLocation = "<%=sLocationName%>";
	  var iFormID = document.frmPaymentLink.iFormID.value;
	  var sLinkName = document.frmPaymentLink.ALinkName.value;
	  if (sLinkName =='') {
		sLinkName = document.frmPaymentLink.AFormName.value;
	  }

	  sLink = "<a target='_EGOVLINK' href='http://www.egovlink.com/" + sLocation + "/payment.asp?paymenttype=" + iFormID + "'>" + sLinkName + "</a>";

	  var oFormField = 'window.opener.document.' + sFormField; 
	  window.opener.insertAtCaret(eval(oFormField), sLink);
	  window.close();
	}

	function buildwebpagelink(sFormField){
	 
	  var sURL = document.frmURL.UrlType.value + document.frmURL.Url.value;
	  var sLinkName = document.frmURL.UrlName.value;
	  if (sLinkName =='') {
		sLinkName = document.frmURL.Url.value;
	  }
	  sLink = "<a target='_EGOVLINK' href='" + sURL + "'>" + sLinkName + "</a>";

	  var oFormField = 'window.opener.document.' + sFormField; 
	  window.opener.insertAtCaret(eval(oFormField), sLink);
	  window.close();
	}

    function myFunction() {
	    alert(document.frmAddArticle.currentfolderpath.value);
	 }
  //-->
  </script>

 <script language="Javascript">
  <!--
    function doPicker() {
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("../picker/default.asp", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }
  //-->
  </script>
</head>

<body bgcolor="#d4d0c8" leftmargin="2" topmargin="0">
<input type="hidden" name="currentpath" >
  <table border="0" cellpadding="3" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
      <td>
        <input type="text" name="currentfolder" style="height:20px; width:250px;" readonly>&nbsp;<a href="#" style="color:#0000ff" onclick="explorer.window.history.back();" name="anchorBack"><img src="images/up.gif" alt="Back" border="0" align="absmiddle"></a>
      </td>
    </tr>
    <tr>
      <td rowspan="3" valign="top">
        <iframe name="menu" width="100" height="265" src="menu.asp"></iframe>
      </td>
    </tr>
    <tr>
      <td valign="top">

        <iframe id="explorer" name="explorer" width="400" height="250" src="loadtree.asp?path=<%="/public_documents300/custom/pub/" & sLocationName %>"></iframe>

        <div id="exdoc">
          <table border="0" cellpadding="0" cellspacing="0" width="400">
          <form name="frmFilePath">
            <tr>
              <td valign="top" style="padding-top:5px;">
			  <table>
				  <!--<tr><td>Link Text:&nbsp;&nbsp;</td><td><input type="text" name="LinkName" style="width:250px; height:20px;" ></td></tr>-->
				  <tr><td>File name:&nbsp;&nbsp;</td><td><input type="text" name="FilePath" style="width:250px; height:20px;" readonly></td></tr>
			  </table>
		<br>
			<% if request("Message") <> "" then %>
				<% =request("Message") %>
			<% end if %>
              </td>
              <td valign="top" align="right" style="padding-top:5px;">
                <input type="button" value="Select" style="width:80px; height:22px;" onClick="javascript:buildlink('<%=request("name")%>');"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                <input type="button" value="Cancel" style="width:80px; height:22px;" onclick="window.close();">
              </td>
            </tr>
            </form>
          </table>
        </div>

        <div id="newdoc" style="display:none">
                    <table border="0" cellpadding="0" cellspacing="0" width="400">
          <form name="frmAddArticle">
		   <input type="hidden" name="currentfolderpath">
		   <input type="hidden" name="iFormID">
            <tr>
              <td valign="top" style="padding-top:5px;">
			  <table>
				  <tr>
				  <td>Link Text:&nbsp;&nbsp;</td><td><input type="text" name="ALinkName" style="width:250px; height:20px;" ></td></tr>
				  <tr><td>Form Name:&nbsp;&nbsp;</td><td><input type="text" name="AFormName" style="width:250px; height:20px;" readonly></td></tr>
			  </table>
		<br>
			<% if request("Message") <> "" then %>
				<% =request("Message") %>
			<% end if %>
              </td>
              <td valign="top" align="right" style="padding-top:5px;">
                <input type="button" value="Add Link" style="width:80px; height:22px;" onClick="javascript:buildactionlink('<%=request("name")%>');"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                <input type="button" value="Cancel" style="width:80px; height:22px;" onclick="window.close();">
              </td>
            </tr>
            </form>
          </table>
        </div>

		        <div id="newpay" style="display:none">
                    <table border="0" cellpadding="0" cellspacing="0" width="400">
          <form name="frmPaymentLink">
		   <input type="hidden" name="currentfolderpath">
		   <input type="hidden" name="iFormID">
            <tr>
              <td valign="top" style="padding-top:5px;">
			  <table>
				  <tr>
				  <td>Link Text:&nbsp;&nbsp;</td><td><input type="text" name="ALinkName" style="width:250px; height:20px;" ></td></tr>
				  <tr><td>Form Name:&nbsp;&nbsp;</td><td><input type="text" name="AFormName" style="width:250px; height:20px;" readonly></td></tr>
			  </table>
		<br>
			<% if request("Message") <> "" then %>
				<% =request("Message") %>
			<% end if %>
              </td>
              <td valign="top" align="right" style="padding-top:5px;">
                <input type="button" value="Add Link" style="width:80px; height:22px;" onClick="javascript:buildpaymentlink('<%=request("name")%>');"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                <input type="button" value="Cancel" style="width:80px; height:22px;" onclick="window.close();">
              </td>
            </tr>
            </form>
          </table>
        </div>

        <div id="newurl" style="display:none">
          <form name="frmURL">
		  <table border="0" cellpadding="0" cellspacing="0" width="400">
            <tr>
              <td valign="top" style="padding-top:5px;">
                Name:&nbsp;&nbsp;<input type="text" name="UrlName" style="width:267px; height:20px;"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                URL:&nbsp;&nbsp;&nbsp;&nbsp;<select name="UrlType">
                  <option value="http://">http://</option>
                  <option value="mailto:">mailto:</option>
                  <option value="ftp://">ftp://</option>
                </select>
                <input type="text" name="Url" style="width:207px; height:20px;">
              </td>
              <td align="right" style="padding-top:5px;">
                    <input type="button" value="Add Link" style="width:80px; height:22px;" onClick="javascript:buildwebpagelink('<%=request("name")%>');"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                <input type="button" value="Cancel" style="width:80px; height:22px;" onclick="window.close();">
              </td>
            </tr>
          </table>
		  </form>
        </div>


      </td>
    </tr>
  </table>


</body>
</html>

<%
Function GetVirtualName(iorgid)
  
  sReturnValue = "UNKNOWN"
  
  Set oRst = Server.CreateObject("ADODB.Recordset")
  sSQL = "SELECT OrgVirtualSiteName FROM Organizations WHere orgid='" &  iorgid & "'"
  oRst.open sSQL,Application("DSN"),3,1
  
  If NOT oRst.EOF THEN
	sReturnValue = oRst("OrgVirtualSiteName")
  END IF

  GetVirtualName = sReturnValue

End Function
%>
