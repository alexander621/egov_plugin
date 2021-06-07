<!-- #include file="../includes/common.asp" //-->
<%
 Dim sLocationName, sDocLocationName

 sLocationName    = trim(GetVirtualName(session("orgid")))
 sDocLocationName = trim(GetDocLocationName(session("orgid")))

'Get the starting folder
 lcl_folderStart = getStartingFolder(request("folderStart"))
%>
<html>
<head>
  <title>Choose File...</title>
  <style type="text/css">
  <!--
    td, input, select {font-family:MS Sans Serif,Tahoma,Arial; font-size:11px;}
  //-->
  </style>
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.0/jquery.min.js"></script>
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
      //window.close();
    }

	function buildlink(sFormField){
		// This is for the documents links
	  var sLocation = "<%=sLocationName%>";
      	  var path=document.frmFilePath.FilePath.value;
	  var sLinkName = document.frmFilePath.LinkName.value;
	  //var sVirtualDir = document.location.pathname;
	  var sVirtualDir = '<%=session("virtualdirectory")%>';
	  
	  var sFolderPath = document.all.currentpath.value;
	  var sUrlOnly = "";

	  /* Used to get virtual directory for specific organization */
	  //sVirtualDir = sVirtualDir.substring(0, sVirtualDir.indexOf("/a"));
	  sFolderPath = sFolderPath.replace(" ","+");


	  if (sLinkName =='') {
		sLinkName = path;
	  }

	  if (path =='') {
		sLink = "<a target='_EGOVLINK' href='" + sLocation + '/' + sVirtualDir + "/docs/menu/home.asp?path=" + sFolderPath + "'>" + sLinkName + "</a>";
		sLink = sLink.replace("/custom/pub","");
		sUrlOnly = sLocation + '/' +sVirtualDir + "/docs/menu/home.asp?path=" + sFolderPath;
		sUrlOnly = sUrlOnly.replace("/custom/pub","");
	  }
	  else
	  {
		sUrlOnly = sLocation + document.all.currentpath.value+ "/" + path;
		sUrlOnly = sUrlOnly.replace("/custom/pub","");

		if (sUrlOnly.indexOf("published_documents") > 0)
		{
		    //When a published document, then create short URL
		    $.ajax({
			type: "POST",
         		url: "shorturl.asp",
			data: {URL: sUrlOnly},
         		success: function(result) {
				sUrlOnly = result;
                  		},
         		async:   false
    		    });          

		}


		sLink = "<a target='_EGOVLINK' href='" + sUrlOnly + "'>" + sLinkName + "</a>";
	  }

	  document.frmFilePath.SiteLink.value = sLink;
	  document.frmFilePath.SiteURL.value = sUrlOnly;


	  sTextLink = sUrlOnly;
	  while (sTextLink.indexOf(" ") > -1) 
	  {
		sTextLink = sTextLink.replace(" ","%20");
	  }


	  //alert('buildlink: ' + sUrlOnly);
	  document.frmFilePath.TextLink.value = sTextLink;
	  /*
	  var oFormField = 'window.opener.document.' + sFormField;
	  window.opener.insertAtCaret(eval(oFormField), sLink);
	  window.close();*/
	}

	function buildactionlink(sFormField){
		// This is for the Action Line Forms
	  var sLocation = "<%=sLocationName%>";
	  var iFormID = document.frmAddArticle.iFormID.value;
	  var sLinkName = document.frmAddArticle.ALinkName.value;
	  var sUrlOnly = "";
	  var sVirtualDir = '<%=session("virtualdirectory")%>';

	  if (sLinkName =='') {
		sLinkName = document.frmAddArticle.AFormName.value;
	  }
	  sLink = sLocation +  '/' + sVirtualDir + "/action.asp?actionid=" + iFormID;
	  //while (sLink.indexOf(" ") > -1) 
		//{
		//	sLink = sLink.replace(" ","%20");
		//}
		sLink = "<a target='_EGOVLINK' href='" + sLink + "'>" + sLinkName + "</a>";
	  //sUrlOnly = "https://www.egovlink.com/" + sLocation + "/action.asp?actionid=" + iFormID;
	  sUrlOnly = sLocation +  '/' +  sVirtualDir + "/action.asp?actionid=" + iFormID;

	  document.frmAddArticle.SiteLink.value = sLink;
	  document.frmAddArticle.SiteURL.value = sUrlOnly;
	  sTextLink = sUrlOnly;
  	  while (sTextLink.indexOf(" ") > -1) 
	  {
		  sTextLink = sTextLink.replace(" ","%20");
	  }
	  //alert('buildactionlink: ' + sUrlOnly);
	  document.frmAddArticle.TextLink.value = sTextLink;
	  /*
	  var oFormField = 'window.opener.document.' + sFormField;
	  window.opener.insertAtCaret(eval(oFormField), sLink);
	  window.close();*/
	}

	function buildpaymentlink(sFormField){
		// This is for the Payment Forms Links
	  var sLocation = "<%=sLocationName%>";
	  var iFormID = document.frmPaymentLink.iFormID.value;
	  var sLinkName = document.frmPaymentLink.ALinkName.value;
	  var sUrlOnly = "";
	  var sVirtualDir = '<%=session("virtualdirectory")%>';

	  if (sLinkName =='') {
		sLinkName = document.frmPaymentLink.AFormName.value;
	  }
		// https://www.egovlink.com/
	  
	  sLink = sLocation +  sVirtualDir + "/payment.asp?paymenttype=" + iFormID;
	  //while (sLink.indexOf(" ") > -1) 
		//{
		//	sLink = sLink.replace(" ","%20");
		//}
		sLink = "<a target='_EGOVLINK' href='" + sLink + "'>" + sLinkName + "</a>";
	  sUrlOnly = sLocation +  sVirtualDir +  "/payment.asp?paymenttype=" + iFormID;

	  document.frmPaymentLink.SiteLink.value = sLink;
	  document.frmPaymentLink.SiteURL.value = sUrlOnly;
	  sTextLink = sUrlOnly;
	  while (sTextLink.indexOf(" ") > -1) 
	  {
		  sTextLink = sTextLink.replace(" ","%20");
	  }
	  //alert('buildpaymentlink: ' + sUrlOnly);
	  document.frmPaymentLink.TextLink.value = sTextLink;
	  /*
	  var oFormField = 'window.opener.document.' + sFormField;
	  window.opener.insertAtCaret(eval(oFormField), sLink);
	  window.close();*/
	}

	function buildwebpagelink(sFormField){
		// This is for the Web Links
	  var sURL = document.frmURL.UrlType.value + document.frmURL.Url.value;
	  var sLinkName = document.frmURL.UrlName.value;

	  if (sLinkName =='') {
		sLinkName = document.frmURL.Url.value;
	  }
	  sLink = "<a target='_EGOVLINK' href='" + sURL + "'>" + sLinkName + "</a>";

	  document.frmURL.SiteLink.value = sLink;
	  sTextLink = sURL;
	  while (sTextLink.indexOf(" ") > -1) 
	  {
		  sTextLink = sTextLink.replace(" ","%20");
	  }
	  //alert('buildwebpagelink: ' + sTextLink);
	  document.frmURL.TextLink.value = sTextLink;
	  document.frmURL.SiteURL.value = sURL;
	  /*
	  var oFormField = 'window.opener.document.' + sFormField;
	  window.opener.insertAtCaret(eval(oFormField), sLink);
	  window.close();*/
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
        <iframe name="menu" width="100" height="265" src="menu.asp?folderStart=<%=lcl_folderStart%>"></iframe>
      </td>
    </tr>
    <tr>
      <td valign="top">
      <%
        lcl_loadtree_url = "loadtree.asp?path=/public_documents300/custom/pub/" & sDocLocationName & lcl_folderStart

        'response.write "<iframe id=""explorer"" name=""explorer"" width=""400"" height=""250"" src=""loadtree.asp?path=/public_documents300/custom/pub/" & sDocLocationName & "/published_documents""></iframe>" & vbcrlf
        response.write "<iframe id=""explorer"" name=""explorer"" width=""400"" height=""250"" src=""" & lcl_loadtree_url & """></iframe>" & vbcrlf
      %>
        <div id="exdoc">
          <table border="0" cellpadding="0" cellspacing="0" width="400">
          <form name="frmFilePath">
            <tr>
              <td valign="top" style="padding-top:5px;">
			  <table>
				  <tr>
				  <td>Link Text:&nbsp;&nbsp;</td><td><input type="text" name="LinkName" style="width:250px; height:20px;" ></td></tr>
				  <tr><td>File name:&nbsp;&nbsp;</td><td><input type="text" name="FilePath" style="width:250px; height:20px;" ></td></tr>
			  </table>
		<br>
			<% if request("Message") <> "" then %>
				<% =request("Message") %>
			<% end if %>
              </td>
              <td valign="top" align="right" style="padding-top:5px;">
                <input type="button" value="Create Link" style="width:80px; height:22px;" onClick="javascript:buildlink('<%=request("name")%>');"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                <input type="button" value="Cancel" style="width:80px; height:22px;" onclick="window.close();">
              </td>
            </tr>
            <tr>
				<td colspan=3>Site link:<br>
				<input type="text" name="SiteLink" style="width:100%; height:20px;">
				</td>
            </tr>
			<tr>
				<td colspan=3>Savvy Link:<br />
				<input type="text" name="SiteURL" style="width:100%; height:20px;">
				</td>
            </tr>
			<tr>
				<td colspan=3>Text Link:<br />
				<input type="text" name="TextLink" style="width:100%; height:20px;">
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
            <tr>
				<td colspan=3>Site link:<br>
				<input type="text" name="SiteLink" style="width:100%; height:20px;">
				</td>
            </tr>
			<tr>
				<td colspan=3>Savvy Link:<br />
				<input type="text" name="SiteURL" style="width:100%; height:20px;">
				</td>
            </tr>
			<tr>
				<td colspan=3>Text Link:<br />
				<input type="text" name="TextLink" style="width:100%; height:20px;">
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
            <tr>
				<td colspan=3>Site link:<br>
				<input type="text" name="SiteLink" style="width:100%; height:20px;">
				</td>
            </tr>
			<tr>
				<td colspan=3>Savvy Link:<br />
				<input type="text" name="SiteURL" style="width:100%; height:20px;">
				</td>
            </tr>
			<tr>
				<td colspan=3>Text Link:<br />
				<input type="text" name="TextLink" style="width:100%; height:20px;">
				</td>
            </tr>
            </form>
          </table>
        </div>

        <div id="newurl"  style="display:none">
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
            <tr>
				<td colspan=3>Site link:<br>
				<input type="text" name="SiteLink" style="width:100%; height:20px;">
				</td>
            </tr>
			<tr>
				<td colspan=3>Savvy Link:<br />
				<input type="text" name="SiteURL" style="width:100%; height:20px;">
				</td>
            </tr>
			<tr>
				<td colspan=3>Text Link:<br />
				<input type="text" name="TextLink" style="width:100%; height:20px;">
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
  'sSQL = "SELECT OrgVirtualSiteName FROM Organizations WHere orgid='" &  iorgid & "'"
  sSQL = "SELECT orgegovwebsiteurl FROM Organizations WHere orgid='" &  iorgid & "'"
  oRst.open sSQL,Application("DSN"),3,1

  If NOT oRst.EOF THEN
	sReturnValue = Trim(oRst("orgegovwebsiteurl"))
	'response.write Trim(oRst("orgegovwebsiteurl")) & "&nbsp;" & Len(sReturnValue) & "&nbsp;" & InstrRev(sReturnValue,"/")
	sReturnValue = Mid(sReturnValue,1,(InstrRev(sReturnValue,"/")-1))
  END If
  oRst.close
  Set oRst = Nothing 

  GetVirtualName = sReturnValue

End Function


Function GetDocLocationName(iorgid)
  
  sReturnValue = "UNKNOWN"
  
  Set oRst = Server.CreateObject("ADODB.Recordset")
  sSQL = "SELECT OrgVirtualSiteName FROM Organizations WHere orgid='" &  iorgid & "'"
  oRst.open sSQL,Application("DSN"),3,1
  
  If NOT oRst.EOF THEN
	sReturnValue = Trim(oRst("OrgVirtualSiteName"))
  END If
  oRst.close
  Set oRst = Nothing 

  GetDocLocationName = sReturnValue

end function
%>
