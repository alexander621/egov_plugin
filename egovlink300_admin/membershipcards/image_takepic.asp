<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="image_upload.asp" -->
<%
'Check to see if the feature is offline
if isFeatureOffline("memberships") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

'option explicit 
Response.Expires     = -1
Server.ScriptTimeout = 600

sLevel = "../" ' Override of value from common.asp

'Determine if this is a demo or not.  demo = Y means that these screens can function without the web camera attached
 lcl_demo = request("demo")

'set up demo variables
 if lcl_demo = "Y" then
    lcl_demo_page_title = " (DEMO)"
    lcl_demo_url        = "&demo=" & lcl_demo
 else
    lcl_demo_page_title = ""
    lcl_demo_url        = ""
 end if

'Retreive the MemeberID
 dim lcl_member_id, lcl_reload_pic, lcl_file_directory
 lcl_member_id  = request("memberid")
 lcl_reload_pic = request("reload_pic")
 lcl_poolpassid = request("poolpassid")

' ****************************************************
'Display all of the files in the directory *** FOR DEBUGGING ONLY ***
 'set MyFile   = server.createobject("scripting.filesystemobject")
 'set MyFolder = MyFile.GetFolder("d:\wwwroot\www.cityegov.com\egovlink300_admin\")
 'set MyFolder = MyFile.GetFolder("d:\wwwroot\www.cityegov.com\egovlink release QA Test Environment 4.0.0\")
 '    for each thing in MyFolder.files
 '	    response.write(thing&"<br />")
 '    next

' ****************************************************
'Depending on the environment the folder will need to be set differently.
'You will need to uncomment out the one you need.
'sSystem = "DVLP"

'Select Case sSystem 
'	Case "DVLP"
'		lcl_file_directory = "c:\www_server_root\egovlink\egovlink release 4.0.0\"
'	Case "TEST"
'      lcl_file_directory = "c:\wwwroot\www.cityegov.com\egovlink release QA Test Environment 4.0.0\"
'	Case "PROD"
'      lcl_file_directory = "c:\wwwroot\www.cityegov.com\"
'	Case Else
'		    lcl_file_directory = "c:\www_server_root\egovlink\egovlink release 4.0.0\"
'End Select 
lcl_file_directory = Application("membershipcard_filedirectory")

'If this is a demo then no actual picture will need to be taken.
'Pull image from the DEMO folder.
 if lcl_demo = "Y" then
'    lcl_file_upload = "egovlink300_admin\images\MembershipCard_Photos\demo"
    lcl_file_upload = "\demo"
 else
'    lcl_file_upload = "egovlink300_admin\images\MembershipCard_Photos\temp"
    lcl_file_upload = "\temp"
 end if

lcl_file_directory = lcl_file_directory & lcl_file_upload

' ****************************************************
' Change the value of the variable below to the pathname
' of a directory with write permissions, for example "C:\Inetpub\wwwroot"
  Dim uploadsDirVar
'  uploadsDirVar = "d:\inetpub\webmailasp\database\tempUploads" 
  uploadsDirVar = lcl_file_directory
' ****************************************************

function OutputForm()
%>
    <form name="frmSend" method="POST" enctype="multipart/form-data" action="image_takepic.asp?memberid=<%=lcl_member_id%>&poolpassid=<%=lcl_poolpassid%>&demo=<%=lcl_demo%>" onSubmit="return onSubmitForm();">
    <span id="instruction_2">&nbsp;&nbsp;<strong>2. </strong>Once the picture has been taken you must now upload the file.  To do this you must enter the filename of 
        the image into the "File Name" field.  You can do this by either typing in the value or by clicking on the "Browse" button and selecting the file 
        itself.  The file you need is identified under the "Browse" button.</span>
    <p>
    <table id="file_name" border="0" cellspacing="0" cellpadding="2" style="width: 380px;">
      <tr>
          <td>
<%
  if request("step2") = "Y" then
     response.write "<p>" & vbcrlf
     response.write "<strong>File name:</strong>&nbsp;<input name=""attach1"" type=""text"" value=""c:/MembershipCard_Photos/" & lcl_member_id & ".jpg"" size=""35"">&nbsp;" & vbcrlf
     response.write "<input type=""button"" value=""Browse"" /><br />" & vbcrlf
     response.write "<div align=""center""><font style=""font-size:20px; color:#FF0000; font-weight:bold;"">c:/MembershipCard_Photos/" & lcl_member_id & ".jpg</font></div>" & vbcrlf
     response.write "</p>" & vbcrlf
  else
     response.write "<p>" & vbcrlf
     response.write "<strong>File name:</strong>&nbsp;<input name=""attach1"" type=""file"" size=""35"" onChange=""javascript:check_instruction3()""/><br />" & vbcrlf
     response.write "<div align=""center"" id=""path_name"">&nbsp;</div>" & vbcrlf
     response.write "</p>" & vbcrlf
  end if
%>
          </td>
      </tr>
    </table>

    <span id="instruction_3">&nbsp;&nbsp;<strong>3. </strong>Once you have selected the image to upload, you can then click on the "Upload Image" button.  
            This will set up the member's ID to be printed.<p></span>

    <table id="upload_image" border="0" cellspacing="0" cellpadding="2" style="width: 380px;">
<%
  if lcl_demo = "Y" then
     response.write "<tr><td><input style=""margin-top:4; left:17px; width:150px; top:363px; height:29px"" type=""button"" value=""Upload Image"" onclick=""javascript:location.href='image_display.asp?memberid=" & lcl_member_id & "&poolpassid=" & lcl_poolpassid & "&demo=Y'"" /></td></tr>" & vbcrlf
  else
     response.write "<tr><td><input style=""margin-top:4; left:17px; width:150px; top:363px; height:29px"" type=""submit"" value=""Upload Image"" /></td></tr>" & vbcrlf
  end if

  response.write "</table>" & vbcrlf
  response.write "</form>" & vbcrlf

end function

'------------------------------------------------------------------------------
function TestEnvironment()
    Dim fso, mf, fileName, testFile, streamTest
    TestEnvironment = ""
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    if not fso.FolderExists(uploadsDirVar) then
       set mf = fso.CreateFolder(lcl_file_directory)
'        TestEnvironment = "<strong>Folder " & uploadsDirVar & " does not exist.</strong><br />The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
'        exit function
    end if
    fileName = uploadsDirVar & "\test.txt"
    on error resume next
    Set testFile = fso.CreateTextFile(fileName, true)
    If Err.Number<>0 then
        TestEnvironment = "<strong>Folder " & uploadsDirVar & " does not have write permissions.</strong><br />The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
        exit function
    end if
    Err.Clear
    testFile.Close
    fso.DeleteFile(fileName)
    If Err.Number<>0 then
        TestEnvironment = "<strong>Folder " & uploadsDirVar & " does not have delete permissions</strong>, although it does have write permissions.<br />Change the permissions for IUSR_<I>computername</I> on this folder."
        exit function
    end if
    Err.Clear
    Set streamTest = Server.CreateObject("ADODB.Stream")
    If Err.Number<>0 then
        TestEnvironment = "<strong>The ADODB object <I>Stream</I> is not available in your server.</strong><br />Check the Requirements page for information about upgrading your ADODB libraries."
        exit function
    end if
    Set streamTest = Nothing
end function

'------------------------------------------------------------------------------
function SaveFiles()
    Dim Upload, fileName, fileSize, ks, i, fileKey

'	Set Upload = New FreeASPUpload
	Set Upload = New ImageUpload
 Upload.Save(uploadsDirVar)

	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then Exit function

    SaveFiles = ""
    ks = Upload.UploadedFiles.keys
    if (UBound(ks) <> -1) then
  	    Session("RELOAD_PIC") = "Y"
       response.redirect "image_display.asp?memberid=" & lcl_member_id & "&poolpassid=" & lcl_poolpassid

'        SaveFiles = "<strong>Files uploaded:</strong> "
'        for each fileKey in Upload.UploadedFiles.keys
'            SaveFiles = SaveFiles & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
'        next
    else
        SaveFiles = "The file name specified in the upload form does not correspond to a valid file in the system."
    end if
'	SaveFiles = SaveFiles & "<br />Enter a number = " & Upload.Form("enter_a_number") & "<br />"
'	SaveFiles = SaveFiles & "Checkbox values = " & Upload.Form("checkbox_values") & "<br />"
end function

%>
<html>
<head>
  <title>Membership Photo Taking System</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

<script language="Javascript">
// Register the Twain interface
function register() {
  var User;
  var Domain;
  var RegCode;

  User    = "Jerry Felix";
  Domain  = "egovlink.com";
  RegCode = "74F85955BBA574EBB983FE2F8EF479EB928CABF18C14ACB8BEC7715B484E95C86BB464D1151069EFBF19BD3094963A4642E78EA1EAD55CC8AFF321B09B62016CFF1AC3EA30FC03413FA72DCBD774E009ABA053A5151405F0355522C0277E053C23868C0CED1DE34766CD5C4EFCBA3EDDB1C6D3EF3685853973CE69067FCCE9EA407A1537D47A04CD6D62398FC2A9AE823E01ABECF8FEE61B443D50FCF443EFB9";

  VSTwain1.Register( User, Domain, RegCode );
}

// Setup the Camera
function init() {
  // hide these fields when the page is first opened.
//  document.getElementById("file_name").style.display     = "none";
//  document.getElementById("upload_image").style.display  = "none";

  // set the color to RED for the first instruction with the page is first opened.
  document.getElementById("instruction_1").style.color   = "#800000";
  document.getElementById("instruction_2").style.color   = "#000000";
  document.getElementById("instruction_3").style.color   = "#000000";

  VSTwain1.StartDevice()
  register();
  VSTwain1.maxImages           = 1;
  VSTwain1.autoCleanBuffer     = 1;
  VSTwain1.disableAfterAcquire = 1;
  VSTwain1.unitOfMeasure       = 0;  // inches
  VSTwain1.pixelType           = 2;  // RGB
  VSTwain1.resolution          = 600;
}

// Start the camera program and take the picture
function WithDialog() {
  init();
  VSTwain1.ShowUI = 1;
  VSTwain1.Acquire();
VSTwain1.DeleteImage(VSTwain1.numImages)
}

// Stop the Twain interface
function onPageUnload() {
  VSTwain1.StopDevice()
}

function GoBack() {
  var lcl_return_page = '<%=session("redirectpage")%>';

  if (lcl_return_page != "") {
      location.href = lcl_return_page;
  } else {
      history.go(-1);
  }
}

function onSubmitForm() {
    var formDOMObj = document.frmSend;
    if (formDOMObj.attach1.value == "")
        alert("Please press the browse button and pick a file.")
    else
        return true;
    return false;
}

function check_instruction3() {
  if(document.frmSend.attach1.value != "") {
     document.getElementById("instruction_2").style.color  = "#000000";
     document.getElementById("instruction_3").style.color  = "#800000";
//     document.getElementById("upload_image").style.display = "block";
  }else{
     document.getElementById("instruction_2").style.color  = "#800000";
     document.getElementById("instruction_3").style.color  = "#000000";
//     document.getElementById("upload_image").style.display = "none";
  }

}

function openWin(page) {
  OpenWin = window.open(page, "new", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes,width=505,height=370,screenX=0,screenY=0");
  if (document.images) {OpenWin.focus();}
}
</script>

<script language="javascript" event="PostScan(flag)" for="VSTwain1">
<!--
     if(flag != 0) {
        if (VSTwain1.errorCode != 0) alert(VSTwain1.errorString)
     } else {
        var imgPath = "c:/MembershipCard_Photos/<%=lcl_member_id%>.jpg";

        if(VSTwain1.SaveImage(0,imgPath) == 0) {
           alert(VSTwain1.errorString);
           alert("Please try to scan black-white or gray images!");
        }else{
           var img = new Image();
           img.src = imgPath;

      		   document.getElementById("file_name").style.display     = "block";
           document.getElementById("instruction_2").style.display = "block";

           document.getElementById("instruction_1").style.color   = "#000000";
           document.getElementById("instruction_2").style.color   = "#800000";

           document.getElementById("button1").value               = "Retake Picture";
           document.getElementById("path_name").innerHTML         = "<font style=\"font-size: 14px; font-style: italic; \">The file can be found here: </font><font style=\"font-size: 20px; color: FF0000; font-weight: bold; \">" + imgPath + "</font>";
        }
    }
//     VSTwain1.CloseDataSource()
// -->
</script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onUnload="onPageUnload()" onLoad="init();">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!-- <object id="VSTwain1" width="1" height="1" classid="CLSID:1169E0CD-9E76-11D7-B1D8-FB63945DE96D" codebase=""></object> -->
<!-- <object id="VSTwain1" width="1" height="1" classid="CLSID:1169E0CD-9E76-11D7-B1D8-FB63945DE96D" codebase="camera_dvlp_driver/VSTwain_new.dll"></object> -->
<object id="VSTwain1" width="1" height="1" classid="CLSID:1169E0CD-9E76-11D7-B1D8-FB63945DE96D" codebase="camera_dvlp_driver/VSTwain.dll"></object>

<!--BEGIN PAGE CONTENT-->
<div id="content">
<div id="centercontent">

  <h3>Taking a Picture<%=lcl_demo_page_title%></h3>
  <a href="javascript:GoBack()"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=Session("RedirectLang")%></a>
  <p>
  <strong>Instructions: </strong><br />

<!-- Instruction 1 -->
  <span id="instruction_1">&nbsp;&nbsp;<strong>1. </strong>Click the "Take Picture" button when you are ready to take the member's picture.<p></span>

  <table border="0" cellspacing="0" cellpadding="0">
<%
  'If this is a demo then simulate the picture being taken
   if lcl_demo = "Y" then
      lcl_onclick = "javascript:openWin('demo_image_takepic.asp?memberid=" & lcl_member_id & "&poolpassid=" & lcl_poolpassid & "&demo=Y')"
   else
      lcl_onclick = "javascript:WithDialog()"
   end if
%>
    <tr><td><input type="button" id="button1" name="button1" style="left: 17px; width: 150px; top: 363px; height: 29px" onclick="<%=lcl_onclick%>" value="Take Picture" /></td></tr>
  </table>
  <p>
<%
'Set up the file input form field

Dim diagnostics
if Request.ServerVariables("REQUEST_METHOD") <> "POST" then
    diagnostics = TestEnvironment()
    if diagnostics<>"" then
        response.write "<div style=""margin-left:20; margin-top:30; margin-right:30; margin-bottom:30;"">"
        response.write diagnostics
        response.write "<p>After you correct this problem, reload the page."
        response.write "</div>"
    else
        response.write "<div style=""margin-left:150"">"
        OutputForm()
        response.write "</div>"
    end if
else
    response.write "<div style=""margin-left:150"">"
    OutputForm()
    response.write SaveFiles()
    response.write "<br /><br /></div>"
end if
%>

</div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
