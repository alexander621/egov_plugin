<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="user_imageupload.asp" -->
<%
'Check to see if the feature is offline
 if isFeatureOffline("registration") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 if not userhaspermission(session("userid"),"create_user_membershipcards") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 dim lcl_userid, lcl_reload_pic, lcl_file_directory, diagnostics, uploadsDirVar

'option explicit 
 Response.Expires     = -1
 Server.ScriptTimeout = 600
 sLevel               = "../"  'Override of value from common.asp

'Retreive the UserID
 lcl_userid     = request("userid")
 lcl_reload_pic = request("reload_pic")

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
 'lcl_file_upload    = "egovlink300_admin\images\MembershipCard_Photos\temp"
 lcl_file_upload    = "\temp"
 lcl_file_directory = lcl_file_directory & lcl_file_upload

'------------------------------------------------------------------------------
' Change the value of the variable below to the pathname
' of a directory with write permissions, for example "C:\Inetpub\wwwroot"
  'uploadsDirVar = "d:\inetpub\webmailasp\database\tempUploads" 
  uploadsDirVar = lcl_file_directory
%>
<html>
<head>
  <title>Membership Photo Taking System</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

<script language="javascript">
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
        var imgPath = "c:/MembershipCard_Photos/<%=lcl_userid%>.jpg";

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

           document.getElementById("takePictureButton").value     = "Retake Picture";
           document.getElementById("path_name").innerHTML         = "<font style=\"font-size: 14px; font-style: italic; \">The file can be found here: </font><font style=\"font-size: 20px; color: FF0000; font-weight: bold; \">" + imgPath + "</font>";
        }
    }
//     VSTwain1.CloseDataSource()
// -->
</script>
<style type="text/css">
  #takePictureButton {
     top:   363px;
     left:  17px;
     width: 150px;
     height:29px;
  }

  #diagnostics {
     margin-top:    30px;
     margin-left:   20px;
     margin-right:  30px;
     margin-bottom: 30px;
  }
</style>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onUnload="onPageUnload()" onLoad="init();">

<!-- <object id="VSTwain1" width="1" height="1" classid="CLSID:1169E0CD-9E76-11D7-B1D8-FB63945DE96D" codebase=""></object> -->
<object id="VSTwain1" width="1" height="1" classid="CLSID:1169E0CD-9E76-11D7-B1D8-FB63945DE96D" codebase="camera_dvlp_driver/VSTwain.dll"></object>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<h3>Taking a Picture</h3>" & vbcrlf
  response.write "<p><input type=""button"" name=""returnButton"" id=""returnButton"" value=""Back to List"" class=""button"" onclick=""location.href='user_search.asp'"" /></p>" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "<strong>Instructions: </strong><br />" & vbcrlf

 'BEGIN: Instruction 1 --------------------------------------------------------
  lcl_onclick = "javascript:WithDialog()"

  'response.write "<span id=""instruction_1"">&nbsp;&nbsp;<strong>1. </strong>Click the ""Take Picture"" button when you are ready to take the member's picture.</span>" & vbcrlf
  response.write "<span id=""instruction_1"">&nbsp;&nbsp;<strong>1. </strong>Take the picture with your picture taking software.</span>" & vbcrlf
  response.write "</p>" & vbcrlf
  'response.write "<p>" & vbcrlf
  'response.write "<input type=""button"" name=""takePictureButton"" id=""takePictureButton"" value=""Take Picture"" class=""button"" onclick=""" & lcl_onclick & """ />" & vbcrlf
  'response.write "</p>" & vbcrlf

 'BEGIN: Set up the file input form field -------------------------------------
  if request.ServerVariables("REQUEST_METHOD") <> "POST" then
     diagnostics = TestEnvironment(lcl_file_directory, uploadsDirVar)

     if diagnostics <> "" then
         response.write "<div id=""diagnostics"">" & vbcrlf
         response.write    diagnostics & vbcrlf
         response.write "  <p>After you correct this problem, reload the page." & vbcrlf
         response.write "</div>" & vbcrlf
     else
         response.write "<div id=""outputform"">" & vbcrlf
                           OutputForm lcl_userid
         response.write "</div>" & vbcrlf
     end if
  else
      response.write "<div id=""outputform"">" & vbcrlf
                        OutputForm lcl_userid
                        SaveFiles lcl_userid, uploadsDirVar
      response.write "</div>" & vbcrlf
  end if
 'END: Set up the file input form field ---------------------------------------

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub OutputForm(iUserID)

  sUserID = 0

  if iUserID <> "" then
     sUserID = clng(iUserID)
  end if

  response.write "<form name=""frmSend"" method=""post"" enctype=""multipart/form-data"" action=""user_takepic.asp?userid=" & sUserID & """ onsubmit=""return onSubmitForm();"">" & vbcrlf
  response.write "  <span id=""instruction_2"">" & vbcrlf
  response.write "    &nbsp;&nbsp;<strong>2. </strong>Once the picture has been taken you must now upload the file.  " & vbcrlf
  response.write "    To do this you must enter the filename of the image into the ""File Name"" field.  You can do this by " & vbcrlf
  response.write "    either typing in the value or by clicking on the ""Browse"" button and selecting the file itself.  " & vbcrlf
  response.write "    The file you need is identified under the ""Browse"" button.</span>" & vbcrlf
  response.write "  <p>" & vbcrlf
  response.write "  <table id=""file_name"" border=""0"" cellspacing=""0"" cellpadding=""2"" style=""width:380px;"">" & vbcrlf
  response.write "    <tr>" & vbcrlf
  response.write "        <td>" & vbcrlf

  if request("step2") = "Y" then
     response.write "            <p>" & vbcrlf
     response.write "              <strong>File name:</strong>&nbsp;<input name=""attach1"" type=""text"" value=""c:/MembershipCard_Photos/" & sUserID & ".jpg"" size=""35"">&nbsp;" & vbcrlf
     response.write "              <input type=""button"" value=""Browse"" /><br />" & vbcrlf
     response.write "              <div align=""center""><font style=""font-size:20px; color:#FF0000; font-weight:bold;"">c:/MembershipCard_Photos/" & sUserID & ".jpg</font></div>" & vbcrlf
     response.write "            </p>" & vbcrlf
  else
     response.write "            <p>" & vbcrlf
     response.write "              <strong>File name:</strong>&nbsp;<input name=""attach1"" type=""file"" size=""35"" onChange=""javascript:check_instruction3()""/><br />" & vbcrlf
     response.write "              <div align=""center"" id=""path_name"">&nbsp;</div>" & vbcrlf
     response.write "            </p>" & vbcrlf
  end if

  response.write "        </td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "  </table>" & vbcrlf
  response.write "  <span id=""instruction_3"">" & vbcrlf
  response.write "    <p>&nbsp;&nbsp;<strong>3. </strong>Once you have selected the image to upload, you can then click on " & vbcrlf
  response.write "    the ""Upload Image"" button.  This will set up the member's ID to be printed.</p>" & vbcrlf
  response.write "  </span>" & vbcrlf
  response.write "  <table id=""upload_image"" border=""0"" cellspacing=""0"" cellpadding=""2"" style=""width:380px;"">" & vbcrlf
  response.write "    <tr><td><input style=""margin-top:4; left:17px; width:150px; top:363px; height:29px"" type=""submit"" class=""button"" value=""Upload Image"" /></td></tr>" & vbcrlf
  response.write "  </table>" & vbcrlf
  response.write "</form>" & vbcrlf

end sub

'------------------------------------------------------------------------------
function TestEnvironment(iFileDirectory, iUploadsDirVar)
  dim fso, mf, fileName, testFile, streamTest, lcl_return, lcl_filedirectory, lcl_uploadsdirvar

   lcl_return        = ""
   lcl_filedirectory = iFileDirectory
   lcl_uploadsdirvar = iUploadsDirVar

   set fso = Server.CreateObject("Scripting.FileSystemObject")

   if not fso.FolderExists(lcl_uploadsdirvar) then
      set mf = fso.CreateFolder(lcl_filedirectory)
      'lcl_return = "<strong>Folder " & uploadsDirVar & " does not exist.</strong><br />The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
      'exit function
   end if

   fileName = lcl_uploadsdirvar & "\test.txt"

   on error resume next
   set testFile = fso.CreateTextFile(fileName, true)

   if Err.Number <> 0 then
      lcl_return = lcl_return & "<strong>Folder " & lcl_uploadsdirvar & " does not have write permissions.</strong><br />" & vbcrlf
      lcl_return = lcl_return & "The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and " & vbcrlf
      lcl_return = lcl_return & "change the value of uploadsDirVar to the pathname of a directory with write permissions." & vbcrlf
      exit function
   end if

   Err.Clear
   testFile.Close
   fso.DeleteFile(fileName)

   If Err.Number<>0 then
      lcl_return = lcl_return & "<strong>Folder " & lcl_uploadsdirvar & " does not have delete permissions</strong>, " & vbcrlf
      lcl_return = lcl_return & "although it does have write permissions.<br />Change the permissions for " & vbcrlf
      lcl_return = lcl_return & "IUSR_<em>computername</em> on this folder."
      exit function
   end if

   Err.Clear
   Set streamTest = Server.CreateObject("ADODB.Stream")

   If Err.Number<>0 then
      lcl_return = lcl_return & "<strong>The ADODB object <em>Stream</em> is not available in your server.</strong><br />" & vbcrlf
      lcl_return = lcl_return & "Check the Requirements page for information about upgrading your ADODB libraries." & vbcrlf
      exit function
   end if

   set streamTest = nothing

   TestEnvironemnt = lcl_return

end function

'------------------------------------------------------------------------------
sub SaveFiles(iUserID, iUploadsDirVar)
  dim Upload, fileName, fileSize, ks, i, fileKey, lcl_uploadsdirvar

  sUserID           = 0
  lcl_uploadsdirvar = iUploadsDirVar

  if iUserID <> "" then
     sUserID = clng(iUserID)
  end if

  'Set Upload = New FreeASPUpload
 	set Upload = New ImageUpload
  Upload.Save(lcl_uploadsdirvar)

	'If something fails inside the script, but the exception is handled
 	if Err.Number = 0 then
     'SaveFiles = ""
     ks = Upload.UploadedFiles.keys

     if (UBound(ks) <> -1) then
  	     session("RELOAD_PIC") = "Y"
        response.redirect "user_displaycard.asp?userid=" & sUserID

        'SaveFiles = "<strong>Files uploaded:</strong> "
        'for each fileKey in Upload.UploadedFiles.keys
        '    SaveFiles = SaveFiles & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
        'next
     else
        response.write "The file name specified in the upload form does not correspond to a valid file in the system." & vbcrlf
     end if
  end if

     'SaveFiles = SaveFiles & "<br />Enter a number = " & Upload.Form("enter_a_number") & "<br />"
     'SaveFiles = SaveFiles & "Checkbox values = " & Upload.Form("checkbox_values") & "<br />"
end sub
%>
