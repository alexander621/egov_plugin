<script language="Javascript">
	<!--

		function uploadtoWordPress(targetID, modalid)
		{
			$("#modalbody" + modalcount).hide();
			$("#modalwait" + modalcount).show();
  			$.ajax({
    				// Your server script to process the upload
    				url: '<%=sHomeWebsiteURL%>',
    				type: 'POST',
				
    				// Form data
    				data: new FormData($('#uploadimg')[0]),
				
    				// Tell jQuery not to process data or worry about content-type
    				// You *must* include these options!
    				cache: false,
    				contentType: false,
    				processData: false,
				
    				// Custom XMLHttpRequest
    				xhr: function () {
      					var myXhr = $.ajaxSettings.xhr();
      					if (myXhr.upload) {
        					// For handling the progress of the upload
        					myXhr.upload.addEventListener('progress', function (e) {
          						if (e.lengthComputable) {
            							$('progress').attr({
              							value: e.loaded,
              							max: e.total,
            						});
          					}
        					}, false);
      					}
      					return myXhr;
    				},
    				error: function() {
	    				alert("Sorry, there was an error");
    				},
    				success: function(data) {
	    				var url = data.split("##EGOVFILEURL##");
	    				//alert(url[1]);
					//document.getElementById(targetID).value = url[1];
					$("#" + targetID).val(url[1]);
					//document.getElementById(targetID + 'pic').src = url[1];
					$("#" + targetID + "pic").attr("src",url[1]);
					if (targetID.indexOf("document") >= 0)
					{
						$("#" + targetID + 'pic').html('<a href="' + url[1] + '" target="_newwindow">View Document</a>&nbsp;&nbsp;');
					}
					//Destroy Modal
					hideModal(modalid);
    				}
  			});
		}

		$( document ).ready(function() {
			refreshImageURLListener();
		});

		function refreshImageURLListener()
		{
			$(".imageurl").change(function() {
				document.getElementById(this.name + 'pic').src = this.value;
			});
		}
	var modalcount = 0;
	function showModal(title, width, height,target) {
		genModal();

		left = (100-width)/2;
		csstop = (100 - height)/2;
		$("#modaltitle" + modalcount).text(title);
		$("#modal" + modalcount).css({"width": width + "%","height": height + "%","left": left + "%","top": csstop + "%"});
		//$("#modalwait" + modalcount).show();
		$("#modal" + modalcount).show();
		$('#modal' + modalcount).draggable({ handle: "#modalheader"+modalcount }); 
		$('#modal' + modalcount).resizable();
		//if (modaltype == "upload")
		//{
			//$('#modalbody' + modalcount).html('<form method="post" id="uploadimg" enctype="multipart/form-data"><input type="hidden" name="egovcode" value="KEY" /><input type="hidden" name="add_from_egov" value="Upload" /><input type="file" name="file" required /></form><input id="uploadbtn" type="button" value="Upload" onClick="uploadtoWordPress(\'' + target + '\',' + modalcount + ');" />');
			//$("#modalwait" + modalcount).hide();
		//}
		//if (modaltype == "pick")
		//{
			//Get Pick Content
			getPickerContent(1, target, modalcount);
		//}
	}
	function genModal()
	{
		modalcount++;
		var modalBody = '<div id="modal' + modalcount +'" class="modal fade">';
    		modalBody += '<div class="modal-dialog">';
        		modalBody += '<div class="modal-content">';
            		modalBody += '<div id="modalheader' + modalcount + '" class="modal-header">';
	    			modalBody += '<div id="closemodal' + modalcount +'" class="closemodal" onclick="hideModal(' + modalcount + ');"><img src="../images/close-icon.png" width="15" height="15" border="0" /></div>';
                 		modalBody += '<h4 class="modal-title" id="modaltitle' + modalcount +'"></h4>  ';
            		modalBody += '</div>';
	    		modalBody += '<div id="modalbody' + modalcount + '" class="modalbody"></div>';
	    		modalBody += '<div id="modalwait' + modalcount +'" class="modalwait"><img src="../images/ajax-loader.gif" /></div>';
        		modalBody += '</div>';
    		modalBody += '</div>';
		modalBody += '</div>';

		//document.body.innerHTML += modalBody;
		$('body').append(modalBody);
	}
	function hideModal(id) {
		$("#modal"+id).remove();
	}
	function hideModalWait(id)
	{
		$("#modalwait"+id).hide();
	}
	function getPickerContent(count, target, modalcount)
	{
		$.get( "<%=sHomeWebsiteURL%>?showmedialib=" + count + "&target=" + target + "&modalcount=" + modalcount, function( data ) {
    			var picks = data.split("###EGOVPICKER###");
 			$( "#modalbody" + modalcount ).html( picks[1] + '<br /><br /><a name="uploadnew"></a><h3>Upload New Image</h3><form method="post" id="uploadimg" enctype="multipart/form-data"><input type="hidden" name="egovcode" value="KEY" /><input type="hidden" name="add_from_egov" value="Upload" /><input type="file" name="file" required /></form><br /><br /><input id="uploadbtn" type="button" class="button" value="Upload and Pick" onClick="uploadtoWordPress(\'' + target + '\',' + modalcount + ');" />');
			$( "#modalbody" + modalcount ).show();
			$( "#modalbody" + modalcount ).scrollTop(0);
			$("#modalwait" + modalcount).hide();
		});
	}
	function selImg(targetID, modalid, url)
	{
		//document.getElementById(targetID).value = url;
		$("#" + targetID).val(url);
		//document.getElementById(targetID + 'pic').src = url;
		$("#" + targetID + "pic").attr("src",url);
		if (targetID.indexOf("document") >= 0)
		{
			$("#" + targetID + 'pic').html('<a href="' + url + '" target="_newwindow">View Document</a>&nbsp;&nbsp;');
		}
		//Destroy Modal
		hideModal(modalid);
	}
	-->
</script>
<style>
.modal{
    border:4px solid #CCC;
    position:fixed;
    left:50%; 
    top:50%;
    border-radius:5px;
    display:none;
    width:50%;
    height:60%;
}
.modalbody
{
	padding:20px;
	overflow:scroll;
	display:none;
}
.comment {
  opacity: .6;
  font-style: italic;
  position: absolute;
  left: 40%;
}
.modal
{
    overflow: hidden;
    background:white;
}
.modal-content {height:100%;display:flex;flex-flow: column;}
.modal-dialog{
    margin-right: 0;
    margin-left: 0;
    height:100%;
}
.modal-header{
  background-color:#2C5F93;
  color:white;
  flex: 0 1 auto;
}
.modal-header h4 { margin: 0;}
.modal-title{
  font-size:16px;padding: 6px;
}
.modal-header .close{
  color:#fff;
}
.modal-body{
  color:#888;
  flex: 1 1 auto;
}
.modal-body p {
  text-align:center;
  padding-top:10px;
}
.closemodal
{
	float:right;
	margin:6px;
}
.modalwait
{
    position:absolute;
    left:50%; 
    top:50%;
    margin: -16px 0 0 -16px;
}

.modalbody .nextprev
{
	padding:0;
	margin:0;
	clear:both;
}
.modalbody .nextprev li
{
	display:inline-block;
}
.modalbody .nextprev li:nth-child(2)
{
	display:inline-block;
	margin-left:50px;
}
.modalbody .imggroup img
{
	margin:2px;
	vertical-align:text-top;
}
.modalbody .imggroup .imgpickbox
{
	display:block;
	text-align:center;
	width:154px;
	height:154px;
	overflow-wrap: break-word;
}
.modalbody .imggroup .imgpickbox.file
{
	height:80px;
	margin-top:20px;
}
.modalbody .imggroup .filename
{
	overflow-wrap:break-word;
	width:154px;
	text-align:center;
}
.modalbody .imggroup a
{
	display:inline-block;
	margin:0 8px 15px 8px;
	position:relative;
	float:left;
}
</style>
