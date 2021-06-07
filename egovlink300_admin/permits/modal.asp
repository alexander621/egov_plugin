<style>
.modaliframe
{
	width:100%;
	height:100%;
}

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
</style>
<script language="Javascript">
	<!--
	var modalcount = 0;
	function showModal(url, title, width, height) {
		genModal();

		left = (100-width)/2;
		csstop = (100 - height)/2;
		$("#modalframe" + modalcount).attr('src',url);
		$("#modaltitle" + modalcount).text(title);
		$("#modal" + modalcount).css({"width": width + "%","height": height + "%","left": left + "%","top": csstop + "%"});
		$("#modalwait" + modalcount).show();
		$("#modal" + modalcount).show();
		$('#modal' + modalcount).draggable({ handle: "#modalheader"+modalcount }); 
		$('#modal' + modalcount).resizable();
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
	    		modalBody += '<iframe class="modaliframe" data-close="' + modalcount + '" name="modalframe' + modalcount + '" id="modalframe' + modalcount +'" class="modal-body" onload="hideModalWait(' + modalcount + ');"></iframe>';
	    		modalBody += '<div id="modalwait' + modalcount +'" class="modalwait"><img src="../images/ajax-loader.gif" /></div>';
        		modalBody += '</div>';
    		modalBody += '</div>';
		modalBody += '</div>';

		//document.body.innerHTML += modalBody;
		$('body').append(modalBody);
	}
	function hideModal(id) {
		if (typeof document.getElementById("modalframe"+id).contentWindow.commonIFrameUpdateFunction === "function")
		{
			document.getElementById("modalframe"+id).contentWindow.commonIFrameUpdateFunction();
		}
		$("#modal"+id).remove();
	}
	function hideModalWait(id)
	{
		$("#modalwait"+id).hide();
	}
	-->
</script>
