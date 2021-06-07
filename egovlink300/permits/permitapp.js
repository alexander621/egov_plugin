
function submitPA()
{
	document.getElementById("submitBtn").disabled = true;

	if (validateForm("permitAppFrm"))
	{

       		var form_data = new FormData(document.getElementById('permitAppFrm'));
		var postToURL = "https://www.egovlink.com/permitcity/permits/permitappsubmit.asp";
	
        	// send it and then alert on the done
        	var request = jQuery.ajax({
                	type: "POST",
                	url: postToURL,
                	data: form_data,
                	contentType: false,
                	processData: false
        	});
	
        	request.done( function( data, textStatus, jqXHR ) {
			//alert(data);
			//NEEDS TO DO SOMETHING WITH THIS SUCCESS
			//alert("SUCCESS!");
			document.getElementById("permitAppFrm").style.display = "none";
			document.getElementById("successmsg").style.display = "block";
			document.getElementById("applicationid").innerHTML = data;
		});
		request.fail(function( jqXHR, textStatus ) {
  			alert( "Sorry, the server encountered an error." );
		});
	}
	else
	{
		document.getElementById("submitBtn").disabled = false;
	}
		
}

function goCollapse()
{
var coll = document.getElementsByClassName("collapsible");
var i;

for (i = 0; i < coll.length; i++) {
  coll[i].addEventListener("click", function() {
    this.classList.toggle("active");
    var content = this.nextElementSibling;
    if (content.style.display === "block") {
      content.style.display = "none";
    } else {
      content.style.display = "block";
    }
  });
}
}
goCollapse();

if (document.getElementById("userid").value == "")
{
	//Need to take user to log in page appropriate for them.
	if (window.location.href.includes("egovlink.com"))
	{
		window.location = "../user_login.asp";
	}
	else
	{
		window.location = "/login/?route=permitapp";

	}
}
else
{
	jQuery("#permitformdiv").show();
	jQuery("#permitAppFrm").show();
}

  		jQuery( function() {
    			jQuery( ".datepicker" ).datepicker({
      			changeMonth: true,
      			changeYear: true
    			});
  		} );
