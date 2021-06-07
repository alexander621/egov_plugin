<%
Response.AddHeader "Access-Control-Allow-Origin", "*"
%>
var newDiv = document.createElement("DIV");
    newDiv.id = "permitformdiv";

    var target = document.scripts[document.scripts.length - 1];

    target.parentElement.insertBefore(newDiv, target);

    var divTarget = document.getElementById("permitformdiv")
    divTarget.style.display = "none";


    function loadXMLDoc(theURL)
    {
        if (window.XMLHttpRequest)
        {
            xmlhttp=new XMLHttpRequest();
        }
        else
        {
            xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
        }
        xmlhttp.onreadystatechange=function()
        {
            if (xmlhttp.readyState==4 && xmlhttp.status==200)
            {
                var myDiv = document.getElementById('permitformdiv')
                myDiv.innerHTML = xmlhttp.responseText;

		//Populate userid
		document.getElementById("userid").value = sessionStorage.getItem("citizenId");

		var script = document.createElement('script');
		script.src = "//www.egovlink.com/permitcity/permits/permitapp.js";
    		target.parentElement.insertBefore(script, target);

		var script2 = document.createElement('script');
		script2.src = "//www.egovlink.com/permitcity/scripts/easyform_enhance.js";
    		target.parentElement.insertBefore(script2, target);

		var script3 = document.createElement('script');
		script3.src = "//code.jquery.com/ui/1.12.1/jquery-ui.js";
    		target.parentElement.insertBefore(script3, target);


            }
        }
        xmlhttp.open("GET", theURL, false);
        xmlhttp.send();
    }



    (function() {
        loadXMLDoc('https://www.egovlink.com/permitcity/permits/permitapplication.asp?js=true');

    })();
