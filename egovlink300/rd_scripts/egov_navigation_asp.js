var iframesearchBox = document.getElementById("iframeclassesSearchBox")

iframesearchBox.style.display = "none";

function iframeSearch()
{
        var lcl_searchphrase = '';

            lcl_searchphrase += document.getElementById("iframetxtsearchphrase").value;

	    searchURL = 'rd_classes/class_list.aspx?keywordSearch=' + lcl_searchphrase;
	    if (window.location.pathname.split("/").length - 1 > 2)
    	    {
		searchURL = '../' + searchURL;
    	    }
	    //alert(window.location.pathname);
            location.href = searchURL;
}

function expandiframeSearchBox() {
    if (iframesearchBox.style.display == "none") {
	iframesearchBox.style.display = "";
    } else {
	setTimeout(function(){iframesearchBox.style.display = 'none'},500);
    }
}

