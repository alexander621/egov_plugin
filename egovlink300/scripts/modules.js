<!--
    var view;
    var opacity;

    function tmptoggleDisplay( obj, viewName ) {
      view = eval("document.all." + viewName);

      if (view.style.display == "none") {
        opacity = 0;
        view.style.display = "";
        setTimeout("fadeUp()", 25);
        obj.src = "../images/arrow_collapse.jpg";
      }
      else {
        opacity = 100;
        fadeDown();
        obj.src = "../images/arrow_expand.jpg";
      }
    }

    function toggleDisplay( obj, viewName ) {
      view = eval("document.all." + viewName);
      
      if (view.style.display == "none") {
        opacity = 0;
        view.style.display = "";
        setTimeout("fadeUp()", 25);
        obj.src = "images/arrow_collapse.jpg";
      }
      else {
        opacity = 100;
        fadeDown();
        obj.src = "images/arrow_expand.jpg";
      }
    }

    function fadeUp() {
      if (opacity <= 100) {
        opacity += 20;
        view.style.filter = "alpha(opacity:" + opacity + ")";
        setTimeout("fadeUp()", 25);
      }
    }

    function fadeDown() {
      if (opacity > 0) {
        opacity -= 20;
        view.style.filter = "alpha(opacity:" + opacity + ")";
        setTimeout("fadeDown()", 25);
      }
      else {
        view.style.display = "none";
      }
    }

    function doCalendar() {
      w = (screen.width - 600)/2;
      h = (screen.height - 550)/2;
      eval('window.open("events/calendar.asp?p=1", "_calendar", "width=600,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }

//This function checks to the field length
//The ability to display the remaining characters is available.  To do so:
//  1. create a hidden field with an ID and NAME of "control_field".  
//     Set the maxlength on this field to the maximum length of the biggest field +1 that you will be checking.
//  2. create a <span> tab with an ID of "message_char_cnt"
//  3. pass in field ID (document.getElementById(...)) to evaluate against the control field

function checkFieldLength(p_value,p_limit,p_display_cnt,p_field_id) {
  lcl_length = p_value.length;
  if(lcl_length <= p_limit) {
     if(p_display_cnt=="Y") {
		document.getElementById('message_char_cnt').innerHTML = p_limit + " character limit.  Characters remaining: " + (p_limit - lcl_length);
	 }
  } else {
     p_field_id.value = document.getElementById("control_field").value.substr(0,p_limit);
     alert("Cannot exceed " + p_limit + " characters.");
  }
}



//These two functions handle the row highlight and un-highlight for result lists when the mouse cursor moves over and off a record
function mouseOverRow( oRow ) {
  oRow.style.backgroundColor='silver';
  oRow.style.cursor='pointer';
}

function mouseOutRow( oRow ) {	
  oRow.style.backgroundColor='';
  oRow.style.cursor='';
}


  //-->