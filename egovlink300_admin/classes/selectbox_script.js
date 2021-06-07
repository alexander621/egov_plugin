function addToList(listField, newText, newValue) {
   if ( ( newValue == "" ) || ( newText == "" ) ) {
      alert("You cannot add blank values!");
   } else {
      var len = listField.length++; // Increase the size of list and return the size
      listField.options[len].value = newValue;
      listField.options[len].text = newText;
      listField.selectedIndex = len; // Highlight the one just entered (shows the user that it was entered)
   } // Ends the check to see if the value entered on the form is empty
}

function fnMoveItem(fromlist) {
	alert(fromlist.options[fromlist.selectedIndex].value);
	alert(fromlist.options[fromlist.selectedIndex].text);
}

function removeFromList(listField) {
   if ( listField.length == -1) {  // If the list is empty
      alert("There are no values which can be removed!");
   } else {
      var selected = listField.selectedIndex;
      if (selected == -1) {
         alert("You must select an entry to be removed!");
      } else {  // Build arrays with the text and values to remain
         var replaceTextArray = new Array(listField.length-1);
         var replaceValueArray = new Array(listField.length-1);
         for (var i = 0; i < listField.length; i++) {
            // Put everything except the selected one into the array
            if ( i < selected) { replaceTextArray[i] = listField.options[i].text; }
            if ( i > selected ) { replaceTextArray[i-1] = listField.options[i].text; }
            if ( i < selected) { replaceValueArray[i] = listField.options[i].value; }
            if ( i > selected ) { replaceValueArray[i-1] = listField.options[i].value; }
         }
         listField.length = replaceTextArray.length;  // Shorten the input list
         for (i = 0; i < replaceTextArray.length; i++) { // Put the array back into the list
            listField.options[i].value = replaceValueArray[i];
            listField.options[i].text = replaceTextArray[i];
         }
		BuildList(listField);
      } // Ends the check to make sure something was selected
   } // Ends the check for there being none in the list
}


function moveUpList(listField) {
   if ( listField.length == -1) {  // If the list is empty
      alert("There are no values which can be moved!");
   } else {
      var selected = listField.selectedIndex;
      if (selected == -1) {
         alert("You must select an entry to be moved!");
      } else {  // Something is selected 
         if ( listField.length == 0 ) {  // If there's only one in the list
            alert("There is only one entry!\nThe one entry will remain in place.");
         } else {  // There's more than one in the list, rearrange the list order
            if ( selected == 0 ) {
               alert("The first entry in the list cannot be moved up.");
            } else {
               // Get the text/value of the one directly above the hightlighted entry as
               // well as the highlighted entry; then flip them
               var moveText1 = listField[selected-1].text;
               var moveText2 = listField[selected].text;
               var moveValue1 = listField[selected-1].value;
               var moveValue2 = listField[selected].value;
               listField[selected].text = moveText1;
               listField[selected].value = moveValue1;
               listField[selected-1].text = moveText2;
               listField[selected-1].value = moveValue2;
               listField.selectedIndex = selected-1; // Select the one that was selected before
			   BuildList(listField);
            }  // Ends the check for selecting one which can be moved
         }  // Ends the check for there only being one in the list to begin with
      }  // Ends the check for there being something selected
   }  // Ends the check for there being none in the list
}


function moveDownList(listField) {
   if ( listField.length == -1) {  // If the list is empty
      alert("There are no values which can be moved!");
   } else {
      var selected = listField.selectedIndex;
      if (selected == -1) {
         alert("You must select an entry to be moved!");
      } else {  // Something is selected 
         if ( listField.length == 0 ) {  // If there's only one in the list
            alert("There is only one entry!\nThe one entry will remain in place.");
         } else {  // There's more than one in the list, rearrange the list order
            if ( selected == listField.length-1 ) {
               alert("The last entry in the list cannot be moved down.");
            } else {
               // Get the text/value of the one directly below the hightlighted entry as
               // well as the highlighted entry; then flip them
               var moveText1 = listField[selected+1].text;
               var moveText2 = listField[selected].text;
               var moveValue1 = listField[selected+1].value;
               var moveValue2 = listField[selected].value;
               listField[selected].text = moveText1;
               listField[selected].value = moveValue1;
               listField[selected+1].text = moveText2;
               listField[selected+1].value = moveValue2;
               listField.selectedIndex = selected+1;// Select the one that was selected before
			   BuildList(listField);
            }  // Ends the check for selecting one which can be moved
         }  // Ends the check for there only being one in the list to begin with
      }  // Ends the check for there being something selected
   }  // Ends the check for there being none in the list
}


function BuildList(sellist){
	var slist;
	slist = " ";
	for (var i = 0; i < sellist.length; i++) {
		slist = slist + sellist.options[i].text + ',';
	}
}
