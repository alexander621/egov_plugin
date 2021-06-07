	var global_valfield;	// retain valfield for timer thread
	// --------------------------------------------
	//                  setfocus
	// Delayed focus setting to get around IE bug
	// --------------------------------------------

	function setFocusDelayed()
	{
	  global_valfield.focus();
	}

	function setfocus(valfield)
	{
	  // save valfield in global variable so value retained when routine exits
	  global_valfield = valfield;
	  setTimeout( 'setFocusDelayed()', 100 );
	}
