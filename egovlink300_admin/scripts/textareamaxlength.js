/*function : setMaxLength(  ) 

version: 1.0.0  
This function sets text below a textarea element that shows the chars used and total chars.
To use this, call this function from the body onload event initially. 
It will do it for all text areas on the page that have a maxlength attribute.

*/  

		function setMaxLength() {
			var x = document.getElementsByTagName('textarea');
			var counter = document.createElement('div');
			counter.className = 'counter';
			for (var i=0;i<x.length;i++) {
				if (x[i].getAttribute('maxlength')) {
					var counterClone = counter.cloneNode(true);
					counterClone.relatedElement = x[i];
					counterClone.innerHTML = '<span>0</span> of '+x[i].getAttribute('maxlength') + ' characters';
					x[i].parentNode.insertBefore(counterClone,x[i].nextSibling);
					x[i].relatedElement = counterClone.getElementsByTagName('span')[0];

					x[i].onkeyup = x[i].onchange = checkMaxLength;
					x[i].onkeyup();
				}
			}
		}

/*function : checkMaxLength(  ) 

version: 1.0.0  
This function checks the current length against the maxlength and enforces the limit.

*/  

		function checkMaxLength() {
			var maxLength = this.getAttribute('maxlength');
			var currentLength = this.value.length;
			if (currentLength > maxLength)
			{
				this.relatedElement.className = 'toomuch';
				this.value = this.value.substring(0, maxLength);
				currentLength = this.value.length;
				this.relatedElement.firstChild.nodeValue = currentLength;
				alert("The size limit has been reached on this field.\nNo more text can be accomodated.");
			}
			else
				this.relatedElement.className = '';
			this.relatedElement.firstChild.nodeValue = currentLength;
			// not innerHTML
		}
