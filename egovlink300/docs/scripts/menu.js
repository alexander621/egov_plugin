function showmenuie5() {
  var rightedge=document.body.clientWidth-event.clientX
  var bottomedge=document.body.clientHeight-event.clientY

  if (rightedge<ie5menu.offsetWidth)
    ie5menu.style.left=document.body.scrollLeft+event.clientX-ie5menu.offsetWidth
  else
    ie5menu.style.left=document.body.scrollLeft+event.clientX

  if (bottomedge<ie5menu.offsetHeight)
    ie5menu.style.top=document.body.scrollTop+event.clientY-ie5menu.offsetHeight
  else
    ie5menu.style.top=document.body.scrollTop+event.clientY

  ie5menu.style.visibility="visible"
  return false
}

function hidemenuie5() {
  ie5menu.style.visibility="hidden"
}

function highlightie5() {
  if (event.srcElement.className=="menuitem") {
    event.srcElement.style.backgroundColor="highlight"
    event.srcElement.style.color="white"
  }
}

function lowlightie5() {
  if (event.srcElement.className=="menuitem") {
    event.srcElement.style.backgroundColor=""
    event.srcElement.style.color="black"
  }
}

function jumptoie5(){
  if (event.srcElement.className=="menuitem")
    window.location=event.srcElement.url
}