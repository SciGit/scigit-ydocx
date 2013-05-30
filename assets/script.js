window.onload = function() {
  function foreach(list, fn) {
    for (var i = 0; i < list.length; i++) {
      fn(list[i]);
    }
  }
  function highlight(modClass, on) {
    foreach(document.body.querySelectorAll('.' + modClass), function(elem) {
      if (on == 1) {
        elem.style.background = 'rgba(0,0,255,0.2)';
      } else {
        elem.style.background = 'rgba(0,0,255,0.1)';
      }
    });
  }
  foreach(document.body.querySelectorAll('.modify'), function(elem) {
    if (elem.classList.length >= 2) {
      elem.onmouseover = function() { highlight(elem.classList[1], 1); };
      elem.onmouseout = function() { highlight(elem.classList[1], 0); };
    }
  });
};