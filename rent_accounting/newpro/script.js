window.onscroll = function() { toggleSticky() };

const navbar = document.getElementById("navbar");
const stickyPoint = navbar.offsetTop;

function toggleSticky() {
  if (window.pageYOffset > stickyPoint) {
    navbar.classList.add("sticky");
  } else {
    navbar.classList.remove("sticky");
  }
}