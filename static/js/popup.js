var btnSalvar = document.getElementById("btnSalvar");
var popup = document.getElementById("popup");

// Quando o botão de salvar for clicado
btnSalvar.onclick = function() {
  // Mostrar o popup
  popup.style.display = "block";
}

// Obtenha o elemento do botão OK
var btnOK = document.getElementById("btnOK");

// Quando o botão OK for clicado
btnOK.onclick = function() {
  // Ocultar o popup
  popup.style.display = "none";
}

// Quando o usuário clicar fora do popup, feche-o
window.onclick = function(event) {
  if (event.target == popup) {
    popup.style.display = "none";
  }
}

// Quando o usuário pressionar a tecla Esc, feche o popup
document.onkeydown = function(event) {
  event = event || window.event;
  if (event.key === "Escape") {
    popup.style.display = "none";
  }
}