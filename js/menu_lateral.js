document.addEventListener("DOMContentLoaded", function () {
  fetch("menu_lateral.html")
    .then(response => {
      if (!response.ok) throw new Error("Arquivo não encontrado");
      return response.text();
    })
    .then(data => {
      document.getElementById("menuLateral").innerHTML = data;
    })
    .catch(error => {
      console.error("Erro ao carregar o menu lateral:", error);
    });
});