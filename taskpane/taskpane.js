document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("myButton");
  if (!btn) { console.error("Bouton introuvable"); return; }
  btn.addEventListener("click", onClick);
});


async function onClick() {

  await fetch("https://d5a9f7d36f6be56ba17119f83f77d6.a6.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/8dac9218fabd4357aa5821373ba16569/triggers/manual/paths/invoke?api-version=1", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      event: "button_clicked",
      page: "taskpane.html",
      timestamp: new Date().toISOString()
    })
  });

  alert("Flux déclenché !");
});
``
