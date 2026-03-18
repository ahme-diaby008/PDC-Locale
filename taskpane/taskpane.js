document.getElementById("myButton").addEventListener("click", async () => {
    await fetch("URL_DE_VOTRE_FLUX_POWER_AUTOMATE", {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            event: "button_clicked",
            user: "demo"
        })
    });
    alert("Action envoyée au flux !");
});
