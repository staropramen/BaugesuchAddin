Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        loadChecklist();
    }
});

// Initiale Checkliste
let checklist = [
    { category: "Formale Anforderungen", point: "Vollständigkeit", status: "Ja", remarks: "", relevance: "Hoch", reportText: "Vollständigkeit erfüllt" },
    { category: "Raumplanung", point: "Zonenplan Brütten", status: "Nein", remarks: "Entspricht nicht BZO", relevance: "Hoch", reportText: "Zonenplan nicht eingehalten" }
];

// Checkliste laden
function loadChecklist() {
    const table = document.getElementById("checklistTable");
    checklist.forEach(item => {
        const row = table.insertRow(-1);
        row.innerHTML = `
            <td><input type="text" class="category" value="${item.category}"></td>
            <td><input type="text" class="point" value="${item.point}"></td>
            <td><select class="status"><option ${item.status === "Ja" ? "selected" : ""}>Ja</option><option ${item.status === "Nein" ? "selected" : ""}>Nein</option><option ${item.status === "In Prüfung" ? "selected" : ""}>In Prüfung</option></select></td>
            <td><input type="text" class="remarks" value="${item.remarks}"></td>
            <td><select class="relevance"><option ${item.relevance === "Hoch" ? "selected" : ""}>Hoch</option><option ${item.relevance === "Mittel" ? "selected" : ""}>Mittel</option><option ${item.relevance === "Niedrig" ? "selected" : ""}>Niedrig</option></select></td>
        `;
    });
}

// Neue Zeile hinzufügen
function addRow() {
    const table = document.getElementById("checklistTable");
    const row = table.insertRow(-1);
    row.innerHTML = `
        <td><input type="text" class="category"></td>
        <td><input type="text" class="point"></td>
        <td><select class="status"><option>Ja</option><option>Nein</option><option>In Prüfung</option></select></td>
        <td><input type="text" class="remarks"></td>
        <td><select class="relevance"><option>Hoch</option><option>Mittel</option><option>Niedrig</option></select></td>
    `;
}

// Report generieren
function generateReport() {
    let reportText = "Baubewilligungs-Report\n\n";
    const table = document.getElementById("checklistTable");
    const rows = table.getElementsByTagName("tr");

    for (let i = 1; i < rows.length; i++) {
        const cells = rows[i].getElementsByTagName("td");
        const category = cells[0].getElementsByTagName("input")[0].value;
        const point = cells[1].getElementsByTagName("input")[0].value;
        const status = cells[2].getElementsByTagName("select")[0].value;
        const remarks = cells[3].getElementsByTagName("input")[0].value;
        const relevance = cells[4].getElementsByTagName("select")[0].value;

        if (status === "Nein" && relevance === "Hoch") {
            reportText += `${category}: ${point}\nBemerkungen: ${remarks}\n${getReportText(category, point)}\n\n`;
        }
    }

    Word.run(async (context) => {
        const body = context.document.body;
        body.insertText(reportText, "End");
        await context.sync();
    }).catch(error => console.log(error));
}

// Report-Text abrufen
function getReportText(category, point) {
    const item = checklist.find(c => c.category === category && c.point === point);
    return item ? item.reportText : "Keine spezifische Meldung";
}