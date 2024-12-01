let workbook = null;
let selectedPlayers = [];

// Charger le fichier Excel
document.getElementById('loadFileButton').addEventListener('click', () => {
    const fileInput = document.getElementById('fileInput');
    if (fileInput.files.length === 0) {
        alert('Veuillez choisir un fichier Excel');
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = e.target.result;
        workbook = XLSX.read(data, { type: 'binary' });

        loadSheetsData();
    };

    reader.readAsBinaryString(file);
});

// Charger les données des feuilles Excel
function loadSheetsData() {
    if (!workbook) return;

    const clubsSheet = workbook.Sheets['Clubs'];
    const coachesSheet = workbook.Sheets['Entraîneurs'];
    const playersSheet = workbook.Sheets['Joueurs'];
    const exercisesSheet = workbook.Sheets['Exercices'];

    const clubsData = XLSX.utils.sheet_to_json(clubsSheet, { header: 1 });
    const coachesData = XLSX.utils.sheet_to_json(coachesSheet, { header: 1 });
    const playersData = XLSX.utils.sheet_to_json(playersSheet, { header: 1 });
    const exercisesData = XLSX.utils.sheet_to_json(exercisesSheet, { header: 1 });

    // Remplir les sélections dynamiques
    populateSelectOptions(clubsData, 'clubSelect');
    populateSelectOptions(coachesData, 'coachSelect');
    populateSelectOptions(exercisesData, 'exerciseSelect');
    populateSelectOptions(playersData, 'playerSelect');
}

// Fonction pour remplir dynamiquement les options d'une liste déroulante
function populateSelectOptions(data, selectId) {
    const selectElement = document.getElementById(selectId);
    data.forEach(row => {
        const option = document.createElement('option');
        option.value = row[0]; // Prenez la première colonne comme valeur
        option.textContent = row[0]; // Affichez la première colonne dans la liste
        selectElement.appendChild(option);
    });
}

// Ajouter un autre joueur
document.getElementById('addPlayer').addEventListener('click', () => {
    const selectedPlayer = document.getElementById('playerSelect').value;
    if (selectedPlayers.includes(selectedPlayer)) {
        alert('Ce joueur a déjà été ajouté.');
        return;
    }
    
    selectedPlayers.push(selectedPlayer);
    alert(`Joueur ${selectedPlayer} ajouté.`);
});

// Valider le formulaire et générer le fichier Excel mis à jour
document.getElementById('submitForm').addEventListener('click', () => {
    const clubName = document.getElementById('clubSelect').value;
    const coachName = document.getElementById('coachSelect').value;
    const day = document.getElementById('daySelect').value;
    const exercise = document.getElementById('exerciseSelect').value;

    if (!clubName || !coachName || !day || !exercise) {
        alert('Veuillez remplir tous les champs.');
        return;
    }

    // Ajouter ces informations dans le fichier Excel
    addDataToExcel(clubName, coachName, day, exercise);

    // Exporter le fichier Excel
    XLSX.writeFile(workbook, 'adec.xlsx');

    // Envoyer le fichier par email
    sendEmailWithAttachment();
});

// Fonction pour ajouter les informations dans le fichier Excel
function addDataToExcel(club, coach, day, exercise) {
    const clubsSheet = workbook.Sheets['Clubs'];
    const newData = [[club, coach, day, exercise]];
    const currentData = XLSX.utils.sheet_to_json(clubsSheet, { header: 1 });
    currentData.push(...newData);
    const updatedSheet = XLSX.utils.aoa_to_sheet(currentData);
    workbook.Sheets['Clubs'] = updatedSheet;
}

// Fonction pour envoyer un email avec le fichier Excel attaché
function sendEmailWithAttachment() {
    const formData = new FormData();
    formData.append('email', 'juempak@gmail.com');
    formData.append('file', new Blob([XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'Club_Sport_Updated.xlsx');

    fetch('/send-email', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (response.ok) {
            alert('Email envoyé avec succès.');
        } else {
            alert('Erreur lors de l\'envoi de l\'email.');
        }
    })
    .catch(error => {
        alert('Erreur lors de l\'envoi de l\'email.');
    });
}
