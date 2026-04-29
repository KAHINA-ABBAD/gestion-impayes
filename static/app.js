document.getElementById('date-reference').valueAsDate = new Date();

document.getElementById('fichier-existant').addEventListener('change', function(e) {
    document.getElementById('fichier-existant-display').value = e.target.files[0]?.name || '';
    if (e.target.files[0]) log('Fichier existant sélectionné : ' + e.target.files[0].name);
});

document.getElementById('fichier-brut').addEventListener('change', function(e) {
    document.getElementById('fichier-brut-display').value = e.target.files[0]?.name || '';
    if (e.target.files[0]) log('Export comptable sélectionné : ' + e.target.files[0].name);
});

document.getElementById('fichier-clotures').addEventListener('change', function(e) {
    document.getElementById('fichier-clotures-display').value = e.target.files[0]?.name || '';
    if (e.target.files[0]) log('Comptes clôturés sélectionnés : ' + e.target.files[0].name);
});

function selectMode(mode) {
    document.querySelectorAll('.mode-option').forEach(opt => opt.classList.remove('active'));
    event.target.closest('.mode-option').classList.add('active');
    document.getElementById('mode-' + mode).checked = true;

    const sectionExistant = document.getElementById('section-existant');
    const btn = document.getElementById('btn-generate');
    if (mode === 'actualiser') {
        sectionExistant.classList.remove('hidden');
        btn.textContent = 'Actualiser le fichier';
    } else {
        sectionExistant.classList.add('hidden');
        btn.textContent = 'Générer le fichier';
    }
    log('Mode : ' + (mode === 'creer' ? 'Création' : 'Actualisation'));
}

function log(message) {
    const logContent = document.getElementById('log-content');
    const ts = new Date().toLocaleTimeString('fr-FR');
    if (logContent.classList.contains('empty')) {
        logContent.classList.remove('empty');
        logContent.textContent = '';
    }
    logContent.textContent += `[${ts}] ${message}\n`;
    logContent.scrollTop = logContent.scrollHeight;
}

function showAlert(message, type = 'info') {
    const alertDiv = document.getElementById('alert-message');
    alertDiv.className = 'alert ' + type;
    alertDiv.textContent = message;
    alertDiv.style.display = 'block';
    if (type === 'success') setTimeout(() => { alertDiv.style.display = 'none'; }, 10000);
}

function updateStatus(text, state = 'ready') {
    document.getElementById('status-text').textContent = text;
    const dot = document.getElementById('status-dot');
    dot.className = 'status-dot';
    if (state === 'processing') dot.classList.add('processing');
    if (state === 'error') dot.classList.add('error');
}

document.getElementById('generate-form').addEventListener('submit', async function(e) {
    e.preventDefault();
    const btnGenerate = document.getElementById('btn-generate');
    const mode = document.querySelector('input[name="mode"]:checked').value;

    const fichierBrut = document.getElementById('fichier-brut').files[0];
    if (!fichierBrut) {
        showAlert('Veuillez sélectionner l\'export comptable (obligatoire)', 'error');
        log('ERREUR : Export comptable manquant');
        return;
    }
    if (mode === 'actualiser' && !document.getElementById('fichier-existant').files[0]) {
        showAlert('Veuillez sélectionner le fichier Excel existant à actualiser', 'error');
        log('ERREUR : Fichier existant manquant');
        return;
    }

    btnGenerate.disabled = true;
    btnGenerate.innerHTML = '<span class="spinner"></span>Génération en cours...';
    updateStatus('Traitement en cours...', 'processing');

    const formData = new FormData(this);
    const dateRef = document.getElementById('date-reference').value;

    log('--- ' + (mode === 'creer' ? 'CRÉATION' : 'ACTUALISATION') + ' ---');
    log('Export : ' + fichierBrut.name);
    if (mode === 'actualiser') {
        const fe = document.getElementById('fichier-existant').files[0];
        log('Fichier existant : ' + fe.name);
    }
    const fc = document.getElementById('fichier-clotures').files[0];
    if (fc) log('Comptes clôturés : ' + fc.name);
    log('Date : ' + dateRef);
    log('Envoi au serveur...');

    try {
        const response = await fetch('/api/generate', { method: 'POST', body: formData });
        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.error || 'Erreur serveur');
        }
        log('Traitement terminé — téléchargement...');
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Suivi_Impayes_${dateRef}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        log('Fichier téléchargé : Suivi_Impayes_' + dateRef + '.xlsx');
        log('--- SUCCES ---');
        showAlert('Fichier généré et téléchargé avec succès.', 'success');
        updateStatus('Fichier généré avec succès', 'ready');
    } catch (error) {
        log('--- ERREUR : ' + error.message + ' ---');
        showAlert('Erreur : ' + error.message, 'error');
        updateStatus('Erreur lors de la génération', 'error');
    } finally {
        btnGenerate.disabled = false;
        btnGenerate.textContent = mode === 'creer' ? 'Générer le fichier' : 'Actualiser le fichier';
    }
});

window.addEventListener('load', async function() {
    try {
        const response = await fetch('/api/health');
        const data = await response.json();
        if (data.status === 'ok') {
            log('Connexion au serveur établie (v' + data.version + ')');
        }
    } catch {
        log('Impossible de se connecter au serveur');
        showAlert('Serveur non disponible. Démarrez le serveur Flask.', 'error');
        updateStatus('Serveur non disponible', 'error');
    }
});
