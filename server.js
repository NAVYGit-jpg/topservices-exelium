const express = require('express');
const XLSX    = require('xlsx');
const path    = require('path');

const app   = express();
const PORT  = 3000;
const EXCEL = path.join(__dirname, 'input', 'devis.xlsx');

app.use(express.json());
app.use(express.static(__dirname));

app.post('/submit-devis', (req, res) => {
  const { nom, entreprise, telephone, email, service, message } = req.body;

  if (!nom || !telephone || !message) {
    return res.status(400).json({ ok: false, error: 'Champs requis manquants.' });
  }

  const date = new Date().toLocaleString('fr-FR', {
    timeZone: 'Africa/Abidjan',
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit'
  });

  try {
    const wb = XLSX.readFile(EXCEL);
    const ws = wb.Sheets[wb.SheetNames[0]];

    // Ordre des colonnes : A=Date  B=Nom  C=Entreprise  D=Telephone  E=Email  F=Service recherché  G=message
    XLSX.utils.sheet_add_aoa(ws, [[
      date,
      nom        || '',
      entreprise || '',
      telephone  || '',
      email      || '',
      service    || '',
      message    || ''
    ]], { origin: -1 });

    XLSX.writeFile(wb, EXCEL);
    console.log(`[devis] ${date} — ${nom} (${telephone})`);
    res.json({ ok: true });
  } catch (err) {
    console.error('Erreur écriture Excel :', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`\n✓ Serveur TOP SERVICES EXELIUM démarré`);
  console.log(`  → http://localhost:${PORT}\n`);
});
