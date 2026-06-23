const fs = require('fs');
const path = require('path');
const ExcelJS = require('./client/node_modules/exceljs');

const inputPath = path.resolve('auction-I4HQQ8-1782148111726_backup.json');
const raw = fs.readFileSync(inputPath, 'utf8');
const data = JSON.parse(raw);
const state = data.state || {};
const teams = state.teams || [];
const players = state.players || [];
const config = state.config || { pointsPerTeam: 25000 };

const COLORS = {
  header: 'FF1E3A5F',
  sold: 'FFdbeafe',
  unsold: 'FFfecaca',
  pending: 'FFfef9c3'
};

const wb = new ExcelJS.Workbook();
wb.creator = 'Auction App';
wb.created = new Date();

const addSheet = (name, plist, color) => {
  const ws = wb.addWorksheet(name);
  ws.columns = [
    { header: 'Team/Status', key: 'team', width: 22 },
    { header: 'Player', key: 'player', width: 28 },
    { header: 'Role', key: 'role', width: 18 },
    { header: 'Base Price', key: 'basePrice', width: 14 },
    { header: 'Sold Price', key: 'soldPrice', width: 14 }
  ];
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.header } };
  ws.getRow(1).alignment = { horizontal: 'center' };
  
  for (const p of plist) {
    ws.addRow({
      team: p.teamName || '',
      player: p.name,
      role: p.role,
      basePrice: Number(p.basePrice || 0),
      soldPrice: Number(p.soldPrice || 0)
    });
    const row = ws.lastRow;
    row.eachCell(c => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
    });
  }
  return ws;
};

// Sold players
const soldPlayers = players
  .filter(p => p.status === 'sold' && p.soldTo)
  .map(p => {
    const t = teams.find(tm => tm.id === p.soldTo);
    return { ...p, teamName: t?.name || p.soldTo };
  });
addSheet('Sold Players', soldPlayers, COLORS.sold);

// Unsold players
const unsoldPlayers = players
  .filter(p => p.status === 'unsold')
  .map(p => ({ ...p, teamName: 'UNSOLD' }));
addSheet('Unsold Players', unsoldPlayers, COLORS.unsold);

// Pending players
const pendingPlayers = players
  .filter(p => p.status === 'pending')
  .map(p => ({ ...p, teamName: 'PENDING' }));
addSheet('Pending Players', pendingPlayers, COLORS.pending);

// Summary sheet
const summary = wb.addWorksheet('Summary');
summary.columns = [
  { header: 'Team', key: 'team', width: 22 },
  { header: 'Players Sold', key: 'count', width: 14 },
  { header: 'Total Spent', key: 'spent', width: 16 },
  { header: 'Budget Left', key: 'budget', width: 14 },
  { header: '% Used', key: 'pct', width: 12 }
];
summary.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
summary.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.header } };
summary.getRow(1).alignment = { horizontal: 'center' };

let totalCount = 0, totalSpent = 0;
const byTeam = new Map();
for (const t of teams) {
  byTeam.set(t.id, []);
}
for (const p of soldPlayers) {
  if (!byTeam.has(p.soldTo)) byTeam.set(p.soldTo, []);
  byTeam.get(p.soldTo).push(p);
}

for (const team of teams) {
  const roster = byTeam.get(team.id) || [];
  const spent = roster.reduce((s, p) => s + Number(p.soldPrice || 0), 0);
  const pct = Math.round((spent / config.pointsPerTeam) * 100);
  totalCount += roster.length;
  totalSpent += spent;
  summary.addRow({
    team: team.name,
    count: roster.length,
    spent,
    budget: Number(team.budget || 0),
    pct: pct + '%'
  });
}

const tr = summary.addRow({ team: 'TOTAL', count: totalCount, spent: totalSpent, budget: '', pct: '' });
tr.font = { bold: true };

const outPath = path.resolve('auction-I4HQQ8-comprehensive.xlsx');
wb.xlsx.writeFile(outPath).then(() => {
  console.log('Created', outPath);
}).catch(e => {
  console.error(e);
  process.exit(1);
});
