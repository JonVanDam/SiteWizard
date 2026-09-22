const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'node_modules', 'xlsx'));

const sites = [
  'draftkings.com', 'fanduel.com', 'betmgm.com', 'caesars.com', 'pokerstars.com',
  'bet365.com', 'williamhill.com', 'ladbrokes.com', 'paddypower.com', 'skybet.com',
  'bwin.com', '888casino.com', 'partypoker.com', 'unibet.com', 'betfair.com',
  'coral.co.uk', 'betway.com', 'pointsbet.com', 'barstoolsportsbook.com', 'wynnbet.com',
  'borgataonline.com', 'hardrockbet.com', 'foxbet.com', 'betrivers.com', 'tipico.com',
  'superbook.com', 'circasports.com', 'bovada.lv', 'mybookie.ag', 'betonline.ag',
  'sportsbetting.ag', 'xbet.com', 'gtbets.eu', 'everygame.eu', 'intertops.eu',
  'betus.com.pa', 'bodog.eu', 'ignitioncasino.eu', 'slots.lv', 'cafecasino.lv',
  'lucky247.com', 'royalpanda.com', 'leovegas.com', 'casumo.com', 'mrgreen.com',
  'betsson.com', 'nordicbet.com', 'comeon.com', 'dafabet.com', 'sbobet.com',
  '12bet.com', 'm88.com', 'w88.com', 'fun88.com', '1xbet.com',
  'melbet.com', 'betwinner.com', 'parimatch.com', 'pinnacle.com', 'betvictor.com',
  'matchbook.com', 'smarkets.com', 'betdaq.com', 'skybettingandgaming.com', 'genting.com',
  'grosvenorcasinos.com', 'mecca.com', 'galacasino.com', 'virgingames.com', 'sportingbet.com',
  'neds.com.au', 'sportsbet.com.au', 'tab.com.au', 'ladbrokes.com.au', 'pointsbetaustralia.com.au',
  'betfair.com.au', 'crownresorts.com', 'starcasino.be', 'napoleongames.be', 'holland-casino.nl',
  'toto.nl', 'jackpotcity.com', 'spinpalace.com', 'europacasino.com', 'rubyfortune.com',
  'platinumplay.com', 'villento.com', 'allslotscasino.com', 'luckynugget.com', 'gunsbet.com',
  'casinoroom.com', 'casinoeuro.com', 'mrplay.com', 'karamba.com', 'rizk.com',
  'yeticasino.com', 'spinit.com', 'wildz.com', 'playamo.com', 'bitstarz.com',
];

const rows = [['URL'], ...sites.map((s) => [s])];
const sheet = XLSX.utils.aoa_to_sheet(rows);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, sheet, 'Sites');

const outPath = path.join(__dirname, 'gambling-sites-test.xlsx');
XLSX.writeFile(workbook, outPath);
console.log(`Wrote ${sites.length} rows to ${outPath}`);
