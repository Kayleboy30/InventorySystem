/**
 * Global Configuration & 13 Tracked Materials
 */
const SHEET_NAMES = {
  INCOMING: 'Incoming',
  OUTGOING: 'Outgoing',
  INVENTORY: 'Inventory'
};

const MATERIAL_COLUMNS = [
  { key: 'FAF', name: 'Financial Agreement Form' },
  { key: 'Passbook', name: 'Passbook' },
  { key: 'Desk Calendar', name: 'Desk Calendar' },
  { key: 'GTR', name: 'Group Treasurer Register' },
  { key: 'Wall Calendar', name: 'Wall Calendar' },
  { key: 'Guide Book', name: 'Guide Book' },
  { key: 'FAF Barmm', name: 'Financial Agreement Form (BARMM)' },
  { key: 'Passbook Barmm', name: 'Passbook (BARMM)' },
  { key: 'GTR Barmm', name: 'Group Treasurer Register (BARMM)' },
  { key: 'Poster', name: 'Poster with Acrylic' },
  { key: 'GTR New', name: 'GTR New' },
  { key: 'Enrolment', name: 'Asa Insurance Enrollment Forms' },
  { key: 'Coverage', name: 'Asa Insurance Coverage' }
];

function formatDateMonth(d) {
  if (!d) return null;
  const date = new Date(d);
  if (isNaN(date.getTime())) return null;
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const yr = String(date.getFullYear()).slice(-2);
  return `${months[date.getMonth()]} ${yr}`;
}