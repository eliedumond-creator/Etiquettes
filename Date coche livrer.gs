function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const colCheckbox = 15; // colonne où se trouve la case à cocher (A = 1)
  const colDate = 18; // colonne où mettre la date (B = 2)

  if (e.range.getColumn() === colCheckbox && e.range.getRow() > 1) {
    if (e.range.isChecked()) {
      sheet.getRange(e.range.getRow(), colDate).setValue(new Date());
    } else {
      sheet.getRange(e.range.getRow(), colDate).clear();
    }
  }
}
