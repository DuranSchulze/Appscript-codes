/**
 * Returns the next team member email from a round-robin list in Map Sheet.
 */
function getNextTeamMember() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
  if (!mapSheet) return null;

  const data = mapSheet.getDataRange().getValues();
  const members = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'TeamMember' && data[i][2]) {
      members.push(data[i][2]); // email is in Value column
    }
  }
  if (members.length === 0) return null;

  const props = PropertiesService.getScriptProperties();
  const lastIndex = parseInt(props.getProperty('ROUND_ROBIN_INDEX') || '0');
  const nextIndex = (lastIndex + 1) % members.length;
  props.setProperty('ROUND_ROBIN_INDEX', nextIndex.toString());
  return members[nextIndex];
}
