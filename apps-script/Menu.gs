/**
 * Adds the "Recruiting Agent" menu to the sheet every time it opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Recruiting Agent')
    .addItem('1. Find contacts (Apollo)', 'findContacts')
    .addItem('2. Write drafts (Claude + Gmail)', 'writeDrafts')
    .addItem('3. Send approved emails', 'sendApproved')
    .addItem('4. Open LinkedIn queue', 'openLinkedInQueue')
    .addSeparator()
    .addItem('Check inbox now (sent / replies / follow-ups)', 'syncInbox')
    .addItem('Rebuild Gmail draft for selected row(s)', 'rebuildSelectedDrafts')
    .addItem('Add selected contact(s) to Referrals', 'addSelectedToReferrals')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Setup')
      .addItem('Set up tabs', 'setupTabs')
      .addItem('Set API keys', 'setApiKeys')
      .addItem('Turn on daily inbox check', 'enableDailyTrigger')
      .addItem('Turn off daily inbox check', 'disableDailyTrigger'))
    .addToUi();
}
