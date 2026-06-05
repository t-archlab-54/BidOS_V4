function onOpen() {
  let menu = SpreadsheetApp.getUi().createMenu(CFG.APP_NAME);

  CFG.UI_MENU.ITEMS.forEach(item => {
    if (item.separator) {
      menu = menu.addSeparator();
    } else {
      menu = menu.addItem(item.label, item.fn);
    }
  });

  menu.addToUi();
}
