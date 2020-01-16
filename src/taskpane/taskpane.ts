/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your OneNote code here
   */
  try {
    await OneNote.run(async context => {
      var page = context.application.getActivePage();
      //var html = "<p><ol><li>Item #1</li><li>Item #4</li></ol></p>";
      //var mainTable = new OneNote.Table();
      var pageContents = page.contents;
      var mainTableParagraph ;

      context.load(pageContents);
      context.sync().then(
        function() {
          var items = pageContents.items;
          items.map(
            function(el) {
              if (el.type == "Outline") {
                //console.log("El: " + el.type);
                var pG = el.outline.paragraphs;
                context.load(pG);
                //pG.load('type');
                var newNote = "<b>Z</b>";
                context.sync().then(function() {
                  var item = pG.getItemAt(0);
                  item.load('type');
                  context.sync().then(function(){
                    console.log("Found something! " + item.type);
                    if (item.type == "Table") {
                      var mainTable = item.table;
                      mainTable.insertRow(0, ["<b>AAA</b>"]);
                      return context.sync();
                    } else {
                      console.log("found new note!" + item.type + " - " + item.richText.getHtml());
                      newNote = newNote + item.toJSON();
                    }
                  });
                });
                console.log("new note: " + newNote);
              }
            }
          );
        }
      );
      //console.log(page.contents.items.forEach(a => {return a;}));
      //console.log("AA: " + mainTableParagraph.type);
      
      //page.addOutline(40, 90 + 50, html);
      return context.sync();
    })
  } catch (error) {
    console.log("Error" + error);
  }

}
