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

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
function logTableCell(context: OneNote.RequestContext, cell: OneNote.TableCell) {

}
function logParagraph(context: OneNote.RequestContext, pg: OneNote.Paragraph) {

}


export async function run() {
    try {
        await OneNote.run(async ctx => {
            var rtHTML;
            var activePage = ctx.application.getActivePage();
            var pageContents = activePage.contents;
            pageContents.load('count');
            await ctx.sync();
            console.log("pageContents count: " + pageContents.count);

            ctx.load(pageContents);
            await ctx.sync();
            pageContents.load('items');
            await ctx.sync();
            var pageContentsItems = pageContents.items;
            for (var i = 0; i < pageContentsItems.length; i++) {
                if (pageContentsItems[i].type == "Outline") {
                    console.log("Outline found. ID: " + pageContentsItems[i].id);
                    var paragraphs = pageContentsItems[i].outline.paragraphs;
                    ctx.load(paragraphs);
                    await ctx.sync();
                    var pgArray = paragraphs.items;
                    for (var i_pg = 0; i_pg < pgArray.length; i_pg++) {
                        if (pgArray[i_pg].type == "Table") {
                            console.log("In table...");
                            var table = pgArray[i_pg].table;
                            ctx.load(table);
                            await ctx.sync();
                            var tableRows = table.rows;
                            ctx.load(tableRows);
                            await ctx.sync();
                            console.log("Rows in table: " + tableRows.count);
                            console.log("");
                            for (var i_row = 0; i_row < tableRows.count; i_row++) {
                                var row = tableRows.items[i_row];
                                ctx.load(row);
                                await ctx.sync();
                                console.log("cellInRow: " + row.id + "::: " + row.cellCount);
                                var cellArray = row.cells;
                                ctx.load(cellArray);
                                await ctx.sync();
                                console.log("cellArray: " + cellArray.count);
                                for (var cellInRow = 0; cellInRow < cellArray.count; cellInRow++) {
                                    var endCell = cellArray.getItem(0);
                                    ctx.load(endCell);
                                    await ctx.sync();
                                    console.log("EndCell id: " + endCell.id);

                                    var paragraphsInCell = endCell.paragraphs;
                                    ctx.load(paragraphsInCell);
                                    await ctx.sync();
                                    console.log("EndCell Paragraphs count: " + paragraphsInCell.count);
                                    for (var i_paragraphsInCell = 0; i_paragraphsInCell < paragraphsInCell.count; i_paragraphsInCell++) {
                                        var itemsInCell = paragraphsInCell.items;

                                        console.log("paragraph type: " + itemsInCell[i_paragraphsInCell].type);


                                        var RichTextInParagraphInCell = itemsInCell[i_paragraphsInCell].richText;
                                        ctx.load(RichTextInParagraphInCell);
                                        await ctx.sync();
                                        var cellRT = RichTextInParagraphInCell.getHtml();
                                        RichTextInParagraphInCell.
                                        await ctx.sync();
                                        console.log("Text in cell: " + RichTextInParagraphInCell.text + ":::" + cellRT);

                                        var ParagraphInCell = itemsInCell[i_paragraphsInCell].outline;
                                        ctx.load(ParagraphInCell);
                                        await ctx.sync();
                                        console.log("Paragraph in cell: " + ParagraphInCell.id );

                                    }
                                }

                            }
                            console.log("row end");


                            console.log("Table end.");
                        }
                        if (pgArray[i_pg].type !== "RichText") {
                            console.log("Pg without RichText");
                            continue;
                        }
                        console.log("Pg found.");
                        ctx.load(pgArray[i_pg]);
                        await ctx.sync();
                        rtHTML = pgArray[i_pg].richText.getHtml();
                        console.log("pg getting html");
                        await ctx.sync();
                        console.log("RichText HTML: " + rtHTML.value);
                        console.log("Pg end.");
                    }
                }

            }
            console.log("123456789");
            return ctx.sync();
        });
    } catch (error) {
        console.log("Error" + error);
    }

}
