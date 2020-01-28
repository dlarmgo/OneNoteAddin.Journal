/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
    if (info.host === Office.HostType.OneNote) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run2;
    }
});

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
function logParagraph(context: OneNote.RequestContext, pg: OneNote.Paragraph) {

}

async function addLogToParagraph(context: OneNote.RequestContext, pg: OneNote.Paragraph, txt: string) {
    console.log("Adding text! " + txt);
    var time = new Date().toLocaleTimeString();
    var day = new Date().toDateString();
    var date = new Date().toLocaleDateString();

    var _text: string;
    context.load(pg);
    await context.sync();
    if (pg.type == "RichText") {
        var richText = pg.richText;
        context.load(richText);
        await context.sync();
        _text = richText.text;
    }
    pg.insertHtmlAsSibling("Before", "UUU");

    var dayVals = day.split(" ");
    var dayStr = dayVals[2] + " " + dayVals[1] + ", " + dayVals[0];

    if (dayStr != _text) {
        pg.insertHtmlAsSibling(OneNote.InsertLocation.before, "<b>" + dayStr + "</b>" );
        pg.insertHtmlAsSibling(OneNote.InsertLocation.before, "&emsp;<sup><i>" + time + "</i></sup>: <b>" + txt + "</b>" );
    } else {
        pg.insertHtmlAsSibling(OneNote.InsertLocation.after, "&emsp;<sup><i>" + time + "</i></sup>: <b>" + txt + "</b>" );
    }
    await context.sync();
}


async function addLogToCell(context: OneNote.RequestContext, cell: OneNote.TableCell, txt: string) {
    var time = new Date().toLocaleTimeString();
    var day = new Date().toDateString();
    var date = new Date().toLocaleDateString();
    var richText;
    var cellParagraphs = cell.paragraphs;
    context.load(cellParagraphs);
    await context.sync();
    var pg = cellParagraphs.items[0];
    context.load(pg);
    await context.sync();
    pg.insertRichTextAsSibling("Before", "ZZZ");
    await context.sync();
}


async function addLogToTable(context: OneNote.RequestContext, table: OneNote.Table, txt: string) {
    var time = new Date().toLocaleTimeString();
    var day = new Date().toDateString();
    var date = new Date().toLocaleDateString();

    var dayVals = day.split(" ");
    var dayStr = dayVals[2] + " " + dayVals[1] + ", " + dayVals[0];

    context.load(table);
    await context.sync();
    
    var rows = table.rows;
    context.load(rows);
    await context.sync();



    var firstRow = rows.items[0];
    context.load(firstRow);
    await context.sync();

    var cells = firstRow.cells;
    context.load(cells);
    await context.sync();

    var cell = cells.items[0];
    context.load(cell);
    await context.sync();

    var cellPgs = cell.paragraphs;
    context.load(cellPgs);
    await context.sync();

    var cellPg = cellPgs.items[0]; 
    context.load(cellPg);
    await context.sync();
    
    if (cellPg.type == "RichText") {
        var cellRt = cellPg.richText;
        context.load(cellRt);
        await context.sync();

        console.log("cellRt.text: " + cellRt.text);
 
        if (dayStr != cellRt.text) {
        console.log("cellRt.text new date: " + cellRt.text);
            table.insertRow(0, [""]);
            await context.sync();
            context.load(table);
            await context.sync();

            var newRows = table.rows;
            context.load(newRows);
            await context.sync();

            var newRow = newRows.items[0];
            context.load(newRow);
            await context.sync();

            var newCells = newRow.cells;
            context.load(newCells);
            await context.sync();

            var newCell = newCells.items[0];
            context.load(newCell);
            await context.sync();

            var newCellPgs = newCell.paragraphs;
            context.load(newCellPgs);
            await context.sync();

            var newCellPg = newCellPgs.items[0]; 
            context.load(newCellPg);
            await context.sync();

            newCellPg.insertHtmlAsSibling(OneNote.InsertLocation.before, "<b>" + dayStr + "</b>" );
            newCellPg.insertHtmlAsSibling(OneNote.InsertLocation.before, "&emsp;<sup><i>" + time + "</i></sup>: <b>" + txt + "</b>" );
        } else {
        console.log("cellRt.text same date: " + cellRt.text);
            cellPg.insertHtmlAsSibling(OneNote.InsertLocation.after, "&emsp;<sup><i>" + time + "</i></sup>: <b>" + txt + "</b>" );
        }
    }
}

export async function run2() {
    try {
        await OneNote.run(async ctx => {
            console.log("HI");
            var pageContents = ctx.application.getActivePage().contents;
            ctx.load(pageContents);
            await ctx.sync();
            console.log("Contents number: " + pageContents.count);
            var pageItems = pageContents.items;
            for (var i = 0; i < pageContents.count; i++) {
                var pageItem = pageContents.items[i];
                ctx.load(pageItem);
                await ctx.sync();
                if (pageItem.type == "Outline") {
                    var outline = pageItem.outline;
                    ctx.load(outline);
                    await ctx.sync();
                    var paragraphs = outline.paragraphs;
                    ctx.load(paragraphs);
                    await ctx.sync();
                    var item = paragraphs.items[0];
                    ctx.load(item);
                    await ctx.sync();
                    if (item.type == "RichText") {
                        var rt = item.richText;
                        ctx.load(rt);
                        await ctx.sync();
                        var text = rt.text;
                        console.log("RichText text: " + text);
                        text = text + " Addded .";
                        console.log("RichText text: " + text);
                        pageItem.load('outline,outline/id,outline/paragraphs,outline/paragraphs/items,outline/paragraphs/items/richText,outline/paragraphs/items/richText/text')
                        await ctx.sync();

                        //addLogToParagraph(ctx, item, "AAA");

                        console.log("VALUE: " + pageItem.outline.paragraphs.items[0].richText.text);
                        var txt = pageItem.outline.paragraphs.items[0].richText.text;
                        txt += "asasa";
                        var html = rt.getHtml(); 
                        await ctx.sync();
                        console.log("VALUE: " + pageItem.outline.paragraphs.items[0].richText.text);
                        //console.log("pageItem length: " + pageItem.outline.paragraphs.getItemAt(0).richText.text);
                    }
                }
            }

            ctx.load(pageContents);
            await ctx.sync();
            var historyTable: OneNote.Table;
            for (var i = 0; i < pageContents.items.length; i++) {
                if (pageContents.items[i].type == "Outline") {
                    var outline = pageContents.items[i].outline;
                    ctx.load(outline);
                    await ctx.sync();
                    var paragraphs = outline.paragraphs;
                    ctx.load(paragraphs);
                    await ctx.sync();
                    if (paragraphs.items[0].type == "Table") {
                        historyTable = paragraphs.items[0].table;
                        break;
                    }
                }
            }

            ctx.load(historyTable);
            await ctx.sync();
            console.log("Table found! " + historyTable.id);

            await addLogToTable(ctx, historyTable, "ZZZ");



        });
    } catch (error) {
        console.log("Error" + error);
    }
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

                                    if (i_row == 0 && cellInRow  == 0) {
                                        for (var it = 0; it < pgArray.length; it++) {
                                            console.log("**************************************************************");
                                            console.log("** before " + paragraphsInCell.count);
                                            var copyItem = pgArray[it];
                                            ctx.load(copyItem);
                                            await ctx.sync();
                                            if (copyItem.type == "RichText") {
                                            console.log("** items length before " + paragraphsInCell.items.length);
                                                paragraphsInCell.items.unshift(copyItem);
                                                //ctx.load(paragraphsInCell);
                                                endCell.set(endCell);
                                                await ctx.sync();
                                            console.log("** items length after " + paragraphsInCell.items.length);


                                            }
                                            console.log("** after " + paragraphsInCell.count);
                                        }

                                    }

                                    for (var i_paragraphsInCell = 0; i_paragraphsInCell < paragraphsInCell.count; i_paragraphsInCell++) {
                                        var itemsInCell = paragraphsInCell.items;

                                        console.log("paragraph type: " + itemsInCell[i_paragraphsInCell].type);


                                        var RichTextInParagraphInCell = itemsInCell[i_paragraphsInCell].richText;
                                        ctx.load(RichTextInParagraphInCell);
                                        await ctx.sync();
                                        var cellRT = RichTextInParagraphInCell.getHtml();
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
