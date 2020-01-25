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
function logTableCell(context: OneNote.RequestContext, cell: OneNote.TableCell) {

}
function logParagraph(context: OneNote.RequestContext, pg: OneNote.Paragraph) {

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
                    //var newParagraph = paragraphs.getItemAt(0);
                    //paragraphs.items.push(newParagraph);
                    for (var i_paragraph = 0; i_paragraph < paragraphs.count; i_paragraph++) {
                        var item = paragraphs.items[i_paragraph];
                        ctx.load(item);
                        await ctx.sync();
                        console.log("i_paragraphs: " + i_paragraph + " type: " + item.type);
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
                            var date = new Date().toLocaleTimeString().slice(0, -3);
                            item.insertHtmlAsSibling(OneNote.InsertLocation.before, date + ":<b>\tAAA</b>" );



                            console.log("VALUE: " + pageItem.outline.paragraphs.items[0].richText.text);
                            var txt = pageItem.outline.paragraphs.items[0].richText.text;
                            txt += "asasa";
                            var html = rt.getHtml(); 
                            await ctx.sync();
                            console.log("VALUE: " + pageItem.outline.paragraphs.items[0].richText.text);
                            //console.log("pageItem length: " + pageItem.outline.paragraphs.getItemAt(0).richText.text);
                            ctx.sync();
                        }
                        //item.delete();
                        //paragraphs.items.pop();
                        //console.log("count: " + paragraphs.count);
                        //pageItem.set(pageItem);
                        //await ctx.sync();


                    }

                    //console.log(outline.id);
                }

            }


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
