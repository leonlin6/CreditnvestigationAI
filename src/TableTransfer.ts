// Example of how you would create a table and add data to it

import * as fs from "fs";
import { Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow, VerticalAlignTable, TextDirection } from "docx";

export const excuteNewDoc = (arrStr) => {
const arr = JSON.parse(arrStr);
const row1 = [new TableRow({
    children: [
        new TableCell({
            children: [new Paragraph({text:'First Name',heading: HeadingLevel.HEADING_1,}), new Paragraph({})],
            verticalAlign: VerticalAlignTable.CENTER,
        }),
        new TableCell({
            children: [new Paragraph({text:'Last Name',heading: HeadingLevel.HEADING_1,}), new Paragraph({})],
            verticalAlign: VerticalAlignTable.CENTER,
        }),
        new TableCell({
            children: [new Paragraph({text:'指數',heading: HeadingLevel.HEADING_1,}), new Paragraph({})],
            verticalAlign: VerticalAlignTable.CENTER,
        }),
    ]
})]

const contentArray = arr.map((item)=>{

    const content = [
        new TableCell({
            children: [
                new Paragraph({
                    text: item.FirstName,
                    heading: HeadingLevel.HEADING_1,
                }),
            ],
        }),
        new TableCell({
            children: [
                new Paragraph({
                    text: item.LastName,
                    heading: HeadingLevel.HEADING_1,
                }),
            ],
        }),
        new TableCell({
            children: [
                new Paragraph({
                    text: item.WealthRank,
                    heading: HeadingLevel.HEADING_1,
                }),
            ],
            verticalAlign: VerticalAlignTable.CENTER,
        })
    ];
        
    return new TableRow({
        children: content
    })
});



// const row2 =[
//     new TableCell({
//         children: [
//             new Paragraph({
//                 text: "保留盈餘合計",
//                 heading: HeadingLevel.HEADING_1,
//             }),
//         ],
//     }),
//     new TableCell({
//         children: [
//             new Paragraph({
//                 text: "1,422,628",
//                 heading: HeadingLevel.HEADING_1,
//             }),
//         ],
//         verticalAlign: VerticalAlignTable.CENTER,
//     }),
    
// ];

const finalArray = row1.concat(contentArray);

const doc = new Document({
    sections: [
        {
            children: [
                new Table({
                    rows: finalArray
                    
                }),
            ],
        },
    ],
});



Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});


}
