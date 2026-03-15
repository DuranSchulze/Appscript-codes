"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.generateTransmittalDocx = void 0;
const docx_1 = require("docx");
const exportText_1 = require("./exportText");
// Helper to convert base64 to Uint8Array safely
const base64ToUint8Array = (base64) => {
    try {
        const data = base64.includes(',') ? base64.split(',')[1] : base64;
        const binaryString = window.atob(data);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        return bytes;
    }
    catch (e) {
        console.error("Failed to convert base64 to bytes", e);
        return new Uint8Array(0);
    }
};
const FONT_FAMILY = "Arial";
const COLOR_BORDER = "CBD5E1";
const COLOR_BG_HEADER = "F8FAFC";
const COLOR_TEXT_PRIMARY = "1E293B";
const COLOR_TEXT_SECONDARY = "475569";
const COLOR_TEXT_MUTED = "94A3B8";
const SIZE_TITLE = 28;
const SIZE_TEXT = 20;
const SIZE_LABEL = 16;
const SIZE_HEADER_SMALL = 14;
/**
 * Creates a floating "Box" for signatories that can be dragged in Word.
 * Uses string literals for anchoring to avoid ESM import errors.
 */
const createFloatingSignatoryBox = (label, name, role, horizontalOffset) => {
    return new docx_1.Table({
        width: { size: 1800, type: docx_1.WidthType.DXA }, // Twips
        float: {
            horizontalAnchor: "margin", // Equivalent to TableAnchorHorizontal.MARGIN
            verticalAnchor: "text", // Equivalent to TableAnchorVertical.TEXT
            absoluteHorizontalPosition: horizontalOffset,
            absoluteVerticalPosition: 0,
        },
        borders: {
            top: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            bottom: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            left: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
        },
        rows: [
            new docx_1.TableRow({
                children: [
                    new docx_1.TableCell({
                        children: [
                            new docx_1.Paragraph({
                                children: [new docx_1.TextRun({ text: label, size: SIZE_LABEL, color: COLOR_TEXT_SECONDARY, font: FONT_FAMILY })],
                                spacing: { after: 100 }
                            }),
                            new docx_1.Paragraph({
                                children: [new docx_1.TextRun({ text: name.toUpperCase(), bold: true, size: SIZE_TEXT, font: FONT_FAMILY, color: COLOR_TEXT_PRIMARY })]
                            }),
                            role ? new docx_1.Paragraph({
                                children: [new docx_1.TextRun({ text: role.toUpperCase(), size: SIZE_HEADER_SMALL, color: COLOR_TEXT_MUTED, font: FONT_FAMILY })]
                            }) : new docx_1.Paragraph({ children: [] })
                        ],
                        margins: { top: 120, bottom: 120, left: 120, right: 120 },
                        shading: { fill: "FFFFFF" }
                    })
                ]
            })
        ]
    });
};
const generateTransmittalDocx = async (data) => {
    let logoRun = new docx_1.TextRun({ text: "[LOGO]", bold: true, size: 24, color: COLOR_TEXT_MUTED });
    if (data.sender.logoBase64) {
        try {
            const imageBytes = base64ToUint8Array(data.sender.logoBase64);
            if (imageBytes.length > 0) {
                logoRun = new docx_1.ImageRun({
                    data: imageBytes,
                    transformation: { width: 150, height: 60 }
                });
            }
        }
        catch (e) {
            console.warn("Logo add failed", e);
        }
    }
    const headerTable = new docx_1.Table({
        width: { size: 5000, type: docx_1.WidthType.PERCENTAGE },
        borders: { top: { style: docx_1.BorderStyle.NONE }, bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE }, right: { style: docx_1.BorderStyle.NONE }, insideVertical: { style: docx_1.BorderStyle.NONE }, insideHorizontal: { style: docx_1.BorderStyle.NONE } },
        rows: [
            new docx_1.TableRow({
                children: [
                    new docx_1.TableCell({
                        children: [new docx_1.Paragraph({ children: [logoRun] })],
                        width: { size: 2000, type: docx_1.WidthType.PERCENTAGE },
                        verticalAlign: docx_1.VerticalAlign.TOP
                    }),
                    new docx_1.TableCell({
                        children: [
                            new docx_1.Paragraph({ text: `Telephone: ${data.sender.telephone}`, alignment: docx_1.AlignmentType.RIGHT, style: "HeaderSmall" }),
                            new docx_1.Paragraph({ text: `Mobile: ${data.sender.mobile}`, alignment: docx_1.AlignmentType.RIGHT, style: "HeaderSmall" }),
                            new docx_1.Paragraph({ text: `Email: ${data.sender.email}`, alignment: docx_1.AlignmentType.RIGHT, style: "HeaderSmall" }),
                            new docx_1.Paragraph({ text: data.sender.addressLine1, alignment: docx_1.AlignmentType.RIGHT, style: "HeaderSmall" }),
                            new docx_1.Paragraph({ text: data.sender.addressLine2, alignment: docx_1.AlignmentType.RIGHT, style: "HeaderSmall" }),
                            new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: data.sender.website, color: "E94E1B", bold: true, font: FONT_FAMILY, size: SIZE_HEADER_SMALL })], alignment: docx_1.AlignmentType.RIGHT }),
                        ],
                        width: { size: 3000, type: docx_1.WidthType.PERCENTAGE },
                        verticalAlign: docx_1.VerticalAlign.TOP
                    })
                ]
            })
        ]
    });
    const title = new docx_1.Paragraph({
        children: [new docx_1.TextRun({ text: "TRANSMITTAL FORM", font: FONT_FAMILY, bold: true, size: SIZE_TITLE, color: "0F172A" })],
        alignment: docx_1.AlignmentType.CENTER,
        spacing: { before: 300, after: 300 }
    });
    const leftData = [
        { label: 'To:', value: data.recipient.to },
        { label: 'Company:', value: data.recipient.company },
        { label: 'Attention:', value: data.recipient.attention },
        { label: 'Address:', value: data.recipient.address },
        { label: 'Contact No:', value: data.recipient.contactNumber },
        { label: 'Email:', value: data.recipient.email }
    ];
    const rightData = [
        { label: 'Project Name:', value: data.project.projectName },
        { label: 'Project No:', value: data.project.projectNumber },
        { label: 'Engagement Ref #:', value: data.project.engagementRef },
        { label: 'Purpose:', value: data.project.purpose },
        { label: 'Transmittal No:', value: data.project.transmittalNumber },
        { label: 'Department:', value: data.project.department },
        { label: 'Date:', value: data.project.date },
        { label: 'Time Generated:', value: data.project.timeGenerated }
    ];
    const maxMetaRows = Math.max(leftData.length, rightData.length);
    const metaTableRows = [];
    const cellMargin = { top: 120, bottom: 120, left: 100, right: 100 };
    for (let i = 0; i < maxMetaRows; i++) {
        const left = leftData[i];
        const right = rightData[i];
        const cells = [];
        if (left) {
            cells.push(new docx_1.TableCell({
                children: [
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({
                                text: left.label,
                                bold: true,
                                size: SIZE_LABEL,
                                color: COLOR_TEXT_SECONDARY,
                                font: FONT_FAMILY,
                            }),
                        ],
                    }),
                ],
                shading: { fill: COLOR_BG_HEADER },
                width: { size: 750, type: docx_1.WidthType.PERCENTAGE },
                verticalAlign: docx_1.VerticalAlign.CENTER,
                margins: cellMargin,
                borders: { right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER } }
            }));
            cells.push(new docx_1.TableCell({
                children: [
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({
                                text: left.value,
                                size: SIZE_TEXT,
                                color: COLOR_TEXT_PRIMARY,
                                font: FONT_FAMILY,
                                bold: true,
                            }),
                        ],
                    }),
                ],
                width: { size: 1750, type: docx_1.WidthType.PERCENTAGE },
                verticalAlign: docx_1.VerticalAlign.CENTER,
                margins: cellMargin,
                borders: { right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER } }
            }));
        }
        else {
            cells.push(new docx_1.TableCell({ children: [], width: { size: 2500, type: docx_1.WidthType.PERCENTAGE }, columnSpan: 2, borders: { bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE }, right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER } } }));
        }
        if (right) {
            cells.push(new docx_1.TableCell({
                children: [
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({
                                text: right.label,
                                bold: true,
                                size: SIZE_LABEL,
                                color: COLOR_TEXT_SECONDARY,
                                font: FONT_FAMILY,
                            }),
                        ],
                    }),
                ],
                shading: { fill: COLOR_BG_HEADER },
                width: { size: 875, type: docx_1.WidthType.PERCENTAGE },
                verticalAlign: docx_1.VerticalAlign.CENTER,
                margins: cellMargin,
                borders: { right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER } }
            }));
            cells.push(new docx_1.TableCell({
                children: [
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({
                                text: right.value,
                                size: SIZE_TEXT,
                                color: COLOR_TEXT_PRIMARY,
                                font: FONT_FAMILY,
                            }),
                        ],
                    }),
                ],
                width: { size: 1625, type: docx_1.WidthType.PERCENTAGE },
                verticalAlign: docx_1.VerticalAlign.CENTER,
                margins: cellMargin
            }));
        }
        else {
            cells.push(new docx_1.TableCell({ children: [], width: { size: 2500, type: docx_1.WidthType.PERCENTAGE }, columnSpan: 2, borders: { bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE }, right: { style: docx_1.BorderStyle.NONE } } }));
        }
        metaTableRows.push(new docx_1.TableRow({ children: cells }));
    }
    const metaTable = new docx_1.Table({
        width: { size: 5000, type: docx_1.WidthType.PERCENTAGE },
        borders: {
            top: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            bottom: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            left: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER }
        },
        rows: metaTableRows
    });
    const tableHeader = new docx_1.TableRow({
        children: [
            createHeaderCell("No. of Items", 10),
            createHeaderCell("QTY", 10),
            createHeaderCell("Document # / Certificate #", 25),
            createHeaderCell("Description", 30),
            createHeaderCell("Remarks", 25),
        ],
        tableHeader: true
    });
    const itemRows = data.items.map(item => new docx_1.TableRow({
        children: [
            createItemCell(item.noOfItems, 10, docx_1.AlignmentType.CENTER),
            createItemCell(item.qty, 10, docx_1.AlignmentType.CENTER),
            createItemCell(item.documentNumber, 25),
            createItemCell(item.description, 30),
            createItemCell(item.remarks, 25),
        ]
    }));
    if (itemRows.length === 0) {
        itemRows.push(new docx_1.TableRow({
            children: [
                new docx_1.TableCell({
                    columnSpan: 5,
                    children: [
                        new docx_1.Paragraph({
                            children: [
                                new docx_1.TextRun({
                                    text: "No items listed.",
                                    italics: true,
                                    color: COLOR_TEXT_MUTED,
                                    size: SIZE_TEXT,
                                    font: FONT_FAMILY,
                                }),
                            ],
                            alignment: docx_1.AlignmentType.CENTER,
                        }),
                    ],
                    margins: { top: 200, bottom: 200, left: 100, right: 100 },
                })
            ]
        }));
    }
    const itemsTable = new docx_1.Table({
        width: { size: 5000, type: docx_1.WidthType.PERCENTAGE },
        rows: [tableHeader, ...itemRows],
        borders: {
            top: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            bottom: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            left: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER }
        }
    });
    // --- MOVABLE SIGNATORIES ---
    const preparedByBox = createFloatingSignatoryBox("Prepared by:", data.signatories.preparedBy, data.signatories.preparedByRole, 0);
    const notedByBox = createFloatingSignatoryBox("Noted by:", data.signatories.notedBy, data.signatories.notedByRole, 2600);
    const timeReleasedBox = createFloatingSignatoryBox("Time Released:", data.signatories.timeReleased, null, 5200);
    const checkChar = (checked) => checked ? "☒" : "☐";
    const transParagraph = new docx_1.Paragraph({
        children: [
            new docx_1.TextRun({ text: "Transmitted via:  ", bold: true, size: SIZE_LABEL, font: FONT_FAMILY, color: "000000" }),
            new docx_1.TextRun({ text: `${checkChar(data.transmissionMethod.personalDelivery)} Personal Delivery    `, size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
            new docx_1.TextRun({ text: `${checkChar(data.transmissionMethod.pickUp)} Pick-up    `, size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
            new docx_1.TextRun({ text: `${checkChar(data.transmissionMethod.grabLalamove)} Courier App    `, size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
            new docx_1.TextRun({ text: `${checkChar(data.transmissionMethod.registeredMail)} Registered Mail`, size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
        ],
        spacing: { before: 200, after: 100 },
        border: { top: { style: docx_1.BorderStyle.SINGLE, space: 5, color: "EEEEEE" }, bottom: { style: docx_1.BorderStyle.SINGLE, space: 5, color: "EEEEEE" }, left: { style: docx_1.BorderStyle.SINGLE, space: 5, color: "EEEEEE" }, right: { style: docx_1.BorderStyle.SINGLE, space: 5, color: "EEEEEE" } },
        shading: { fill: COLOR_BG_HEADER }
    });
    const notesParagraph = data.notes ? [
        new docx_1.Paragraph({
            children: [
                new docx_1.TextRun({ text: "Notes / Instructions:", bold: true, size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
            ],
            spacing: { before: 200, after: 50 }
        }),
        new docx_1.Table({
            width: { size: 5000, type: docx_1.WidthType.PERCENTAGE },
            borders: {
                top: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
                bottom: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
                left: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
                right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            },
            rows: [
                new docx_1.TableRow({
                    children: [
                        new docx_1.TableCell({
                            children: [
                                new docx_1.Paragraph({
                                    children: [
                                        new docx_1.TextRun({
                                            text: data.notes,
                                            size: SIZE_TEXT,
                                            font: FONT_FAMILY,
                                            color: COLOR_TEXT_PRIMARY,
                                        }),
                                    ],
                                }),
                            ],
                            margins: { top: 100, bottom: 100, left: 100, right: 100 }
                        })
                    ]
                })
            ]
        })
    ] : [];
    const createRecvCell = (label, value) => {
        return new docx_1.TableCell({
            children: [
                new docx_1.Paragraph({
                    children: [new docx_1.TextRun({ text: label, bold: true, size: SIZE_LABEL, color: COLOR_TEXT_SECONDARY, font: FONT_FAMILY })]
                }),
                new docx_1.Paragraph({
                    children: [new docx_1.TextRun({ text: value || " ", size: SIZE_TEXT, font: FONT_FAMILY, color: COLOR_TEXT_PRIMARY, bold: true })],
                    border: { bottom: { style: docx_1.BorderStyle.SINGLE, size: 4, color: COLOR_BORDER } },
                    spacing: { before: 50 }
                })
            ],
            margins: { top: 120, bottom: 120, left: 100, right: 100 }
        });
    };
    const receivedByTable = new docx_1.Table({
        width: { size: 5000, type: docx_1.WidthType.PERCENTAGE },
        borders: {
            top: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            bottom: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            left: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            insideVertical: { style: docx_1.BorderStyle.NONE },
            insideHorizontal: { style: docx_1.BorderStyle.NONE }
        },
        rows: [
            new docx_1.TableRow({ children: [createRecvCell("Received by:", data.receivedBy.name), createRecvCell("Date Received:", data.receivedBy.date)] }),
            new docx_1.TableRow({ children: [createRecvCell("Time Received:", data.receivedBy.time), createRecvCell("Remarks:", data.receivedBy.remarks)] })
        ]
    });
    const disclaimer = new docx_1.Paragraph({
        children: [
            new docx_1.TextRun({
                text: data.footerNotes.disclaimer,
                italics: true,
                size: 14,
                color: COLOR_TEXT_MUTED,
                font: FONT_FAMILY,
            }),
        ],
        alignment: docx_1.AlignmentType.CENTER,
        spacing: { before: 200 }
    });
    const doc = new docx_1.Document({
        styles: {
            paragraphStyles: [
                {
                    id: "HeaderSmall",
                    name: "Header Small",
                    run: { font: FONT_FAMILY, size: SIZE_HEADER_SMALL, color: COLOR_TEXT_SECONDARY }
                }
            ]
        },
        sections: [{
                properties: {
                    page: {
                        margin: {
                            top: (0, docx_1.convertInchesToTwip)(1.0),
                            bottom: (0, docx_1.convertInchesToTwip)(1.0),
                            left: (0, docx_1.convertInchesToTwip)(1.0),
                            right: (0, docx_1.convertInchesToTwip)(1.0),
                        }
                    }
                },
                children: [
                    headerTable,
                    title,
                    metaTable,
                    new docx_1.Paragraph({ text: "", spacing: { after: 200 } }),
                    itemsTable,
                    new docx_1.Paragraph({ text: "", spacing: { after: 400 } }),
                    // Signature Row Paragraph
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({ text: "SIGNATORIES:", bold: true, size: SIZE_LABEL, color: COLOR_TEXT_MUTED })
                        ],
                        spacing: { after: 1200 }
                    }),
                    preparedByBox,
                    notedByBox,
                    timeReleasedBox,
                    new docx_1.Paragraph({ text: "", spacing: { before: 200 } }),
                    transParagraph,
                    ...notesParagraph,
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({
                                text: data.footerNotes.acknowledgement,
                                italics: true,
                                size: 16,
                                color: COLOR_TEXT_SECONDARY,
                                font: FONT_FAMILY,
                            }),
                        ],
                        spacing: { after: 100, before: 100 },
                    }),
                    receivedByTable,
                    disclaimer
                ]
            }]
    });
    return await docx_1.Packer.toBlob(doc);
};
exports.generateTransmittalDocx = generateTransmittalDocx;
function createHeaderCell(text, widthPercent) {
    return new docx_1.TableCell({
        children: [
            new docx_1.Paragraph({
                children: [
                    new docx_1.TextRun({
                        text,
                        bold: true,
                        size: SIZE_LABEL,
                        font: FONT_FAMILY,
                        color: COLOR_TEXT_SECONDARY,
                    }),
                ],
                alignment: docx_1.AlignmentType.CENTER,
            }),
        ],
        shading: { fill: COLOR_BG_HEADER },
        width: { size: widthPercent * 50, type: docx_1.WidthType.PERCENTAGE },
        verticalAlign: docx_1.VerticalAlign.CENTER,
        margins: { top: 100, bottom: 100, left: 60, right: 60 },
        borders: {
            top: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            bottom: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            left: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
        }
    });
}
function createItemCell(text, widthPercent, align = docx_1.AlignmentType.LEFT) {
    return new docx_1.TableCell({
        children: [
            new docx_1.Paragraph({
                children: [
                    new docx_1.TextRun({
                        text: (0, exportText_1.formatExportText)(text),
                        size: SIZE_TEXT,
                        font: FONT_FAMILY,
                        color: COLOR_TEXT_PRIMARY,
                    }),
                ],
                alignment: align,
                wordWrap: true,
            }),
        ],
        width: { size: widthPercent * 50, type: docx_1.WidthType.PERCENTAGE },
        verticalAlign: docx_1.VerticalAlign.CENTER,
        margins: { top: 120, bottom: 120, left: 60, right: 60 },
        borders: {
            top: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            bottom: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            left: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
        }
    });
}
