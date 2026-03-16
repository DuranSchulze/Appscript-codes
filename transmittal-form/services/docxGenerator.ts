import { 
    Document, 
    Packer, 
    Paragraph, 
    TextRun, 
    Table, 
    TableRow, 
    TableCell, 
    WidthType, 
    AlignmentType, 
    BorderStyle, 
    ImageRun,
    VerticalAlign,
    convertInchesToTwip
} from "docx";
import { AppData } from "../types";
import { formatExportText } from "./exportText";

// Helper to convert base64 to Uint8Array safely
const base64ToUint8Array = (base64: string) => {
    try {
        const data = base64.includes(',') ? base64.split(',')[1] : base64;
        const binaryString = window.atob(data);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        return bytes;
    } catch (e) {
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

const createWrappedParagraph = (
    text: string,
    options: {
        alignment?: AlignmentValue;
        bold?: boolean;
        size?: number;
        color?: string;
        font?: string;
        style?: string;
        spacing?: { before?: number; after?: number };
        border?: any;
    } = {},
) =>
    new Paragraph({
        children: [
            new TextRun({
                text: formatExportText(text),
                bold: options.bold,
                size: options.size,
                color: options.color,
                font: options.font ?? FONT_FAMILY,
            }),
        ],
        alignment: options.alignment,
        style: options.style,
        spacing: options.spacing,
        border: options.border,
        wordWrap: true,
    });

/**
 * Creates a floating "Box" for signatories that can be dragged in Word.
 * Uses string literals for anchoring to avoid ESM import errors.
 */
const createFloatingSignatoryBox = (label: string, name: string, role: string | null, horizontalOffset: number) => {
    return new Table({
        width: { size: 1800, type: WidthType.DXA }, // Twips
        float: {
            horizontalAnchor: "margin", // Equivalent to TableAnchorHorizontal.MARGIN
            verticalAnchor: "text",     // Equivalent to TableAnchorVertical.TEXT
            absoluteHorizontalPosition: horizontalOffset,
            absoluteVerticalPosition: 0,
        },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            left: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            right: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [
                            new Paragraph({ 
                                children: [new TextRun({ text: label, size: SIZE_LABEL, color: COLOR_TEXT_SECONDARY, font: FONT_FAMILY })],
                                spacing: { after: 100 }
                            }),
                            new Paragraph({ 
                                children: [new TextRun({ text: name.toUpperCase(), bold: true, size: SIZE_TEXT, font: FONT_FAMILY, color: COLOR_TEXT_PRIMARY })]
                            }),
                            role ? new Paragraph({ 
                                children: [new TextRun({ text: role.toUpperCase(), size: SIZE_HEADER_SMALL, color: COLOR_TEXT_MUTED, font: FONT_FAMILY })]
                            }) : new Paragraph({ children: [] })
                        ],
                        margins: { top: 120, bottom: 120, left: 120, right: 120 },
                        shading: { fill: "FFFFFF" }
                    })
                ]
            })
        ]
    });
};

export const generateTransmittalDocx = async (data: AppData): Promise<Blob> => {
    let logoRun: ImageRun | TextRun = new TextRun({ text: "[LOGO]", bold: true, size: 24, color: COLOR_TEXT_MUTED });
    if (data.sender.logoBase64) {
        try {
            const imageBytes = base64ToUint8Array(data.sender.logoBase64);
            if (imageBytes.length > 0) {
                logoRun = new ImageRun({
                    data: imageBytes,
                    transformation: { width: 150, height: 60 }
                });
            }
        } catch(e) { console.warn("Logo add failed", e); }
    }

    const headerTable = new Table({
        width: { size: 5000, type: WidthType.PERCENTAGE },
        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE } },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ children: [logoRun] })],
                        width: { size: 2000, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.TOP
                    }),
                    new TableCell({
                        children: [
                            createWrappedParagraph(`Telephone: ${data.sender.telephone}`, { alignment: AlignmentType.RIGHT, style: "HeaderSmall", size: SIZE_HEADER_SMALL, color: COLOR_TEXT_SECONDARY }),
                            createWrappedParagraph(`Mobile: ${data.sender.mobile}`, { alignment: AlignmentType.RIGHT, style: "HeaderSmall", size: SIZE_HEADER_SMALL, color: COLOR_TEXT_SECONDARY }),
                            createWrappedParagraph(`Email: ${data.sender.email}`, { alignment: AlignmentType.RIGHT, style: "HeaderSmall", size: SIZE_HEADER_SMALL, color: COLOR_TEXT_SECONDARY }),
                            createWrappedParagraph(data.sender.addressLine1, { alignment: AlignmentType.RIGHT, style: "HeaderSmall", size: SIZE_HEADER_SMALL, color: COLOR_TEXT_SECONDARY }),
                            createWrappedParagraph(data.sender.addressLine2, { alignment: AlignmentType.RIGHT, style: "HeaderSmall", size: SIZE_HEADER_SMALL, color: COLOR_TEXT_SECONDARY }),
                            createWrappedParagraph(data.sender.website, { alignment: AlignmentType.RIGHT, bold: true, size: SIZE_HEADER_SMALL, color: "E94E1B" }),
                        ],
                        width: { size: 3000, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.TOP
                    })
                ]
            })
        ]
    });

    const title = new Paragraph({
        children: [new TextRun({ text: "TRANSMITTAL FORM", font: FONT_FAMILY, bold: true, size: SIZE_TITLE, color: "0F172A" })],
        alignment: AlignmentType.CENTER,
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

    for(let i = 0; i < maxMetaRows; i++) {
        const left = leftData[i];
        const right = rightData[i];
        const cells = [];
        if (left) {
            cells.push(new TableCell({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
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
                width: { size: 750, type: WidthType.PERCENTAGE },
                verticalAlign: VerticalAlign.CENTER,
                margins: cellMargin,
                borders: { right: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER } }
            }));
            cells.push(new TableCell({
                children: [
                    createWrappedParagraph(left.value, {
                        size: SIZE_TEXT,
                        color: COLOR_TEXT_PRIMARY,
                        bold: true,
                    }),
                ],
                width: { size: 1750, type: WidthType.PERCENTAGE },
                verticalAlign: VerticalAlign.CENTER,
                margins: cellMargin,
                borders: { right: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER } }
            }));
        } else {
             cells.push(new TableCell({ children: [], width: { size: 2500, type: WidthType.PERCENTAGE }, columnSpan: 2, borders: { bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER } } }));
        }

        if (right) {
             cells.push(new TableCell({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
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
                width: { size: 875, type: WidthType.PERCENTAGE },
                verticalAlign: VerticalAlign.CENTER,
                margins: cellMargin,
                borders: { right: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER } }
            }));
             cells.push(new TableCell({
                children: [
                    createWrappedParagraph(right.value, {
                        size: SIZE_TEXT,
                        color: COLOR_TEXT_PRIMARY,
                    }),
                ],
                width: { size: 1625, type: WidthType.PERCENTAGE },
                verticalAlign: VerticalAlign.CENTER,
                margins: cellMargin
            }));
        } else {
             cells.push(new TableCell({ children: [], width: { size: 2500, type: WidthType.PERCENTAGE }, columnSpan: 2, borders: { bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } } }));
        }
        metaTableRows.push(new TableRow({ children: cells }));
    }

    const metaTable = new Table({
        width: { size: 5000, type: WidthType.PERCENTAGE },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            left: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            right: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            insideVertical: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER }
        },
        rows: metaTableRows
    });

    const tableHeader = new TableRow({
        children: [
            createHeaderCell("No. of Items", 10),
            createHeaderCell("QTY", 10),
            createHeaderCell("Document # / Certificate #", 25),
            createHeaderCell("Description", 30),
            createHeaderCell("Remarks", 25),
        ],
        tableHeader: true
    });

    const itemRows = data.items.map(item => new TableRow({
        children: [
            createItemCell(item.noOfItems, 10, AlignmentType.CENTER),
            createItemCell(item.qty, 10, AlignmentType.CENTER),
            createItemCell(item.documentNumber, 25),
            createItemCell(item.description, 30),
            createItemCell(item.remarks, 25),
        ]
    }));

    if (itemRows.length === 0) {
        itemRows.push(new TableRow({
            children: [
                new TableCell({
                    columnSpan: 5,
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "No items listed.",
                                    italics: true,
                                    color: COLOR_TEXT_MUTED,
                                    size: SIZE_TEXT,
                                    font: FONT_FAMILY,
                                }),
                            ],
                            alignment: AlignmentType.CENTER,
                        }),
                    ],
                    margins: { top: 200, bottom: 200, left: 100, right: 100 },
                })
            ]
        }));
    }

    const itemsTable = new Table({
        width: { size: 5000, type: WidthType.PERCENTAGE },
        rows: [tableHeader, ...itemRows],
        borders: {
             top: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
             bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
             left: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
             right: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
             insideVertical: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
             insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER }
        }
    });

    // --- MOVABLE SIGNATORIES ---
    const preparedByBox = createFloatingSignatoryBox("Prepared by:", data.signatories.preparedBy, data.signatories.preparedByRole, 0);
    const notedByBox = createFloatingSignatoryBox("Noted by:", data.signatories.notedBy, data.signatories.notedByRole, 2600);
    const timeReleasedBox = createFloatingSignatoryBox("Time Released:", data.signatories.timeReleased, null, 5200);

    const checkChar = (checked: boolean) => checked ? "☒" : "☐";
    const transParagraph = new Paragraph({
        children: [
            new TextRun({ text: "Transmitted via:  ", bold: true, size: SIZE_LABEL, font: FONT_FAMILY, color: "000000" }),
            new TextRun({ text: `${checkChar(data.transmissionMethod.personalDelivery)} Personal Delivery    `, size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
            new TextRun({ text: `${checkChar(data.transmissionMethod.pickUp)} Pick-up    `, size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
            new TextRun({ text: `${checkChar(data.transmissionMethod.grabLalamove)} Courier App    `, size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
            new TextRun({ text: `${checkChar(data.transmissionMethod.registeredMail)} Registered Mail`, size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
        ],
        spacing: { before: 200, after: 100 },
        border: { top: { style: BorderStyle.SINGLE, space: 5, color: "EEEEEE" }, bottom: { style: BorderStyle.SINGLE, space: 5, color: "EEEEEE" }, left: { style: BorderStyle.SINGLE, space: 5, color: "EEEEEE" }, right: { style: BorderStyle.SINGLE, space: 5, color: "EEEEEE" } },
        shading: { fill: COLOR_BG_HEADER }
    });

    const notesParagraph = data.notes ? [
        new Paragraph({
            children: [
                new TextRun({ text: "Notes / Instructions:", bold: true, size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
            ],
            spacing: { before: 200, after: 50 }
        }),
        new Table({
            width: { size: 5000, type: WidthType.PERCENTAGE },
            borders: {
                top: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
                bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
                left: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
                right: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            children: [
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: formatExportText(data.notes),
                                            size: SIZE_TEXT,
                                            font: FONT_FAMILY,
                                            color: COLOR_TEXT_PRIMARY,
                                        }),
                                    ],
                                    wordWrap: true,
                                }),
                            ],
                            margins: { top: 100, bottom: 100, left: 100, right: 100 }
                        })
                    ]
                })
            ]
        })
    ] : [];

    const createRecvCell = (label: string, value: string) => {
        return new TableCell({
            children: [
                new Paragraph({
                    children: [new TextRun({ text: label, bold: true, size: SIZE_LABEL, color: COLOR_TEXT_SECONDARY, font: FONT_FAMILY })]
                }),
                new Paragraph({
                    children: [new TextRun({ text: formatExportText(value || " "), size: SIZE_TEXT, font: FONT_FAMILY, color: COLOR_TEXT_PRIMARY, bold: true })],
                    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER } },
                    spacing: { before: 50 },
                    wordWrap: true,
                })
            ],
            margins: { top: 120, bottom: 120, left: 100, right: 100 }
        });
    }

    const receivedByTable = new Table({
        width: { size: 5000, type: WidthType.PERCENTAGE },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            left: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            right: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            insideVertical: { style: BorderStyle.NONE },
            insideHorizontal: { style: BorderStyle.NONE }
        },
        rows: [
            new TableRow({ children: [createRecvCell("Received by:", data.receivedBy.name), createRecvCell("Date Received:", data.receivedBy.date)] }),
            new TableRow({ children: [createRecvCell("Time Received:", data.receivedBy.time), createRecvCell("Remarks:", data.receivedBy.remarks)] })
        ]
    });

    const disclaimer = new Paragraph({
        children: [
            new TextRun({
                text: formatExportText(data.footerNotes.disclaimer),
                italics: true,
                size: 14,
                color: COLOR_TEXT_MUTED,
                font: FONT_FAMILY,
            }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 200 },
        wordWrap: true,
    });

    const doc = new Document({
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
                        top: convertInchesToTwip(1.0),
                        bottom: convertInchesToTwip(1.0),
                        left: convertInchesToTwip(1.0),
                        right: convertInchesToTwip(1.0),
                    }
                }
            },
            children: [
                headerTable,
                title,
                metaTable,
                new Paragraph({ text: "", spacing: { after: 200 } }),
                itemsTable,
                new Paragraph({ text: "", spacing: { after: 400 } }),
                // Signature Row Paragraph
                new Paragraph({ 
                    children: [
                        new TextRun({ text: "SIGNATORIES:", bold: true, size: SIZE_LABEL, color: COLOR_TEXT_MUTED })
                    ],
                    spacing: { after: 1200 } 
                }),
                preparedByBox,
                notedByBox,
                timeReleasedBox,
                new Paragraph({ text: "", spacing: { before: 200 } }),
                transParagraph,
                ...notesParagraph,
                new Paragraph({
                    children: [
                        new TextRun({
                            text: formatExportText(data.footerNotes.acknowledgement),
                            italics: true,
                            size: 16,
                            color: COLOR_TEXT_SECONDARY,
                            font: FONT_FAMILY,
                        }),
                    ],
                    spacing: { after: 100, before: 100 },
                    wordWrap: true,
                }),
                receivedByTable,
                disclaimer
            ]
        }]
    });

    return await Packer.toBlob(doc);
};

function createHeaderCell(text: string, widthPercent: number) {
    return new TableCell({
        children: [
            new Paragraph({
                children: [
                    new TextRun({
                        text,
                        bold: true,
                        size: SIZE_LABEL,
                        font: FONT_FAMILY,
                        color: COLOR_TEXT_SECONDARY,
                    }),
                ],
                alignment: AlignmentType.CENTER,
            }),
        ],
        shading: { fill: COLOR_BG_HEADER },
        width: { size: widthPercent * 50, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 100, bottom: 100, left: 60, right: 60 },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            left: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            right: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
        }
    });
}

type AlignmentValue = (typeof AlignmentType)[keyof typeof AlignmentType];

function createItemCell(
    text: string,
    widthPercent: number,
    align: AlignmentValue = AlignmentType.LEFT,
) {
    return new TableCell({
        children: [
            new Paragraph({
                children: [
                    new TextRun({
                        text: formatExportText(text),
                        size: SIZE_TEXT,
                        font: FONT_FAMILY,
                        color: COLOR_TEXT_PRIMARY,
                    }),
                ],
                alignment: align,
                wordWrap: true,
            }),
        ],
        width: { size: widthPercent * 50, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 120, bottom: 120, left: 60, right: 60 },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            left: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
            right: { style: BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
        }
    });
}
