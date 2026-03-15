"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.generateTransmittalDocx = void 0;
var docx_1 = require("docx");
var exportText_1 = require("./exportText");
// Helper to convert base64 to Uint8Array safely
var base64ToUint8Array = function (base64) {
    try {
        var data = base64.includes(',') ? base64.split(',')[1] : base64;
        var binaryString = window.atob(data);
        var len = binaryString.length;
        var bytes = new Uint8Array(len);
        for (var i = 0; i < len; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        return bytes;
    }
    catch (e) {
        console.error("Failed to convert base64 to bytes", e);
        return new Uint8Array(0);
    }
};
var FONT_FAMILY = "Arial";
var COLOR_BORDER = "CBD5E1";
var COLOR_BG_HEADER = "F8FAFC";
var COLOR_TEXT_PRIMARY = "1E293B";
var COLOR_TEXT_SECONDARY = "475569";
var COLOR_TEXT_MUTED = "94A3B8";
var SIZE_TITLE = 28;
var SIZE_TEXT = 20;
var SIZE_LABEL = 16;
var SIZE_HEADER_SMALL = 14;
/**
 * Creates a floating "Box" for signatories that can be dragged in Word.
 * Uses string literals for anchoring to avoid ESM import errors.
 */
var createFloatingSignatoryBox = function (label, name, role, horizontalOffset) {
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
var generateTransmittalDocx = function (data) { return __awaiter(void 0, void 0, void 0, function () {
    var logoRun, imageBytes, headerTable, title, leftData, rightData, maxMetaRows, metaTableRows, cellMargin, i, left, right, cells, metaTable, tableHeader, itemRows, itemsTable, preparedByBox, notedByBox, timeReleasedBox, checkChar, transParagraph, notesParagraph, createRecvCell, receivedByTable, disclaimer, doc;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0:
                logoRun = new docx_1.TextRun({ text: "[LOGO]", bold: true, size: 24, color: COLOR_TEXT_MUTED });
                if (data.sender.logoBase64) {
                    try {
                        imageBytes = base64ToUint8Array(data.sender.logoBase64);
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
                headerTable = new docx_1.Table({
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
                                        new docx_1.Paragraph({ text: "Telephone: ".concat(data.sender.telephone), alignment: docx_1.AlignmentType.RIGHT, style: "HeaderSmall" }),
                                        new docx_1.Paragraph({ text: "Mobile: ".concat(data.sender.mobile), alignment: docx_1.AlignmentType.RIGHT, style: "HeaderSmall" }),
                                        new docx_1.Paragraph({ text: "Email: ".concat(data.sender.email), alignment: docx_1.AlignmentType.RIGHT, style: "HeaderSmall" }),
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
                title = new docx_1.Paragraph({
                    children: [new docx_1.TextRun({ text: "TRANSMITTAL FORM", font: FONT_FAMILY, bold: true, size: SIZE_TITLE, color: "0F172A" })],
                    alignment: docx_1.AlignmentType.CENTER,
                    spacing: { before: 300, after: 300 }
                });
                leftData = [
                    { label: 'To:', value: data.recipient.to },
                    { label: 'Company:', value: data.recipient.company },
                    { label: 'Attention:', value: data.recipient.attention },
                    { label: 'Address:', value: data.recipient.address },
                    { label: 'Contact No:', value: data.recipient.contactNumber },
                    { label: 'Email:', value: data.recipient.email }
                ];
                rightData = [
                    { label: 'Project Name:', value: data.project.projectName },
                    { label: 'Project No:', value: data.project.projectNumber },
                    { label: 'Engagement Ref #:', value: data.project.engagementRef },
                    { label: 'Purpose:', value: data.project.purpose },
                    { label: 'Transmittal No:', value: data.project.transmittalNumber },
                    { label: 'Department:', value: data.project.department },
                    { label: 'Date:', value: data.project.date },
                    { label: 'Time Generated:', value: data.project.timeGenerated }
                ];
                maxMetaRows = Math.max(leftData.length, rightData.length);
                metaTableRows = [];
                cellMargin = { top: 120, bottom: 120, left: 100, right: 100 };
                for (i = 0; i < maxMetaRows; i++) {
                    left = leftData[i];
                    right = rightData[i];
                    cells = [];
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
                metaTable = new docx_1.Table({
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
                tableHeader = new docx_1.TableRow({
                    children: [
                        createHeaderCell("No. of Items", 10),
                        createHeaderCell("QTY", 10),
                        createHeaderCell("Document # / Certificate #", 25),
                        createHeaderCell("Description", 30),
                        createHeaderCell("Remarks", 25),
                    ],
                    tableHeader: true
                });
                itemRows = data.items.map(function (item) { return new docx_1.TableRow({
                    children: [
                        createItemCell(item.noOfItems, 10, docx_1.AlignmentType.CENTER),
                        createItemCell(item.qty, 10, docx_1.AlignmentType.CENTER),
                        createItemCell(item.documentNumber, 25),
                        createItemCell(item.description, 30),
                        createItemCell(item.remarks, 25),
                    ]
                }); });
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
                itemsTable = new docx_1.Table({
                    width: { size: 5000, type: docx_1.WidthType.PERCENTAGE },
                    rows: __spreadArray([tableHeader], itemRows, true),
                    borders: {
                        top: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
                        bottom: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
                        left: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
                        right: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
                        insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER },
                        insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 2, color: COLOR_BORDER }
                    }
                });
                preparedByBox = createFloatingSignatoryBox("Prepared by:", data.signatories.preparedBy, data.signatories.preparedByRole, 0);
                notedByBox = createFloatingSignatoryBox("Noted by:", data.signatories.notedBy, data.signatories.notedByRole, 2600);
                timeReleasedBox = createFloatingSignatoryBox("Time Released:", data.signatories.timeReleased, null, 5200);
                checkChar = function (checked) { return checked ? "☒" : "☐"; };
                transParagraph = new docx_1.Paragraph({
                    children: [
                        new docx_1.TextRun({ text: "Transmitted via:  ", bold: true, size: SIZE_LABEL, font: FONT_FAMILY, color: "000000" }),
                        new docx_1.TextRun({ text: "".concat(checkChar(data.transmissionMethod.personalDelivery), " Personal Delivery    "), size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
                        new docx_1.TextRun({ text: "".concat(checkChar(data.transmissionMethod.pickUp), " Pick-up    "), size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
                        new docx_1.TextRun({ text: "".concat(checkChar(data.transmissionMethod.grabLalamove), " Courier App    "), size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
                        new docx_1.TextRun({ text: "".concat(checkChar(data.transmissionMethod.registeredMail), " Registered Mail"), size: SIZE_LABEL, font: FONT_FAMILY, color: COLOR_TEXT_SECONDARY }),
                    ],
                    spacing: { before: 200, after: 100 },
                    border: { top: { style: docx_1.BorderStyle.SINGLE, space: 5, color: "EEEEEE" }, bottom: { style: docx_1.BorderStyle.SINGLE, space: 5, color: "EEEEEE" }, left: { style: docx_1.BorderStyle.SINGLE, space: 5, color: "EEEEEE" }, right: { style: docx_1.BorderStyle.SINGLE, space: 5, color: "EEEEEE" } },
                    shading: { fill: COLOR_BG_HEADER }
                });
                notesParagraph = data.notes ? [
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
                createRecvCell = function (label, value) {
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
                receivedByTable = new docx_1.Table({
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
                disclaimer = new docx_1.Paragraph({
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
                doc = new docx_1.Document({
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
                            children: __spreadArray(__spreadArray([
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
                                transParagraph
                            ], notesParagraph, true), [
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
                            ], false)
                        }]
                });
                return [4 /*yield*/, docx_1.Packer.toBlob(doc)];
            case 1: return [2 /*return*/, _a.sent()];
        }
    });
}); };
exports.generateTransmittalDocx = generateTransmittalDocx;
function createHeaderCell(text, widthPercent) {
    return new docx_1.TableCell({
        children: [
            new docx_1.Paragraph({
                children: [
                    new docx_1.TextRun({
                        text: text,
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
function createItemCell(text, widthPercent, align) {
    if (align === void 0) { align = docx_1.AlignmentType.LEFT; }
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
