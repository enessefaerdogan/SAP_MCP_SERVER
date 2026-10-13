import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";

import { createRequire } from "node:module";
import * as path from "node:path";
import { existsSync } from "node:fs";

import {
    Document,
    Packer,
    Paragraph,
    HeadingLevel,
    Table,
    TableRow,
    TableCell,
    TextRun,
    WidthType
} from "docx";

import { artifactStore } from "../services/artifact-store.js";
import { Logger } from "../utils/logger.js";


/*
 * ============================================================
 * REQUIRE
 * ============================================================
 */

const require =
    createRequire(import.meta.url);


/*
 * ============================================================
 * PDFMAKE - NODE.JS SERVER SIDE
 * ============================================================
 *
 * pdfmake 0.3.x:
 *
 * require("pdfmake")
 * pdfmake.addFonts(...)
 * await pdfmake.createPdf(...).getBuffer()
 *
 * Browser build / callback API kullanılmıyor.
 * ============================================================
 */

const pdfmake: any =
    require("pdfmake");


/*
 * pdfmake paketinin root klasörünü bul.
 */

const pdfmakeRoot =
    path.dirname(
        require.resolve(
            "pdfmake/package.json"
        )
    );


/*
 * pdfmake paket sürümüne / publish yapısına göre
 * Roboto iki farklı konumda bulunabilir.
 *
 * İkisini de kontrol ediyoruz.
 */

const robotoCandidates = [

    path.join(
        pdfmakeRoot,
        "fonts",
        "Roboto"
    ),

    path.join(
        pdfmakeRoot,
        "build",
        "fonts",
        "Roboto"
    )
];


const robotoDirectory =
    robotoCandidates.find(
        directory =>
            existsSync(
                path.join(
                    directory,
                    "Roboto-Regular.ttf"
                )
            )
    );


if (!robotoDirectory) {

    throw new Error(
        `Roboto fonts could not be found inside pdfmake package. pdfmakeRoot=${pdfmakeRoot}`
    );
}


/*
 * Türkçe karakter desteği için Unicode font.
 */

pdfmake.addFonts({

    Roboto: {

        normal:
            path.join(
                robotoDirectory,
                "Roboto-Regular.ttf"
            ),

        bold:
            path.join(
                robotoDirectory,
                "Roboto-Medium.ttf"
            ),

        italics:
            path.join(
                robotoDirectory,
                "Roboto-Italic.ttf"
            ),

        bolditalics:
            path.join(
                robotoDirectory,
                "Roboto-MediumItalic.ttf"
            )
    }
});


/*
 * ============================================================
 * REPORT BLOCK SCHEMAS
 * ============================================================
 */

const headingBlockSchema =
    z.object({

        type:
            z.literal(
                "heading"
            ),

        text:
            z.string(),

        level:
            z.number()
                .min(1)
                .max(3)
                .optional()
    });


const paragraphBlockSchema =
    z.object({

        type:
            z.literal(
                "paragraph"
            ),

        text:
            z.string()
    });


const keyValueBlockSchema =
    z.object({

        type:
            z.literal(
                "keyValue"
            ),

        items:
            z.array(

                z.object({

                    label:
                        z.string(),

                    value:
                        z.string()
                })
            )
    });


const tableBlockSchema =
    z.object({

        type:
            z.literal(
                "table"
            ),

        title:
            z.string()
                .optional(),

        columns:
            z.array(
                z.string()
            ),

        rows:
            z.array(

                z.array(
                    z.string()
                )
            )
    });


const bulletsBlockSchema =
    z.object({

        type:
            z.literal(
                "bullets"
            ),

        items:
            z.array(
                z.string()
            )
    });


const reportBlockSchema =
    z.discriminatedUnion(
        "type",
        [
            headingBlockSchema,
            paragraphBlockSchema,
            keyValueBlockSchema,
            tableBlockSchema,
            bulletsBlockSchema
        ]
    );


type ReportBlock =
    z.infer<
        typeof reportBlockSchema
    >;


/*
 * ============================================================
 * ARTIFACT TOOL REGISTRY
 * ============================================================
 */

export class ArtifactToolRegistry {

    constructor(
        private mcpServer: McpServer,
        private logger: Logger
    ) {}


    /*
     * ========================================================
     * REGISTER TOOLS
     * ========================================================
     */

    registerTools(): void {

        this.mcpServer.registerTool(

            "create-report-artifact",

            {

                title:
                    "Create Professional Report Artifact",

                description:
                    "Creates a professionally formatted PDF, Word or text report. " +
                    "Use structured blocks such as headings, paragraphs, key-value sections, " +
                    "tables and bullet lists. Do not send Markdown formatting inside text.",


                inputSchema: {

                    format:
                        z.enum([
                            "pdf",
                            "docx",
                            "txt"
                        ]),

                    fileName:
                        z.string()
                            .optional(),

                    title:
                        z.string(),

                    subtitle:
                        z.string()
                            .optional(),

                    summary:
                        z.string()
                            .optional(),

                    blocks:
                        z.array(
                            reportBlockSchema
                        )
                }
            },


            async ({
                format,
                fileName,
                title,
                subtitle,
                summary,
                blocks
            }) => {

                try {

                    this.logger.info(
                        `📄 Creating structured ${format} report: ${title}`
                    );


                    let buffer: Buffer;
                    let mimeType: string;
                    let finalFileName: string;


                    /*
                     * ================================================
                     * PDF
                     * ================================================
                     */

                    if (
                        format ===
                        "pdf"
                    ) {

                        buffer =
                            await this.createPdf(
                                title,
                                subtitle,
                                summary,
                                blocks
                            );


                        mimeType =
                            "application/pdf";


                        finalFileName =
                            this.normalizeFileName(
                                fileName ||
                                "report.pdf",
                                ".pdf"
                            );
                    }


                    /*
                     * ================================================
                     * DOCX
                     * ================================================
                     */

                    else if (
                        format ===
                        "docx"
                    ) {

                        buffer =
                            await this.createDocx(
                                title,
                                subtitle,
                                summary,
                                blocks
                            );


                        mimeType =
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";


                        finalFileName =
                            this.normalizeFileName(
                                fileName ||
                                "report.docx",
                                ".docx"
                            );
                    }


                    /*
                     * ================================================
                     * TXT
                     * ================================================
                     */

                    else {

                        buffer =
                            this.createTxt(
                                title,
                                subtitle,
                                summary,
                                blocks
                            );


                        mimeType =
                            "text/plain; charset=utf-8";


                        finalFileName =
                            this.normalizeFileName(
                                fileName ||
                                "report.txt",
                                ".txt"
                            );
                    }


                    /*
                     * ================================================
                     * STORE ARTIFACT
                     * ================================================
                     */

                    const artifact =
                        artifactStore.create(
                            finalFileName,
                            mimeType,
                            buffer
                        );


                    const result = {

                        type:
                            "artifact",

                        artifactId:
                            artifact.id,

                        fileName:
                            artifact.fileName,

                        mimeType:
                            artifact.mimeType,

                        size:
                            artifact.size
                    };


                    this.logger.info(
                        `✅ Artifact created: ${artifact.id} (${artifact.fileName}, ${artifact.size} bytes)`
                    );


                    /*
                     * Binary veya base64 Claude'ya gitmiyor.
                     *
                     * Sadece metadata.
                     */

                    return {

                        content: [
                            {
                                type:
                                    "text",

                                text:
                                    JSON.stringify(
                                        result
                                    )
                            }
                        ]
                    };


                } catch (error) {

                    this.logger.error(
                        `❌ Failed to create report artifact: ${
                            error instanceof Error
                                ? error.stack || error.message
                                : String(error)
                        }`
                    );


                    throw error;
                }
            }
        );


        this.logger.info(
            "📎 Registered professional report artifact tool"
        );
    }


    /*
     * ============================================================
     * PDF GENERATOR
     * ============================================================
     */

    private async createPdf(
        title: string,
        subtitle: string | undefined,
        summary: string | undefined,
        blocks: ReportBlock[]
    ): Promise<Buffer> {

        /*
         * Çok kolonlu tablolar landscape olsun.
         */

        const tableBlocks =
            blocks.filter(
                (
                    block
                ): block is Extract<
                    ReportBlock,
                    {
                        type: "table";
                    }
                > =>
                    block.type ===
                    "table"
            );


        const maxColumnCount =
            tableBlocks.reduce(
                (
                    maximum,
                    block
                ) =>
                    Math.max(
                        maximum,
                        block.columns.length
                    ),
                0
            );


        const pageOrientation =
            maxColumnCount >= 6
                ? "landscape"
                : "portrait";


        const content: any[] =
            [];


        /*
         * ========================================================
         * TITLE
         * ========================================================
         */

        content.push({

            text:
                this.cleanText(
                    title
                ),

            style:
                "reportTitle"
        });


        /*
         * ========================================================
         * SUBTITLE
         * ========================================================
         */

        if (subtitle) {

            content.push({

                text:
                    this.cleanText(
                        subtitle
                    ),

                style:
                    "reportSubtitle"
            });
        }


        /*
         * ========================================================
         * SUMMARY
         * ========================================================
         */

        if (summary) {

            content.push({

                table: {

                    widths: [
                        "*"
                    ],

                    body: [
                        [
                            {
                                text:
                                    this.cleanText(
                                        summary
                                    ),

                                style:
                                    "summaryText"
                            }
                        ]
                    ]
                },

                layout: {

                    fillColor:
                        () =>
                            "#F5F8FA",

                    hLineColor:
                        () =>
                            "#D9E2E8",

                    vLineColor:
                        () =>
                            "#D9E2E8",

                    paddingLeft:
                        () =>
                            12,

                    paddingRight:
                        () =>
                            12,

                    paddingTop:
                        () =>
                            10,

                    paddingBottom:
                        () =>
                            10
                },

                margin: [
                    0,
                    8,
                    0,
                    18
                ]
            });
        }


        /*
         * ========================================================
         * STRUCTURED BLOCKS
         * ========================================================
         */

        for (
            const block
            of blocks
        ) {

            /*
             * ====================================================
             * HEADING
             * ====================================================
             */

            if (
                block.type ===
                "heading"
            ) {

                const style =
                    block.level === 1
                        ? "heading1"
                        : block.level === 3
                            ? "heading3"
                            : "heading2";


                content.push({

                    text:
                        this.cleanText(
                            block.text
                        ),

                    style
                });


                continue;
            }


            /*
             * ====================================================
             * PARAGRAPH
             * ====================================================
             */

            if (
                block.type ===
                "paragraph"
            ) {

                content.push({

                    text:
                        this.cleanText(
                            block.text
                        ),

                    style:
                        "paragraph"
                });


                continue;
            }


            /*
             * ====================================================
             * BULLETS
             * ====================================================
             */

            if (
                block.type ===
                "bullets"
            ) {

                content.push({

                    ul:
                        block.items.map(
                            item =>
                                this.cleanText(
                                    item
                                )
                        ),

                    style:
                        "bulletList",

                    margin: [
                        8,
                        4,
                        0,
                        14
                    ]
                });


                continue;
            }


            /*
             * ====================================================
             * KEY / VALUE
             * ====================================================
             */

            if (
                block.type ===
                "keyValue"
            ) {

                const body =
                    block.items.map(
                        item => [

                            {
                                text:
                                    this.cleanText(
                                        item.label
                                    ),

                                bold:
                                    true,

                                fontSize:
                                    9,

                                color:
                                    "#354A5F"
                            },

                            {
                                text:
                                    this.cleanText(
                                        item.value
                                    ),

                                fontSize:
                                    9
                            }
                        ]
                    );


                content.push({

                    table: {

                        widths: [
                            "35%",
                            "*"
                        ],

                        body
                    },


                    layout: {

                        fillColor:
                            (
                                rowIndex: number
                            ) =>
                                rowIndex % 2 === 0
                                    ? "#FAFBFC"
                                    : null,

                        hLineColor:
                            () =>
                                "#E5E8EB",

                        vLineColor:
                            () =>
                                "#E5E8EB",

                        paddingLeft:
                            () =>
                                8,

                        paddingRight:
                            () =>
                                8,

                        paddingTop:
                            () =>
                                6,

                        paddingBottom:
                            () =>
                                6
                    },


                    margin: [
                        0,
                        6,
                        0,
                        18
                    ]
                });


                continue;
            }


            /*
             * ====================================================
             * TABLE
             * ====================================================
             */

            if (
                block.type ===
                "table"
            ) {

                if (
                    block.title
                ) {

                    content.push({

                        text:
                            this.cleanText(
                                block.title
                            ),

                        style:
                            "tableTitle"
                    });
                }


                /*
                 * Header
                 */

                const headerRow =
                    block.columns.map(
                        column => ({

                            text:
                                this.cleanText(
                                    column
                                ),

                            bold:
                                true,

                            fontSize:
                                8.5,

                            color:
                                "#FFFFFF"
                        })
                    );


                /*
                 * Rows
                 */

                const dataRows =
                    block.rows.map(

                        row =>

                            block.columns.map(
                                (
                                    _,
                                    index
                                ) => ({

                                    text:
                                        this.cleanText(
                                            row[index] ??
                                            ""
                                        ),

                                    fontSize:
                                        8
                                })
                            )
                    );


                content.push({

                    table: {

                        headerRows:
                            1,

                        dontBreakRows:
                            true,

                        widths:
                            block.columns.map(
                                () =>
                                    "*"
                            ),

                        body: [
                            headerRow,
                            ...dataRows
                        ]
                    },


                    layout: {

                        fillColor:
                            (
                                rowIndex: number
                            ) => {

                                if (
                                    rowIndex ===
                                    0
                                ) {

                                    return (
                                        "#354A5F"
                                    );
                                }


                                return (
                                    rowIndex % 2 === 0
                                        ? "#F7F9FA"
                                        : "#FFFFFF"
                                );
                            },


                        hLineColor:
                            () =>
                                "#D9DDE2",

                        vLineColor:
                            () =>
                                "#D9DDE2",

                        hLineWidth:
                            () =>
                                0.6,

                        vLineWidth:
                            () =>
                                0.6,

                        paddingLeft:
                            () =>
                                6,

                        paddingRight:
                            () =>
                                6,

                        paddingTop:
                            () =>
                                5,

                        paddingBottom:
                            () =>
                                5
                    },


                    margin: [
                        0,
                        5,
                        0,
                        20
                    ]
                });


                continue;
            }
        }


        /*
         * ========================================================
         * DOCUMENT DEFINITION
         * ========================================================
         */

        const documentDefinition: any = {

            pageSize:
                "A4",

            pageOrientation,

            pageMargins: [
                40,
                45,
                40,
                50
            ],


            info: {

                title:
                    this.cleanText(
                        title
                    ),

                author:
                    "SAP AI / MCP",

                subject:
                    subtitle
                        ? this.cleanText(
                            subtitle
                        )
                        : "SAP AI generated report"
            },


            content,


            /*
             * ====================================================
             * FOOTER
             * ====================================================
             */

            footer: (
                currentPage: number,
                pageCount: number
            ) => {

                return {

                    columns: [

                        {
                            text:
                                "SAP AI / MCP",

                            alignment:
                                "left"
                        },

                        {
                            text:
                                `${currentPage} / ${pageCount}`,

                            alignment:
                                "right"
                        }
                    ],

                    margin: [
                        40,
                        12,
                        40,
                        0
                    ],

                    fontSize:
                        8,

                    color:
                        "#6A6D70"
                };
            },


            /*
             * ====================================================
             * STYLES
             * ====================================================
             */

            styles: {

                reportTitle: {

                    fontSize:
                        22,

                    bold:
                        true,

                    color:
                        "#1D2D3E",

                    margin: [
                        0,
                        0,
                        0,
                        6
                    ]
                },


                reportSubtitle: {

                    fontSize:
                        11,

                    color:
                        "#6A6D70",

                    margin: [
                        0,
                        0,
                        0,
                        18
                    ]
                },


                summaryText: {

                    fontSize:
                        10,

                    lineHeight:
                        1.3,

                    color:
                        "#32363A"
                },


                heading1: {

                    fontSize:
                        17,

                    bold:
                        true,

                    color:
                        "#1D2D3E",

                    margin: [
                        0,
                        18,
                        0,
                        8
                    ]
                },


                heading2: {

                    fontSize:
                        14,

                    bold:
                        true,

                    color:
                        "#354A5F",

                    margin: [
                        0,
                        16,
                        0,
                        7
                    ]
                },


                heading3: {

                    fontSize:
                        11.5,

                    bold:
                        true,

                    color:
                        "#354A5F",

                    margin: [
                        0,
                        12,
                        0,
                        5
                    ]
                },


                paragraph: {

                    fontSize:
                        10,

                    lineHeight:
                        1.3,

                    color:
                        "#32363A",

                    margin: [
                        0,
                        3,
                        0,
                        10
                    ]
                },


                bulletList: {

                    fontSize:
                        10,

                    lineHeight:
                        1.25,

                    color:
                        "#32363A"
                },


                tableTitle: {

                    fontSize:
                        11,

                    bold:
                        true,

                    color:
                        "#354A5F",

                    margin: [
                        0,
                        10,
                        0,
                        5
                    ]
                }
            },


            defaultStyle: {

                font:
                    "Roboto",

                fontSize:
                    10,

                color:
                    "#32363A"
            }
        };


        /*
         * ========================================================
         * GENERATE PDF
         * ========================================================
         */

        this.logger.info(
            `🧾 Starting PDF render (orientation=${pageOrientation}, Roboto=${robotoDirectory})`
        );


        /*
         * pdfmake 0.3.x API.
         */

        const pdf =
            pdfmake.createPdf(
                documentDefinition
            );


        const generatedBuffer =
            await pdf.getBuffer();


        const buffer =
            Buffer.isBuffer(
                generatedBuffer
            )
                ? generatedBuffer
                : Buffer.from(
                    generatedBuffer
                );


        if (
            buffer.length ===
            0
        ) {

            throw new Error(
                "pdfmake returned an empty PDF buffer."
            );
        }


        this.logger.info(
            `✅ PDF render completed: ${buffer.length} bytes`
        );


        return buffer;
    }


    /*
     * ============================================================
     * DOCX GENERATOR
     * ============================================================
     */

    private async createDocx(
        title: string,
        subtitle: string | undefined,
        summary: string | undefined,
        blocks: ReportBlock[]
    ): Promise<Buffer> {

        const children:
            Array<
                Paragraph |
                Table
            > =
            [];


        /*
         * TITLE
         */

        children.push(

            new Paragraph({

                text:
                    this.cleanText(
                        title
                    ),

                heading:
                    HeadingLevel.TITLE
            })
        );


        /*
         * SUBTITLE
         */

        if (
            subtitle
        ) {

            children.push(

                new Paragraph({

                    children: [

                        new TextRun({

                            text:
                                this.cleanText(
                                    subtitle
                                ),

                            italics:
                                true
                        })
                    ]
                })
            );
        }


        /*
         * SUMMARY
         */

        if (
            summary
        ) {

            children.push(

                new Paragraph({

                    text:
                        this.cleanText(
                            summary
                        )
                })
            );
        }


        /*
         * BLOCKS
         */

        for (
            const block
            of blocks
        ) {

            /*
             * ====================================================
             * HEADING
             * ====================================================
             */

            if (
                block.type ===
                "heading"
            ) {

                const heading =
                    block.level === 1
                        ? HeadingLevel.HEADING_1
                        : block.level === 3
                            ? HeadingLevel.HEADING_3
                            : HeadingLevel.HEADING_2;


                children.push(

                    new Paragraph({

                        text:
                            this.cleanText(
                                block.text
                            ),

                        heading
                    })
                );


                continue;
            }


            /*
             * ====================================================
             * PARAGRAPH
             * ====================================================
             */

            if (
                block.type ===
                "paragraph"
            ) {

                children.push(

                    new Paragraph({

                        text:
                            this.cleanText(
                                block.text
                            )
                    })
                );


                continue;
            }


            /*
             * ====================================================
             * BULLETS
             * ====================================================
             */

            if (
                block.type ===
                "bullets"
            ) {

                for (
                    const item
                    of block.items
                ) {

                    children.push(

                        new Paragraph({

                            text:
                                this.cleanText(
                                    item
                                ),

                            bullet: {
                                level:
                                    0
                            }
                        })
                    );
                }


                continue;
            }


            /*
             * ====================================================
             * KEY VALUE TABLE
             * ====================================================
             */

            if (
                block.type ===
                "keyValue"
            ) {

                children.push(

                    new Table({

                        width: {

                            size:
                                100,

                            type:
                                WidthType.PERCENTAGE
                        },


                        rows:
                            block.items.map(

                                item =>

                                    new TableRow({

                                        children: [

                                            new TableCell({

                                                children: [

                                                    new Paragraph({

                                                        children: [

                                                            new TextRun({

                                                                text:
                                                                    this.cleanText(
                                                                        item.label
                                                                    ),

                                                                bold:
                                                                    true
                                                            })
                                                        ]
                                                    })
                                                ]
                                            }),


                                            new TableCell({

                                                children: [

                                                    new Paragraph({

                                                        text:
                                                            this.cleanText(
                                                                item.value
                                                            )
                                                    })
                                                ]
                                            })
                                        ]
                                    })
                            )
                    })
                );


                continue;
            }


            /*
             * ====================================================
             * DATA TABLE
             * ====================================================
             */

            if (
                block.type ===
                "table"
            ) {

                /*
                 * Title
                 */

                if (
                    block.title
                ) {

                    children.push(

                        new Paragraph({

                            children: [

                                new TextRun({

                                    text:
                                        this.cleanText(
                                            block.title
                                        ),

                                    bold:
                                        true
                                })
                            ]
                        })
                    );
                }


                /*
                 * Header
                 */

                const header =
                    new TableRow({

                        tableHeader:
                            true,

                        children:
                            block.columns.map(

                                column =>

                                    new TableCell({

                                        children: [

                                            new Paragraph({

                                                children: [

                                                    new TextRun({

                                                        text:
                                                            this.cleanText(
                                                                column
                                                            ),

                                                        bold:
                                                            true
                                                    })
                                                ]
                                            })
                                        ]
                                    })
                            )
                    });


                /*
                 * Rows
                 */

                const rows =
                    block.rows.map(

                        row =>

                            new TableRow({

                                children:
                                    block.columns.map(
                                        (
                                            _,
                                            index
                                        ) =>

                                            new TableCell({

                                                children: [

                                                    new Paragraph({

                                                        text:
                                                            this.cleanText(
                                                                row[index] ??
                                                                ""
                                                            )
                                                    })
                                                ]
                                            })
                                    )
                            })
                    );


                children.push(

                    new Table({

                        width: {

                            size:
                                100,

                            type:
                                WidthType.PERCENTAGE
                        },

                        rows: [
                            header,
                            ...rows
                        ]
                    })
                );


                continue;
            }
        }


        /*
         * DOCUMENT
         */

        const document =
            new Document({

                sections: [
                    {
                        children
                    }
                ]
            });


        const buffer =
            Buffer.from(

                await Packer.toBuffer(
                    document
                )
            );


        this.logger.info(
            `✅ DOCX render completed: ${buffer.length} bytes`
        );


        return buffer;
    }


    /*
     * ============================================================
     * TXT GENERATOR
     * ============================================================
     */

    private createTxt(
        title: string,
        subtitle: string | undefined,
        summary: string | undefined,
        blocks: ReportBlock[]
    ): Buffer {

        const lines:
            string[] =
            [];


        /*
         * TITLE
         */

        const cleanedTitle =
            this.cleanText(
                title
            );


        lines.push(
            cleanedTitle
        );


        lines.push(

            "=".repeat(

                Math.max(
                    cleanedTitle.length,
                    3
                )
            )
        );


        /*
         * SUBTITLE
         */

        if (
            subtitle
        ) {

            lines.push(
                this.cleanText(
                    subtitle
                )
            );
        }


        lines.push(
            ""
        );


        /*
         * SUMMARY
         */

        if (
            summary
        ) {

            lines.push(
                this.cleanText(
                    summary
                ),
                ""
            );
        }


        /*
         * BLOCKS
         */

        for (
            const block
            of blocks
        ) {

            /*
             * HEADING
             */

            if (
                block.type ===
                "heading"
            ) {

                const heading =
                    this.cleanText(
                        block.text
                    );


                lines.push(
                    heading
                );


                lines.push(

                    "-".repeat(

                        Math.max(
                            heading.length,
                            3
                        )
                    )
                );


                lines.push(
                    ""
                );
            }


            /*
             * PARAGRAPH
             */

            else if (
                block.type ===
                "paragraph"
            ) {

                lines.push(
                    this.cleanText(
                        block.text
                    ),
                    ""
                );
            }


            /*
             * BULLETS
             */

            else if (
                block.type ===
                "bullets"
            ) {

                block.items.forEach(

                    item =>

                        lines.push(
                            `- ${this.cleanText(item)}`
                        )
                );


                lines.push(
                    ""
                );
            }


            /*
             * KEY VALUE
             */

            else if (
                block.type ===
                "keyValue"
            ) {

                block.items.forEach(

                    item =>

                        lines.push(
                            `${this.cleanText(item.label)}: ${this.cleanText(item.value)}`
                        )
                );


                lines.push(
                    ""
                );
            }


            /*
             * TABLE
             */

            else if (
                block.type ===
                "table"
            ) {

                if (
                    block.title
                ) {

                    lines.push(
                        this.cleanText(
                            block.title
                        )
                    );
                }


                lines.push(

                    block.columns

                        .map(
                            column =>
                                this.cleanText(
                                    column
                                )
                        )

                        .join(
                            " | "
                        )
                );


                lines.push(

                    block.columns

                        .map(
                            () =>
                                "---"
                        )

                        .join(
                            " | "
                        )
                );


                block.rows.forEach(

                    row =>

                        lines.push(

                            block.columns

                                .map(
                                    (
                                        _,
                                        index
                                    ) =>
                                        this.cleanText(
                                            row[index] ??
                                            ""
                                        )
                                )

                                .join(
                                    " | "
                                )
                        )
                );


                lines.push(
                    ""
                );
            }
        }


        const buffer =
            Buffer.from(
                lines.join(
                    "\n"
                ),
                "utf8"
            );


        this.logger.info(
            `✅ TXT render completed: ${buffer.length} bytes`
        );


        return buffer;
    }


    /*
     * ============================================================
     * CLEAN TEXT
     * ============================================================
     *
     * Tool schema structured olduğu için burada Markdown
     * istemiyoruz.
     *
     * Claude yine de **bold**, # heading veya `code`
     * gönderirse PDF içine karakter olarak basılmasını engeller.
     * ============================================================
     */

    private cleanText(
        value: string
    ): string {

        return String(
            value ??
            ""
        )

            /*
             * Markdown heading
             */
            .replace(
                /^#{1,6}\s+/gm,
                ""
            )

            /*
             * Bold
             */
            .replace(
                /\*\*(.*?)\*\*/g,
                "$1"
            )

            .replace(
                /__(.*?)__/g,
                "$1"
            )

            /*
             * Inline code
             */
            .replace(
                /`([^`]+)`/g,
                "$1"
            )

            /*
             * Markdown links
             * [Text](url) -> Text
             */
            .replace(
                /\[([^\]]+)\]\([^)]+\)/g,
                "$1"
            )

            .trim();
    }


    /*
     * ============================================================
     * FILE NAME
     * ============================================================
     */

    private normalizeFileName(
        fileName: string,
        extension: string
    ): string {

        let result =
            fileName

                .replace(
                    /[\\/:*?"<>|]/g,
                    "_"
                )

                .trim();


        if (
            !result
        ) {

            result =
                `report${extension}`;
        }


        if (
            !result
                .toLowerCase()
                .endsWith(
                    extension
                )
        ) {

            result +=
                extension;
        }


        return result;
    }
}