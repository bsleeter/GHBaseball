const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageBreak, PageNumber, LevelFormat
} = require("docx");

const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

// Content width for US Letter with 1" margins = 9360 DXA
const CONTENT_WIDTH = 9360;

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "1F3864", type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 22 })]
    })]
  });
}

function dataCell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text, font: "Arial", size: 22, bold: !!opts.bold })]
    })]
  });
}

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Arial", size: 24 } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "1F3864" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "333333" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 }
      }
    ]
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "bullets2",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "bullets3",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "bullets4",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "bullets5",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "bullets6",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "bullets7",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "bullets8",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  sections: [
    // ===== TITLE PAGE =====
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: [
        new Paragraph({ spacing: { before: 4000 }, children: [] }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "Gig Harbor High School", size: 52, bold: true, font: "Arial", color: "1F3864" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [new TextRun({ text: "Baseball Field Renovation Proposal", size: 44, bold: true, font: "Arial", color: "2E75B6" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 20 } },
          children: []
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "Presented to the Peninsula School District", size: 28, font: "Arial", color: "444444" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "Prepared by: Gig Harbor High School Baseball Booster Club", size: 28, font: "Arial", color: "444444" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "April 2026", size: 28, font: "Arial", color: "444444" })]
        }),
      ]
    },

    // ===== EXECUTIVE SUMMARY =====
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: "GHHS Baseball Field Renovation Proposal", italics: true, size: 18, color: "888888", font: "Arial" })]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "Page ", size: 18, font: "Arial", color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: "888888" })]
          })]
        })
      },
      children: [
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Executive Summary")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The Gig Harbor High School baseball field is currently in such poor condition that the varsity team is unable to play home games on campus. All home games are played at Sehmel Homestead Park, an off-campus facility. This situation undermines school pride, increases logistical costs, and deprives our student-athletes of the experience of playing on their home field.",
            size: 22
          })]
        }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The Gig Harbor High School Baseball Booster Club has approximately $30,000 in funds currently committed toward field renovation. However, the scope of improvements necessary to restore the field to safe and playable standards will require significantly more investment. The Booster Club is requesting a partnership with the Peninsula School District to develop and fund a comprehensive renovation plan.",
            size: 22
          })]
        }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "This document outlines the current field conditions, proposed improvements organized in a phased approach, estimated costs, and a funding partnership framework. Our goal is to work collaboratively with the district to create a plan that restores the GHHS baseball field to a safe, functional facility that our athletes, school, and community can be proud of.",
            size: 22
          })]
        }),

        // ===== SECTION 1: CURRENT CONDITIONS =====
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("1. Current Field Conditions")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The GHHS baseball field has multiple areas that require significant attention. The following assessment details each area of concern:",
            size: 22
          })]
        }),

        // Infield
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Infield")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The infield accumulates standing water after rain events and does not drain properly. The current surface lacks proper infield mix material, resulting in unsafe and unplayable conditions. The infield requires a complete rebuild, including laser leveling to establish proper drainage grades, installation of appropriate infield mix, and top dressing to maintain a safe, consistent playing surface.",
            size: 22
          })]
        }),

        // Perimeter Fencing
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Perimeter Fencing")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The existing perimeter fencing is old, deteriorated, and open in numerous areas. Gaps in the fencing create safety hazards and fail to properly secure the field. The entire perimeter fencing system needs to be replaced with new, regulation-appropriate fencing.",
            size: 22
          })]
        }),

        // Warning Track
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Warning Track")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The current warning track is composed of large rock material that presents a serious safety risk to players. A warning track is designed to alert outfielders that they are approaching the fence; the current material could cause injury during play. It must be replaced with proper warning track material, typically a crushed brick or stone dust mix that provides appropriate footing and tactile warning.",
            size: 22
          })]
        }),

        // Outfield Fence
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Outfield Fence")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The outfield fence consists of portable fencing that is in poor condition. Portable fencing does not provide the stability, safety, or professional appearance expected of a high school varsity field. Permanent or high-quality semi-permanent fencing at regulation distance is needed.",
            size: 22
          })]
        }),

        // Bullpens
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Bullpens (Home & Visitor)")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "Both the home and visitor bullpen areas need to be completely rebuilt. This includes proper mound construction, adequate drainage, protective fencing or enclosures, and appropriate surfacing to allow pitchers to warm up safely.",
            size: 22
          })]
        }),

        // ===== SECTION 2: PROPOSED IMPROVEMENTS =====
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("2. Proposed Improvements")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "To manage costs and prioritize the most critical needs, we propose a phased approach to the renovation. This allows work to begin on the highest-priority items while additional funding is secured for subsequent phases.",
            size: 22
          })]
        }),

        // Phase 1
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Phase 1: Critical Safety & Playability (Highest Priority)")] }),
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: "This phase addresses the most urgent safety concerns and the core playability of the field:", size: 22 })]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 60 },
          children: [new TextRun({ text: "Complete infield renovation: strip existing surface, laser level the subgrade to establish proper drainage slope, install appropriate infield mix, and apply top dressing", size: 22 })]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 60 },
          children: [new TextRun({ text: "Warning track replacement: remove all hazardous rock material and install proper crushed brick/stone dust warning track material around the outfield perimeter", size: 22 })]
        }),
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 200 },
          children: [new TextRun({ text: "Estimated cost: $40,000 \u2013 $65,000", bold: true, size: 22 })]
        }),

        // Phase 2
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Phase 2: Fencing & Field Boundaries")] }),
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: "This phase addresses field security and boundary definition:", size: 22 })]
        }),
        new Paragraph({
          numbering: { reference: "bullets2", level: 0 },
          spacing: { after: 60 },
          children: [new TextRun({ text: "Replace all perimeter fencing with new chain-link or comparable fencing material", size: 22 })]
        }),
        new Paragraph({
          numbering: { reference: "bullets2", level: 0 },
          spacing: { after: 60 },
          children: [new TextRun({ text: "Install permanent or high-quality semi-permanent outfield fence at regulation distance", size: 22 })]
        }),
        new Paragraph({
          numbering: { reference: "bullets2", level: 0 },
          spacing: { after: 200 },
          children: [new TextRun({ text: "Estimated cost: $25,000 \u2013 $50,000", bold: true, size: 22 })]
        }),

        // Phase 3
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Phase 3: Bullpens & Finishing Improvements")] }),
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: "This phase completes the renovation with the remaining infrastructure:", size: 22 })]
        }),
        new Paragraph({
          numbering: { reference: "bullets3", level: 0 },
          spacing: { after: 60 },
          children: [new TextRun({ text: "Rebuild home and visitor bullpens with proper mound construction, drainage, and protective enclosures", size: 22 })]
        }),
        new Paragraph({
          numbering: { reference: "bullets3", level: 0 },
          spacing: { after: 200 },
          children: [new TextRun({ text: "Estimated cost: $15,000 \u2013 $30,000", bold: true, size: 22 })]
        }),

        // Cost Summary Table
        new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 300 }, children: [new TextRun("Cost Summary")] }),
        new Table({
          width: { size: CONTENT_WIDTH, type: WidthType.DXA },
          columnWidths: [4680, 4680],
          rows: [
            new TableRow({
              children: [
                headerCell("Phase", 4680),
                headerCell("Estimated Cost", 4680),
              ]
            }),
            new TableRow({
              children: [
                dataCell("Phase 1: Critical Safety & Playability", 4680),
                dataCell("$40,000 \u2013 $65,000", 4680, { align: AlignmentType.RIGHT }),
              ]
            }),
            new TableRow({
              children: [
                dataCell("Phase 2: Fencing & Field Boundaries", 4680, { shading: "F2F2F2" }),
                dataCell("$25,000 \u2013 $50,000", 4680, { shading: "F2F2F2", align: AlignmentType.RIGHT }),
              ]
            }),
            new TableRow({
              children: [
                dataCell("Phase 3: Bullpens & Finishing", 4680),
                dataCell("$15,000 \u2013 $30,000", 4680, { align: AlignmentType.RIGHT }),
              ]
            }),
            new TableRow({
              children: [
                dataCell("Total Estimated Project Cost", 4680, { shading: "1F3864", bold: true }),
                new TableCell({
                  borders,
                  width: { size: 4680, type: WidthType.DXA },
                  shading: { fill: "1F3864", type: ShadingType.CLEAR },
                  margins: cellMargins,
                  children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "$80,000 \u2013 $145,000", bold: true, color: "FFFFFF", font: "Arial", size: 22 })]
                  })]
                }),
              ]
            }),
            new TableRow({
              children: [
                dataCell("Booster Club Funds Available", 4680, { shading: "E2EFDA" }),
                dataCell("~$30,000", 4680, { shading: "E2EFDA", align: AlignmentType.RIGHT, bold: true }),
              ]
            }),
            new TableRow({
              children: [
                dataCell("Remaining Funding Needed", 4680, { shading: "FCE4EC" }),
                dataCell("$50,000 \u2013 $115,000", 4680, { shading: "FCE4EC", align: AlignmentType.RIGHT, bold: true }),
              ]
            }),
          ]
        }),

        // ===== SECTION 3: FUNDING =====
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("3. Funding & Partnership Proposal")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The Gig Harbor High School Baseball Booster Club has approximately $30,000 in funds currently available and committed to this project. We recognize that the total project cost will require significantly more investment and are requesting a partnership with the Peninsula School District to bridge the funding gap.",
            size: 22
          })]
        }),
        new Paragraph({
          spacing: { after: 100 },
          children: [new TextRun({ text: "We propose exploring the following funding sources together:", size: 22 })]
        }),
        new Paragraph({
          numbering: { reference: "bullets4", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "District Capital Improvement Budget: ", bold: true, size: 22 }),
            new TextRun({ text: "Allocate district capital improvement or athletic facility funds to the project", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets4", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "State Athletic Facility Grants: ", bold: true, size: 22 }),
            new TextRun({ text: "Research and apply for state-level grants available for public school athletic facility improvements", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets4", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Community Fundraising: ", bold: true, size: 22 }),
            new TextRun({ text: "Organize fundraising campaigns within the Gig Harbor community, including events, online campaigns, and donor outreach", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets4", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Corporate Sponsorships: ", bold: true, size: 22 }),
            new TextRun({ text: "Pursue sponsorship opportunities with local and regional businesses", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets4", level: 0 },
          spacing: { after: 200 },
          children: [
            new TextRun({ text: "Volunteer Labor: ", bold: true, size: 22 }),
            new TextRun({ text: "Utilize volunteer labor from the booster club, families, and community members where appropriate to reduce project costs", size: 22 })
          ]
        }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The Booster Club is open to a phased approach over multiple years if a single-year budget allocation is not feasible. Our priority is to begin with Phase 1 improvements as soon as possible to address the most critical safety and playability concerns.",
            size: 22
          })]
        }),

        // ===== SECTION 4: BENEFITS =====
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("4. Benefits of Renovation")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "Investing in the GHHS baseball field renovation will provide meaningful, lasting benefits to students, the school, and the broader community:",
            size: 22
          })]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("Student Safety")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The current warning track material and deteriorated fencing present real safety hazards to student-athletes. Renovation directly addresses these risks and provides a safe environment for play.",
            size: 22
          })]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("Home Games on Campus")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "Restoring the field allows the varsity team to play home games at GHHS for the first time. This builds school pride, increases student and community attendance, reduces transportation costs and logistics, and generates potential revenue from concessions and gate fees.",
            size: 22
          })]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("Facility Equity")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "Other athletic facilities within the district have received improvements and investment. The baseball field has been overlooked and deserves equitable attention to ensure all student-athletes have access to safe, quality facilities regardless of their sport.",
            size: 22
          })]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("Long-Term Community Asset")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "A renovated baseball field serves the community well beyond the varsity season. It can be used for junior varsity and freshman games, physical education classes, youth baseball programs, and community recreation. This is a long-term investment in school and community infrastructure.",
            size: 22
          })]
        }),

        // ===== SECTION 5: NATURAL SURFACE VS. TURF =====
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("5. Natural Surface vs. Synthetic Turf: Why This Approach")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "A fully turfed and lighted baseball field is an aspirational long-term goal that would serve the school and community well. However, it is important to consider the practical realities of that approach compared to the natural surface renovation proposed in this document.",
            size: 22
          })]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("The Synthetic Turf Alternative")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "A full synthetic turf baseball field with lighting is a major capital project. Turf field installations at the high school level typically range from $800,000 to over $1.5 million, depending on scope, site preparation, and lighting infrastructure. A project of this scale would almost certainly require passage of a school bond measure, which involves a multi-year process of planning, community engagement, ballot placement, and voter approval. Even under favorable circumstances, a turf field project could take five to ten years or longer from initial concept to completion.",
            size: 22
          })]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Benefits of a Natural Surface Field")] }),
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: "A properly built and maintained natural surface baseball field offers significant advantages:", size: 22 })]
        }),
        new Paragraph({
          numbering: { reference: "bullets6", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Aesthetics and Tradition: ", bold: true, size: 22 }),
            new TextRun({ text: "Natural grass and dirt infields are the traditional playing surface for baseball at every level, from youth leagues through the major leagues. A well-maintained natural field is widely regarded as the most attractive and authentic playing environment.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets6", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Player Safety and Performance: ", bold: true, size: 22 }),
            new TextRun({ text: "Natural surfaces provide predictable ball bounce and player footing. Many coaches, players, and sports medicine professionals prefer natural grass for reduced heat exposure and more natural playing conditions.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets6", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Cost-Effective Solution: ", bold: true, size: 22 }),
            new TextRun({ text: "The proposed renovation represents a fraction of the cost of a synthetic turf installation, making it achievable without a bond measure and within a realistic near-term budget.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets6", level: 0 },
          spacing: { after: 200 },
          children: [
            new TextRun({ text: "Achievable Timeline: ", bold: true, size: 22 }),
            new TextRun({ text: "With funding in place, the proposed natural surface renovation can be completed in a single summer/fall construction window, meaning student-athletes could be playing on their home field as early as the following spring season. This stands in stark contrast to the years-long timeline required for a turf project.", size: 22 })
          ]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("A Comparison")] }),
        new Table({
          width: { size: CONTENT_WIDTH, type: WidthType.DXA },
          columnWidths: [2340, 3510, 3510],
          rows: [
            new TableRow({
              children: [
                headerCell("", 2340),
                headerCell("Natural Surface (Proposed)", 3510),
                headerCell("Synthetic Turf Alternative", 3510),
              ]
            }),
            new TableRow({
              children: [
                dataCell("Estimated Cost", 2340, { bold: true }),
                dataCell("$80,000 \u2013 $145,000", 3510),
                dataCell("$800,000 \u2013 $1,500,000+", 3510),
              ]
            }),
            new TableRow({
              children: [
                dataCell("Timeline to Play", 2340, { bold: true, shading: "F2F2F2" }),
                dataCell("1 summer/fall season", 3510, { shading: "F2F2F2" }),
                dataCell("5\u201310+ years (bond required)", 3510, { shading: "F2F2F2" }),
              ]
            }),
            new TableRow({
              children: [
                dataCell("Funding Mechanism", 2340, { bold: true }),
                dataCell("Booster funds + district partnership", 3510),
                dataCell("School bond measure (voter approval)", 3510),
              ]
            }),
            new TableRow({
              children: [
                dataCell("Aesthetics", 2340, { bold: true, shading: "F2F2F2" }),
                dataCell("Traditional, authentic baseball look", 3510, { shading: "F2F2F2" }),
                dataCell("Modern but less traditional", 3510, { shading: "F2F2F2" }),
              ]
            }),
            new TableRow({
              children: [
                dataCell("Ongoing Maintenance", 2340, { bold: true }),
                dataCell("Annual (booster-funded)", 3510),
                dataCell("Lower annual, but costly replacement every 8\u201312 years", 3510),
              ]
            }),
          ]
        }),
        new Paragraph({
          spacing: { before: 200, after: 200 },
          children: [new TextRun({
            text: "The natural surface renovation is the right approach for right now. It gets our student-athletes back on their home field in the near term while being fiscally responsible. A turf conversion can remain a long-term aspiration that the district may pursue in the future through the bond process.",
            size: 22, italics: true
          })]
        }),

        // ===== SECTION 6: ONGOING MAINTENANCE & REVENUE =====
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("6. Ongoing Maintenance & Field Revenue")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "A natural surface baseball field requires consistent annual maintenance to remain in top condition. The Baseball Booster Club understands this responsibility and is prepared to take it on, provided a fair partnership framework is established with the district.",
            size: 22
          })]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Annual Maintenance Requirements")] }),
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: "Keeping the renovated field at a high standard will require ongoing investment in materials and equipment:", size: 22 })]
        }),
        new Paragraph({
          numbering: { reference: "bullets7", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Infield Mix Replenishment: ", bold: true, size: 22 }),
            new TextRun({ text: "Infield mix material must be added annually to maintain proper depth, consistency, and drainage performance.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets7", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Top Dressing: ", bold: true, size: 22 }),
            new TextRun({ text: "Regular top dressing of the infield skin is necessary to keep the surface smooth, level, and safe for play.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets7", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Specialty Equipment: ", bold: true, size: 22 }),
            new TextRun({ text: "Proper field maintenance requires specialized equipment including infield drags, grooming tools, mound/plate clay, line markers, and tarps. Some of this equipment represents a one-time purchase, while consumable materials are an annual cost.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets7", level: 0 },
          spacing: { after: 200 },
          children: [
            new TextRun({ text: "General Upkeep: ", bold: true, size: 22 }),
            new TextRun({ text: "Mound and home plate area maintenance, warning track grooming, fence repairs, and seasonal preparation all require regular attention.", size: 22 })
          ]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Booster Club Commitment to Maintenance")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The GHHS Baseball Booster Club is willing to take on the responsibility of funding and coordinating the annual maintenance of the field. This includes purchasing materials, acquiring and maintaining equipment, and organizing volunteer work crews for routine upkeep. Our goal is to ensure the field remains a facility that the school, its students, and the community can be proud of.",
            size: 22
          })]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Field Rental Revenue Proposal")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "To sustain the ongoing maintenance costs, we respectfully request that the district designate any revenue generated from rental of the baseball field directly back to the baseball program. This arrangement would create a self-sustaining maintenance model:",
            size: 22
          })]
        }),
        new Paragraph({
          numbering: { reference: "bullets8", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Off-Season Events & Tournaments: ", bold: true, size: 22 }),
            new TextRun({ text: "The baseball program would organize and manage off-season events, including youth tournaments, camps, and community rentals. The booster club would provide volunteer staff to run these events, reducing operational costs.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets8", level: 0 },
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Revenue Reinvestment: ", bold: true, size: 22 }),
            new TextRun({ text: "All rental revenue would be reinvested directly into field maintenance, ensuring the facility is kept to a high standard year after year without placing an additional burden on the district budget.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "bullets8", level: 0 },
          spacing: { after: 200 },
          children: [
            new TextRun({ text: "Accountability & Stewardship: ", bold: true, size: 22 }),
            new TextRun({ text: "The baseball program would serve as the primary steward of the facility, ensuring that events are well-managed, the field is protected during use, and maintenance standards are upheld.", size: 22 })
          ]
        }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "This model aligns the interests of the baseball program, the school, and the district. The booster club is incentivized to maintain the field at the highest level because it directly serves our athletes, and the district benefits from a well-maintained facility with no ongoing maintenance cost obligation. It is a true partnership.",
            size: 22
          })]
        }),

        // ===== SECTION 7: NEXT STEPS =====
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("7. Requested Next Steps")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({
            text: "The Booster Club respectfully requests the following actions to move this project forward:",
            size: 22
          })]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [
            new TextRun({ text: "Schedule a Field Assessment Meeting: ", bold: true, size: 22 }),
            new TextRun({ text: "Bring together district facilities staff, athletics administration, and Booster Club representatives to conduct an on-site assessment of the current field conditions.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [
            new TextRun({ text: "Obtain Professional Quotes: ", bold: true, size: 22 }),
            new TextRun({ text: "Engage qualified contractors to provide detailed estimates for each phase of the proposed renovation work.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [
            new TextRun({ text: "Develop a Timeline & Funding Plan: ", bold: true, size: 22 }),
            new TextRun({ text: "Collaborate on a realistic project timeline and identify specific funding sources and allocation strategies.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 100 },
          children: [
            new TextRun({ text: "Identify Available Grants & District Funds: ", bold: true, size: 22 }),
            new TextRun({ text: "Research and pursue any available state grants, district capital improvement funds, or other funding sources applicable to athletic facility improvements.", size: 22 })
          ]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 300 },
          children: [
            new TextRun({ text: "Establish a Project Committee: ", bold: true, size: 22 }),
            new TextRun({ text: "Form a committee with representatives from the district, school administration, athletics department, and Booster Club to oversee the project from planning through completion.", size: 22 })
          ]
        }),

        new Paragraph({
          spacing: { after: 200 },
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 12 } },
          children: [new TextRun({
            text: "The Gig Harbor High School Baseball Booster Club is committed to being an active, contributing partner in this effort. We are ready to invest our funds, our time, and our volunteer energy alongside the district to make this renovation a reality for our student-athletes and community.",
            size: 22, italics: true, color: "333333"
          })]
        }),

        // ===== APPENDIX =====
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Appendix")] }),
        new Paragraph({
          spacing: { after: 200 },
          children: [
            new TextRun({ text: "Appendix A: ", bold: true, size: 22 }),
            new TextRun({ text: "Photo documentation of current field conditions (to be attached)", size: 22, italics: true, color: "666666" })
          ]
        }),
        new Paragraph({
          spacing: { after: 200 },
          children: [
            new TextRun({ text: "Appendix B: ", bold: true, size: 22 }),
            new TextRun({ text: "Professional contractor estimates (to be obtained)", size: 22, italics: true, color: "666666" })
          ]
        }),
      ]
    }
  ]
});

const outputPath = "/Users/bsleeter/Documents/Documents - Benjamin's MacBook Pro/GH Baseball/Docs/GHHS Baseball Field Renovation Proposal.docx";

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log("Document created: " + outputPath);
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
