const fs = require('fs');
const path = require('path');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  ShadingType,
  ImageRun,
  convertInchesToTwip,
  PageBreak
} = require('docx');

// Image paths
const IMG_PLAYBOOK = 'C:/Users/AYUSH RAI/.gemini/antigravity/brain/36f60c86-87a0-4071-bdd7-879eab938421/.user_uploaded/media_1791970649624_f5540e9f.png';
const IMG_FIT = 'C:/Users/AYUSH RAI/.gemini/antigravity/brain/36f60c86-87a0-4071-bdd7-879eab938421/.user_uploaded/media_1791973272877_dd2e7399.png';
const IMG_GRID = 'C:/Users/AYUSH RAI/.gemini/antigravity/brain/36f60c86-87a0-4071-bdd7-879eab938421/.user_uploaded/media_1791973275709_5cb06f41.png';
const IMG_CHANNEL = 'C:/Users/AYUSH RAI/.gemini/antigravity/brain/36f60c86-87a0-4071-bdd7-879eab938421/.user_uploaded/media_1791973278260_9cdb993e.png';
const IMG_BEH = 'C:/Users/AYUSH RAI/.gemini/antigravity/brain/36f60c86-87a0-4071-bdd7-879eab938421/.user_uploaded/media_1791973280398_7fc48454.png';
const IMG_PAGE = 'C:/Users/AYUSH RAI/.gemini/antigravity/brain/36f60c86-87a0-4071-bdd7-879eab938421/.user_uploaded/media_1791973282897_4d22bf40.png';

// Theme Colors
const COLOR_PRIMARY = '1B365D';   // Deep Corporate Navy
const COLOR_SECONDARY = '2B6CB0'; // Medium Slate Blue
const COLOR_ACCENT = '2C7A57';    // Deep Forest Green
const COLOR_TEXT = '2D3748';      // Dark Neutral Slate
const COLOR_MUTED = '718096';     // Medium Gray
const COLOR_LIGHT_BG = 'F7FAFC';  // Light background tint
const COLOR_BORDER = 'CBD5E0';    // Subtle border gray
const COLOR_WHITE = 'FFFFFF';
const COLOR_HIGHLIGHT = 'EBF8FF'; // Soft Blue Callout

function createTitle(text) {
  return new Paragraph({
    text: text,
    heading: HeadingLevel.TITLE,
    spacing: { before: 240, after: 120 },
    alignment: AlignmentType.CENTER,
    run: {
      font: 'Calibri',
      size: 48,
      bold: true,
      color: COLOR_PRIMARY,
    }
  });
}

function createSubtitle(text) {
  return new Paragraph({
    text: text,
    spacing: { before: 0, after: 360 },
    alignment: AlignmentType.CENTER,
    run: {
      font: 'Calibri',
      size: 26,
      color: COLOR_SECONDARY,
      italics: true,
    }
  });
}

function createH1(text) {
  return new Paragraph({
    text: text,
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR_PRIMARY }
    },
    run: {
      font: 'Calibri',
      size: 32,
      bold: true,
      color: COLOR_PRIMARY,
    }
  });
}

function createH2(text) {
  return new Paragraph({
    text: text,
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    run: {
      font: 'Calibri',
      size: 26,
      bold: true,
      color: COLOR_SECONDARY,
    }
  });
}

function createH3(text) {
  return new Paragraph({
    text: text,
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 80 },
    run: {
      font: 'Calibri',
      size: 22,
      bold: true,
      color: COLOR_TEXT,
    }
  });
}

function createP(text, options = {}) {
  return new Paragraph({
    spacing: { before: options.before || 60, after: options.after || 100 },
    alignment: options.align || AlignmentType.LEFT,
    children: [
      new TextRun({
        text: text,
        font: 'Calibri',
        size: options.size || 22,
        color: options.color || COLOR_TEXT,
        bold: options.bold || false,
        italics: options.italics || false,
      })
    ]
  });
}

function createBullet(title, description) {
  return new Paragraph({
    bullet: { level: 0 },
    spacing: { before: 40, after: 60 },
    children: [
      new TextRun({
        text: title + ': ',
        font: 'Calibri',
        size: 22,
        bold: true,
        color: COLOR_PRIMARY,
      }),
      new TextRun({
        text: description,
        font: 'Calibri',
        size: 22,
        color: COLOR_TEXT,
      })
    ]
  });
}

function createCallout(title, bodyText) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 100, type: WidthType.PERCENTAGE },
            shading: { type: ShadingType.CLEAR, fill: COLOR_HIGHLIGHT },
            borders: {
              left: { style: BorderStyle.SINGLE, size: 24, color: COLOR_SECONDARY },
              top: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
            },
            margins: { top: 140, bottom: 140, left: 180, right: 180 },
            children: [
              new Paragraph({
                spacing: { after: 60 },
                children: [
                  new TextRun({
                    text: '★ ' + title,
                    font: 'Calibri',
                    size: 22,
                    bold: true,
                    color: COLOR_PRIMARY,
                  })
                ]
              }),
              new Paragraph({
                spacing: { after: 0 },
                children: [
                  new TextRun({
                    text: bodyText,
                    font: 'Calibri',
                    size: 21,
                    color: COLOR_TEXT,
                    italics: true,
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });
}

function createImageBox(imagePath, caption, targetWidth = 540) {
  const buf = fs.readFileSync(imagePath);
  // calculate height keeping aspect ratio
  const origW = buf.readUInt32BE(16);
  const origH = buf.readUInt32BE(20);
  const targetHeight = Math.round((origH / origW) * targetWidth);

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 80 },
      children: [
        new ImageRun({
          data: buf,
          type: 'png',
          transformation: {
            width: targetWidth,
            height: targetHeight,
          }
        })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [
        new TextRun({
          text: 'Figure: ' + caption,
          font: 'Calibri',
          size: 19,
          italics: true,
          color: COLOR_MUTED,
        })
      ]
    })
  ];
}

async function buildDoc() {
  const doc = new Document({
    creator: 'Commercial & Sales Operations Team',
    title: 'Court Lead MQL & Cadence Monitoring Plan',
    description: 'Executive Strategy & Implementation Blueprint for Top Management',
    styles: {
      default: {
        document: {
          run: {
            font: 'Calibri',
            size: 22,
            color: COLOR_TEXT,
          }
        }
      }
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(1),
              bottom: convertInchesToTwip(1),
              left: convertInchesToTwip(1),
              right: convertInchesToTwip(1),
            }
          }
        },
        children: [
          // HEADER / TITLE BLOCK
          createTitle('EXECUTIVE STRATEGY & OPERATIONAL BLUEPRINT'),
          createSubtitle('End-to-End Court Lead Qualification, Fit × Intent Scoring & Cadence Monitoring System'),
          
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 80, after: 300 },
            children: [
              new TextRun({ text: 'Prepared for: ', bold: true, color: COLOR_PRIMARY }),
              new TextRun({ text: 'Executive Leadership & Commercial Steering Committee  |  ' }),
              new TextRun({ text: 'Platform: ', bold: true, color: COLOR_PRIMARY }),
              new TextRun({ text: 'Frappe CRM Integration  |  ' }),
              new TextRun({ text: 'Date: ', bold: true, color: COLOR_PRIMARY }),
              new TextRun({ text: 'October 2026' })
            ]
          }),

          createCallout(
            'Strategic Executive Summary',
            'This blueprint presents an automated, data-driven revenue engine for our sports infrastructure & acrylic court surfacing business. By combining explicit buyer profile data (Fit) with real-time digital engagement signals (Intent), the system automatically categorizes every inbound inquiry into a 3×3 priority matrix, enforces rigid SLA response standards, orchestrates multi-channel cadences (Phone, WhatsApp, AI Voice), and holds territory dealers strictly accountable.'
          ),

          new Paragraph({ spacing: { after: 200 } }),

          // SECTION 1: BUSINESS CONTEXT
          createH1('1. Business Context: Transforming Inbound Lead Operations'),
          createP(
            'In the high-value sports flooring and court construction industry, deals range from ₹5 Lakhs to ₹1+ Crore. However, sales operations face three major bottlenecks when relying on manual processes:'
          ),
          createBullet(
            'Lead Dilution & Rep Misalignment',
            'Sales representatives spend up to 40% of their time calling low-budget private backyard owners, students, or vendors, while high-value institutional sports contractors and clubs sit waiting in the queue.'
          ),
          createBullet(
            'The "Speed-to-Lead" Crisis',
            'Industry data shows that responding within 4 hours increases conversion by over 300%. Without automation, hot inquiries often wait 24 to 72 hours before first contact.'
          ),
          createBullet(
            'The Dealer "Black Hole"',
            'Because territory dealers execute site surveys and quotes, deals frequently stall after sales qualification. Without real-time cadence monitoring, management has zero visibility into whether dealers accepted the lead, conducted the site survey, or sent the quotation on time.'
          ),

          createP(
            'The solution is a closed-loop system connecting our public web portal (AST), Frappe CRM, an automated Fit × Intent Scoring Engine, and a 24/7 Cadence & SLA Watchdog.'
          ),

          new Paragraph({ spacing: { after: 200 } }),

          // SECTION 2: THE LEAD SCORING FRAMEWORK
          createH1('2. The Lead Scoring Framework: Fit × Intent Methodology'),
          createP(
            'Every prospective buyer is evaluated across two independent, non-negotiable axes: Fit (Who they are) and Intent (How urgently they are looking to buy).'
          ),

          createH2('A. Fit Score (0–50): "Is this the kind of buyer we want?"'),
          createP(
            'Fit evaluates whether the prospective client matches our Ideal Customer Profile (ICP). It is calculated from explicit form inputs, organizational vetting, and project specifications.'
          ),

          ...createImageBox(
            IMG_FIT,
            'Fit Score Evaluation Signals & Weighting Table'
          ),

          createP('Key Strategic Rules for Fit Evaluation:', { bold: true, color: COLOR_PRIMARY }),
          createBullet(
            'Tier 1 Priority (+20 pts)',
            'Sports infrastructure contractors, schools/universities, municipal sports bodies, real estate developers, and sports clubs receive top weighting due to high repeat procurement and large ticket size.'
          ),
          createBullet(
            'Authority Weighting (+10 / +6 pts)',
            'Direct decision-makers (owners, project directors, procurement heads) are fast-tracked compared to general influencers (facility staff, students).'
          ),
          createBullet(
            'Scale Thresholds (+10 pts)',
            'Projects specifying 3+ courts or ≥ 2,000 m² automatically gain maximum volume bonus.'
          ),
          createBullet(
            'Geographic Serviceability (+5 / -15 pts)',
            'Projects inside our active dealer service regions gain +5 points. Inquiries outside our service network are penalized -15 points to prevent sales rep distraction.'
          ),
          createBullet(
            'Immediate Disqualification',
            'Competitors, job seekers, and students are assigned -50 points and automatically disqualified from sales outreach.'
          ),

          new Paragraph({ children: [new PageBreak()] }),

          createH2('B. Intent Score (0–50): "How ready are they to buy right now?"'),
          createP(
            'Intent measures buyer urgency based on active digital behavior rather than stated interest. The calculation follows a 4-step sequential formula:'
          ),

          createH3('Step 1: Deepest Page Reached (Base Intent: Up to 25 pts)'),
          createP(
            'To prevent visitors from inflating scores by aimlessly clicking, only the single deepest page reached in a session is counted.'
          ),

          ...createImageBox(
            IMG_PAGE,
            'Step A: Deepest Page Reached Scoring Tiers'
          ),

          createBullet(
            'High-Intent Conversions (+25 pts)',
            'Visiting Request a Quote, Site-Visit Request, or WhatsApp Click-to-Chat signals immediate procurement intention.'
          ),
          createBullet(
            'Evaluation Pages (+10 pts)',
            'Reviewing court specifications, system technical datasheets, or case studies indicates serious evaluation.'
          ),

          createH3('Step 2: High-Value Buyer Behaviors (Added Points)'),
          createP(
            'Every high-intent action performed on our digital platforms adds cumulative points.'
          ),

          ...createImageBox(
            IMG_BEH,
            'Step B: Behavioral Signals and Fast-Track Triggers'
          ),

          createBullet(
            'Interactive Estimator (+12 pts)',
            'Prospective clients who run numbers on our Court Cost Estimator demonstrate active budgeting.'
          ),
          createBullet(
            'Spec Sheet Downloads (+10 pts)',
            'Architects and contractors downloading technical certification sheets represent active bid preparation.'
          ),
          createBullet(
            'High-Velocity Return (+8 pts)',
            'Prospects returning to the website within 7 days across 2+ sessions indicate active decision cycles.'
          ),
          createBullet(
            'RFQ / Tender Document Attached',
            'Immediately triggers a Fast-Track override, bypassing standard nurture loops straight to Hot Priority.'
          ),

          new Paragraph({ children: [new PageBreak()] }),

          createH3('Step 3 & 4: Acquisition Channel Multiplier & Time Decay'),
          createP(
            'Raw intent points are adjusted based on where the visitor originated and decayed over time to ensure freshness.'
          ),

          ...createImageBox(
            IMG_CHANNEL,
            'Steps C & D: Acquisition Channel Multipliers and Inactivity Decay'
          ),

          createBullet(
            'Commercial Search Urgency (1.5×)',
            'Google searches containing buying keywords ("court contractor near me", "acrylic flooring cost") carry a 1.5× multiplier, recognizing high commercial intent.'
          ),
          createBullet(
            'Informational Search (0.8×)',
            'Searches for general information ("tennis court dimensions") are tempered with a 0.8× multiplier.'
          ),
          createBullet(
            'Referral & Direct Trust (1.3×)',
            'Direct URL visits, client referrals, and dealer introductions carry a 1.3× boost due to pre-established trust.'
          ),
          createBullet(
            'Automated Inactivity Decay (0.75× / 30 Days)',
            'If a prospect stops engaging, their intent score automatically decays by 25% every 30 days. Stale leads naturally fall out of active queues.'
          ),

          new Paragraph({ spacing: { after: 200 } }),

          // SECTION 3: THE 3x3 MATRIX
          createH1('3. The 3×3 Decision Grid: Automated Lead Routing'),
          createP(
            'By crossing the Fit Score (Axis A/B/C) with the Intent Score (Axis 1/2/3), the system automatically routes every lead to an exact operational owner and response timeline.'
          ),

          ...createImageBox(
            IMG_GRID,
            'The 3×3 Grid: Operational Lead Routing Matrix'
          ),

          createP('Grid Categorization Breakdown for Leadership:', { bold: true, color: COLOR_PRIMARY }),
          createBullet(
            'Grid Cell A1 [HOT]: Fit A (≥35) + Intent 1 (≥35)',
            'High-value target with immediate buying intent. SLA: First call within 4 hours. Dedicated attention from Inside Sales + Sales Team Lead involvement.'
          ),
          createBullet(
            'Grid Cells A2 & B1 [MQL]: High Fit / Moderate Intent OR Moderate Fit / High Intent',
            'Institutional buyer or urgent commercial inquiry. SLA: First call within 24 hours.'
          ),
          createBullet(
            'Grid Cell B2 [MQL Qualify]: Fit B (25–34) + Intent 2 (20–34)',
            'Moderate opportunity requiring discovery by inside sales. SLA: Call within 48 hours.'
          ),
          createBullet(
            'Grid Cell C1 [Dealer / Self-Serve]: Fit C (<25) + Intent 1 (≥35)',
            'Typically private individuals or single residential courts. Route directly to territory dealer within 24 hours or provide self-serve court pricing.'
          ),
          createBullet(
            'Grid Cells A3, B3, C2 [Nurture]: Fit A/B/C + Intent 3 (<20)',
            'No sales calls. Added to automated WhatsApp and email nurture campaigns. Rescored automatically when they show fresh activity.'
          ),
          createBullet(
            'Grid Cell C3 [No Action]',
            'Low fit and low intent. Archived to preserve sales team focus.'
          ),

          new Paragraph({ children: [new PageBreak()] }),

          // SECTION 4: THE FOLLOW-UP SYSTEM & CADENCE MONITORING
          createH1('4. The Follow-Up System & Cadence Monitoring Timeline'),
          createP(
            'This operational matrix represents the exact standard operating procedure (SOP) that will be monitored and enforced by the automated CRM system.'
          ),

          ...createImageBox(
            IMG_PLAYBOOK,
            'Complete Playbook Matrix: Multi-Touch Cadences, Dealer SLAs, and Milestone Timelines'
          ),

          createH2('Detailed Breakdown of the Playbook Timelines:'),

          createH3('1. Hot · A1 or Fast-Track Playbook (Top Priority)'),
          createBullet('First Response SLA', 'Mandatory phone contact within 4 hours of form submission.'),
          createBullet('Contact Cadence (6 Touches over 7 Days)', 
            'Day 0: Call 1 (Inside Sales) → Day 0: Call 2 + WhatsApp intro with full court catalogue & project photos → Day 1: Call 3 → Day 2: AI Voice Bot follow-up call + WhatsApp → Day 4: Call 4 + email with nearby completed court case study → Day 7: Direct call from Sales Team Lead.'
          ),
          createBullet('Recycle Deadline', 'Day 10. If prospect remains completely unresponsive after 6 touches, automatically transferred to Nurture (prevents reps hoarding dead leads).'),
          createBullet('Dealer Acceptance SLA', 'Assigned territory dealer must formally accept and acknowledge the lead within 24 hours.'),
          createBullet('Site Visit & Quote SLA', '3 days for on-site base survey + 3 days to deliver the official BOQ/Quotation (Total: 6 days).'),
          createBullet('Client Follow-up SLA', 'Inside sales checks in with client within 2 days after quotation submission.'),
          createBullet('Dealer Progress Update', 'Dealer must submit status updates every 3 days until PO or outcome.'),

          createH3('2. MQL · A2 or B1 Playbook (Standard High Priority)'),
          createBullet('First Response SLA', 'Contact within 24 hours.'),
          createBullet('Contact Cadence (5 Touches over 9 Days)', 
            'Day 0: Call 1 → Day 1: Call 2 + WhatsApp intro → Day 2: AI Voice Bot call → Day 5: Call 3 + Case Study email → Day 9: Final qualification call.'
          ),
          createBullet('Recycle Deadline', 'Day 14. Automatically moves to Nurture if unreached.'),
          createBullet('Dealer & Quote Timelines', 'Dealer accepts in 24 hrs. 7 days for site survey + 5 days for quote (Total: 12 days). Client follow-up in 3 days; dealer updates every 7 days.'),

          createH3('3. MQL · B2 Playbook (Inside Sales Qualifies)'),
          createBullet('First Response SLA', 'Contact within 48 hours.'),
          createBullet('Contact Cadence (5 Touches over 13 Days)', 
            'Day 0: Call 1 → Day 1: AI Bot call → Day 3: Call 2 + WhatsApp → Day 7: Email + WhatsApp with Court Cost Guide → Day 13: Final call.'
          ),
          createBullet('Recycle Deadline', 'Day 21. Handed over to automated email nurture if unresponsive.'),
          createBullet('Dealer & Quote Timelines', 'Dealer accepts in 48 hrs. 10 days for site survey + 7 days for quote. Client follow-up in 4 days; dealer updates every 7 days.'),

          createH3('4. C1 · Dealer / Self-Serve'),
          createBullet('Turnaround', 'Transferred directly to dealer; dealer must initiate call within 24 hours. Inside sales conducts an audit call with the dealer on Day 7 to verify customer contact.'),

          new Paragraph({ children: [new PageBreak()] }),

          // SECTION 5: HOW MANAGEMENT MONITORS THIS
          createH1('5. How Management Will Monitor & Enforce This in Real Time'),
          createP(
            'To ensure this plan does not remain "on paper", the Frappe CRM system incorporates four active monitoring mechanisms:'
          ),

          createH2('A. Automated SLA Countdown Clocks & Escalation'),
          createP(
            'Every lead enters the CRM with pre-calculated deadline timestamps (`first_response_due`, `dealer_accept_due`, `quote_due`).'
          ),
          createBullet(
            'Rep First Response Alarm',
            'If an A1 lead is not contacted within 4 hours, the CRM flags the card in red and immediately sends a priority escalation alert to the Sales Team Lead on Slack and WhatsApp.'
          ),
          createBullet(
            'Dealer 24-Hour Acceptance Watchdog',
            'When a dealer receives an SQL, a 24-hour timer starts. If the dealer fails to accept the lead within the window, an automated SMS/WhatsApp warning is dispatched. If unaccepted at 36 hours, the system automatically revokes the lead and reassigns it to the alternate dealer in that territory.'
          ),
          createBullet(
            'Dealer Stagnation Alert',
            'If 3 days (Hot) or 7 days (MQL) elapse without a logged update from the dealer, the system sends an automated WhatsApp prompt: "Status update required for Lead [Client Name] — Site Visit / Quote pending."'
          ),

          createH2('B. Intelligent Cadence Automation (Rep + WhatsApp + AI Voice Bot)'),
          createBullet(
            'Zero-Friction Logging',
            'When a rep logs a call outcome (e.g., "No Answer"), the system automatically schedules the exact next touch in the playbook, eliminating manual calendar guessing.'
          ),
          createBullet(
            'Multi-Channel Touch Execution',
            'On Day 0, the system triggers the WhatsApp API to dispatch the official AST Court Catalogue and project gallery. On Day 2, an AI Voice Agent initiates the follow-up call, seamlessly qualifying the prospect and updating the CRM record.'
          ),
          createBullet(
            'Instant Stage Advancement',
            'The moment a customer answers and confirms project details, the lead breaks out of the "Attempting Contact" cadence and transitions into "Connected · Discovery" for SQL qualification.'
          ),

          createH2('C. Monsoon Seasonality Buffer (Sports Flooring Industry Specific)'),
          createP(
            'Outdoor acrylic court surfacing cannot cure in rainy conditions. In our industry, unmonitored leads often stall during monsoon months and are mistakenly written off as lost.'
          ),
          createBullet(
            'Automated Monsoon Detection',
            'The system cross-references the project city with regional Indian monsoon windows (e.g. West Coast: June–Oct; North: July–Sept; Tamil Nadu: Oct–Dec).'
          ),
          createBullet(
            'Monsoon Slot Booking Workflow',
            'Instead of flagging delays as dealer failure, the CRM automatically switches the deal into a "Monsoon Booking" milestone, prompting sales to lock in a 20% advance payment to reserve an installation slot for post-monsoon execution.'
          ),

          new Paragraph({ spacing: { after: 200 } }),

          // SECTION 6: EXECUTIVE KPI DASHBOARD
          createH1('6. Executive KPI Dashboard & Management Metrics'),
          createP(
            'Top management will monitor commercial performance through four core executive metrics rendered directly in Frappe CRM:'
          ),

          createBullet(
            'Speed to Lead (Median First Response Time)',
            'Target: < 4 hours for Hot leads; < 24 hours for MQLs. Monitored weekly by rep and channel.'
          ),
          createBullet(
            'Cadence Compliance Rate (% On-Time Touches)',
            'Percentage of playbook touches executed on the scheduled day. Target: ≥ 90% team compliance.'
          ),
          createBullet(
            'Dealer Performance & Conversion Index',
            'Ranking territory dealers by: (1) On-time acceptance rate, (2) Days to quote delivery, and (3) Win rate. Dealers with higher compliance receive priority lead routing.'
          ),
          createBullet(
            'Touch Conversion Attribution',
            'Visibility into which touch closes the customer (e.g., Call 1 vs Day 2 AI Bot vs Day 7 Sales Lead Call), allowing continuous optimization of marketing spend.'
          ),

          new Paragraph({ spacing: { after: 200 } }),

          // SECTION 7: IMPLEMENTATION ROADMAP
          createH1('7. Implementation Roadmap & Rollout Plan'),
          createP(
            'The rollout is organized into four clean operational phases:'
          ),
          createBullet(
            'Phase 1: CRM Centralization & Intake (Weeks 1–2)',
            'Deploy Frappe CRM Lead module; connect web forms and WhatsApp webhook directly to CRM with custom fields for Fit, Intent, and Grid Cell.'
          ),
          createBullet(
            'Phase 2: Automated Scoring & Routing (Weeks 3–4)',
            'Implement Fit & Intent calculation engine, acquisition channel multipliers, and the 3×3 Grid assignment logic.'
          ),
          createBullet(
            'Phase 3: Cadence & Dealer SLA Watchdogs (Weeks 5–6)',
            'Activate automated 15-minute SLA background checks, dealer acceptance timers, auto-reassignment rules, and rep daily worklists.'
          ),
          createBullet(
            'Phase 4: AI Voice & Multi-Channel Integrations (Weeks 7–8)',
            'Deploy WhatsApp automated catalogue delivery and Day 2 AI Voice Agent follow-ups.'
          ),

          new Paragraph({ spacing: { after: 240 } }),

          createCallout(
            'Expected Strategic Outcomes',
            'Upon full rollout, the organization expects: (1) 300% faster response times for institutional buyers, (2) Zero leads dropped in dealer handoffs, (3) 40% reduction in sales rep administrative overhead, and (4) An estimated 25–35% increase in lead-to-opportunity conversion across acrylic sports flooring projects.'
          )
        ]
      }
    ]
  });

  const buffer = await Packer.toBuffer(doc);
  const outPath = 'C:/PACE/MQL/Court_Lead_MQL_Executive_Plan.docx';
  fs.writeFileSync(outPath, buffer);
  console.log('Successfully created Word document at:', outPath, 'Bytes:', buffer.length);

  // Also copy to artifact directory for persistent link
  const artifactDir = 'C:/Users/AYUSH RAI/.gemini/antigravity/brain/36f60c86-87a0-4071-bdd7-879eab938421/';
  const artifactOut = path.join(artifactDir, 'Court_Lead_MQL_Executive_Plan.docx');
  fs.writeFileSync(artifactOut, buffer);
  console.log('Successfully copied to artifact directory at:', artifactOut);
}

buildDoc().catch(err => {
  console.error('Error generating document:', err);
  process.exit(1);
});
