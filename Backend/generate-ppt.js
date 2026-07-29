const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 16:9 widescreen format

// Define custom color palette
const COLORS = {
    bg: "0D1117",      // Dark background
    card: "161B22",    // Sleek card color
    accent: "FF2D78",  // Theme Pink
    text: "E6EDF3",    // Light gray text
    muted: "7D8590",   // Muted gray
    blue: "8AB4F8"     // Highlight blue
};

// Load academic details configuration if available
let config = {
    studentName: "Your Name",
    rollNumber: "Your Roll Number",
    branch: "Computer Science & Engineering",
    collegeName: "Your College Name",
    guideName: "Professor's Name",
    guideDesignation: "Assistant Professor",
    projectTitle: "AI INTERVIEW PLAN\n& RESUME ANALYZER"
};

const configPaths = [
    path.join(__dirname, "../ppt-config.json"),
    path.join(__dirname, "ppt-config.json"),
    path.join(process.cwd(), "ppt-config.json")
];

for (const p of configPaths) {
    if (fs.existsSync(p)) {
        try {
            const data = JSON.parse(fs.readFileSync(p, "utf8"));
            config = { ...config, ...data };
            console.log(`Loaded presentation config from: ${p}`);
            break;
        } catch (e) {
            console.error(`Failed to parse config at ${p}:`, e);
        }
    }
}

// Helper function to apply dark theme layout to a slide
function createStandardSlide(title) {
    let slide = pres.addSlide();
    
    // Background color
    slide.background = { fill: COLORS.bg };
    
    // Top border line
    slide.addShape(pres.ShapeType.rect, {
        x: 0, y: 0, w: 13.33, h: 0.1, // Widescreen 16:9 is 13.33 x 7.5 inches
        fill: { color: COLORS.accent }
    });
    
    // Slide Title
    slide.addText(title, {
        x: 0.5, y: 0.4, w: 12.33, h: 0.7,
        fontSize: 24,
        bold: true,
        color: COLORS.accent,
        fontFace: "Segoe UI"
    });
    
    // Sub-title divider line
    slide.addShape(pres.ShapeType.rect, {
        x: 0.5, y: 1.1, w: 12.33, h: 0.02,
        fill: { color: "2A3348" }
    });
    
    // Footer page info
    slide.addText("AI Interview Planner & Resume Analyzer", {
        x: 0.5, y: 7.1, w: 6.0, h: 0.3,
        fontSize: 10,
        color: COLORS.muted,
        fontFace: "Segoe UI"
    });
    
    return slide;
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 1: Title Slide (Custom Academic & Dark Theme)
// ─────────────────────────────────────────────────────────────────────────────
let slide1 = pres.addSlide();
slide1.background = { fill: COLORS.bg };

// Decorative central card box
slide1.addShape(pres.ShapeType.rect, {
    x: 1.5, y: 0.8, w: 10.33, h: 5.6,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 2 }
});

// Top accent bar inside card
slide1.addShape(pres.ShapeType.rect, {
    x: 1.5, y: 0.8, w: 10.33, h: 0.15,
    fill: { color: COLORS.accent }
});

// Title Text
slide1.addText(config.projectTitle, {
    x: 1.5, y: 1.2, w: 10.33, h: 1.3,
    fontSize: 34,
    bold: true,
    color: COLORS.accent,
    fontFace: "Segoe UI",
    align: "center"
});

// Subtitle
slide1.addText("A Premium AI-Powered Interview Preparation & Roadmap Platform", {
    x: 1.5, y: 2.6, w: 10.33, h: 0.4,
    fontSize: 15,
    color: COLORS.text,
    fontFace: "Segoe UI",
    align: "center"
});

// Divider line between title and presenter info
slide1.addShape(pres.ShapeType.rect, {
    x: 2.5, y: 3.1, w: 8.33, h: 0.02,
    fill: { color: "2A3348" }
});

// Presenter Details (Left Column)
slide1.addText(
    [
        { text: "SUBMITTED BY:\n", options: { bold: true, color: COLORS.accent, fontSize: 11, fontFace: "Segoe UI" } },
        { text: `Name: `, options: { color: COLORS.muted, fontSize: 13, fontFace: "Segoe UI" } },
        { text: `${config.studentName}\n`, options: { color: COLORS.text, fontSize: 13, bold: true, fontFace: "Segoe UI" } },
        { text: `Roll No: `, options: { color: COLORS.muted, fontSize: 13, fontFace: "Segoe UI" } },
        { text: `${config.rollNumber}\n`, options: { color: COLORS.text, fontSize: 13, fontFace: "Segoe UI" } },
        { text: `Branch: `, options: { color: COLORS.muted, fontSize: 13, fontFace: "Segoe UI" } },
        { text: `${config.branch}`, options: { color: COLORS.text, fontSize: 13, fontFace: "Segoe UI" } }
    ],
    { x: 2.2, y: 3.4, w: 4.2, h: 1.8, lineSpacing: 22 }
);

// Guide Details (Right Column)
slide1.addText(
    [
        { text: "GUIDED BY:\n", options: { bold: true, color: COLORS.blue, fontSize: 11, fontFace: "Segoe UI" } },
        { text: `Supervisor: `, options: { color: COLORS.muted, fontSize: 13, fontFace: "Segoe UI" } },
        { text: `${config.guideName}\n`, options: { color: COLORS.text, fontSize: 13, bold: true, fontFace: "Segoe UI" } },
        { text: `Designation: `, options: { color: COLORS.muted, fontSize: 13, fontFace: "Segoe UI" } },
        { text: `${config.guideDesignation}\n`, options: { color: COLORS.text, fontSize: 13, fontFace: "Segoe UI" } },
        { text: "Department: ", options: { color: COLORS.muted, fontSize: 13, fontFace: "Segoe UI" } },
        { text: "Computer Science", options: { color: COLORS.text, fontSize: 13, fontFace: "Segoe UI" } }
    ],
    { x: 6.8, y: 3.4, w: 4.2, h: 1.8, lineSpacing: 22 }
);

// College Name (Bottom Row)
slide1.addText(config.collegeName, {
    x: 1.5, y: 5.4, w: 10.33, h: 0.6,
    fontSize: 14,
    bold: true,
    color: COLORS.accent,
    fontFace: "Segoe UI",
    align: "center"
});


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 2: Executive Summary
// ─────────────────────────────────────────────────────────────────────────────
let slide2 = createStandardSlide("Executive Summary");
slide2.addText(
    "Traditional job hunting involves sending generic resumes to highly specific job roles. Job seekers struggle to identify skill deficiencies and lack target-driven study strategies.\n\n" +
    "This platform bridges that gap using generative AI to analyze target Job Descriptions (JD) against a candidate's Resume or Profile. " +
    "It compiles an active study program consisting of precise QA cards and a structured daily roadmap, enabling candidates to successfully match expectations.",
    { x: 0.8, y: 1.8, w: 5.5, h: 4.5, fontSize: 16, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 24 }
);

// Problem / Solution Box
slide2.addShape(pres.ShapeType.rect, {
    x: 7.0, y: 1.8, w: 5.5, h: 4.5,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide2.addText("PLATFORM VALUE PROPOSITION", {
    x: 7.3, y: 2.1, w: 4.9, h: 0.4,
    fontSize: 14, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slide2.addText(
    "•  100% Tailored Preparation: Eliminates generic web search questions.\n" +
    "•  Actionable Feedback: Identifies exact technical and behavioral skill gaps.\n" +
    "•  Interactive Roadmap: Generates daily training tasks, complete with topic focuses.\n" +
    "•  Portable Outputs: Compiles print-ready PDF resumes styled for ATS evaluation.",
    { x: 7.3, y: 2.7, w: 4.9, h: 3.3, fontSize: 13, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 22 }
);


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 3: Core Features
// ─────────────────────────────────────────────────────────────────────────────
let slide3 = createStandardSlide("Core Features");
const features = [
    {
        title: "Resume Processing",
        desc: "Drag-and-drop file upload. The PDF parser extracts plain text to match and calculate a dynamic profile score against the target JD."
    },
    {
        title: "Tailored QA Cards",
        desc: "Generates custom Technical and Behavioral questions. Includes detailed interviewer intentions and comprehensive model answers."
    },
    {
        title: "Interactive Roadmap",
        desc: "Provides a day-by-day preparation calendar, focusing on closing skill gaps and reinforcing strengths before the interview."
    }
];

features.forEach((feat, idx) => {
    let cardX = 0.5 + idx * 4.2;
    slide3.addShape(pres.ShapeType.rect, {
        x: cardX, y: 1.8, w: 3.9, h: 4.5,
        fill: { color: COLORS.card },
        line: { color: "2A3348", width: 1 }
    });
    slide3.addShape(pres.ShapeType.rect, {
        x: cardX, y: 1.8, w: 3.9, h: 0.1,
        fill: { color: COLORS.accent }
    });
    slide3.addText(`0${idx + 1}`, {
        x: cardX + 0.3, y: 2.1, w: 3.3, h: 0.5,
        fontSize: 22, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
    });
    slide3.addText(feat.title, {
        x: cardX + 0.3, y: 2.7, w: 3.3, h: 0.5,
        fontSize: 18, bold: true, color: COLORS.text, fontFace: "Segoe UI"
    });
    slide3.addText(feat.desc, {
        x: cardX + 0.3, y: 3.4, w: 3.3, h: 2.6,
        fontSize: 13, color: COLORS.muted, fontFace: "Segoe UI", lineSpacing: 20
    });
});


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 4: System Architecture
// ─────────────────────────────────────────────────────────────────────────────
let slide4 = createStandardSlide("System Architecture");
slide4.addText(
    "The application is structured as a decoupled client-server architecture built on a modern JavaScript stack. Express orchestrates the request routing while React handles state and responsive layout rendering.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

const archBlocks = [
    { title: "React / Vite Frontend", desc: "User inputs (JD & PDF resume) are collected. Axios forwards them to the API. Handles local JWT cookies and displays reports.", color: "2A3348" },
    { title: "Express / Node API", desc: "Manages endpoint access, handles file upload streams via Multer, registers users, hashes credentials, and sanitizes input data.", color: "2A3348" },
    { title: "MongoDB Database", desc: "Persists user documents, session history, and cached interview plans (storing parsed text and generated JSON objects).", color: "2A3348" },
    { title: "Google Gemini AI", desc: "Processes inputs using gemini-3.5-flash and responds with structured data matching our strictly validated Zod Schema.", color: "ff2d78" }
];

archBlocks.forEach((block, idx) => {
    let blockX = 0.5 + idx * 3.1;
    slide4.addShape(pres.ShapeType.rect, {
        x: blockX, y: 2.3, w: 2.9, h: 4.2,
        fill: { color: COLORS.card },
        line: { color: block.color, width: 2 }
    });
    slide4.addText(block.title, {
        x: blockX + 0.2, y: 2.6, w: 2.5, h: 0.8,
        fontSize: 16, bold: true, color: block.color === "ff2d78" ? COLORS.accent : COLORS.blue, fontFace: "Segoe UI", align: "center"
    });
    slide4.addText(block.desc, {
        x: blockX + 0.2, y: 3.6, w: 2.5, h: 2.6,
        fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 18, align: "center"
    });
});


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 5: System Workflow & Data Flow
// ─────────────────────────────────────────────────────────────────────────────
let slideWorkflow = createStandardSlide("System Workflow & Data Flow");
slideWorkflow.addText(
    "The application processes client data through a secure, pipeline-based flow. Here is how input documents are transformed into structured, actionable interview roadmaps.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

const steps = [
    { title: "1. Document Upload", desc: "User uploads PDF resume & enters Job Description on the React client.", active: false },
    { title: "2. Backend Parsing", desc: "Express parses the PDF using pdf-parse and cleans the extracted text.", active: false },
    { title: "3. Gemini API Call", desc: "Text is sent to Gemini 3.5 Flash along with a strict Zod JSON response schema.", active: true },
    { title: "4. Report Creation", desc: "The structured JSON response is persisted to MongoDB and returned to the client.", active: false }
];

steps.forEach((step, idx) => {
    let blockX = 0.5 + idx * 3.1;
    slideWorkflow.addShape(pres.ShapeType.rect, {
        x: blockX, y: 2.3, w: 2.9, h: 4.2,
        fill: { color: COLORS.card },
        line: { color: step.active ? COLORS.accent : "2A3348", width: step.active ? 2 : 1 }
    });
    
    // Step number bubble
    slideWorkflow.addText(`0${idx + 1}`, {
        x: blockX + 0.2, y: 2.6, w: 2.5, h: 0.4,
        fontSize: 18, bold: true, color: step.active ? COLORS.accent : COLORS.blue, fontFace: "Segoe UI", align: "center"
    });
    
    slideWorkflow.addText(step.title, {
        x: blockX + 0.2, y: 3.1, w: 2.5, h: 0.6,
        fontSize: 15, bold: true, color: COLORS.text, fontFace: "Segoe UI", align: "center"
    });
    
    slideWorkflow.addText(step.desc, {
        x: blockX + 0.2, y: 3.9, w: 2.5, h: 2.3,
        fontSize: 12, color: COLORS.muted, fontFace: "Segoe UI", lineSpacing: 18, align: "center"
    });
    
    // Add connector arrow between blocks if not the last block
    if (idx < 3) {
        slideWorkflow.addText("→", {
            x: blockX + 2.85, y: 4.0, w: 0.4, h: 0.5,
            fontSize: 24, bold: true, color: COLORS.muted, align: "center"
        });
    }
});


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 6: System Requirements & Software Stack
// ─────────────────────────────────────────────────────────────────────────────
let slideRequirements = createStandardSlide("System Requirements & Software Stack");
slideRequirements.addText(
    "Developing and running the AI Interview Plan & Resume Analyzer requires a compatible hardware and software stack to support concurrent web clients and AI processing.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

// Hardware Requirements Card
slideRequirements.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.3, w: 5.8, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slideRequirements.addText("HARDWARE SPECIFICATIONS", {
    x: 0.8, y: 2.6, w: 5.2, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slideRequirements.addText(
    "•  Processor: Dual-core Intel i5 or AMD Ryzen 5 (Minimum 2.0 GHz)\n" +
    "•  Memory: 8 GB RAM minimum (16 GB recommended for concurrent Puppeteer browser rendering)\n" +
    "•  Storage: 500 MB of free storage space for dependencies and local code\n" +
    "•  Network: Active high-speed internet connection for API integration with Gemini and MongoDB Atlas.",
    { x: 0.8, y: 3.2, w: 5.2, h: 3.0, fontSize: 13, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 22 }
);

// Software Requirements Card
slideRequirements.addShape(pres.ShapeType.rect, {
    x: 6.8, y: 2.3, w: 6.0, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slideRequirements.addText("SOFTWARE ENVIRONMENT & STACK", {
    x: 7.1, y: 2.6, w: 5.4, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.blue, fontFace: "Segoe UI"
});
slideRequirements.addText(
    "•  Operating System: Windows 10/11, macOS Big Sur+, or Ubuntu 20.04+ LTS\n" +
    "•  Runtime: Node.js (v18.x or v20.x LTS) & npm (v9.x or v10.x)\n" +
    "•  Database: MongoDB Server (v6.0+) or MongoDB Atlas Cloud Account\n" +
    "•  Client Core: React 18, Vite Bundler, SCSS styling, React Router DOM v6\n" +
    "•  Server Core: Express.js, Mongoose ODM, Zod Validator, JWT Cookies\n" +
    "•  AI Service: @google/genai SDK (Gemini API with gemini-3.5-flash)\n" +
    "•  PDF Compiler: puppeteer-core & @sparticuz/chromium lambda-binary",
    { x: 7.1, y: 3.2, w: 5.4, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 5: Frontend Design & UI Styling
// ─────────────────────────────────────────────────────────────────────────────
let slide5 = createStandardSlide("Frontend Design & Custom SCSS/CSS Theme");
slide5.addText(
    "A clean, premium UI is key to user retention. The application implements custom styling (SCSS) built from the ground up without third-party frameworks like Tailwind, ensuring complete control over the layout.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

slide5.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.3, w: 5.8, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide5.addText("VISUAL DESIGN SYSTEM", {
    x: 0.8, y: 2.6, w: 5.2, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slide5.addText(
    "•  Dark Page Canvas: Hex #0D1117 reduces eye strain during long prep sessions.\n" +
    "•  Elevation Cards: Hex #161B22 stands out, separating forms and reports.\n" +
    "•  Accent Pink: Hex #FF2D78 highlights key CTAs, buttons, and matching metrics.\n" +
    "•  Typography Stack: System-UI fallback stack ensures instant font loads and clean readability.",
    { x: 0.8, y: 3.2, w: 5.2, h: 3.0, fontSize: 13, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 22 }
);

slide5.addShape(pres.ShapeType.rect, {
    x: 6.8, y: 2.3, w: 6.0, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide5.addText("COMPOSABLE COMPONENT ARCHITECTURE", {
    x: 7.1, y: 2.6, w: 5.4, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.blue, fontFace: "Segoe UI"
});
slide5.addText(
    "•  Drag-and-Drop Area: Interactive drag zone listens to HTML5 file listeners.\n" +
    "•  Responsive Grid: Side-by-side split input forms stack dynamically on mobile.\n" +
    "•  Q&A Accordion: Accordion cards reveal Intention and Model Answer on click.\n" +
    "•  Sidebar Statistics: Score circle and color-coded skill tags update reactively.",
    { x: 7.1, y: 3.2, w: 5.4, h: 3.0, fontSize: 13, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 22 }
);


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 6: Backend & REST API Design
// ─────────────────────────────────────────────────────────────────────────────
let slide6 = createStandardSlide("Backend REST API Architecture");
slide6.addText(
    "The backend serves as a stateless RESTful service. Routes are logically structured into specific router groups, utilizing middleware for security, file handling, and cross-site request validation.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

slide6.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.3, w: 5.8, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide6.addText("AUTHENTICATION ENDPOINTS", {
    x: 0.8, y: 2.6, w: 5.2, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slide6.addText(
    "•  POST /api/auth/register : Registers a new user, hashes password, saves to DB, returns JWT cookie.\n" +
    "•  POST /api/auth/login : Validates credentials, signs 24h token cookie.\n" +
    "•  GET /api/auth/logout : Clears cookie, blacklists the token in DB.\n" +
    "•  POST /api/auth/reset-password : Self-service recovery via email/username verification.",
    { x: 0.8, y: 3.2, w: 5.2, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);

slide6.addShape(pres.ShapeType.rect, {
    x: 6.8, y: 2.3, w: 6.0, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide6.addText("INTERVIEW & UTILITIES ENDPOINTS", {
    x: 7.1, y: 2.6, w: 5.4, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.blue, fontFace: "Segoe UI"
});
slide6.addText(
    "•  POST /api/interview/generate : Processes Multer file inputs, calls PDF parser, communicates with Gemini, saves and returns report.\n" +
    "•  GET /api/interview/report/:id : Retrieves full structured data for the report.\n" +
    "•  GET /api/interview/reports : Fetches all historical reports (lightweight details).\n" +
    "•  POST /api/interview/resume/pdf/:id : Backend export route utilizing Puppeteer to compile HTML resume structure into an ATS PDF.",
    { x: 7.1, y: 3.2, w: 5.4, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 7: AI Service Integration
// ─────────────────────────────────────────────────────────────────────────────
let slide7 = createStandardSlide("AI Service Integration & Prompt Security");
slide7.addText(
    "Integrating LLMs into production requires strict control over input parameters, structured output formats, and error handling for rate limits or credential exceptions.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

slide7.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.3, w: 5.8, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide7.addText("GEMINI 3.5 FLASH SCHEMA CONTROL", {
    x: 0.8, y: 2.6, w: 5.2, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slide7.addText(
    "•  Strict Response Schema: The system defines standard schemas using Zod validation. Gemini is configured to output strict JSON matching that format.\n" +
    "•  Zod to JSON Translation: We convert schemas dynamically using zod-to-json-schema to configure the responseMimeType: 'application/json' API parameter.\n" +
    "•  Direct Parsing: Eliminates Markdown backtick wrappers, allowing direct parsing via JSON.parse() on API response text.",
    { x: 0.8, y: 3.2, w: 5.2, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);

slide7.addShape(pres.ShapeType.rect, {
    x: 6.8, y: 2.3, w: 6.0, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide7.addText("ROBUST RETRIES & QUOTA TOLERANCE", {
    x: 7.1, y: 2.6, w: 5.4, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.blue, fontFace: "Segoe UI"
});
slide7.addText(
    "•  API Key Rotation: Supports a comma-separated list of keys, rotating to the next key if a rate limit or quota exception is reached.\n" +
    "•  Exponential Backoff: Implement automated delays (with random jitter) on rate limits before trying again, preventing network dropouts.\n" +
    "•  Permission Handlers: Automatically intercepts 403 status errors (e.g. account suspension) to expose clear feedback to users.",
    { x: 7.1, y: 3.2, w: 5.4, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 8: PDF Compilation Engine (Cloud Fix)
// ─────────────────────────────────────────────────────────────────────────────
let slide8 = createStandardSlide("PDF Generation Engine & Cloud Caching Fix");
slide8.addText(
    "Exporting ATS-compliant resumes requires rendering HTML layouts into PDFs. However, standard headless browsers like Puppeteer crash when deployed to minimal cloud platforms like Render.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

slide8.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.3, w: 5.8, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide8.addText("THE RENDER DEPLOYMENT CHALLENGE", {
    x: 0.8, y: 2.6, w: 5.2, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slide8.addText(
    "•  Missing OS Libraries: Render's default Node Linux images lack required libraries (e.g., libatk, libnss3, libcups), causing Puppeteer to fail to launch.\n" +
    "•  Ephemeral Browser Binaries: Default Puppeteer downloads Chromium to a global ~/.cache directory, which gets ignored during Render's build-to-run container migration, causing 'browser not found' errors.",
    { x: 0.8, y: 3.2, w: 5.2, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);

slide8.addShape(pres.ShapeType.rect, {
    x: 6.8, y: 2.3, w: 6.0, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "ff2d78", width: 2 }
});
slide8.addText("THE ROBUST SOLUTIONS APPLIED", {
    x: 7.1, y: 2.6, w: 5.4, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.blue, fontFace: "Segoe UI"
});
slide8.addText(
    "•  Migrated to core dependencies: Replaced puppeteer package with puppeteer-core and @sparticuz/chromium.\n" +
    "•  Lambda-Optimized Binary: @sparticuz/chromium contains pre-compiled, statically linked binaries requiring no OS libraries.\n" +
    "•  Cross-Platform Compatibility Check: Implements an isLocal check. In production, it launches the pre-compiled binary. In local Windows/macOS, it automatically searches the system application paths for your local Chrome app.",
    { x: 7.1, y: 3.2, w: 5.4, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 9: Mongoose Database Schemas
// ─────────────────────────────────────────────────────────────────────────────
let slide9 = createStandardSlide("Database Models & Schema Definitions");
slide9.addText(
    "We use MongoDB Atlas to store persistent objects. Strict validation is managed using Mongoose schemas. User relations are mapped using standard ObjectIds.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

slide9.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.3, w: 3.9, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide9.addText("USER SCHEMA", {
    x: 0.8, y: 2.6, w: 3.3, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slide9.addText(
    "•  username: String\n" +
    "   - unique: true\n" +
    "   - required: true\n" +
    "•  email: String\n" +
    "   - unique: true\n" +
    "   - required: true\n" +
    "•  password: String\n" +
    "   - hashed (bcryptjs)\n" +
    "   - required: true",
    { x: 0.8, y: 3.2, w: 3.3, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Courier New", lineSpacing: 20 }
);

slide9.addShape(pres.ShapeType.rect, {
    x: 4.7, y: 2.3, w: 8.1, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide9.addText("INTERVIEW REPORT SCHEMA", {
    x: 5.0, y: 2.6, w: 7.5, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.blue, fontFace: "Segoe UI"
});
slide9.addText(
    "•  user: mongoose.Schema.Types.ObjectId (Reference to users collection)\n" +
    "•  title: String (Target job position title)\n" +
    "•  jobDescription: String (Required job duties & stack details)\n" +
    "•  resume: String (Extracted candidate resume plain text)\n" +
    "•  selfDescription: String (Optional user candidate details)\n" +
    "•  matchScore: Number (Calculated match fit percentage: 0 to 100)\n" +
    "•  technicalQuestions: Array of { question, intention, answer }\n" +
    "•  behavioralQuestions: Array of { question, intention, answer }\n" +
    "•  skillGaps: Array of { skill, severity: ['low', 'medium', 'high'] }\n" +
    "•  preparationPlan: Array of { day: Number, focus: String, tasks: Array[String] }",
    { x: 5.0, y: 3.2, w: 7.5, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Courier New", lineSpacing: 18 }
);


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 10: Authentication Security & Recovery
// ─────────────────────────────────────────────────────────────────────────────
let slide10 = createStandardSlide("Authentication Security & Recovery Flow");
slide10.addText(
    "User authentication must remain secure across different subdomains. The platform uses JSON Web Tokens (JWT) stored in HTTP-only cookies, combined with security handlers.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

slide10.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.3, w: 5.8, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide10.addText("CROSS-SITE COOKIE HANDLING", {
    x: 0.8, y: 2.6, w: 5.2, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slide10.addText(
    "•  Secure JWT Cookie: Signed token with 24-hour expiration containing user identity claims is saved inside HTTP-only cookies, blocking JS script theft.\n" +
    "•  Production Cross-Site Settings: In production, cookies are configured with secure: true and sameSite: 'none' properties, enabling Safari/Chrome to validate sessions when frontend and backend are hosted on separate domains.\n" +
    "•  Token Blacklist Store: When users log out, the token is added to a blacklist database store, preventing replay attacks.",
    { x: 0.8, y: 3.2, w: 5.2, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);

slide10.addShape(pres.ShapeType.rect, {
    x: 6.8, y: 2.3, w: 6.0, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide10.addText("SELF-SERVICE PASSWORD RECOVERY", {
    x: 7.1, y: 2.6, w: 5.4, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.blue, fontFace: "Segoe UI"
});
slide10.addText(
    "•  Verification: Since no external email SMTP service is configured, users reset credentials by verifying both their Email and unique Username.\n" +
    "•  Secure Overwrite: If verified, the server generates a new hash via bcryptjs and safely saves the updated password to the user document, deleting/clearing the previous credentials from the database.",
    { x: 7.1, y: 3.2, w: 5.4, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 11: Project Screenshot Proof - Dashboard
// ─────────────────────────────────────────────────────────────────────────────
let slide11 = createStandardSlide("Project Proof: Dashboard Interface");
slide11.addText(
    "Proof of the functional local frontend interface showing the drag-and-drop resume upload zone, target job description form, and model selection.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

const imgPath1 = "C:\\Users\\Never_Mind\\.gemini\\antigravity-ide\\brain\\d167864c-25cd-435d-b055-b2576a8b9ae5\\initial_page_load_1785070358594.png";
if (fs.existsSync(imgPath1)) {
    slide11.addImage({
        path: imgPath1,
        x: 2.66, y: 2.2, w: 8.0, h: 4.5
    });
} else {
    slide11.addShape(pres.ShapeType.rect, {
        x: 2.66, y: 2.2, w: 8.0, h: 4.5,
        fill: { color: COLORS.card },
        line: { color: COLORS.accent, width: 1 }
    });
    slide11.addText("Screenshot: initial_page_load.png", {
        x: 2.66, y: 4.2, w: 8.0, h: 0.5,
        fontSize: 14, color: COLORS.muted, align: "center"
    });
}


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 12: Project Screenshot Proof - Main App Setup
// ─────────────────────────────────────────────────────────────────────────────
let slide12 = createStandardSlide("Project Proof: Resume Planning Form");
slide12.addText(
    "Proof showing the landing form state with populated details, character limits, and system parameters ready for AI evaluation.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

const imgPath2 = "C:\\Users\\Never_Mind\\.gemini\\antigravity-ide\\brain\\d167864c-25cd-435d-b055-b2576a8b9ae5\\media__1785145741721.png";
if (fs.existsSync(imgPath2)) {
    slide12.addImage({
        path: imgPath2,
        x: 2.66, y: 2.2, w: 8.0, h: 4.5
    });
} else {
    slide12.addShape(pres.ShapeType.rect, {
        x: 2.66, y: 2.2, w: 8.0, h: 4.5,
        fill: { color: COLORS.card },
        line: { color: COLORS.accent, width: 1 }
    });
    slide12.addText("Screenshot: media__1785145741721.png", {
        x: 2.66, y: 4.2, w: 8.0, h: 0.5,
        fontSize: 14, color: COLORS.muted, align: "center"
    });
}


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 13: Mobile Responsive Layout Updates
// ─────────────────────────────────────────────────────────────────────────────
let slide13 = createStandardSlide("Mobile Responsiveness Improvements");
slide13.addText(
    "Mobile user experience was significantly improved by adding responsive media queries (@media (max-width: 900px) and @media (max-width: 768px)) to adapt pages for mobile phone viewports.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

slide13.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.3, w: 3.9, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide13.addText("AUTH SCREEN FIXES", {
    x: 0.8, y: 2.6, w: 3.3, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slide13.addText(
    "•  Container widths changed from rigid min-width: 350px to responsive width: 100% and max-width: 380px.\n" +
    "•  Prevents form clipping on narrow displays like iPhone SE (320px width).\n" +
    "•  Input fields, labels, and password toggles scale fluidly.",
    { x: 0.8, y: 3.2, w: 3.3, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);

slide13.addShape(pres.ShapeType.rect, {
    x: 4.7, y: 2.3, w: 3.9, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide13.addText("DASHBOARD LAYOUT", {
    x: 5.0, y: 2.6, w: 3.3, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.blue, fontFace: "Segoe UI"
});
slide13.addText(
    "•  Collapses side-by-side flexbox columns (Job description vs resume) into a single vertical scroll.\n" +
    "•  The vertical dividers are hidden to maximize horizontal layout space.\n" +
    "•  Action buttons scale to 100% width for easier thumb tapping on mobile screens.",
    { x: 5.0, y: 3.2, w: 3.3, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);

slide13.addShape(pres.ShapeType.rect, {
    x: 8.9, y: 2.3, w: 3.9, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide13.addText("REPORT DETAILS VIEW", {
    x: 9.2, y: 2.6, w: 3.3, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slide13.addText(
    "•  Exchanges the 3-column desktop layout (nav, content, sidebar) for a vertical order.\n" +
    "•  Pushes Match Score and Gaps to the top using CSS order: 2.\n" +
    "•  Converts vertical navigation buttons into horizontal top tabs.",
    { x: 9.2, y: 3.2, w: 3.3, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 14: Deployed Environment Configuration
// ─────────────────────────────────────────────────────────────────────────────
let slide14 = createStandardSlide("Render Deployment & Active Hosts");
slide14.addText(
    "The application is fully hosted and running on Render cloud infrastructure, configured with automatic GitHub build triggers.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

slide14.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.3, w: 5.8, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide14.addText("FRONTEND STATIC HOSTING", {
    x: 0.8, y: 2.6, w: 5.2, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slide14.addText(
    "•  Service Type: Render Static Site\n" +
    "•  Live Web Address: https://genairesume-ats.onrender.com\n" +
    "•  Build Command: cd Frontend && npm install && npm run build\n" +
    "•  Publish Directory: Frontend/dist\n" +
    "•  Security: Automatically served over HTTPS with TLS 1.3.",
    { x: 0.8, y: 3.2, w: 5.2, h: 3.0, fontSize: 13, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 22 }
);

slide14.addShape(pres.ShapeType.rect, {
    x: 6.8, y: 2.3, w: 6.0, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide14.addText("BACKEND WEB SERVICE HOSTING", {
    x: 7.1, y: 2.6, w: 5.4, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.blue, fontFace: "Segoe UI"
});
slide14.addText(
    "•  Service Type: Render Web Service\n" +
    "•  Live Web API: https://resume-maker-project-okii.onrender.com\n" +
    "•  Build Command: cd Backend && npm install\n" +
    "•  Start Command: cd Backend && node server.js\n" +
    "•  CORS whitelist: Allows cross-origin requests specifically from the frontend site.",
    { x: 7.1, y: 3.2, w: 5.4, h: 3.0, fontSize: 13, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 22 }
);


// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 15: Conclusion & Future Scope
// ─────────────────────────────────────────────────────────────────────────────
let slide15 = createStandardSlide("Conclusion & Future Work");
slide15.addText(
    "The AI Interview Plan & Resume Analyzer is fully implemented, verified, responsive, and deployed in production. It offers immense value to job applicants looking for structured preparation.",
    { x: 0.5, y: 1.5, w: 12.33, h: 0.6, fontSize: 14, color: COLORS.text, fontFace: "Segoe UI" }
);

slide15.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.3, w: 5.8, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide15.addText("PROJECT SUMMARY", {
    x: 0.8, y: 2.6, w: 5.2, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.accent, fontFace: "Segoe UI"
});
slide15.addText(
    "•  Successfully resolved the complex Render Puppeteer library crash by implementing a serverless core with Sparticuz Chromium.\n" +
    "•  Enhanced security with HTTP-only cookies and a verification-based password recovery flow.\n" +
    "•  Upgraded models to gemini-3.5-flash to prevent deprecated model lockouts.\n" +
    "•  Polished UI layouts for 100% responsiveness on mobile devices.",
    { x: 0.8, y: 3.2, w: 5.2, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);

slide15.addShape(pres.ShapeType.rect, {
    x: 6.8, y: 2.3, w: 6.0, h: 4.2,
    fill: { color: COLORS.card },
    line: { color: "2A3348", width: 1 }
});
slide15.addText("FUTURE WORK MAP", {
    x: 7.1, y: 2.6, w: 5.4, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.blue, fontFace: "Segoe UI"
});
slide15.addText(
    "•  AI Voice Interviews: Integrating speech-to-text to perform mock oral interviews directly in the browser.\n" +
    "•  ATS Formatting Checker: Scanning PDF resume layouts and giving recommendations to optimize for automated parser checkers.\n" +
    "•  Real-Time Progress Tracking: Saving preparation checklists showing daily completed roadmap tasks.\n\n" +
    "Thank you! Questions & Discussion.",
    { x: 7.1, y: 3.2, w: 5.4, h: 3.0, fontSize: 12, color: COLORS.text, fontFace: "Segoe UI", lineSpacing: 20 }
);


// ─────────────────────────────────────────────────────────────────────────────
// SAVE PRESENTATION
// ─────────────────────────────────────────────────────────────────────────────
const outFileName = "AI_Interview_Plan_Presentation.pptx";
const localOutPath = path.join(__dirname, outFileName);
const workspaceOutPath = path.join("f:\\New folder\\Sheryians coding\\GEN AI project 01", outFileName);

pres.writeFile({ fileName: localOutPath })
    .then(() => {
        console.log(`Saved locally to: ${localOutPath}`);
        // Copy to workspace root so user can easily download/view it
        fs.copyFileSync(localOutPath, workspaceOutPath);
        console.log(`Copied to workspace root: ${workspaceOutPath}`);
    })
    .catch(err => {
        console.error("Failed to generate PPTX:", err);
    });
