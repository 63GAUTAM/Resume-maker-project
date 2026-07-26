const { GoogleGenAI } = require("@google/genai")
const { z } = require("zod")
const { zodToJsonSchema } = require("zod-to-json-schema")
const puppeteer = require("puppeteer")

// Parse API keys list from comma-separated env values, with fallbacks
function getApiKeys() {
    // Check GOOGLE_GENAI_API_KEYS first, then GOOGLE_GENAI_API_KEY, then GEMINI_API_KEY
    const keysRaw = process.env.GOOGLE_GENAI_API_KEYS || process.env.GOOGLE_GENAI_API_KEY || process.env.GEMINI_API_KEY || ""

    // Split by comma, trim spaces, filter empty values
    return keysRaw.split(",")
        .map(key => key.trim())
        .filter(key => key.length > 0)
}

// Keep track of the last request timestamp globally to enforce sequential delays
let lastRequestTime = 0

// Keep track of the last known working key index globally across requests
let currentKeyIndex = 0

/**
 * Auto-rotating fallback wrapper around client.models.generateContent
 */
async function generateContentWithRotationAndBackoff({ model, contents, config }, maxRetriesPerKey = 2) {
    const apiKeys = getApiKeys()
    if (apiKeys.length === 0) {
        throw new Error("No Google GenAI API keys configured. Please check your Backend/.env configuration.")
    }

    // Explicitly enforce gemini-3.5-flash unless explicitly specified
    const allowedModels = ["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.5-pro", "gemini-3.6-flash", "gemini-3.1-pro-preview"]
    let selectedModel = model
    if (!selectedModel || !allowedModels.includes(selectedModel)) {
        selectedModel = "gemini-3.5-flash"
    }

    let attempt = 0
    const totalMaxAttempts = apiKeys.length * maxRetriesPerKey
    const errorsList = []

    while (attempt < totalMaxAttempts) {
        // Calculate the target index dynamically using global index and current attempt offset
        const targetIndex = (currentKeyIndex + attempt) % apiKeys.length
        const activeKey = apiKeys[targetIndex]

        // Enforce 1.2s sequential request spacing between consecutive requests
        const now = Date.now()
        const timeSinceLast = now - lastRequestTime
        if (timeSinceLast < 1200) {
            const delay = 1200 - timeSinceLast
            console.log(`[GenAI] Enforcing sequential delay: waiting ${delay}ms before next request...`)
            await new Promise(resolve => setTimeout(resolve, delay))
        }

        // Log diagnostic line before each API call
        const timestamp = new Date().toISOString()
        const maskedKey = activeKey.length > 5 ? `...${activeKey.slice(-5)}` : activeKey
        console.log(`[${timestamp}] Calling Gemini API. Model: ${selectedModel}, Key: ${maskedKey}`)

        // Update timestamp just before making the request
        lastRequestTime = Date.now()

        const client = new GoogleGenAI({ apiKey: activeKey })

        try {
            const response = await client.models.generateContent({
                model: selectedModel,
                contents,
                config
            })

            // Update the global key index to this working key
            currentKeyIndex = targetIndex
            return response
        } catch (err) {
            attempt++

            const errorStr = err.message || ""
            errorsList.push(`[Key ${targetIndex}] ${errorStr}`)

            const isRateLimit = err.statusCode === 429 || err.status === 429 ||
                errorStr.includes("429") || errorStr.includes("RESOURCE_EXHAUSTED") ||
                errorStr.includes("quota") || errorStr.includes("Quota")

            const isKeyAuthError = err.statusCode === 403 || err.status === 403 || err.status === "PERMISSION_DENIED" ||
                errorStr.includes("API key not valid") || errorStr.includes("API_KEY_INVALID") ||
                errorStr.includes("denied access") || errorStr.includes("PERMISSION_DENIED") || errorStr.includes("403")

            if (isRateLimit) {
                console.warn(`[GenAI] Rate limited or quota exhausted on Key Index: ${targetIndex}. Rotating key and waiting 2 seconds before retry...`)
                // Wait exactly 2 seconds before the next key attempt
                await new Promise(resolve => setTimeout(resolve, 2000))
            } else if (isKeyAuthError) {
                console.warn(`[GenAI] Invalid or restricted API Key detected at Index: ${targetIndex}. Rotating key...`)
            } else {
                // For structural errors (e.g., bad arguments, schema validation errors), rethrow immediately
                console.error(`[GenAI] Structural error on Key Index: ${targetIndex}:`, errorStr)
                throw err
            }
        }
    }

    throw new Error(`[GenAI] Request failed. Errors: ${errorsList.join(" | ")}`)
}


const interviewReportSchema = z.object({
    matchScore: z.number().describe("A score between 0 and 100 indicating how well the candidate's profile matches the job describe"),
    technicalQuestions: z.array(z.object({
        question: z.string().describe("The technical question can be asked in the interview"),
        intention: z.string().describe("The intention of interviewer behind asking this question"),
        answer: z.string().describe("How to answer this question, what points to cover, what approach to take etc.")
    })).describe("Technical questions that can be asked in the interview along with their intention and how to answer them"),
    behavioralQuestions: z.array(z.object({
        question: z.string().describe("The technical question can be asked in the interview"),
        intention: z.string().describe("The intention of interviewer behind asking this question"),
        answer: z.string().describe("How to answer this question, what points to cover, what approach to take etc.")
    })).describe("Behavioral questions that can be asked in the interview along with their intention and how to answer them"),
    skillGaps: z.array(z.object({
        skill: z.string().describe("The skill which the candidate is lacking"),
        severity: z.enum(["low", "medium", "high"]).describe("The severity of this skill gap, i.e. how important is this skill for the job and how much it can impact the candidate's chances")
    })).describe("List of skill gaps in the candidate's profile along with their severity"),
    preparationPlan: z.array(z.object({
        day: z.number().describe("The day number in the preparation plan, starting from 1"),
        focus: z.string().describe("The main focus of this day in the preparation plan, e.g. data structures, system design, mock interviews etc."),
        tasks: z.array(z.string()).describe("List of tasks to be done on this day to follow the preparation plan, e.g. read a specific book or article, solve a set of problems, watch a video etc.")
    })).describe("A day-wise preparation plan for the candidate to follow in order to prepare for the interview effectively"),
    title: z.string().describe("The title of the job for which the interview report is generated"),
})

async function generateInterviewReport({ resume, selfDescription, jobDescription, model }) {


    const prompt = `Generate an interview report for a candidate with the following details:
                        Resume: ${resume}
                        Self Description: ${selfDescription}
                        Job Description: ${jobDescription}
`

    const selectedModel = model || "gemini-3.5-flash";

    const response = await generateContentWithRotationAndBackoff({
        model: selectedModel,
        contents: prompt,
        config: {
            responseMimeType: "application/json",
            responseSchema: zodToJsonSchema(interviewReportSchema),
        }
    })
    console.log(`[Gemini API] Requesting model: ${selectedModel}`);

    return JSON.parse(response.text)


}



async function generatePdfFromHtml(htmlContent) {
    const browser = await puppeteer.launch()
    const page = await browser.newPage();
    await page.setContent(htmlContent, { waitUntil: "networkidle0" })

    const pdfBuffer = await page.pdf({
        format: "A4", margin: {
            top: "20mm",
            bottom: "20mm",
            left: "15mm",
            right: "15mm"
        }
    })

    await browser.close()

    return pdfBuffer
}

async function generateResumePdf({ resume, selfDescription, jobDescription }) {

    const resumePdfSchema = z.object({
        html: z.string().describe("The HTML content of the resume which can be converted to PDF using any library like puppeteer")
    })

    const prompt = `Generate resume for a candidate with the following details:
                        Resume: ${resume}
                        Self Description: ${selfDescription}
                        Job Description: ${jobDescription}

                        the response should be a JSON object with a single field "html" which contains the HTML content of the resume which can be converted to PDF using any library like puppeteer.
                        The resume should be tailored for the given job description and should highlight the candidate's strengths and relevant experience. The HTML content should be well-formatted and structured, making it easy to read and visually appealing.
                        The content of resume should be not sound like it's generated by AI and should be as close as possible to a real human-written resume.
                        you can highlight the content using some colors or different font styles but the overall design should be simple and professional.
                        The content should be ATS friendly, i.e. it should be easily parsable by ATS systems without losing important information.
                        The resume should not be so lengthy, it should ideally be 1-2 pages long when converted to PDF. Focus on quality rather than quantity and make sure to include all the relevant information that can increase the candidate's chances of getting an interview call for the given job description.
                    `

    const response = await generateContentWithRotationAndBackoff({
        model: "gemini-3.5-flash",
        contents: prompt,
        config: {
            responseMimeType: "application/json",
            responseSchema: zodToJsonSchema(resumePdfSchema),
        }
    })


    const jsonContent = JSON.parse(response.text)

    const pdfBuffer = await generatePdfFromHtml(jsonContent.html)

    return pdfBuffer

}

module.exports = { generateInterviewReport, generateResumePdf }