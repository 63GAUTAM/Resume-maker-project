const express = require("express")
const cookieParser = require("cookie-parser")
const cors = require("cors")

const app = express()

app.use(express.json())
app.use(cookieParser())
const allowedOrigins = [
    process.env.FRONTEND_URL
].filter(Boolean)

app.use(cors({
    origin: function (origin, callback) {
        if (!origin) return callback(null, true);
        // Allow localhost with any port in development
        if (/^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)) {
            return callback(null, true);
        }
        // Allow production frontend url
        if (allowedOrigins.includes(origin) || origin === process.env.FRONTEND_URL) {
            return callback(null, true);
        }
        return callback(new Error("Not allowed by CORS"));
    },
    credentials: true
}))

/* require all the routes here */
const authRouter = require("./routes/auth.routes")
const interviewRouter = require("./routes/interview.routes")


/* using all the routes here */
app.use("/api/auth", authRouter)
app.use("/api/interview", interviewRouter)


/* Global error handling middleware */
app.use((err, req, res, next) => {
    console.error(err.stack)

    let statusCode = err.statusCode || err.status || 500
    let message = err.message || "Internal server error"

    // Try parsing the error if it contains a JSON error from Google GenAI
    let isQuotaError = false
    let isKeyInvalid = false
    let isProjectDenied = false

    try {
        const parsed = JSON.parse(err.message)
        if (parsed.error) {
            message = parsed.error.message || message
            if (parsed.error.status === "RESOURCE_EXHAUSTED" || parsed.error.code === 429) {
                isQuotaError = true
            }
            if (parsed.error.status === "INVALID_ARGUMENT" || parsed.error.code === 400) {
                if (message.includes("API key not valid") || message.includes("API_KEY_INVALID")) {
                    isKeyInvalid = true
                }
            }
            if (parsed.error.status === "PERMISSION_DENIED" || parsed.error.code === 403) {
                isProjectDenied = true
            }
        }
    } catch (_) { }

    // Fallback checks on message string
    if (message.includes("API key not valid") || message.includes("API_KEY_INVALID")) {
        isKeyInvalid = true
    }
    if (
        message.includes("denied access") ||
        message.includes("PERMISSION_DENIED") ||
        message.includes("403")
    ) {
        isProjectDenied = true
    }
    if (
        !isProjectDenied &&
        !isKeyInvalid &&
        (message.includes("quota") ||
        message.includes("RESOURCE_EXHAUSTED") ||
        message.includes("exhausted") ||
        statusCode === 429)
    ) {
        isQuotaError = true
    }

    if (isKeyInvalid) {
        statusCode = 401
        message = "Invalid Google Gemini API key. Please check GOOGLE_GENAI_API_KEYS in Backend/.env"
    } else if (isProjectDenied) {
        statusCode = 403
        message = "Your Google Gemini API key project has been denied access. Please check your project and billing status in Google AI Studio or Google Cloud Console."
    } else if (isQuotaError) {
        statusCode = 429
        message = "Your Google Gemini API quota has been exhausted. Please check your plan/billing details or replace the GOOGLE_GENAI_API_KEYS in Backend/.env with a new key."
    }

    res.status(statusCode).json({
        message: message
    })
})

module.exports = app