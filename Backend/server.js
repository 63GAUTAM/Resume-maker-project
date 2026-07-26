require("dotenv").config()
const app = require("./src/app")
const connectToDB = require("./src/config/database")

// Validate required environment variables before starting
const requiredEnvVars = ["MONGO_URI", "JWT_SECRET"]
for (const envVar of requiredEnvVars) {
    if (!process.env[envVar]) {
        console.error(`FATAL: Missing required environment variable: ${envVar}`)
        process.exit(1)
    }
}

// Ensure at least one Google GenAI API key variable is configured
const hasGeminiKey = process.env.GOOGLE_GENAI_API_KEYS || process.env.GOOGLE_GENAI_API_KEY || process.env.GEMINI_API_KEY
if (!hasGeminiKey) {
    console.error("FATAL: Missing required environment variable: GOOGLE_GENAI_API_KEYS, GOOGLE_GENAI_API_KEY, or GEMINI_API_KEY")
    process.exit(1)
}

connectToDB()

const PORT = process.env.PORT || 3000

app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`)
})