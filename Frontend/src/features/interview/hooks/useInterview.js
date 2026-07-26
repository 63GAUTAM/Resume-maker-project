import { getAllInterviewReports, generateInterviewReport, getInterviewReportById, generateResumePdf } from "../services/interview.api"
import { useContext, useEffect, useCallback } from "react"
import { InterviewContext } from "../interview.context"
import { useParams } from "react-router"


export const useInterview = () => {

    const context = useContext(InterviewContext)
    const { interviewId } = useParams()

    if (!context) {
        throw new Error("useInterview must be used within an InterviewProvider")
    }

    const { loading, setLoading, report, setReport, reports, setReports } = context

    const generateReport = useCallback(async ({ jobDescription, selfDescription, resumeFile, model }) => {
        setLoading(true)
        try {
            const response = await generateInterviewReport({ jobDescription, selfDescription, resumeFile, model })
            setReport(response.interviewReport)
            return { success: true, interviewReport: response.interviewReport }
        } catch (error) {
            console.error("Failed to generate report:", error)
            const message = error.response?.data?.message || error.message || "Failed to generate interview report"
            return { success: false, error: message }
        } finally {
            setLoading(false)
        }
    }, [setLoading, setReport])

    const getReportById = useCallback(async (id) => {
        setLoading(true)
        try {
            const response = await getInterviewReportById(id)
            setReport(response.interviewReport)
            return response.interviewReport
        } catch (error) {
            console.error("Failed to get report:", error)
            return null
        } finally {
            setLoading(false)
        }
    }, [setLoading, setReport])

    const getReports = useCallback(async () => {
        setLoading(true)
        try {
            const response = await getAllInterviewReports()
            setReports(response.interviewReports)
            return response.interviewReports
        } catch (error) {
            console.error("Failed to get reports:", error)
            return null
        } finally {
            setLoading(false)
        }
    }, [setLoading, setReports])

    const getResumePdf = useCallback(async (interviewReportId) => {
        setLoading(true)
        try {
            const response = await generateResumePdf({ interviewReportId })
            const url = window.URL.createObjectURL(new Blob([ response ], { type: "application/pdf" }))
            const link = document.createElement("a")
            link.href = url
            link.setAttribute("download", `resume_${interviewReportId}.pdf`)
            document.body.appendChild(link)
            link.click()
            link.remove()
        }
        catch (error) {
            console.error("Failed to download resume PDF:", error)
        } finally {
            setLoading(false)
        }
    }, [setLoading])

    useEffect(() => {
        if (interviewId) {
            getReportById(interviewId)
        } else {
            getReports()
        }
    }, [ interviewId, getReportById, getReports ])

    return { loading, report, reports, generateReport, getReportById, getReports, getResumePdf }

}