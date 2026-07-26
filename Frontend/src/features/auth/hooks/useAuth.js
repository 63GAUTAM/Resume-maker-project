import { useContext, useEffect, useState, useCallback } from "react";
import { AuthContext } from "../auth.context";
import { login, register, logout, getMe } from "../services/auth.api";



export const useAuth = () => {

    const context = useContext(AuthContext)
    const { user, setUser, loading, setLoading } = context
    const [ error, setError ] = useState(null)


    const handleLogin = useCallback(async ({ email, password }) => {
        setLoading(true)
        setError(null)
        try {
            const data = await login({ email, password })
            setUser(data.user)
            return data
        } catch (err) {
            const message = err.response?.data?.message || "Login failed. Please try again."
            setError(message)
            throw err
        } finally {
            setLoading(false)
        }
    }, [setLoading, setUser])

    const handleRegister = useCallback(async ({ username, email, password }) => {
        setLoading(true)
        setError(null)
        try {
            const data = await register({ username, email, password })
            setUser(data.user)
            return data
        } catch (err) {
            const message = err.response?.data?.message || "Registration failed. Please try again."
            setError(message)
            throw err
        } finally {
            setLoading(false)
        }
    }, [setLoading, setUser])

    const handleLogout = useCallback(async () => {
        setLoading(true)
        setError(null)
        try {
            await logout()
            setUser(null)
        } catch (err) {
            const message = err.response?.data?.message || "Logout failed."
            setError(message)
        } finally {
            setLoading(false)
        }
    }, [setLoading, setUser])

    useEffect(() => {

        const getAndSetUser = async () => {
            try {
                const data = await getMe()
                setUser(data.user)
            } catch (err) { } finally {
                setLoading(false)
            }
        }

        getAndSetUser()

    }, [])

    return { user, loading, error, handleRegister, handleLogin, handleLogout }
}