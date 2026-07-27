import React, { useState } from 'react'
import { useNavigate, Link } from 'react-router'
import "../auth.form.scss"
import { useAuth } from '../hooks/useAuth'

const ForgotPassword = () => {
    const { loading, handleResetPassword, error } = useAuth()
    const navigate = useNavigate()

    const [ email, setEmail ] = useState("")
    const [ username, setUsername ] = useState("")
    const [ newPassword, setNewPassword ] = useState("")
    const [ showPassword, setShowPassword ] = useState(false)
    const [ successMessage, setSuccessMessage ] = useState("")

    const handleSubmit = async (e) => {
        e.preventDefault()
        try {
            await handleResetPassword({ email, username, newPassword })
            setSuccessMessage("Password reset successfully! Redirecting to login...")
            setTimeout(() => {
                navigate('/login')
            }, 2500)
        } catch (err) {
            // Error is handled in the hook
        }
    }

    if (loading) {
        return (<main><h1>Processing reset...</h1></main>)
    }

    return (
        <main>
            <div className="form-container">
                <h1>Reset Password</h1>
                <p style={{ textAlign: 'left', color: '#7d8590', marginBottom: '0.5rem', fontSize: '0.85rem' }}>
                    Enter your email, username, and a new password to reset it.
                </p>
                {error && <p className="error-message" style={{ color: '#ef4444', marginBottom: '1rem', fontSize: '0.9rem', textAlign: 'center' }}>{error}</p>}
                {successMessage && <p className="success-message" style={{ color: '#10b981', marginBottom: '1rem', fontSize: '0.9rem', textAlign: 'center' }}>{successMessage}</p>}
                
                <form onSubmit={handleSubmit}>
                    <div className="input-group">
                        <label htmlFor="email">Email Address</label>
                        <input
                            onChange={(e) => { setEmail(e.target.value) }}
                            value={email}
                            type="email" id="email" name='email' placeholder='Enter your email address' required />
                    </div>
                    <div className="input-group">
                        <label htmlFor="username">Username</label>
                        <input
                            onChange={(e) => { setUsername(e.target.value) }}
                            value={username}
                            type="text" id="username" name='username' placeholder='Enter your username' required />
                    </div>
                    <div className="input-group">
                        <label htmlFor="newPassword">New Password</label>
                        <input
                            onChange={(e) => { setNewPassword(e.target.value) }}
                            value={newPassword}
                            type={showPassword ? "text" : "password"} 
                            id="newPassword" 
                            name='newPassword' 
                            placeholder='Enter new password' 
                            required 
                        />
                        <button
                            type="button"
                            className="password-toggle"
                            onClick={() => setShowPassword(show => !show)}
                            title={showPassword ? "Hide password" : "Show password"}
                        >
                            {showPassword ? (
                                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></svg>
                            ) : (
                                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
                            )}
                        </button>
                    </div>
                    <button className='button primary-button'>Reset Password</button>
                </form>
                <p>Back to <Link to={"/login"}>Login</Link></p>
            </div>
        </main>
    )
}

export default ForgotPassword
