import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Box, TextField, Button, Typography, Alert, Chip } from '@mui/material'
import { LANGUAGES, changeLanguage, currentLang } from '../../i18n'

interface Props {
  onLogin: (token: string, user: { id: string; username: string; can_admin: boolean }) => void
}

export default function LoginPage({ onLogin }: Props) {
  const { t } = useTranslation()
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)
  const [lang, setLang] = useState(currentLang())

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setError('')
    setLoading(true)
    try {
      const resp = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password }),
      })
      const data = await resp.json()
      if (data.error) {
        setError(data.error)
      } else if (data.token) {
        localStorage.setItem('pebble_token', data.token)
        localStorage.setItem('pebble_user', JSON.stringify(data.user))
        onLogin(data.token, data.user)
      }
    } catch (err) {
      setError(t('auth.connectionError'))
    } finally {
      setLoading(false)
    }
  }

  return (
    <Box sx={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100vh', bgcolor: '#f5f5f5' }}>
      <Box component="form" onSubmit={handleSubmit} sx={{ width: 340, p: 4, bgcolor: '#fff', borderRadius: 2, boxShadow: 3 }}>
        <Typography variant="h5" sx={{ mb: 3, textAlign: 'center', fontWeight: 600 }}>Pebble</Typography>
        <Box sx={{ display: 'flex', justifyContent: 'center', gap: 0.5, mb: 2 }}>
          {LANGUAGES.map(l => (
            <Chip
              key={l.code}
              label={l.label}
              size="small"
              variant={lang === l.code ? 'filled' : 'outlined'}
              color={lang === l.code ? 'primary' : 'default'}
              onClick={() => { changeLanguage(l.code); setLang(l.code) }}
              sx={{ fontSize: 11, cursor: 'pointer' }}
            />
          ))}
        </Box>
        {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}
        <TextField
          label={t('auth.login')} fullWidth value={username}
          onChange={e => setUsername(e.target.value)}
          autoComplete="username" name="username"
          sx={{ mb: 2 }} autoFocus
        />
        <TextField
          label={t('auth.password')} fullWidth type="password" value={password}
          onChange={e => setPassword(e.target.value)}
          autoComplete="current-password" name="password"
          sx={{ mb: 3 }}
        />
        <Button type="submit" variant="contained" fullWidth disabled={loading || !username || !password}>
          {loading ? t('auth.loggingIn') : t('auth.loginBtn')}
        </Button>
      </Box>
    </Box>
  )
}
