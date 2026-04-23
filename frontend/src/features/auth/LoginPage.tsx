import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Box, TextField, Button, Typography, Alert, Select, MenuItem } from '@mui/material'
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
        <Box sx={{ display: 'flex', justifyContent: 'center', mb: 2 }}>
          <Select
            size="small"
            value={lang}
            onChange={e => { const v = e.target.value; changeLanguage(v); setLang(v) }}
            sx={{ fontSize: 12, height: 32, minWidth: 100 }}
          >
            {LANGUAGES.map(l => (
              <MenuItem key={l.code} value={l.code} sx={{ fontSize: 12 }}>{l.label} — {l.name}</MenuItem>
            ))}
          </Select>
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
