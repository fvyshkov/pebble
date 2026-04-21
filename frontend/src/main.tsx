declare const __APP_VERSION__: string
declare const __BUILD_TIME__: string
console.log(`%cPebble v${__APP_VERSION__} (${__BUILD_TIME__})`, 'color: #1976d2; font-weight: bold')

import React from 'react'
import ReactDOM from 'react-dom/client'
import { ThemeProvider, createTheme, CssBaseline } from '@mui/material'
import App from './App'

const theme = createTheme({
  palette: {
    primary: { main: '#1976d2' },
    background: { default: '#fafafa' },
  },
  typography: {
    fontSize: 13,
  },
  components: {
    MuiButton: { defaultProps: { size: 'small' } },
    MuiTextField: { defaultProps: { size: 'small', variant: 'outlined' } },
    MuiIconButton: { defaultProps: { size: 'small' } },
  },
})

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <App />
    </ThemeProvider>
  </React.StrictMode>,
)
