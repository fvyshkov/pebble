import { useRef } from 'react'
import { Box, IconButton, Button, Tooltip } from '@mui/material'
import CloseOutlined from '@mui/icons-material/CloseOutlined'
import DownloadOutlined from '@mui/icons-material/DownloadOutlined'

interface Props {
  html: string
  title: string
  onClose: () => void
}

export default function PresentationPanel({ html, title, onClose }: Props) {
  const iframeRef = useRef<HTMLIFrameElement>(null)

  const handleDownload = () => {
    const iframe = iframeRef.current
    if (!iframe?.contentWindow) return
    iframe.contentWindow.print()
  }

  return (
    <Box sx={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <Box sx={{
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        px: 2, py: 1, borderBottom: '1px solid #e0e0e0', background: '#fafafa',
      }}>
        <Box sx={{ fontWeight: 500, fontSize: 14 }}>{title}</Box>
        <Box sx={{ display: 'flex', gap: 0.5 }}>
          <Tooltip title="Скачать / Печать (PDF)">
            <Button
              size="small" variant="outlined"
              startIcon={<DownloadOutlined />}
              onClick={handleDownload}
              sx={{ textTransform: 'none', fontSize: 12 }}
            >
              PDF
            </Button>
          </Tooltip>
          <IconButton size="small" onClick={onClose}>
            <CloseOutlined fontSize="small" />
          </IconButton>
        </Box>
      </Box>
      <Box sx={{ flex: 1, overflow: 'hidden' }}>
        <iframe
          ref={iframeRef}
          srcDoc={html}
          style={{ width: '100%', height: '100%', border: 'none' }}
          title={title}
        />
      </Box>
    </Box>
  )
}
