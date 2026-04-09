import InboxOutlined from '@mui/icons-material/InboxOutlined'
import { Typography } from '@mui/material'

export default function EmptyState({ text = 'Выберите элемент в дереве' }: { text?: string }) {
  return (
    <div className="empty-state">
      <InboxOutlined sx={{ fontSize: 48, opacity: 0.3 }} />
      <Typography variant="body2" color="textSecondary">{text}</Typography>
    </div>
  )
}
