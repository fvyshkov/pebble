import { useState, useMemo } from 'react'
import { Dialog, DialogTitle, DialogContent, TextField, Box, IconButton, Tooltip } from '@mui/material'
import * as Icons from '@mui/icons-material'

const ICON_NAMES: string[] = [
  'AccountBalance', 'AccountCircle', 'AccountTree', 'Add', 'Agriculture',
  'Analytics', 'Apartment', 'Assessment', 'Assignment', 'AttachMoney',
  'BarChart', 'Bookmark', 'Build', 'Business', 'BusinessCenter',
  'Calculate', 'Calendar', 'CalendarMonth', 'CalendarToday', 'Campaign',
  'Category', 'CheckCircle', 'Cloud', 'Code', 'CreditCard',
  'Dashboard', 'DataUsage', 'DateRange', 'Description', 'Dns',
  'Domain', 'DonutLarge', 'DoubleArrow', 'Download', 'Drafts',
  'Eco', 'Edit', 'Email', 'Engineering', 'Euro',
  'Event', 'Extension', 'Face', 'Factory', 'Favorite',
  'FileCopy', 'FilterList', 'Flag', 'FlightTakeoff', 'Folder',
  'FolderOpen', 'Functions', 'Gavel', 'GridView', 'Group',
  'Groups', 'Handyman', 'HealthAndSafety', 'Home', 'Hub',
  'Inventory', 'Key', 'Label', 'Landscape', 'Language',
  'Layers', 'Leaderboard', 'LightMode', 'ListAlt', 'LocalAtm',
  'LocalOffer', 'LocalShipping', 'LocationOn', 'Lock', 'Loyalty',
  'ManageAccounts', 'Map', 'Memory', 'MergeType', 'MonetizationOn',
  'Monitoring', 'NorthEast', 'Notifications', 'Paid', 'Palette',
  'PeopleAlt', 'Percent', 'Person', 'PieChart', 'Pin',
  'PlayArrow', 'Policy', 'PriceChange', 'Public', 'QueryStats',
  'Receipt', 'Redeem', 'Repeat', 'Report', 'Rocket',
  'Rule', 'Savings', 'Schedule', 'School', 'Science',
  'Search', 'Security', 'Sell', 'Settings', 'Shield',
  'ShoppingCart', 'ShowChart', 'Speed', 'StackedBarChart', 'Star',
  'Storage', 'Store', 'Summarize', 'Support', 'Sync',
  'TableChart', 'Tag', 'Task', 'TextSnippet', 'Timeline',
  'Token', 'TrendingUp', 'Tune', 'Upload', 'Verified',
  'ViewList', 'Visibility', 'Wallet', 'Warehouse', 'Work',
]

const OUTLINED_NAMES = ICON_NAMES.map(n => n + 'Outlined')

interface Props {
  open: boolean
  onClose: () => void
  onSelect: (iconName: string) => void
}

export default function IconPickerDialog({ open, onClose, onSelect }: Props) {
  const [search, setSearch] = useState('')

  const filtered = useMemo(() => {
    const q = search.toLowerCase()
    return OUTLINED_NAMES.filter(n => n.toLowerCase().includes(q))
  }, [search])

  return (
    <Dialog open={open} onClose={onClose} maxWidth="sm" fullWidth>
      <DialogTitle>Выбор иконки</DialogTitle>
      <DialogContent>
        <TextField
          autoFocus fullWidth placeholder="Поиск..." value={search}
          onChange={e => setSearch(e.target.value)} sx={{ mb: 2 }}
        />
        <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 0.5, maxHeight: 400, overflow: 'auto' }}>
          {filtered.map(name => {
            const Icon = (Icons as any)[name]
            if (!Icon) return null
            return (
              <Tooltip key={name} title={name.replace('Outlined', '')}>
                <IconButton onClick={() => { onSelect(name); onClose() }} sx={{ p: 1 }}>
                  <Icon fontSize="small" />
                </IconButton>
              </Tooltip>
            )
          })}
        </Box>
      </DialogContent>
    </Dialog>
  )
}
