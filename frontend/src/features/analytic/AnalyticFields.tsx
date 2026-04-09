import { useState, useEffect } from 'react'
import {
  Box, Typography, Table, TableHead, TableBody, TableRow, TableCell,
  TextField, Select, MenuItem, IconButton,
} from '@mui/material'
import AddOutlined from '@mui/icons-material/AddOutlined'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import * as api from '../../api'
import type { AnalyticField, Analytic } from '../../types'
import { transliterate } from '../../utils/transliterate'

const DATA_TYPES = [
  { value: 'string', label: 'Строка' },
  { value: 'number', label: 'Число' },
  { value: 'percent', label: '%' },
  { value: 'money', label: 'Деньги' },
  { value: 'date', label: 'Дата' },
]

interface Props {
  analyticId: string
}

export default function AnalyticFields({ analyticId }: Props) {
  const [fields, setFields] = useState<AnalyticField[]>([])
  const [isPeriods, setIsPeriods] = useState(false)

  const load = async () => {
    const [fs, analytic] = await Promise.all([
      api.listFields(analyticId),
      api.getAnalytic(analyticId),
    ])
    setFields(fs)
    setIsPeriods(!!analytic.is_periods)
  }

  useEffect(() => { load() }, [analyticId])

  const handleAdd = async () => {
    await api.createField(analyticId, { name: 'Новое поле', data_type: 'string', sort_order: fields.length })
    load()
  }

  const handleUpdate = async (f: AnalyticField, updates: Partial<AnalyticField>) => {
    const updated = { ...f, ...updates }
    if (updates.name && !updates.code) {
      updated.code = transliterate(updates.name)
    }
    await api.updateField(analyticId, f.id, updated)
    load()
  }

  const handleDelete = async (fid: string) => {
    await api.deleteField(analyticId, fid)
    load()
  }

  return (
    <Box sx={{ mb: 3, maxWidth: 600 }}>
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 1 }}>
        <Typography variant="subtitle1">Поля</Typography>
        {!isPeriods && (
          <IconButton size="small" onClick={handleAdd}><AddOutlined fontSize="small" /></IconButton>
        )}
      </Box>
      {fields.length > 0 && (
        <Table size="small">
          <TableHead>
            <TableRow>
              <TableCell>Название</TableCell>
              <TableCell>Код</TableCell>
              <TableCell>Тип</TableCell>
              <TableCell width={40}></TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {fields.map(f => (
              <TableRow key={f.id}>
                <TableCell>
                  <TextField
                    key={f.id + f.name}
                    defaultValue={f.name} fullWidth variant="standard"
                    disabled={isPeriods}
                    onBlur={e => e.target.value !== f.name && handleUpdate(f, { name: e.target.value })}
                    InputProps={{ disableUnderline: !isPeriods }}
                  />
                </TableCell>
                <TableCell>
                  <Typography variant="body2" sx={{ fontFamily: 'monospace', color: '#666' }}>{f.code}</Typography>
                </TableCell>
                <TableCell>
                  <Select
                    value={f.data_type} variant="standard" disabled={isPeriods}
                    onChange={e => handleUpdate(f, { data_type: e.target.value as any })}
                  >
                    {DATA_TYPES.map(dt => <MenuItem key={dt.value} value={dt.value}>{dt.label}</MenuItem>)}
                  </Select>
                </TableCell>
                <TableCell>
                  {!isPeriods && (
                    <IconButton size="small" onClick={() => handleDelete(f.id)}>
                      <DeleteOutlineOutlined sx={{ fontSize: 16 }} />
                    </IconButton>
                  )}
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      )}
    </Box>
  )
}
