import { useState, useEffect } from 'react'
import { TextField, Typography, Box } from '@mui/material'
import * as api from '../../api'
import type { Model } from '../../types'
import { usePending } from '../../store/PendingContext'

interface Props {
  modelId: string
  onRefresh: () => void
}

export default function ModelSettings({ modelId, onRefresh }: Props) {
  const [model, setModel] = useState<Model | null>(null)
  const { addOp, getOverrides } = usePending()

  useEffect(() => {
    api.getModelTree(modelId).then(m => {
      const overrides = getOverrides(`model:${modelId}`)
      setModel({
        id: m.id,
        name: overrides?.name ?? m.name,
        description: overrides?.description ?? m.description,
        created_at: m.created_at,
        updated_at: m.updated_at,
      })
    })
  }, [modelId])

  if (!model) return null

  const change = (field: string, value: string) => {
    const updated = { ...model, [field]: value }
    setModel(updated)
    addOp({
      key: `model:${modelId}`,
      type: 'updateModel',
      id: modelId,
      data: { name: updated.name, description: updated.description },
    })
  }

  return (
    <Box sx={{ maxWidth: 500 }}>
      <Typography variant="h6" sx={{ mb: 2 }}>Модель</Typography>
      <TextField
        label="Название" fullWidth value={model.name}
        onChange={e => change('name', e.target.value)} sx={{ mb: 2 }}
      />
      <TextField
        label="Описание" fullWidth multiline rows={3} value={model.description}
        onChange={e => change('description', e.target.value)}
      />
    </Box>
  )
}
