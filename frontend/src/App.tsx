import { useState, useCallback, useEffect } from 'react'
import { IconButton, Tooltip, Badge, Select, MenuItem, FormControl } from '@mui/material'
import SaveOutlined from '@mui/icons-material/SaveOutlined'
import GridOnOutlined from '@mui/icons-material/GridOnOutlined'
import PeopleOutlined from '@mui/icons-material/PeopleOutlined'
import type { TreeSelection } from './types'
import LeftPanel from './panels/LeftPanel'
import CenterPanel from './panels/CenterPanel'
import Splitter from './components/Splitter'
import UsersDialog from './components/UsersDialog'
import PivotGrid from './features/sheet/PivotGrid'
import { PendingProvider, usePending } from './store/PendingContext'
import * as api from './api'
import './App.css'

function SaveButton() {
  const { isDirty, flush } = usePending()
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); if (isDirty) flush() }
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [isDirty, flush])
  return (
    <Tooltip title={isDirty ? 'Сохранить (Ctrl+S)' : 'Нет изменений'}>
      <span>
        <IconButton size="small" disabled={!isDirty} onClick={flush} sx={{ color: isDirty ? '#1976d2' : undefined }}>
          <Badge variant="dot" color="error" invisible={!isDirty}><SaveOutlined fontSize="small" /></Badge>
        </IconButton>
      </span>
    </Tooltip>
  )
}

function AppInner() {
  const [selection, setSelection] = useState<TreeSelection | null>(null)
  const [leftWidth, setLeftWidth] = useState(280)
  const [refreshKey, setRefreshKey] = useState(0)
  const [showPivot, setShowPivot] = useState(false)
  const [showUsers, setShowUsers] = useState(false)
  const [expandAfterCreate, setExpandAfterCreate] = useState<any>(null)
  const [users, setUsers] = useState<any[]>([])
  const [currentUserId, setCurrentUserId] = useState('')

  useEffect(() => {
    api.listUsers().then(u => {
      setUsers(u)
      if (u.length > 0 && !currentUserId) setCurrentUserId(u[0].id)
    })
  }, [showUsers])

  const onRefresh = useCallback(() => setRefreshKey(k => k + 1), [])

  const onCreated = useCallback((info: { modelId: string; folder: 'sheets' | 'analytics'; id: string; type: 'sheet' | 'analytic' }) => {
    setExpandAfterCreate({ modelId: info.modelId, folder: info.folder, selectId: info.id, selectType: info.type })
    setSelection({ type: info.type, id: info.id, modelId: info.modelId })
    setRefreshKey(k => k + 1)
  }, [])

  const isSheetSelected = selection?.type === 'sheet'

  return (
    <PendingProvider onFlushed={onRefresh}>
      <div className="app-root">
        <div className="app-toolbar">
          <SaveButton />
          <Tooltip title="Просмотр / ввод данных">
            <span>
              <IconButton size="small" disabled={!isSheetSelected} onClick={() => setShowPivot(true)}>
                <GridOnOutlined fontSize="small" />
              </IconButton>
            </span>
          </Tooltip>
          <div style={{ flex: 1 }} />

          {/* User selector */}
          {users.length > 0 && (
            <FormControl size="small" sx={{ minWidth: 120 }}>
              <Select
                value={currentUserId}
                onChange={e => setCurrentUserId(e.target.value)}
                variant="standard"
                disableUnderline
                sx={{ fontSize: 12 }}
              >
                {users.map(u => <MenuItem key={u.id} value={u.id} sx={{ fontSize: 12 }}>{u.username}</MenuItem>)}
              </Select>
            </FormControl>
          )}

          <Tooltip title="Пользователи">
            <IconButton size="small" onClick={() => setShowUsers(true)}>
              <PeopleOutlined fontSize="small" />
            </IconButton>
          </Tooltip>
        </div>
        <div className="app-body">
          <div style={{ width: leftWidth, minWidth: 180, flexShrink: 0 }}>
            <LeftPanel
              selection={selection} onSelect={setSelection}
              refreshKey={refreshKey} expandAfterCreate={expandAfterCreate} onCreated={onCreated}
            />
          </div>
          <Splitter onResize={d => setLeftWidth(w => Math.max(180, w + d))} />
          <CenterPanel selection={selection} onRefresh={onRefresh} />
        </div>
        {showPivot && isSheetSelected && (
          <PivotGrid sheetId={selection.id} modelId={selection.modelId} currentUserId={currentUserId} onClose={() => setShowPivot(false)} />
        )}
        <UsersDialog open={showUsers} onClose={() => setShowUsers(false)} />
      </div>
    </PendingProvider>
  )
}

export default function App() { return <AppInner /> }
