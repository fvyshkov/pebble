import React, { useCallback, useRef } from 'react'

interface Props {
  onResize: (delta: number) => void
}

export default function Splitter({ onResize }: Props) {
  const startX = useRef(0)

  const onMouseDown = useCallback((e: React.MouseEvent) => {
    e.preventDefault()
    startX.current = e.clientX
    const onMove = (ev: MouseEvent) => {
      const delta = ev.clientX - startX.current
      startX.current = ev.clientX
      onResize(delta)
    }
    const onUp = () => {
      document.removeEventListener('mousemove', onMove)
      document.removeEventListener('mouseup', onUp)
    }
    document.addEventListener('mousemove', onMove)
    document.addEventListener('mouseup', onUp)
  }, [onResize])

  return <div className="splitter" onMouseDown={onMouseDown} />
}
