import { useRef, useLayoutEffect } from 'react'
import { Box, Typography, IconButton } from '@mui/material'
import CloseOutlined from '@mui/icons-material/CloseOutlined'
import * as am5 from '@amcharts/amcharts5'
import * as am5xy from '@amcharts/amcharts5/xy'
import * as am5percent from '@amcharts/amcharts5/percent'
import am5themes_Animated from '@amcharts/amcharts5/themes/Animated'

export interface ChartConfig {
  title: string
  chart_type: 'line' | 'bar' | 'pie' | 'area'
  data: Record<string, any>[]
  series: { field: string; name: string }[]
  category_field?: string
}

interface ChartPanelProps {
  config: ChartConfig
  onClose: () => void
}

export default function ChartPanel({ config, onClose }: ChartPanelProps) {
  const chartRef = useRef<HTMLDivElement>(null)
  const rootRef = useRef<am5.Root | null>(null)

  useLayoutEffect(() => {
    if (!chartRef.current) return
    // Dispose previous
    if (rootRef.current) {
      rootRef.current.dispose()
    }

    const root = am5.Root.new(chartRef.current)
    rootRef.current = root
    root.setThemes([am5themes_Animated.new(root)])

    const categoryField = config.category_field || 'category'

    if (config.chart_type === 'pie') {
      // Pie chart
      const chart = root.container.children.push(
        am5percent.PieChart.new(root, { layout: root.verticalLayout })
      )
      const series = chart.series.push(
        am5percent.PieSeries.new(root, {
          valueField: config.series[0]?.field || 'value',
          categoryField,
        })
      )
      series.data.setAll(config.data)
      series.appear(1000, 100)
      chart.appear(1000, 100)
    } else {
      // XY chart: line, bar, area
      const chart = root.container.children.push(
        am5xy.XYChart.new(root, {
          panX: true,
          panY: false,
          wheelX: 'panX',
          wheelY: 'zoomX',
          layout: root.verticalLayout,
        })
      )

      const xRenderer = am5xy.AxisRendererX.new(root, { minGridDistance: 80 })
      xRenderer.labels.template.setAll({ rotation: -45, centerY: am5.percent(50), centerX: am5.percent(100), paddingRight: 8, fontSize: 11, oversizedBehavior: 'truncate', maxWidth: 120 })
      const xAxis = chart.xAxes.push(
        am5xy.CategoryAxis.new(root, {
          categoryField,
          renderer: xRenderer,
          tooltip: am5.Tooltip.new(root, {}),
        })
      )
      xAxis.data.setAll(config.data)

      const yAxis = chart.yAxes.push(
        am5xy.ValueAxis.new(root, {
          renderer: am5xy.AxisRendererY.new(root, {}),
        })
      )

      for (const s of config.series) {
        let seriesInstance: am5xy.XYSeries
        if (config.chart_type === 'bar') {
          seriesInstance = chart.series.push(
            am5xy.ColumnSeries.new(root, {
              name: s.name,
              xAxis,
              yAxis,
              valueYField: s.field,
              categoryXField: categoryField,
              tooltip: am5.Tooltip.new(root, { labelText: '{name}: {valueY}' }),
            })
          )
        } else {
          seriesInstance = chart.series.push(
            am5xy.LineSeries.new(root, {
              name: s.name,
              xAxis,
              yAxis,
              valueYField: s.field,
              categoryXField: categoryField,
              tooltip: am5.Tooltip.new(root, { labelText: '{name}: {valueY}' }),
            })
          )
          if (config.chart_type === 'area') {
            (seriesInstance as am5xy.LineSeries).fills.template.setAll({
              fillOpacity: 0.3,
              visible: true,
            })
          }
        }
        seriesInstance.data.setAll(config.data)
        seriesInstance.appear(1000)
      }

      // Legend
      if (config.series.length > 1) {
        const legend = chart.children.push(am5.Legend.new(root, { centerX: am5.percent(50), x: am5.percent(50) }))
        legend.data.setAll(chart.series.values)
      }

      // Cursor
      chart.set('cursor', am5xy.XYCursor.new(root, {}))
      chart.appear(1000, 100)
    }

    return () => {
      root.dispose()
      rootRef.current = null
    }
  }, [config])

  return (
    <Box sx={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0, height: '100%' }}>
      <Box sx={{
        display: 'flex', alignItems: 'center', px: 2, py: 1,
        borderBottom: '1px solid #e0e0e0', background: '#fff',
      }}>
        <Typography sx={{ flex: 1, fontSize: 14, fontWeight: 600 }}>
          {config.title || 'График'}
        </Typography>
        <IconButton size="small" onClick={onClose}><CloseOutlined fontSize="small" /></IconButton>
      </Box>
      <div ref={chartRef} style={{ flex: 1, minHeight: 300, width: '100%' }} />
    </Box>
  )
}
