import * as React from "react";
import {
  Chart as ChartJS,
  ArcElement,
  Tooltip,
  Legend,
  Title,
  CategoryScale,
  LinearScale,
  BarElement,
  PointElement,
  LineElement,
  BarController,
  LineController,
  PieController,
  DoughnutController
} from "chart.js";
import { Bar, Line, Pie, Doughnut } from 'react-chartjs-2';

// Register required Chart.js components
ChartJS.register(
  ArcElement, 
  Tooltip, 
  Legend, 
  Title, 
  CategoryScale, 
  LinearScale, 
  BarElement, 
  PointElement, 
  LineElement,
  BarController,
  LineController,
  PieController,
  DoughnutController
);

type AdaptiveCardChartType = "Chart.Donut" | "Chart.Doughnut" | "Chart.Pie" | "Chart.VerticalBar" | "Chart.Bar" | "Chart.HorizontalBar" | "Chart.Line" | "Chart.Gauge";
type ChartType = 'line' | 'bar' | 'pie' | 'doughnut';

interface ChartDataset {
  label?: string;
  data: number[];
  backgroundColor?: string | string[];
  borderColor?: string | string[];
  borderWidth?: number;
  fill?: boolean;
  tension?: number;
  hoverOffset?: number;
  [key: string]: any; // Allow additional properties for Chart.js v4
}

interface ChartData {
  labels: string[];
  datasets: ChartDataset[];
}

interface ChartDataItem {
  legend?: string; // for pie/doughnut/bar/line
  value?: number;  // for pie/doughnut/bar/line
  x?: string;      // for bar/line
  y?: number;      // for bar/line
}

interface CustomChartRendererProps {
  title: string;
  data: ChartDataItem[];
  type: AdaptiveCardChartType;
  xAxisTitle?: string;
  yAxisTitle?: string;
}

const colorPalette = [
  "#FF6384", "#36A2EB", "#FFCE56", "#4BC0C0",
  "#9966FF", "#FF9F40", "#8BC34A", "#E91E63"
];

export const CustomChartRenderer: React.FC<CustomChartRendererProps> = ({ title, data, type, xAxisTitle, yAxisTitle }) => {
  // Convert Adaptive Card chart type to internal type
  const getInternalChartType = (chartType: AdaptiveCardChartType): ChartType => {
    switch (chartType) {
      case "Chart.Donut":
      case "Chart.Doughnut":
        return "doughnut";
      case "Chart.Pie":
        return "pie";
      case "Chart.VerticalBar":
      case "Chart.Bar":
      case "Chart.HorizontalBar":
        return "bar";
      case "Chart.Line":
        return "line";
      default:
        return "bar";
    }
  };

  const chartType = getInternalChartType(type);

  // Process data based on chart type
  const processData = (): ChartData => {
    if (chartType === 'bar' || chartType === 'line') {
      const labels = data.map(item => item.x || item.legend || '');
      const values = data.map(item => item.y || item.value || 0);
      
      return {
        labels,
        datasets: [{
          label: title,
          data: values,
          backgroundColor: chartType === 'line' ? 'rgba(54, 162, 235, 0.5)' : colorPalette.map(color => `${color}CC`),
          borderColor: chartType === 'line' ? colorPalette[0] : colorPalette.map(color => color),
          borderWidth: 2,
          fill: chartType === 'line' ? false : true,
          tension: chartType === 'line' ? 0.1 : undefined
        }]
      };
    } else {
      // For pie and doughnut charts
      const labels = data.map(item => item.legend || '');
      const values = data.map(item => item.value || 0);
      
      return {
        labels,
        datasets: [{
          data: values,
          backgroundColor: colorPalette.slice(0, values.length).map(color => `${color}DD`),
          borderColor: colorPalette.slice(0, values.length),
          borderWidth: 1,
          hoverOffset: 4
        }]
      };
    }
  };

  const baseOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      title: {
        display: !!title,
        text: title,
        font: { size: 16 }
      },
      legend: {
        position: "right" as const
      },
      tooltip: {
        enabled: true
      }
    }
  } as const;

  const processedData = processData();
  
  // Chart.js does not support gauge natively; show a message for Chart.Gauge
  if (type === "Chart.Gauge") {
    return (
      <div style={{ 
        height: '100%', 
        display: 'flex', 
        alignItems: 'center', 
        justifyContent: 'center',
        padding: '20px',
        color: '#605e5c',
        textAlign: 'center'
      }}>
        Gauge charts are not supported in this component.
      </div>
    );
  }

  const chartOptions = type === "Chart.HorizontalBar" ? {
    ...baseOptions,
    indexAxis: 'y' as const,
    scales: {
      x: {
        title: { 
          display: !!xAxisTitle, 
          text: xAxisTitle || '',
          font: { weight: 'bold' as const }
        }
      },
      y: {
        title: { 
          display: !!yAxisTitle, 
          text: yAxisTitle || '',
          font: { weight: 'bold' as const }
        }
      }
    }
  } : chartType === 'bar' || chartType === 'line' ? {
    ...baseOptions,
    scales: {
      x: {
        title: { 
          display: !!xAxisTitle, 
          text: xAxisTitle || '',
          font: { weight: 'bold' as const }
        }
      },
      y: {
        title: { 
          display: !!yAxisTitle, 
          text: yAxisTitle || '',
          font: { weight: 'bold' as const }
        },
        beginAtZero: true
      }
    }
  } : baseOptions;

  const renderChart = (): JSX.Element => {
    switch (chartType) {
      case "doughnut":
        return <Doughnut data={processedData} options={chartOptions} />;
      case "pie":
        return <Pie data={processedData} options={chartOptions} />;
      case "bar":
        return <Bar data={processedData} options={chartOptions} />;
      case "line":
        return <Line data={processedData} options={chartOptions} />;
      default:
        return <Bar data={processedData} options={chartOptions} />;
    }
  };

  return (
    <div style={{ 
      width: '100%', 
      height: '100%', 
      position: 'relative',
      display: 'flex',
      justifyContent: 'center',
      alignItems: 'center'
    }}>
      {renderChart()}
    </div>
  );
};

export default CustomChartRenderer;
