import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import * as AdaptiveCards from 'adaptivecards';
import { Template, IEvaluationContext } from 'adaptivecards-templating';
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

// Set global Chart.js font defaults to Segoe UI
ChartJS.defaults.font.family = '"Segoe UI", system-ui, sans-serif';

// Chart types
type AdaptiveCardChartType = "Chart.Donut" | "Chart.Doughnut" | "Chart.Pie" | "Chart.VerticalBar" | "Chart.Bar" | "Chart.HorizontalBar" | "Chart.Line" | "Chart.Gauge";
type ChartType = 'line' | 'bar' | 'pie' | 'doughnut';

// interface for all adaptive card elements
interface AdaptiveCardElement {
  type: string;
  size?: 'Small' | 'Medium' | 'Large';
  weight?: 'Normal' | 'Bold' | 'Bolder';
  color?: 'Dark' | 'Light' | 'Accent';
  wrap?: boolean;
  text?: string;
  title?: string;
  data?: ChartDataItem[];
  xAxisTitle?: string;
  yAxisTitle?: string;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  label?: string;
  placeholder?: string;
  value?: string;
  maxLength?: number;
  isRequired?: boolean;
  choices?: Array<{ title: string; value: string }>;
  style?: 'expanded' | 'compact';
  isMultiSelect?: boolean;
  columns?: AdaptiveCardElement[];
  items?: AdaptiveCardElement[];
  actions?: AdaptiveCardAction[];
  url?: string;
  altText?: string;
  horizontalAlignment?: 'Left' | 'Center' | 'Right';
  spacing?: 'None' | 'Small' | 'Default' | 'Medium' | 'Large' | 'ExtraLarge';
  separator?: boolean;
  height?: 'Auto' | 'Stretch';
  // ProgressBar properties
  min?: number;
  max?: number;
  progressValue?: number;
  // CompoundButton properties
  primaryText?: string;
  secondaryText?: string;
  iconProps?: { iconName?: string; };
  disabled?: boolean;
  checked?: boolean;
  onClick?: () => void;
  // Table properties
  tableColumns?: Array<{
    title?: string;
    width?: string | number;
    horizontalCellContentAlignment?: 'Left' | 'Center' | 'Right';
    verticalCellContentAlignment?: 'Top' | 'Center' | 'Bottom';
  }>;
  rows?: Array<{
    cells?: Array<{
      type?: string;
      text?: string;
      items?: AdaptiveCardElement[];
      horizontalAlignment?: 'Left' | 'Center' | 'Right';
      verticalAlignment?: 'Top' | 'Center' | 'Bottom';
      [key: string]: unknown;
    }>;
    [key: string]: unknown;
  }>;
  gridStyle?: 'Default' | 'Emphasis' | 'Accent' | 'Good' | 'Attention' | 'Warning' | 'Light' | 'Dark';
  showGridLines?: boolean;
  firstRowAsHeaders?: boolean;
  [key: string]: unknown;
}

interface AdaptiveCardAction {
  type: string;
  title?: string;
  url?: string;
  data?: unknown;
  content?: AdaptiveCardElement; // For Action.Popover content
  [key: string]: unknown;
}

interface ChartDataset {
  label?: string;
  data: number[];
  backgroundColor?: string | string[];
  borderColor?: string | string[];
  borderWidth?: number;
  fill?: boolean;
  tension?: number;
  hoverOffset?: number;
  [key: string]: unknown;
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

interface AdaptiveCardContentRendererProps {
  cardJson: string;
  data?: unknown;
  onActionExecute?: (action: AdaptiveCardAction | AdaptiveCards.Action) => void;
  useNativeRenderer?: boolean; // Flag to choose between custom and native renderer
}

interface ParsedCard {
  type?: string;
  version?: string;
  body?: AdaptiveCardElement[];
  actions?: AdaptiveCardAction[];
  $schema?: string;
}

const colorPalette = [
  "#FF6384", "#36A2EB", "#FFCE56", "#4BC0C0",
  "#9966FF", "#FF9F40", "#8BC34A", "#E91E63"
];

// Custom Chart Renderer Component
const CustomChartRenderer: React.FC<{
  title: string;
  data: ChartDataItem[];
  type: AdaptiveCardChartType;
  xAxisTitle?: string;
  yAxisTitle?: string;
}> = ({ title, data, type, xAxisTitle, yAxisTitle }) => {
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
        font: { 
          size: 16,
          family: '"Segoe UI", system-ui, sans-serif'
        }
      },
      legend: {
        position: "right" as const,
        labels: {
          font: {
            family: '"Segoe UI", system-ui, sans-serif'
          }
        }
      },
      tooltip: {
        enabled: true,
        titleFont: {
          family: '"Segoe UI", system-ui, sans-serif'
        },
        bodyFont: {
          family: '"Segoe UI", system-ui, sans-serif'
        }
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
          font: { 
            weight: 'bold' as const,
            family: '"Segoe UI", system-ui, sans-serif'
          }
        },
        ticks: {
          font: {
            family: '"Segoe UI", system-ui, sans-serif'
          }
        }
      },
      y: {
        title: { 
          display: !!yAxisTitle, 
          text: yAxisTitle || '',
          font: { 
            weight: 'bold' as const,
            family: '"Segoe UI", system-ui, sans-serif'
          }
        },
        ticks: {
          font: {
            family: '"Segoe UI", system-ui, sans-serif'
          }
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
          font: { 
            weight: 'bold' as const,
            family: '"Segoe UI", system-ui, sans-serif'
          }
        },
        ticks: {
          font: {
            family: '"Segoe UI", system-ui, sans-serif'
          }
        }
      },
      y: {
        title: { 
          display: !!yAxisTitle, 
          text: yAxisTitle || '',
          font: { 
            weight: 'bold' as const,
            family: '"Segoe UI", system-ui, sans-serif'
          }
        },
        ticks: {
          font: {
            family: '"Segoe UI", system-ui, sans-serif'
          }
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

// Native Adaptive Card Renderer Component
const NativeAdaptiveCardRenderer: React.FC<{
  cardJson: string;
  data?: unknown;
  onActionExecute?: (action: AdaptiveCards.Action) => void;
}> = ({ cardJson, data, onActionExecute }) => {
  const cardContainerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!cardContainerRef.current || !cardJson) {
      return;
    }

    try {
      // Create an AdaptiveCard instance
      const adaptiveCard = new AdaptiveCards.AdaptiveCard();

      // Set the host config for styling
      adaptiveCard.hostConfig = new AdaptiveCards.HostConfig({
        spacing: {
          small: 4,
          default: 8,
          medium: 16,
          large: 24,
          extraLarge: 32,
          padding: 16
        },
        separator: {
          lineThickness: 1,
          lineColor: "#EEEEEE"
        },
        supportsInteractivity: true,
        fontFamily: "Segoe UI, system-ui, sans-serif",
        fontSizes: {
          small: 12,
          default: 14,
          medium: 17,
          large: 21,
          extraLarge: 26
        },
        fontWeights: {
          lighter: 200,
          default: 400,
          bolder: 600
        },
        containerStyles: {
          default: {
            backgroundColor: "#FFFFFF",
            foregroundColors: {
              default: {
                default: "#333333",
                subtle: "#EE333333"
              },
              accent: {
                default: "#0078D4",
                subtle: "#880078D4"
              },
              attention: {
                default: "#CC3300",
                subtle: "#DDCC3300"
              },
              good: {
                default: "#54A254",
                subtle: "#DD54A254"
              },
              warning: {
                default: "#E69500",
                subtle: "#DDE69500"
              }
            }
          },
          emphasis: {
            backgroundColor: "#F3F2F1",
            foregroundColors: {
              default: {
                default: "#333333",
                subtle: "#EE333333"
              },
              accent: {
                default: "#0078D4",
                subtle: "#880078D4"
              }
            }
          }
        },
        imageSizes: {
          small: 40,
          medium: 80,
          large: 160
        },
        actions: {
          maxActions: 5,
          spacing: "default",
          buttonSpacing: 10,
          showCard: {
            actionMode: "inline",
            inlineTopMargin: 16
          },
          actionsOrientation: "horizontal",
          actionAlignment: "left"
        },
        adaptiveCard: {
          allowCustomStyle: false
        },
        imageSet: {
          imageSize: "medium",
          maxImageHeight: 100
        },
        media: {
          defaultPoster: "",
          allowInlinePlayback: false
        },
        factSet: {
          title: {
            color: "default",
            size: "default",
            isSubtle: false,
            weight: "bolder",
            wrap: true,
            maxWidth: 150
          },
          value: {
            color: "default",
            size: "default",
            isSubtle: false,
            weight: "default",
            wrap: true
          },
          spacing: 10
        }
      });

      // Set up action handling
      if (onActionExecute) {
        adaptiveCard.onExecuteAction = onActionExecute;
      }

      let cardPayload: unknown;

      // If data is provided, use templating
      if (data) {
        const template = new Template(JSON.parse(cardJson));
        const context: IEvaluationContext = {
          $root: data
        };
        cardPayload = template.expand(context);
      } else {
        cardPayload = JSON.parse(cardJson);
      }

      // Parse the card payload
      adaptiveCard.parse(cardPayload);

      // Clear the container and render the card
      cardContainerRef.current.innerHTML = '';
      const renderedCard = adaptiveCard.render();
      
      if (renderedCard) {
        cardContainerRef.current.appendChild(renderedCard);
      }

    } catch (error) {
      console.error('Error rendering AdaptiveCard:', error);
      
      // Show error message in the container
      if (cardContainerRef.current) {
        cardContainerRef.current.innerHTML = `
          <div style="padding: 16px; color: #a4262c; background-color: #fdf2f2; border: 1px solid #d13438; border-radius: 2px;">
            <strong>Error rendering card:</strong><br/>
            ${error instanceof Error ? error.message : 'Unknown error'}
          </div>
        `;
      }
    }
  }, [cardJson, data, onActionExecute]);

  return (
    <div 
      ref={cardContainerRef}
      style={{
        width: '100%',
        height: '100%',
        overflow: 'auto'
      }}
    />
  );
};

// Memoized Text Components
const TextBlock = React.memo(({ element }: { element: AdaptiveCardElement }) => {
  const fontSize = element.size === 'Large' ? '20px' : 
                  element.size === 'Medium' ? '16px' : 
                  element.size === 'Small' ? '12px' : '14px';
  
  const fontWeight = element.weight === 'Bolder' ? 'bold' : 
                    element.weight === 'Bold' ? '600' : 'normal';
  
  const color = element.color === 'Dark' ? '#323130' : 
                element.color === 'Light' ? '#605e5c' : 
                element.color === 'Accent' ? '#0078d4' : '#323130';

  const textAlign = element.horizontalAlignment === 'Center' ? 'center' :
                   element.horizontalAlignment === 'Right' ? 'right' : 'left';

  return (
    <div style={{
      fontSize,
      fontWeight,
      color,
      textAlign,
      marginBottom: element.spacing === 'Large' ? '16px' : '8px',
      marginTop: element.separator ? '16px' : '0',
      borderTop: element.separator ? '1px solid #e1dfdd' : 'none',
      paddingTop: element.separator ? '16px' : '0',
      wordWrap: element.wrap ? 'break-word' : 'normal'
    }}>
      {element.text}
    </div>
  );
});

const TextRun = React.memo(({ element }: { element: AdaptiveCardElement }) => {
  const fontSize = element.size === 'Large' ? '20px' : 
                  element.size === 'Medium' ? '16px' : 
                  element.size === 'Small' ? '12px' : '14px';
  
  const fontWeight = element.weight === 'Bolder' ? 'bold' : 
                    element.weight === 'Bold' ? '600' : 'normal';
  
  const color = element.color === 'Dark' ? '#323130' : 
                element.color === 'Light' ? '#605e5c' : 
                element.color === 'Accent' ? '#0078d4' : '#323130';

  const textDecoration = element.underline ? 'underline' : 
                        element.strikethrough ? 'line-through' : 'none';

  return (
    <span style={{
      fontSize,
      fontWeight,
      color,
      fontStyle: element.italic ? 'italic' : 'normal',
      textDecoration
    }}>
      {element.text}
    </span>
  );
});

// Input Components
const InputText = React.memo(({ element }: { element: AdaptiveCardElement }) => (
  <div style={{ marginBottom: '12px' }}>
    {element.label && (
      <label style={{ 
        display: 'block', 
        marginBottom: '4px', 
        fontSize: '14px',
        fontWeight: '600',
        color: '#323130' 
      }}>
        {element.label}
        {element.isRequired && <span style={{ color: '#a4262c' }}>*</span>}
      </label>
    )}
    <input
      type="text"
      placeholder={element.placeholder || ''}
      defaultValue={element.value || ''}
      maxLength={element.maxLength}
      required={element.isRequired}
      style={{
        width: '100%',
        padding: '8px 12px',
        border: '1px solid #8a8886',
        borderRadius: '2px',
        fontSize: '14px',
        fontFamily: '"Segoe UI", system-ui, sans-serif',
        outline: 'none'
      }}
    />
  </div>
));

const InputNumber = React.memo(({ element }: { element: AdaptiveCardElement }) => (
  <div style={{ marginBottom: '12px' }}>
    {element.label && (
      <label style={{ 
        display: 'block', 
        marginBottom: '4px', 
        fontSize: '14px',
        fontWeight: '600',
        color: '#323130' 
      }}>
        {element.label}
        {element.isRequired && <span style={{ color: '#a4262c' }}>*</span>}
      </label>
    )}
    <input
      type="number"
      placeholder={element.placeholder || ''}
      defaultValue={element.value || ''}
      required={element.isRequired}
      style={{
        width: '100%',
        padding: '8px 12px',
        border: '1px solid #8a8886',
        borderRadius: '2px',
        fontSize: '14px',
        fontFamily: '"Segoe UI", system-ui, sans-serif',
        outline: 'none'
      }}
    />
  </div>
));

const InputChoiceSet = React.memo(({ element }: { element: AdaptiveCardElement }) => {
  const isExpanded = element.style === 'expanded';
  
  if (isExpanded) {
    return (
      <div style={{ marginBottom: '12px' }}>
        {element.label && (
          <label style={{ 
            display: 'block', 
            marginBottom: '8px', 
            fontSize: '14px',
            fontWeight: '600',
            color: '#323130' 
          }}>
            {element.label}
            {element.isRequired && <span style={{ color: '#a4262c' }}>*</span>}
          </label>
        )}
        {element.choices?.map((choice, index) => (
          <div key={index} style={{ marginBottom: '8px' }}>
            <label style={{ display: 'flex', alignItems: 'center', cursor: 'pointer' }}>
              <input
                type={element.isMultiSelect ? 'checkbox' : 'radio'}
                name={`choice-${element.title || 'choice'}`}
                value={choice.value}
                style={{ marginRight: '8px' }}
              />
              <span style={{ fontSize: '14px', color: '#323130' }}>{choice.title}</span>
            </label>
          </div>
        ))}
      </div>
    );
  }

  return (
    <div style={{ marginBottom: '12px' }}>
      {element.label && (
        <label style={{ 
          display: 'block', 
          marginBottom: '4px', 
          fontSize: '14px',
          fontWeight: '600',
          color: '#323130' 
        }}>
          {element.label}
          {element.isRequired && <span style={{ color: '#a4262c' }}>*</span>}
        </label>
      )}
      <select
        required={element.isRequired}
        multiple={element.isMultiSelect}
        style={{
          width: '100%',
          padding: '8px 12px',
          border: '1px solid #8a8886',
          borderRadius: '2px',
          fontSize: '14px',
          fontFamily: '"Segoe UI", system-ui, sans-serif',
          backgroundColor: '#ffffff',
          outline: 'none'
        }}
      >
        {element.choices?.map((choice, index) => (
          <option key={index} value={choice.value}>
            {choice.title}
          </option>
        ))}
      </select>
    </div>
  );
});

// Container Components
const Container = React.memo(({ element, renderElement }: { 
  element: AdaptiveCardElement; 
  renderElement: (elem: AdaptiveCardElement) => JSX.Element;
}) => (
  <div style={{
    padding: '12px',
    border: '1px solid #e1dfdd',
    borderRadius: '4px',
    marginBottom: '12px',
    backgroundColor: '#faf9f8'
  }}>
    {element.items?.map((item, index) => (
      <div key={index}>
        {renderElement(item)}
      </div>
    ))}
  </div>
));

const ColumnSet = React.memo(({ element, renderElement }: { 
  element: AdaptiveCardElement; 
  renderElement: (elem: AdaptiveCardElement) => JSX.Element;
}) => (
  <div style={{
    display: 'flex',
    gap: '12px',
    marginBottom: '12px',
    flexWrap: 'wrap'
  }}>
    {element.columns?.map((column, index) => (
      <div key={index} style={{ flex: 1, minWidth: '200px' }}>
        {column.items?.map((item, itemIndex) => (
          <div key={itemIndex}>
            {renderElement(item)}
          </div>
        ))}
      </div>
    ))}
  </div>
));

const Image = React.memo(({ element }: { element: AdaptiveCardElement }) => (
  <div style={{ 
    marginBottom: '12px',
    textAlign: element.horizontalAlignment === 'Center' ? 'center' :
               element.horizontalAlignment === 'Right' ? 'right' : 'left'
  }}>
    <img
      src={element.url || ''}
      alt={element.altText || 'Image'}
      style={{
        maxWidth: '100%',
        height: 'auto',
        borderRadius: '4px'
      }}
    />
  </div>
));

const ChartRenderer = React.memo(({ element }: { element: AdaptiveCardElement }) => (
  <div style={{ height: '300px', marginBottom: '16px' }}>
    <CustomChartRenderer
      title={element.title || ''}
      data={element.data || []}
      type={(element.type as AdaptiveCardChartType) || 'Chart.Bar'}
      xAxisTitle={element.xAxisTitle || ''}
      yAxisTitle={element.yAxisTitle || ''}
    />
  </div>
));

const ProgressBar = React.memo(({ element }: { element: AdaptiveCardElement }) => {
  const min = element.min ?? 0;
  const max = element.max ?? 100;
  const value = element.progressValue ?? 0;
  
  // Ensure value is within bounds
  const clampedValue = Math.min(Math.max(value, min), max);
  const percentage = max > min ? ((clampedValue - min) / (max - min)) * 100 : 0;
  
  return (
    <div style={{ 
      marginBottom: '12px',
      marginTop: element.separator ? '16px' : '0',
      borderTop: element.separator ? '1px solid #e1dfdd' : 'none',
      paddingTop: element.separator ? '16px' : '0'
    }}>
      {element.title && (
        <div style={{ 
          marginBottom: '8px', 
          fontSize: '14px',
          fontWeight: '600',
          color: '#323130'
        }}>
          {element.title}
        </div>
      )}
      <div style={{
        width: '100%',
        height: '8px',
        backgroundColor: '#f3f2f1',
        borderRadius: '4px',
        overflow: 'hidden',
        border: '1px solid #edebe9'
      }}>
        <div style={{
          width: `${percentage}%`,
          height: '100%',
          backgroundColor: '#0078d4',
          borderRadius: '3px',
          transition: 'width 0.3s ease-in-out'
        }} />
      </div>
      <div style={{
        display: 'flex',
        justifyContent: 'space-between',
        marginTop: '4px',
        fontSize: '12px',
        color: '#605e5c'
      }}>
        <span>{min}</span>
        <span>{clampedValue} / {max}</span>
        <span>{max}</span>
      </div>
    </div>
  );
});

// CompoundButton Component
const CompoundButton = React.memo(({ element }: { element: AdaptiveCardElement }) => {
  const [isPressed, setIsPressed] = useState(false);
  const [isHovered, setIsHovered] = useState(false);
  
  const handleClick = () => {
    if (element.onClick) {
      element.onClick();
    }
    // Add ripple effect
    setIsPressed(true);
    setTimeout(() => setIsPressed(false), 150);
  };

  const handleMouseEnter = () => setIsHovered(true);
  const handleMouseLeave = () => {
    setIsHovered(false);
    setIsPressed(false);
  };

  // Get button style based on element properties
  const getButtonStyle = (): React.CSSProperties => {
    const baseStyle: React.CSSProperties = {
      display: 'flex',
      alignItems: 'center',
      padding: '12px 16px',
      border: '1px solid #8a8886',
      borderRadius: '4px',
      backgroundColor: '#ffffff',
      cursor: element.disabled ? 'not-allowed' : 'pointer',
      fontFamily: '"Segoe UI", system-ui, sans-serif',
      fontSize: '14px',
      color: element.disabled ? '#a6a6a6' : '#323130',
      transition: 'all 0.2s ease',
      marginBottom: element.spacing === 'Large' ? '16px' : '8px',
      marginTop: element.separator ? '16px' : '0',
      borderTop: element.separator ? '2px solid #e1dfdd' : undefined,
      paddingTop: element.separator ? '16px' : '12px',
      opacity: element.disabled ? 0.6 : 1,
      minHeight: '52px',
      width: '100%',
      textAlign: 'left' as const,
      position: 'relative' as const,
      overflow: 'hidden' as const
    };

    // Apply hover and pressed states
    if (!element.disabled) {
      if (isPressed) {
        baseStyle.backgroundColor = '#f3f2f1';
        baseStyle.borderColor = '#605e5c';
        baseStyle.transform = 'scale(0.98)';
      } else if (isHovered) {
        baseStyle.backgroundColor = '#f8f8f8';
        baseStyle.borderColor = '#605e5c';
        baseStyle.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
      }
    }

    // Handle checked state
    if (element.checked) {
      baseStyle.backgroundColor = '#deecf9';
      baseStyle.borderColor = '#0078d4';
      baseStyle.color = '#005a9e';
    }

    // Handle different button styles
    if (element.style) {
      switch (element.style.toLowerCase()) {
        case 'emphasis':
          baseStyle.backgroundColor = element.checked ? '#c7e0f4' : '#f3f2f1';
          baseStyle.borderColor = '#8a8886';
          break;
        case 'positive':
        case 'good':
          baseStyle.backgroundColor = element.checked ? '#c6efce' : '#f3f9f4';
          baseStyle.borderColor = '#107c10';
          baseStyle.color = element.disabled ? '#a6a6a6' : '#107c10';
          break;
        case 'attention':
        case 'warning':
          baseStyle.backgroundColor = element.checked ? '#fff4ce' : '#fffef5';
          baseStyle.borderColor = '#ffb900';
          baseStyle.color = element.disabled ? '#a6a6a6' : '#8a6e00';
          break;
        case 'destructive':
          baseStyle.backgroundColor = element.checked ? '#fde7e9' : '#fef6f6';
          baseStyle.borderColor = '#d13438';
          baseStyle.color = element.disabled ? '#a6a6a6' : '#d13438';
          break;
      }
    }

    return baseStyle;
  };

  // Render icon if provided
  const renderIcon = () => {
    if (element.iconProps?.iconName) {
      return (
        <div style={{
          marginRight: '12px',
          fontSize: '16px',
          color: 'inherit',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          width: '20px',
          height: '20px'
        }}>
          {/* Simple icon representation - you can replace with actual icon component */}
          <span title={element.iconProps.iconName}>
            {getIconSymbol(element.iconProps.iconName)}
          </span>
        </div>
      );
    }
    return null;
  };

  // Simple icon symbol mapping - you can extend this or use actual icon library
  const getIconSymbol = (iconName: string): string => {
    const iconMap: { [key: string]: string } = {
      'Add': '+',
      'Edit': '✏️',
      'Delete': '🗑️',
      'Save': '💾',
      'Cancel': '❌',
      'Search': '🔍',
      'Filter': '🔽',
      'Sort': '↕️',
      'Settings': '⚙️',
      'Info': 'ℹ️',
      'Warning': '⚠️',
      'Error': '❌',
      'Success': '✅',
      'Download': '⬇️',
      'Upload': '⬆️',
      'Print': '🖨️',
      'Mail': '📧',
      'Phone': '📞',
      'Calendar': '📅',
      'Clock': '🕐',
      'Location': '📍',
      'Home': '🏠',
      'Folder': '📁',
      'File': '📄',
      'Image': '🖼️',
      'Video': '🎥',
      'Audio': '🔊',
      'Link': '🔗',
      'Share': '📤',
      'Forward': '➡️',
      'Back': '⬅️',
      'Up': '⬆️',
      'Down': '⬇️',
      'More': '⋯',
      'Close': '✖️',
      'Minimize': '🔽',
      'Maximize': '🔼'
    };
    return iconMap[iconName] || '●';
  };

  return (
    <button
      style={getButtonStyle()}
      onClick={handleClick}
      onMouseEnter={handleMouseEnter}
      onMouseLeave={handleMouseLeave}
      disabled={element.disabled}
      title={element.title}
      aria-pressed={element.checked}
      aria-disabled={element.disabled}
    >
      {/* Ripple effect */}
      {isPressed && (
        <div
          style={{
            position: 'absolute',
            top: '50%',
            left: '50%',
            width: '100px',
            height: '100px',
            borderRadius: '50%',
            backgroundColor: 'rgba(0, 120, 212, 0.3)',
            transform: 'translate(-50%, -50%)',
            animation: 'ripple 0.6s linear',
            pointerEvents: 'none'
          }}
        />
      )}
      
      {/* Button content */}
      <div style={{ display: 'flex', alignItems: 'center', width: '100%', zIndex: 1 }}>
        {renderIcon()}
        <div style={{ flex: 1 }}>
          {/* Primary text */}
          <div style={{
            fontWeight: '600',
            fontSize: '14px',
            lineHeight: '20px',
            marginBottom: element.secondaryText ? '2px' : '0'
          }}>
            {element.primaryText || element.text || element.title || 'Button'}
          </div>
          
          {/* Secondary text */}
          {element.secondaryText && (
            <div style={{
              fontSize: '12px',
              lineHeight: '16px',
              color: element.disabled ? '#a6a6a6' : '#605e5c',
              fontWeight: '400'
            }}>
              {element.secondaryText}
            </div>
          )}
        </div>
        
        {/* Checked indicator */}
        {element.checked && (
          <div style={{
            marginLeft: '8px',
            color: '#0078d4',
            fontSize: '14px',
            fontWeight: 'bold'
          }}>
            ✓
          </div>
        )}
      </div>
      
      <style>{`
        @keyframes ripple {
          to {
            transform: translate(-50%, -50%) scale(4);
            opacity: 0;
          }
        }
      `}</style>
    </button>
  );
});

// Table Component
const Table = React.memo(({ element, renderElement }: { 
  element: AdaptiveCardElement; 
  renderElement: (elem: AdaptiveCardElement) => JSX.Element;
}) => {
  const getGridStyle = (gridStyle: string = 'Default') => {
    const styleMap: { [key: string]: React.CSSProperties } = {
      'Default': {
        backgroundColor: '#ffffff',
        borderColor: '#e1dfdd'
      },
      'Emphasis': {
        backgroundColor: '#f3f2f1',
        borderColor: '#d2d0ce'
      },
      'Accent': {
        backgroundColor: '#deecf9',
        borderColor: '#0078d4'
      },
      'Good': {
        backgroundColor: '#dff6dd',
        borderColor: '#107c10'
      },
      'Attention': {
        backgroundColor: '#fff4ce',
        borderColor: '#ffb900'
      },
      'Warning': {
        backgroundColor: '#fed9cc',
        borderColor: '#d83b01'
      },
      'Light': {
        backgroundColor: '#faf9f8',
        borderColor: '#edebe9'
      },
      'Dark': {
        backgroundColor: '#323130',
        borderColor: '#605e5c'
      }
    };
    return styleMap[gridStyle] || styleMap['Default'];
  };

  const getAlignment = (alignment?: string): React.CSSProperties['textAlign'] => {
    switch (alignment?.toLowerCase()) {
      case 'center': return 'center';
      case 'right': return 'right';
      default: return 'left';
    }
  };

  const getVerticalAlignment = (alignment?: string): React.CSSProperties['verticalAlign'] => {
    switch (alignment?.toLowerCase()) {
      case 'center': return 'middle';
      case 'bottom': return 'bottom';
      default: return 'top';
    }
  };

  const gridStyle = getGridStyle(element.gridStyle);
  const showGridLines = element.showGridLines !== false; // Default to true
  const firstRowAsHeaders = element.firstRowAsHeaders !== false; // Default to true

  const tableStyle: React.CSSProperties = {
    width: '100%',
    borderCollapse: 'collapse' as const,
    fontFamily: '"Segoe UI", system-ui, sans-serif',
    fontSize: '14px',
    marginBottom: element.spacing === 'Large' ? '16px' : '8px',
    marginTop: element.separator ? '16px' : '0',
    borderTop: element.separator ? '2px solid #e1dfdd' : 'none',
    ...gridStyle,
    border: showGridLines ? `1px solid ${gridStyle.borderColor}` : 'none'
  };

  const headerCellStyle: React.CSSProperties = {
    padding: '12px 16px',
    fontWeight: '600',
    backgroundColor: element.gridStyle === 'Dark' ? '#484644' : '#f8f7f6',
    color: element.gridStyle === 'Dark' ? '#ffffff' : '#323130',
    border: showGridLines ? `1px solid ${gridStyle.borderColor}` : 'none',
    textAlign: 'left',
    verticalAlign: 'middle'
  };

  const cellStyle: React.CSSProperties = {
    padding: '12px 16px',
    border: showGridLines ? `1px solid ${gridStyle.borderColor}` : 'none',
    color: element.gridStyle === 'Dark' ? '#ffffff' : '#323130',
    verticalAlign: 'top'
  };

  const renderCellContent = (cell: any): JSX.Element => {
    if (cell.items && Array.isArray(cell.items)) {
      // Cell contains adaptive card elements
      return (
        <div>
          {cell.items.map((item: AdaptiveCardElement, index: number) => (
            <div key={index} style={{ marginBottom: index < cell.items.length - 1 ? '4px' : '0' }}>
              {renderElement(item)}
            </div>
          ))}
        </div>
      );
    } else if (cell.text) {
      // Cell contains plain text
      return <span>{cell.text}</span>;
    } else if (cell.type === 'TextBlock') {
      // Cell is a TextBlock element
      return renderElement(cell as AdaptiveCardElement);
    } else {
      // Fallback to string representation
      return <span>{String(cell)}</span>;
    }
  };

  return (
    <div style={{ 
      overflowX: 'auto',
      marginBottom: '12px',
      marginTop: element.separator ? '16px' : '0',
      borderTop: element.separator ? '2px solid #e1dfdd' : 'none',
      paddingTop: element.separator ? '16px' : '0'
    }}>
      <table style={tableStyle}>
        {/* Table Headers */}
        {firstRowAsHeaders && element.tableColumns && element.tableColumns.length > 0 && (
          <thead>
            <tr>
              {element.tableColumns.map((column, index) => (
                <th
                  key={index}
                  style={{
                    ...headerCellStyle,
                    width: column.width || 'auto',
                    textAlign: getAlignment(column.horizontalCellContentAlignment),
                    verticalAlign: getVerticalAlignment(column.verticalCellContentAlignment)
                  }}
                >
                  {column.title || `Column ${index + 1}`}
                </th>
              ))}
            </tr>
          </thead>
        )}
        
        {/* Table Body */}
        <tbody>
          {element.rows?.map((row, rowIndex) => {
            // Skip first row if it's used as headers
            if (firstRowAsHeaders && rowIndex === 0 && !element.tableColumns) {
              return (
                <tr key={rowIndex}>
                  {row.cells?.map((cell, cellIndex) => (
                    <th
                      key={cellIndex}
                      style={{
                        ...headerCellStyle,
                        textAlign: getAlignment(cell.horizontalAlignment),
                        verticalAlign: getVerticalAlignment(cell.verticalAlignment)
                      }}
                    >
                      {renderCellContent(cell)}
                    </th>
                  ))}
                </tr>
              );
            }

            return (
              <tr key={rowIndex} style={{
                backgroundColor: rowIndex % 2 === 0 ? 'transparent' : 
                  (element.gridStyle === 'Dark' ? '#3c3a39' : '#faf9f8')
              }}>
                {row.cells?.map((cell, cellIndex) => (
                  <td
                    key={cellIndex}
                    style={{
                      ...cellStyle,
                      textAlign: getAlignment(cell.horizontalAlignment),
                      verticalAlign: getVerticalAlignment(cell.verticalAlignment)
                    }}
                  >
                    {renderCellContent(cell)}
                  </td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>

      {/* Empty state */}
      {(!element.rows || element.rows.length === 0) && (
        <div style={{
          padding: '20px',
          textAlign: 'center',
          color: element.gridStyle === 'Dark' ? '#ffffff' : '#605e5c',
          fontStyle: 'italic',
          backgroundColor: gridStyle.backgroundColor,
          border: showGridLines ? `1px solid ${gridStyle.borderColor}` : 'none',
          borderRadius: '4px'
        }}>
          No data available
        </div>
      )}
    </div>
  );
});

// Action Components
const PopoverRenderer = React.memo(({ action, renderElement }: { 
  action: AdaptiveCardAction; 
  renderElement: (elem: AdaptiveCardElement) => JSX.Element;
}) => {
  const [open, setOpen] = useState(false);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const popoverRef = useRef<HTMLDivElement>(null);

  // Close popover when clicking outside or pressing escape
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (popoverRef.current && !popoverRef.current.contains(event.target as Node)) {
        setOpen(false);
      }
    };

    const handleEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setOpen(false);
      }
    };

    if (open) {
      document.addEventListener('mousedown', handleClickOutside);
      document.addEventListener('keydown', handleEscape);
      // Prevent body scroll when modal is open
      document.body.style.overflow = 'hidden';
    }

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
      document.removeEventListener('keydown', handleEscape);
      // Restore body scroll when modal is closed
      document.body.style.overflow = '';
    };
  }, [open]);

  const handleButtonClick = () => {
    setOpen(!open);
  };

  const handleCloseClick = () => {
    setOpen(false);
  };

  const popoverContent = open ? createPortal(
    <>
      {/* Modal Backdrop */}
      <div
        style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: 'rgba(0, 0, 0, 0.4)',
          zIndex: 9999,
          animation: 'backdropFadeIn 0.2s ease-out forwards'
        }}
      />
      {/* Modal Content */}
      <div
        ref={popoverRef}
        style={{
          position: 'fixed',
          top: '50%',
          left: '50%',
          transform: 'translate(-50%, -50%)',
          width: '80vw',
          maxWidth: '800px',
          minWidth: '400px',
          maxHeight: '90vh',
          backgroundColor: '#ffffff',
          border: '1px solid #8a8886',
          borderRadius: '8px',
          boxShadow: '0 16px 48px rgba(0,0,0,0.25), 0 0 0 1px rgba(0,0,0,0.05)',
          zIndex: 10000,
          fontFamily: '"Segoe UI", system-ui, sans-serif',
          animation: 'modalFadeIn 0.2s ease-out forwards',
          display: 'flex',
          flexDirection: 'column'
        }}
      >
        <style>{`
          @keyframes backdropFadeIn {
            from {
              opacity: 0;
            }
            to {
              opacity: 1;
            }
          }
          @keyframes modalFadeIn {
            from {
              opacity: 0;
              transform: translate(-50%, -50%) scale(0.9);
            }
            to {
              opacity: 1;
              transform: translate(-50%, -50%) scale(1);
            }
          }
        `}</style>
        
        {/* Header with Close Button */}
        <div style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          padding: '16px 20px',
          borderBottom: '1px solid #edebe9',
          flexShrink: 0
        }}>
          <h3 style={{
            margin: 0,
            fontSize: '18px',
            fontWeight: '600',
            color: '#323130',
            fontFamily: '"Segoe UI", system-ui, sans-serif'
          }}>
            {action.title || 'Popover Content'}
          </h3>
          <button
            onClick={handleCloseClick}
            style={{
              background: 'none',
              border: 'none',
              fontSize: '20px',
              cursor: 'pointer',
              color: '#605e5c',
              padding: '4px',
              borderRadius: '2px',
              lineHeight: 1,
              fontFamily: '"Segoe UI", system-ui, sans-serif',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              width: '32px',
              height: '32px',
              transition: 'all 0.2s ease'
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#f3f2f1';
              e.currentTarget.style.color = '#323130';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = 'transparent';
              e.currentTarget.style.color = '#605e5c';
            }}
            title="Close popover"
            aria-label="Close popover"
          >
            ×
          </button>
        </div>
        
        {/* Content Area */}
        <div style={{
          padding: '20px',
          overflowY: 'auto',
          flex: 1,
          minHeight: 0 // Important for flex child to be scrollable
        }}>
          {action.content && renderElement(action.content)}
          {!action.content && (
            <div style={{ 
              color: '#605e5c', 
              fontStyle: 'italic',
              fontFamily: '"Segoe UI", system-ui, sans-serif',
              textAlign: 'center',
              padding: '40px 0'
            }}>
              No content available for this popover
            </div>
          )}
        </div>
      </div>
    </>,
    document.body
  ) : null;

  return (
    <>
      <button
        ref={buttonRef}
        onClick={handleButtonClick}
        style={{
          padding: '8px 16px',
          backgroundColor: open ? '#c73e00' : '#d83b01',
          color: '#ffffff',
          border: '1px solid #d83b01',
          borderRadius: '2px',
          fontSize: '14px',
          cursor: 'pointer',
          fontFamily: '"Segoe UI", system-ui, sans-serif',
          fontWeight: '600',
          transition: 'all 0.2s ease',
          transform: open ? 'scale(0.98)' : 'scale(1)',
          boxShadow: open ? 'inset 0 2px 4px rgba(0,0,0,0.1)' : 'none'
        }}
        title={`Action.Popover: ${action.title}`}
        aria-expanded={open}
        aria-haspopup="dialog"
      >
        {action.title}
      </button>
      {popoverContent}
    </>
  );
});

const ActionSet = React.memo(({ element, onActionExecute, renderElement }: { 
  element: AdaptiveCardElement; 
  onActionExecute?: (action: AdaptiveCardAction) => void;
  renderElement?: (elem: AdaptiveCardElement) => JSX.Element;
}) => {
  const handleActionClick = (action: AdaptiveCardAction) => {
    switch (action.type) {
      case 'Action.OpenUrl':
        if (action.url) {
          window.open(action.url, '_blank', 'noopener,noreferrer');
        }
        break;
      case 'Action.ShowCard':
        // Handle show card action - could expand inline or show in modal
        console.log('ShowCard action triggered:', action);
        onActionExecute?.(action);
        break;
      case 'Action.Execute':
        // Handle execute action - typically for custom functionality
        console.log('Execute action triggered:', action);
        onActionExecute?.(action);
        break;
      case 'Action.Popover':
        // Popover is handled by PopoverRenderer component directly
        console.log('Popover action triggered:', action);
        break;
      case 'Action.Submit':
      default:
        // Handle submit and other actions through the callback
        onActionExecute?.(action);
        break;
    }
  };

  const getActionButtonStyle = (actionType: string) => {
    switch (actionType) {
      case 'Action.Submit':
        return {
          backgroundColor: '#0078d4',
          color: '#ffffff',
          border: '1px solid #0078d4'
        };
      case 'Action.OpenUrl':
        return {
          backgroundColor: '#107c10',
          color: '#ffffff',
          border: '1px solid #107c10'
        };
      case 'Action.Execute':
        return {
          backgroundColor: '#8764b8',
          color: '#ffffff',
          border: '1px solid #8764b8'
        };
      case 'Action.ShowCard':
        return {
          backgroundColor: '#ca5010',
          color: '#ffffff',
          border: '1px solid #ca5010'
        };
      case 'Action.Popover':
        return {
          backgroundColor: '#d83b01',
          color: '#ffffff',
          border: '1px solid #d83b01'
        };
      default:
        return {
          backgroundColor: '#f3f2f1',
          color: '#323130',
          border: '1px solid #8a8886'
        };
    }
  };

  return (
    <div style={{ 
      display: 'flex', 
      gap: '8px', 
      marginTop: '16px',
      flexWrap: 'wrap'
    }}>
      {element.actions?.map((action, index) => {
        // Special handling for Action.Popover
        if (action.type === 'Action.Popover') {
          return renderElement ? (
            <PopoverRenderer
              key={index}
              action={action}
              renderElement={renderElement}
            />
          ) : (
            <div key={index} style={{ color: '#a4262c', fontSize: '12px' }}>
              Popover requires renderElement function
            </div>
          );
        }

        // Regular button rendering for other actions
        return (
          <button
            key={index}
            onClick={() => handleActionClick(action)}
            style={{
              padding: '8px 16px',
              borderRadius: '2px',
              fontSize: '14px',
              cursor: 'pointer',
              fontFamily: '"Segoe UI", system-ui, sans-serif',
              fontWeight: '600',
              transition: 'all 0.2s ease',
              ...getActionButtonStyle(action.type)
            }}
            title={`${action.type}: ${action.title}`}
          >
            {action.title}
          </button>
        );
      })}
    </div>
  );
});

const UnsupportedElement = React.memo(({ type }: { type: string }) => (
  <div style={{ 
    padding: '8px', 
    backgroundColor: '#fff4ce', 
    border: '1px solid #ffb900',
    borderRadius: '4px',
    marginBottom: '8px',
    fontSize: '12px',
    color: '#323130'
  }}>
    Unsupported element type: {type}
  </div>
));

export const AdaptiveCardContentRenderer: React.FC<AdaptiveCardContentRendererProps> = React.memo(({ 
  cardJson, 
  data,
  onActionExecute,
  useNativeRenderer = false
}) => {
  const [parsedCard, setParsedCard] = React.useState<ParsedCard | null>(null);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    try {
      if (cardJson) {
        const card = JSON.parse(cardJson) as ParsedCard;
        setParsedCard(card);
        setError(null);
      }
    } catch (err) {
      setError('Failed to parse Adaptive Card JSON');
      console.error('Error parsing Adaptive Card:', err);
    }
  }, [cardJson]);

  const renderElement = React.useCallback((element: AdaptiveCardElement): JSX.Element => {
    switch (element.type) {
      // Text Elements
      case 'TextBlock':
        return <TextBlock element={element} />;
      
      case 'TextRun':
        return <TextRun element={element} />;
      
      // Input Elements
      case 'Input.Text':
        return <InputText element={element} />;
      
      case 'Input.Number':
        return <InputNumber element={element} />;
      
      case 'Input.ChoiceSet':
        return <InputChoiceSet element={element} />;
      
      // Container Elements
      case 'Container':
        return <Container element={element} renderElement={renderElement} />;
      
      case 'ColumnSet':
        return <ColumnSet element={element} renderElement={renderElement} />;
      
      // Media Elements
      case 'Image':
        return <Image element={element} />;
      
      // Progress Elements
      case 'ProgressBar':
        return <ProgressBar element={element} />;
      
      // Button Elements
      case 'CompoundButton':
        return <CompoundButton element={element} />;
      
      // Table Elements
      case 'Table':
        return <Table element={element} renderElement={renderElement} />;
      
      // Chart Elements
      case 'Chart':
      case 'Chart.Bar':
      case 'Chart.Line':
      case 'Chart.Pie':
      case 'Chart.Donut':
      case 'Chart.Doughnut':
      case 'Chart.VerticalBar':
      case 'Chart.HorizontalBar':
      case 'Chart.Gauge':
        return <ChartRenderer element={element} />;
      
      // Action Elements
      case 'ActionSet':
        return <ActionSet element={element} onActionExecute={onActionExecute as (action: AdaptiveCardAction) => void} renderElement={renderElement} />;
      
      default:
        return <UnsupportedElement type={element.type} />;
    }
  }, [onActionExecute]);

  // Use native renderer if specified
  if (useNativeRenderer) {
    return (
      <NativeAdaptiveCardRenderer 
        cardJson={cardJson} 
        data={data} 
        onActionExecute={onActionExecute as (action: AdaptiveCards.Action) => void} 
      />
    );
  }

  if (error) {
    return (
      <div style={{ 
        padding: '20px', 
        textAlign: 'center', 
        color: '#a4262c',
        backgroundColor: '#fdf3f4',
        border: '1px solid #a4262c',
        borderRadius: '4px'
      }}>
        {error}
      </div>
    );
  }

  if (!parsedCard) {
    return (
      <div style={{ 
        padding: '20px', 
        textAlign: 'center', 
        color: '#605e5c' 
      }}>
        Loading...
      </div>
    );
  }

  return (
    <div style={{ 
      padding: '16px',
      height: '100%',
      overflow: 'auto'
    }}>
      {/* Render card body */}
      {parsedCard.body?.map((element: AdaptiveCardElement, index: number) => (
        <React.Fragment key={index}>
          {renderElement(element)}
        </React.Fragment>
      ))}
      
      {/* Render card actions */}
      {parsedCard.actions && parsedCard.actions.length > 0 && (
        <ActionSet 
          element={{ actions: parsedCard.actions } as AdaptiveCardElement} 
          onActionExecute={onActionExecute as (action: AdaptiveCardAction) => void} 
          renderElement={renderElement}
        />
      )}
    </div>
  );
});

export default AdaptiveCardContentRenderer;
