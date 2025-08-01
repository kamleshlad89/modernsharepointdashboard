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

// Action Components
const PopoverRenderer = React.memo(({ action, renderElement }: { 
  action: AdaptiveCardAction; 
  renderElement: (elem: AdaptiveCardElement) => JSX.Element;
}) => {
  const [open, setOpen] = useState(false);
  const [position, setPosition] = useState({ top: 0, left: 0 });
  const buttonRef = useRef<HTMLButtonElement>(null);
  const popoverRef = useRef<HTMLDivElement>(null);

  // Calculate popover position relative to button
  const calculatePosition = () => {
    if (buttonRef.current) {
      const buttonRect = buttonRef.current.getBoundingClientRect();
      const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
      const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
      
      // Position below the button with some margin
      const top = buttonRect.bottom + scrollTop + 8;
      const left = buttonRect.left + scrollLeft;
      
      // Ensure popover doesn't go off-screen
      const viewportWidth = window.innerWidth;
      const popoverWidth = 350;
      const adjustedLeft = Math.min(left, viewportWidth - popoverWidth - 20);
      
      setPosition({
        top: top,
        left: Math.max(10, adjustedLeft) // Minimum 10px from left edge
      });
    }
  };

  // Close popover when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (popoverRef.current && !popoverRef.current.contains(event.target as Node) &&
          buttonRef.current && !buttonRef.current.contains(event.target as Node)) {
        setOpen(false);
      }
    };

    const handleEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setOpen(false);
      }
    };

    const handleScroll = () => {
      if (open) {
        calculatePosition();
      }
    };

    const handleResize = () => {
      if (open) {
        calculatePosition();
      }
    };

    if (open) {
      document.addEventListener('mousedown', handleClickOutside);
      document.addEventListener('keydown', handleEscape);
      window.addEventListener('scroll', handleScroll, true);
      window.addEventListener('resize', handleResize);
      calculatePosition();
    }

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
      document.removeEventListener('keydown', handleEscape);
      window.removeEventListener('scroll', handleScroll, true);
      window.removeEventListener('resize', handleResize);
    };
  }, [open]);

  const handleButtonClick = () => {
    setOpen(!open);
  };

  const popoverContent = open ? createPortal(
    <div
      ref={popoverRef}
      style={{
        position: 'absolute',
        top: `${position.top}px`,
        left: `${position.left}px`,
        width: '400px',
        maxWidth: '90vw',
        padding: '16px',
        backgroundColor: '#ffffff',
        border: '1px solid #8a8886',
        borderRadius: '4px',
        boxShadow: '0 8px 24px rgba(0,0,0,0.15), 0 0 0 1px rgba(0,0,0,0.05)',
        zIndex: 10000,
        maxHeight: '400px',
        overflowY: 'auto',
        fontFamily: '"Segoe UI", system-ui, sans-serif',
        animation: 'popoverFadeIn 0.15s ease-out forwards'
      }}
    >
      <style>{`
        @keyframes popoverFadeIn {
          from {
            opacity: 0;
            transform: translateY(-8px) scale(0.95);
          }
          to {
            opacity: 1;
            transform: translateY(0) scale(1);
          }
        }
      `}</style>
      {action.content && renderElement(action.content)}
      {!action.content && (
        <div style={{ 
          color: '#605e5c', 
          fontStyle: 'italic',
          fontFamily: '"Segoe UI", system-ui, sans-serif'
        }}>
          No content available for this popover
        </div>
      )}
    </div>,
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
