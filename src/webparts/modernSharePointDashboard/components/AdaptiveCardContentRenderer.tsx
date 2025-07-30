import * as React from 'react';
import { CustomChartRenderer } from './CustomChartRenderer';

interface AdaptiveCardElement {
  type: string;
  [key: string]: any;
}

interface AdaptiveCardContentRendererProps {
  cardJson: string;
  onActionExecute?: (action: any) => void;
}

export const AdaptiveCardContentRenderer: React.FC<AdaptiveCardContentRendererProps> = ({ 
  cardJson, 
  onActionExecute 
}) => {
  const [parsedCard, setParsedCard] = React.useState<any>(null);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    try {
      const card = JSON.parse(cardJson);
      setParsedCard(card);
      setError(null);
    } catch (err) {
      setError('Failed to parse Adaptive Card JSON');
      console.error('Error parsing Adaptive Card:', err);
    }
  }, [cardJson]);

  const handleActionClick = (action: any): void => {
    if (action.type === 'Action.OpenUrl' && action.url) {
      window.open(action.url, '_blank');
    }
    
    if (onActionExecute) {
      onActionExecute(action);
    }
  };

  const renderElement = (element: AdaptiveCardElement): JSX.Element => {
    switch (element.type) {
      case 'TextBlock':
        return renderTextBlock(element);
      
      case 'RichTextBlock':
        return renderRichTextBlock(element);
      
      case 'Image':
        return renderImage(element);
      
      case 'ImageSet':
        return renderImageSet(element);
      
      case 'Media':
        return renderMedia(element);
      
      case 'Table':
        return renderTable(element);
      
      case 'FactSet':
        return renderFactSet(element);
      
      case 'ActionSet':
        return renderActionSet(element);
      
      case 'ColumnSet':
        return renderColumnSet(element);
      
      case 'Container':
        return renderContainer(element);
      
      case 'Input.Text':
        return renderInputText(element);
      
      case 'Input.Number':
        return renderInputNumber(element);
      
      case 'Input.Date':
        return renderInputDate(element);
      
      case 'Input.Time':
        return renderInputTime(element);
      
      case 'Input.Toggle':
        return renderInputToggle(element);
      
      case 'Input.ChoiceSet':
        return renderInputChoiceSet(element);
      
      case 'TextRun':
        return renderTextRun(element);
      
      case 'Chart':
      case 'Chart.Bar':
      case 'Chart.Line':
      case 'Chart.Pie':
      case 'Chart.Donut':
      case 'Chart.Doughnut':
        return renderChart(element);
      
      default:
        return (
          <div style={{ 
            padding: '8px', 
            backgroundColor: '#fff4ce', 
            border: '1px solid #ffb900',
            borderRadius: '4px',
            marginBottom: '8px',
            fontSize: '12px',
            color: '#323130'
          }}>
            Unsupported element type: {element.type}
          </div>
        );
    }
  };

  const renderTextBlock = (element: any): JSX.Element => {
    const fontSize = element.size === 'Large' ? '20px' : 
                    element.size === 'Medium' ? '16px' : 
                    element.size === 'Small' ? '12px' : '14px';
    
    const fontWeight = element.weight === 'Bolder' ? 'bold' : 
                      element.weight === 'Bold' ? '600' : 'normal';
    
    const color = element.color === 'Dark' ? '#323130' : 
                  element.color === 'Light' ? '#605e5c' : 
                  element.color === 'Accent' ? '#0078d4' : '#323130';

    return (
      <div style={{
        fontSize,
        fontWeight,
        color,
        marginBottom: '8px',
        wordWrap: element.wrap ? 'break-word' : 'normal'
      }}>
        {element.text}
      </div>
    );
  };

  const renderRichTextBlock = (element: any): JSX.Element => {
    return (
      <div style={{ marginBottom: '8px' }}>
        {element.inlines?.map((inline: any, index: number) => (
          <span key={index}>
            {inline.type === 'TextRun' ? renderTextRun(inline) : inline.text || ''}
          </span>
        ))}
      </div>
    );
  };

  const renderTextRun = (element: any): JSX.Element => {
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
  };

  const renderImage = (element: any): JSX.Element => {
    const size = element.size === 'Small' ? '32px' : 
                 element.size === 'Medium' ? '48px' : 
                 element.size === 'Large' ? '64px' : '40px';
    
    const alignment = element.horizontalAlignment === 'Center' ? 'center' : 
                      element.horizontalAlignment === 'Right' ? 'flex-end' : 'flex-start';

    return (
      <div style={{ display: 'flex', justifyContent: alignment, marginBottom: '8px' }}>
        <img 
          src={element.url} 
          alt={element.altText || 'Image'} 
          style={{ 
            width: size, 
            height: size, 
            objectFit: 'contain' 
          }} 
        />
      </div>
    );
  };

  const renderImageSet = (element: any): JSX.Element => {
    return (
      <div style={{ 
        display: 'flex', 
        flexWrap: 'wrap', 
        gap: '8px', 
        marginBottom: '8px' 
      }}>
        {element.images?.map((image: any, index: number) => (
          <div key={index}>
            {renderImage(image)}
          </div>
        ))}
      </div>
    );
  };

  const renderMedia = (element: any): JSX.Element => {
    const poster = element.poster;
    const sources = element.sources || [];
    
    return (
      <div style={{ marginBottom: '16px' }}>
        {sources.length > 0 ? (
          <video 
            controls 
            poster={poster}
            style={{ 
              width: '100%', 
              maxWidth: '100%',
              height: 'auto'
            }}
          >
            {sources.map((source: any, index: number) => (
              <source key={index} src={source.url} type={source.mimeType} />
            ))}
            Your browser does not support the video tag.
          </video>
        ) : (
          <div style={{ 
            padding: '20px', 
            textAlign: 'center', 
            backgroundColor: '#f3f2f1',
            border: '1px solid #edebe9',
            borderRadius: '4px',
            color: '#605e5c'
          }}>
            Media content unavailable
          </div>
        )}
      </div>
    );
  };

  const renderActionSet = (element: any): JSX.Element => {
    return (
      <div style={{ 
        display: 'flex', 
        gap: '8px', 
        marginTop: '12px',
        flexWrap: 'wrap' 
      }}>
        {element.actions?.map((action: any, index: number) => (
          <button
            key={index}
            onClick={() => handleActionClick(action)}
            style={{
              padding: '8px 16px',
              border: action.style === 'positive' ? '1px solid #107c10' : '1px solid #0078d4',
              backgroundColor: action.style === 'positive' ? '#107c10' : '#0078d4',
              color: 'white',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '14px',
              fontWeight: '500'
            }}
          >
            {action.title}
          </button>
        ))}
      </div>
    );
  };

  const renderTableCell = (cell: any): JSX.Element => {
    return (
      <td style={{ 
        padding: '8px', 
        verticalAlign: 'top',
        borderBottom: '1px solid #edebe9' 
      }}>
        {cell.items?.map((item: any, index: number) => (
          <div key={index}>
            {renderElement(item)}
          </div>
        ))}
      </td>
    );
  };

  const renderTable = (element: any): JSX.Element => {
    const rows = element.rows || [];

    return (
      <div style={{ marginBottom: '16px', overflowX: 'auto' }}>
        <table style={{ 
          width: '100%', 
          borderCollapse: 'collapse',
          backgroundColor: 'white'
        }}>
          <tbody>
            {rows.map((row: any, rowIndex: number) => (
              <tr key={rowIndex}>
                {row.cells?.map((cell: any, cellIndex: number) => (
                  <React.Fragment key={cellIndex}>
                    {renderTableCell(cell)}
                  </React.Fragment>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  const renderFactSet = (element: any): JSX.Element => {
    const facts = element.facts || [];
    
    return (
      <div style={{ marginBottom: '16px' }}>
        <table style={{ 
          width: '100%', 
          borderCollapse: 'collapse'
        }}>
          <tbody>
            {facts.map((fact: any, index: number) => (
              <tr key={index}>
                <td style={{ 
                  padding: '4px 8px 4px 0', 
                  fontWeight: '600',
                  color: '#323130',
                  verticalAlign: 'top',
                  width: '30%'
                }}>
                  {fact.title}:
                </td>
                <td style={{ 
                  padding: '4px 0',
                  color: '#605e5c',
                  verticalAlign: 'top'
                }}>
                  {fact.value}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  const renderChart = (element: any): JSX.Element => {
    return (
      <div style={{ height: '300px', marginBottom: '16px' }}>
        <CustomChartRenderer
          title={element.title || ''}
          data={element.data || []}
          type={element.type || 'Chart.Bar'}
          xAxisTitle={element.xAxisTitle || ''}
          yAxisTitle={element.yAxisTitle || ''}
        />
      </div>
    );
  };

  const renderColumnSet = (element: any): JSX.Element => {
    return (
      <div style={{ 
        display: 'flex', 
        gap: '16px', 
        marginBottom: '16px',
        flexWrap: 'wrap' 
      }}>
        {element.columns?.map((column: any, index: number) => (
          <div 
            key={index} 
            style={{ 
              flex: column.width || 1,
              minWidth: '100px' 
            }}
          >
            {column.items?.map((item: any, itemIndex: number) => (
              <div key={itemIndex}>
                {renderElement(item)}
              </div>
            ))}
          </div>
        ))}
      </div>
    );
  };

  const renderContainer = (element: any): JSX.Element => {
    return (
      <div style={{ 
        padding: '12px',
        border: '1px solid #edebe9',
        borderRadius: '4px',
        backgroundColor: '#faf9f8',
        marginBottom: '12px'
      }}>
        {element.items?.map((item: any, index: number) => (
          <div key={index}>
            {renderElement(item)}
          </div>
        ))}
      </div>
    );
  };

  const renderInputText = (element: any): JSX.Element => {
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
          </label>
        )}
        <input
          type="text"
          placeholder={element.placeholder || ''}
          defaultValue={element.value || ''}
          maxLength={element.maxLength}
          style={{
            width: '100%',
            padding: '8px 12px',
            border: '1px solid #8a8886',
            borderRadius: '2px',
            fontSize: '14px',
            fontFamily: 'inherit'
          }}
        />
      </div>
    );
  };

  const renderInputNumber = (element: any): JSX.Element => {
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
          </label>
        )}
        <input
          type="number"
          placeholder={element.placeholder || ''}
          defaultValue={element.value || ''}
          min={element.min}
          max={element.max}
          style={{
            width: '100%',
            padding: '8px 12px',
            border: '1px solid #8a8886',
            borderRadius: '2px',
            fontSize: '14px',
            fontFamily: 'inherit'
          }}
        />
      </div>
    );
  };

  const renderInputDate = (element: any): JSX.Element => {
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
          </label>
        )}
        <input
          type="date"
          defaultValue={element.value || ''}
          min={element.min}
          max={element.max}
          style={{
            width: '100%',
            padding: '8px 12px',
            border: '1px solid #8a8886',
            borderRadius: '2px',
            fontSize: '14px',
            fontFamily: 'inherit'
          }}
        />
      </div>
    );
  };

  const renderInputTime = (element: any): JSX.Element => {
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
          </label>
        )}
        <input
          type="time"
          defaultValue={element.value || ''}
          min={element.min}
          max={element.max}
          style={{
            width: '100%',
            padding: '8px 12px',
            border: '1px solid #8a8886',
            borderRadius: '2px',
            fontSize: '14px',
            fontFamily: 'inherit'
          }}
        />
      </div>
    );
  };

  const renderInputToggle = (element: any): JSX.Element => {
    return (
      <div style={{ marginBottom: '12px' }}>
        <label style={{ 
          display: 'flex', 
          alignItems: 'center',
          cursor: 'pointer',
          fontSize: '14px',
          color: '#323130' 
        }}>
          <input
            type="checkbox"
            defaultChecked={element.value === element.valueOn}
            style={{
              marginRight: '8px',
              width: '16px',
              height: '16px'
            }}
          />
          {element.title || element.label}
        </label>
      </div>
    );
  };

  const renderInputChoiceSet = (element: any): JSX.Element => {
    const isMultiSelect = element.isMultiSelect || false;
    const style = element.style || 'compact';
    const choices = element.choices || [];

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
          </label>
        )}
        
        {style === 'expanded' ? (
          <div>
            {choices.map((choice: any, index: number) => (
              <label key={index} style={{ 
                display: 'flex', 
                alignItems: 'center',
                marginBottom: '4px',
                cursor: 'pointer',
                fontSize: '14px',
                color: '#323130' 
              }}>
                <input
                  type={isMultiSelect ? 'checkbox' : 'radio'}
                  name={element.id || 'choiceset'}
                  value={choice.value}
                  defaultChecked={choice.value === element.value}
                  style={{
                    marginRight: '8px',
                    width: '16px',
                    height: '16px'
                  }}
                />
                {choice.title}
              </label>
            ))}
          </div>
        ) : (
          <select
            multiple={isMultiSelect}
            defaultValue={element.value}
            style={{
              width: '100%',
              padding: '8px 12px',
              border: '1px solid #8a8886',
              borderRadius: '2px',
              fontSize: '14px',
              fontFamily: 'inherit',
              minHeight: isMultiSelect ? '100px' : 'auto'
            }}
          >
            {choices.map((choice: any, index: number) => (
              <option key={index} value={choice.value}>
                {choice.title}
              </option>
            ))}
          </select>
        )}
      </div>
    );
  };

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
      {parsedCard.body?.map((element: AdaptiveCardElement, index: number) => (
        <div key={index}>
          {renderElement(element)}
        </div>
      ))}
    </div>
  );
};

export default AdaptiveCardContentRenderer;
