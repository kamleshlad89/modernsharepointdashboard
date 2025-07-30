import * as React from 'react';
import { useState, useEffect } from 'react';
import { AdaptiveCardRenderer } from './AdaptiveCardRenderer';
import { AdaptiveCardContentRenderer } from './AdaptiveCardContentRenderer';
import { CustomChartRenderer } from './CustomChartRenderer';
import styles from './ModernSharePointDashboard.module.scss';

interface ICardData {
  id: number;
  title: string;
  cardViewJSON: string;
}

interface CardComponentProps {
  cardData: ICardData;
}

export const CardComponent: React.FC<CardComponentProps> = ({ cardData }) => {
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [contentType, setContentType] = useState<'chart' | 'adaptiveCard' | 'adaptiveCardContent' | 'error' | 'empty'>('empty');
  const [parsedChartData, setParsedChartData] = useState<any>(null);
  
  useEffect(() => {
    if (!cardData.cardViewJSON) {
      setContentType('empty');
      setIsLoading(false);
      return;
    }

    try {
      // Parse the CardViewJSON from MasterCardList
      const json = JSON.parse(cardData.cardViewJSON);
      
      // Check if it's a direct chart configuration (Chart.* format from MasterCardList)
      if (json.type && typeof json.type === 'string' && json.type.startsWith('Chart.')) {
        setContentType('chart');
        setParsedChartData({
          type: json.type,
          title: json.title || cardData.title || '',
          data: json.data || [],
          xAxisTitle: json.xAxisTitle || '',
          yAxisTitle: json.yAxisTitle || ''
        });
      }
      // Check if it's an Adaptive Card with embedded Chart element
      else if (json.body && Array.isArray(json.body)) {
        // Look for Chart elements within the Adaptive Card body
        const findChartElement = (items: any[]): any => {
          for (const item of items) {
            // Check for Chart element (either type: "Chart" or type starting with "Chart.")
            if (item.type === 'Chart' || (item.type && typeof item.type === 'string' && item.type.startsWith('Chart.'))) {
              return item;
            }
            // Recursively search in nested structures
            if (item.items && Array.isArray(item.items)) {
              const found = findChartElement(item.items);
              if (found) return found;
            }
            if (item.body && Array.isArray(item.body)) {
              const found = findChartElement(item.body);
              if (found) return found;
            }
            if (item.columns && Array.isArray(item.columns)) {
              for (const column of item.columns) {
                if (column.items && Array.isArray(column.items)) {
                  const found = findChartElement(column.items);
                  if (found) return found;
                }
              }
            }
          }
          return null;
        };

        const chartElement = findChartElement(json.body);
        if (chartElement) {
          setContentType('chart');
          setParsedChartData({
            type: chartElement.type || chartElement.chartType || 'Chart.Bar', // Use the Chart.Donut type directly
            title: chartElement.title || cardData.title || '',
            data: chartElement.data || [],
            xAxisTitle: chartElement.xAxisTitle || '',
            yAxisTitle: chartElement.yAxisTitle || ''
          });
        } else {
          // Check if it contains tables, images, or other complex elements
          const hasComplexElements = json.body.some((item: any) => 
            item.type === 'Table' || 
            item.type === 'ImageSet' || 
            item.type === 'ColumnSet' || 
            item.type === 'Container' ||
            item.type === 'Media' ||
            item.type === 'FactSet' ||
            item.type === 'RichTextBlock' ||
            item.type === 'Input.Text' ||
            item.type === 'Input.Number' ||
            item.type === 'Input.Date' ||
            item.type === 'Input.Time' ||
            item.type === 'Input.Toggle' ||
            item.type === 'Input.ChoiceSet' ||
            (item.type === 'ActionSet' && item.actions?.length > 0)
          );
          
          if (hasComplexElements || json.type === 'AdaptiveCard' || json.version) {
            // Use the comprehensive renderer for complex Adaptive Cards
            setContentType('adaptiveCardContent');
          } else {
            // Use standard Adaptive Card renderer for simple cards
            setContentType('adaptiveCard');
          }
        }
      }
      // Check if it's a standard Adaptive Card
      else if (json.type === 'AdaptiveCard' || json.version) {
        setContentType('adaptiveCard');
      }
      else {
        setContentType('error');
      }
      
    } catch (error) {
      console.error('Error parsing CardViewJSON from MasterCardList:', error);
      setContentType('error');
    }
    
    setIsLoading(false);
  }, [cardData.cardViewJSON, cardData.title]);

  const handleActionExecute = (action: any): void => {
    // Handle custom actions here (e.g., button clicks, submit actions)
  };

  const renderCardContent = (): JSX.Element => {
    // Show loading state
    if (isLoading) {
      return (
        <div style={{ padding: '20px', textAlign: 'center', color: '#605e5c' }}>
          Loading...
        </div>
      );
    }

    // Show empty state
    if (contentType === 'empty') {
      return (
        <div style={{ padding: '20px', textAlign: 'center', color: '#605e5c' }}>
          No data available
        </div>
      );
    }

    // Show error state
    if (contentType === 'error') {
      return (
        <div style={{ padding: '20px', textAlign: 'center', color: '#a4262c' }}>
          Error loading card content
        </div>
      );
    }

    // Render chart using CustomChartRenderer with parsed data from CardViewJSON
    if (contentType === 'chart' && parsedChartData) {
      return (
        <div style={{ height: '100%', minHeight: '300px', padding: '16px' }}>
          <CustomChartRenderer
            title={parsedChartData.title}
            data={parsedChartData.data}
            type={parsedChartData.type}
            xAxisTitle={parsedChartData.xAxisTitle}
            yAxisTitle={parsedChartData.yAxisTitle}
          />
        </div>
      );
    }

    // Render comprehensive Adaptive Card content (tables, images, etc.)
    if (contentType === 'adaptiveCardContent') {
      return (
        <div style={{ height: '100%', width: '100%' }}>
          <AdaptiveCardContentRenderer
            cardJson={cardData.cardViewJSON}
            onActionExecute={handleActionExecute}
          />
        </div>
      );
    }

    // Render standard Adaptive Card
    if (contentType === 'adaptiveCard') {
      return (
        <div style={{ height: '100%', width: '100%' }}>
          <AdaptiveCardRenderer
            cardJson={cardData.cardViewJSON}
            onActionExecute={handleActionExecute}
          />
        </div>
      );
    }

    // Fallback
    return (
      <div style={{ padding: '20px', textAlign: 'center', color: '#a4262c' }}>
        Unsupported content type
      </div>
    );
  };

  return (
    <div className={styles.cardContainer}>
      {renderCardContent()}
      {/* <div style={{
        height: '100%',
        border: '1px solid #edebe9',
        borderRadius: '2px',
        backgroundColor: '#ffffff',
        overflow: 'auto',
        boxShadow: '0 1.6px 3.6px 0 rgba(0,0,0,0.132), 0 0.3px 0.9px 0 rgba(0,0,0,0.108)'
      }}>
        
      </div> */}
    </div>
  );
};

export default CardComponent;
