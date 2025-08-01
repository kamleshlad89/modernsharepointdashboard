import * as React from 'react';
import { useState, useEffect } from 'react';
import * as AdaptiveCards from 'adaptivecards';
import { TooltipHost, TooltipDelay } from '@fluentui/react/lib/Tooltip';
import { AdaptiveCardContentRenderer } from './AdaptiveCardContentRenderer';
import styles from './ModernSharePointDashboard.module.scss';

interface ICardData {
  id: number;
  title: string;
  cardViewJSON: string;
  CardTooltip?: string;
}

interface IChartData {
  title?: string;
  data?: unknown[];
  type?: string;
  xAxisTitle?: string;
  yAxisTitle?: string;
}

interface CardComponentProps {
  cardData: ICardData;
}

export const CardComponent: React.FC<CardComponentProps> = ({ cardData }) => {
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [contentType, setContentType] = useState<'chart' | 'adaptiveCard' | 'adaptiveCardContent' | 'error' | 'empty'>('empty');
  const [parsedChartData, setParsedChartData] = useState<IChartData | null>(null);
  
  useEffect(() => {
    if (!cardData.cardViewJSON) {
      // setContentType('empty');
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
        const findChartElement = (items: unknown[]): unknown => {
          for (const item of items) {
            // Check for Chart element (either type: "Chart" or type starting with "Chart.")
            if (typeof item === 'object' && item !== null && 'type' in item) {
              const typedItem = item as { type: string; [key: string]: unknown };
              if (typedItem.type === 'Chart' || (typedItem.type && typeof typedItem.type === 'string' && typedItem.type.startsWith('Chart.'))) {
                return typedItem;
              }
            }
            // Recursively search in nested structures
            if (typeof item === 'object' && item !== null && 'items' in item) {
              const itemWithItems = item as { items: unknown };
              if (Array.isArray(itemWithItems.items)) {
                const found = findChartElement(itemWithItems.items);
                if (found) return found;
              }
            }
            if (typeof item === 'object' && item !== null && 'body' in item) {
              const itemWithBody = item as { body: unknown };
              if (Array.isArray(itemWithBody.body)) {
                const found = findChartElement(itemWithBody.body);
                if (found) return found;
              }
            }
            if (typeof item === 'object' && item !== null && 'columns' in item) {
              const itemWithColumns = item as { columns: unknown };
              if (Array.isArray(itemWithColumns.columns)) {
                for (const column of itemWithColumns.columns) {
                  if (typeof column === 'object' && column !== null && 'items' in column) {
                    const columnWithItems = column as { items: unknown };
                    if (Array.isArray(columnWithItems.items)) {
                      const found = findChartElement(columnWithItems.items);
                      if (found) return found;
                    }
                  }
                }
              }
            }
          }
          return null;
        };

        const chartElement = findChartElement(json.body);
        if (chartElement) {
          const typedChartElement = chartElement as { 
            type?: string; 
            chartType?: string; 
            title?: string; 
            data?: unknown[]; 
            xAxisTitle?: string; 
            yAxisTitle?: string; 
          };
          setContentType('chart');
          setParsedChartData({
            type: typedChartElement.type || typedChartElement.chartType || 'Chart.Bar', // Use the Chart.Donut type directly
            title: typedChartElement.title || cardData.title || '',
            data: typedChartElement.data || [],
            xAxisTitle: typedChartElement.xAxisTitle || '',
            yAxisTitle: typedChartElement.yAxisTitle || ''
          });
        } else {
          // Check if it contains tables, images, or other complex elements
          const hasComplexElements = json.body.some((item: unknown) => {
            if (typeof item !== 'object' || item === null || !('type' in item)) {
              return false;
            }
            const typedItem = item as { type: string; actions?: unknown[] };
            return typedItem.type === 'Table' || 
              typedItem.type === 'ImageSet' || 
              typedItem.type === 'ColumnSet' || 
              typedItem.type === 'Container' ||
              typedItem.type === 'Media' ||
              typedItem.type === 'FactSet' ||
              typedItem.type === 'RichTextBlock' ||
              typedItem.type === 'Input.Text' ||
              typedItem.type === 'Input.Number' ||
              typedItem.type === 'Input.Date' ||
              typedItem.type === 'Input.Time' ||
              typedItem.type === 'Input.Toggle' ||
              typedItem.type === 'Input.ChoiceSet' ||
              (typedItem.type === 'ActionSet' && Array.isArray(typedItem.actions) && typedItem.actions.length > 0);
          });
          
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

  const handleActionExecute = (action: AdaptiveCards.Action): void => {
    // Handle custom actions here (e.g., button clicks, submit actions)
  };

  // Render the card content based on the content type
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

    // Render chart using unified AdaptiveCardContentRenderer
    if (contentType === 'chart' && parsedChartData) {
      return (
        <div style={{ height: '100%', minHeight: '300px', padding: '16px' }}>
          <AdaptiveCardContentRenderer
            cardJson={cardData.cardViewJSON}
            onActionExecute={handleActionExecute}
            useNativeRenderer={false}
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
          <AdaptiveCardContentRenderer
            cardJson={cardData.cardViewJSON}
            onActionExecute={handleActionExecute}
            useNativeRenderer={true}
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
    <TooltipHost 
      content={cardData.CardTooltip || cardData.title || 'Dashboard Card'}
      delay={TooltipDelay.medium}
      styles={{
        root: { height: '100%', width: '100%' }
      }}
    >
      <div 
        className={styles.cardContainer}
        style={{ 
          cursor: 'pointer',
          transition: 'transform 0.2s ease, box-shadow 0.2s ease'
        }}
        onMouseEnter={(e) => {
          e.currentTarget.style.transform = 'translateY(-2px)';
          e.currentTarget.style.boxShadow = '0 4px 12px rgba(0,0,0,0.15)';
        }}
        onMouseLeave={(e) => {
          e.currentTarget.style.transform = 'translateY(0)';
          e.currentTarget.style.boxShadow = '';
        }}
      >
        {renderCardContent()}
      </div>
    </TooltipHost>
  );
};

export default CardComponent;
