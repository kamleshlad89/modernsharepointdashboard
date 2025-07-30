import * as React from 'react';
import { useEffect, useRef } from 'react';
import * as AdaptiveCards from 'adaptivecards';
import { Template, IEvaluationContext } from 'adaptivecards-templating';


interface AdaptiveCardRendererProps {
  cardJson: string;
  data?: any;
  onActionExecute?: (action: AdaptiveCards.Action) => void;
}

export const AdaptiveCardRenderer: React.FC<AdaptiveCardRendererProps> = ({ 
  cardJson, 
  data, 
  onActionExecute 
}) => {
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

      let cardPayload: any;

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

export default AdaptiveCardRenderer;
