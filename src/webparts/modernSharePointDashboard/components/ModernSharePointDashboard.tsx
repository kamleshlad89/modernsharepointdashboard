import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './ModernSharePointDashboard.module.scss';
import type { IModernSharePointDashboardProps } from './IModernSharePointDashboardProps';
import { CommandButton, PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Panel } from '@fluentui/react/lib/Panel';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { Icon } from '@fluentui/react/lib/Icon';
import { useBoolean } from '@fluentui/react-hooks';
import { DndProvider } from 'react-dnd';
import { HTML5Backend } from 'react-dnd-html5-backend';
import { useDrag, useDrop } from 'react-dnd';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { CardComponent } from './CardComponent';

interface ICard {
  id: number;
  title: string;
  order: number;
  visible: boolean;
  selected: boolean;
  fixed: boolean;
  defaultOrder: number;
  cardViewJSON: string;
}

interface ISharePointListItem {
  Id: number;
  Title: string;
  Fixed: string;
  DefaultOrder: number;
  CardViewJSON: string;
  [key: string]: any;
}

// Draggable Card Component for Customize Panel
const DraggableCard: React.FC<{
  card: ICard;
  index: number;
  moveCard: (dragIndex: number, hoverIndex: number) => void;
  onSelectionChange: (cardId: number, selected: boolean) => void;
  maxSelectionReached: boolean;
}> = ({ card, index, moveCard, onSelectionChange, maxSelectionReached }) => {
  const ref = React.useRef<HTMLDivElement>(null);

  const [{ isDragging }, drag] = useDrag({
    type: 'CARD',
    item: () => ({ type: 'CARD', id: card.id, index }),
    collect: (monitor) => ({
      isDragging: monitor.isDragging(),
    }),
    canDrag: () => !card.fixed, // Prevent dragging fixed cards
  });

  const [, drop] = useDrop({
    accept: 'CARD',
    hover: (item: { type: string; id: number; index: number }, monitor) => {
      if (!ref.current) {
        return;
      }
      
      const dragIndex = item.index;
      const hoverIndex = index;
      
      // Don't replace items with themselves
      if (dragIndex === hoverIndex) {
        return;
      }
      
      // Determine rectangle on screen
      const hoverBoundingRect = ref.current?.getBoundingClientRect();
      
      // Get vertical middle
      const hoverMiddleY = (hoverBoundingRect.bottom - hoverBoundingRect.top) / 2;
      
      // Determine mouse position
      const clientOffset = monitor.getClientOffset();
      
      // Get pixels to the top
      const hoverClientY = clientOffset!.y - hoverBoundingRect.top;
      
      // Only perform the move when the mouse has crossed half of the items height
      // When dragging downwards, only move when the cursor is below 50%
      // When dragging upwards, only move when the cursor is above 50%
      
      // Dragging downwards
      if (dragIndex < hoverIndex && hoverClientY < hoverMiddleY) {
        return;
      }
      
      // Dragging upwards
      if (dragIndex > hoverIndex && hoverClientY > hoverMiddleY) {
        return;
      }
      
      // Time to actually perform the action
      moveCard(dragIndex, hoverIndex);
      
      // Note: we're mutating the monitor item here!
      // Generally it's better to avoid mutations,
      // but it's good here for the sake of performance
      // to avoid expensive index searches.
      item.index = hoverIndex;
    },
  });

  const handleCheckboxChange = (checked: boolean): void => {
    onSelectionChange(card.id, checked);
  };

  // Connect drag and drop to the same element
  drag(drop(ref));

  return (
    <div
      ref={ref}
      style={{
        opacity: isDragging ? 0.5 : 1,
        cursor: card.fixed ? 'not-allowed' : 'move',
        padding: '12px',
        margin: '8px 0',
        backgroundColor: card.fixed ? '#faf9f8' : '#f3f2f1',
        border: card.fixed ? '1px solid #d2d0ce' : '1px solid #edebe9',
        borderRadius: '2px',
        display: 'flex',
        alignItems: 'center',
        gap: '12px',
        position: 'relative'
      }}
    >
      {card.fixed && (
        <Icon
          iconName="Lock"
          style={{
            position: 'absolute',
            top: '8px',
            right: '8px',
            color: '#a19f9d',
            fontSize: '12px'
          }}
        />
      )}
      <Checkbox
        checked={card.selected}
        onChange={(ev, checked) => handleCheckboxChange(checked || false)}
        disabled={card.fixed || (!card.selected && maxSelectionReached)}
      />
      <span style={{ color: card.fixed ? '#a19f9d' : '#323130', paddingRight: card.fixed ? '20px' : '0' }}>
        {card.title} {card.fixed}
      </span>
    </div>
  );
};

const ModernSharePointDashboard: React.FC<IModernSharePointDashboardProps> = (props) => {
  const [cards, setCards] = useState<ICard[]>([]);
  const [isCustomizePanelOpen, { setTrue: openCustomizePanel, setFalse: dismissCustomizePanel }] = useBoolean(false);

  const loadCardsFromSharePoint = async (): Promise<void> => {
    try {
      console.log('Loading cards from SharePoint MasterCardList...');
      
      // Make actual SharePoint REST API call
      const response: SPHttpClientResponse = await props.context.spHttpClient.get(
        `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('MasterCardList')/items?$select=Id,Title,Fixed,DefaultOrder,CardViewJSON&$orderby=DefaultOrder`,
        SPHttpClient.configurations.v1
      );
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      
      const data = await response.json();
      const items: ISharePointListItem[] = data.value;
      
      // Transform SharePoint data to our card format
      const cardsData: ICard[] = items.map((item: ISharePointListItem, index: number) => {
        const isFixed = item.Fixed === 'Yes';
        return {
          id: item.Id,
          title: item.Title,
          order: index + 1,
          visible: index < 4, // First 4 cards visible by default
          selected: isFixed ? true : index < 4, // Fixed cards are always selected, others follow default logic
          fixed: isFixed,
          defaultOrder: item.DefaultOrder || index + 1,
          cardViewJSON: item.CardViewJSON || ''
        };
      });
      
      setCards(cardsData);
    } catch (error) {
      console.error('Failed to load cards from SharePoint:', error);
      // Fallback to empty array if SharePoint call fails
      setCards([]);
    }
  };

  useEffect(() => {
    void loadCardsFromSharePoint();
  }, []);

  const moveCard = (dragIndex: number, hoverIndex: number): void => {
    const newCards = [...cards];
    const dragCard = newCards[dragIndex];
    
    // Remove the dragged card from its original position
    newCards.splice(dragIndex, 1);
    // Insert it at the new position
    newCards.splice(hoverIndex, 0, dragCard);
    
    // Update the order property for all cards
    const updatedCards = newCards.map((card, index) => ({
      ...card,
      order: index + 1
    }));
    
    setCards(updatedCards);
  };

  const handleSelectionChange = (cardId: number, selected: boolean): void => {
    const newCards = cards.map(card => {
      if (card.id === cardId && !card.fixed) { // Only allow changes for non-fixed cards
        return { ...card, selected };
      }
      return card;
    });
    setCards(newCards);
  };

  const selectedCount = cards.filter(card => card.selected).length;
  const fixedCount = cards.filter(card => card.fixed).length;
  const selectableCount = cards.filter(card => !card.fixed && card.selected).length;
  const maxSelectableCards = 4 - fixedCount; // Reduce max selection by number of fixed cards
  const maxSelectionReached = selectableCount >= maxSelectableCards;

  const handleSave = async (): Promise<void> => {
    // Update visibility based on selection
    const updatedCards = cards.map(card => ({
      ...card,
      visible: card.selected
    }));
    setCards(updatedCards);

    try {
      // Get selected cards in order
      const selectedCards = cards
        .filter(card => card.selected)
        .sort((a, b) => a.order - b.order)
        .map((card, index) => ({
          id: card.id,
          title: card.title,
          order: index + 1,
          cardViewJSON: card.cardViewJSON
        }));

      // Create JSON schema with user settings
      const userSettingsJSON = {
        timestamp: new Date().toISOString(),
        selectedCards: selectedCards,
        totalSelected: selectedCards.length,
        cardDetails: selectedCards.map(card => ({
          cardId: card.id,
          cardTitle: card.title,
          displayOrder: card.order,
          cardViewJSON: JSON.parse(card.cardViewJSON || '{}')
        }))
      };

      // Get current user information
      const currentUser = props.context.pageContext.user;
      const userPrincipalName = currentUser.loginName;
      const userName = currentUser.displayName;

      // Check if UserSettingsList exists and if user settings already exist
      let existingItems: any[] = [];
      let listExists = true;
      
      try {
        const existingSettingsResponse = await props.context.spHttpClient.get(
          `${props.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('UserSettingsList')/items?$filter=UserID eq '${userPrincipalName}'&$top=1`,
          SPHttpClient.configurations.v1
        );

        if (existingSettingsResponse.ok) {
          const existingData = await existingSettingsResponse.json();
          existingItems = existingData.value || [];
        } else {
          console.log('UserSettingsList might not exist or is empty');
          listExists = false;
        }
      } catch (listError) {
        console.log('Error accessing UserSettingsList, it might not exist:', listError);
        listExists = false;
      }

      // Prepare the item data
      const itemData = {
        '__metadata': { 'type': 'SP.Data.UserSettingsListListItem' },
        'Title': userName,
        'UserID': userPrincipalName,
        'PersonalisedCards': JSON.stringify(userSettingsJSON)
      };

      if (listExists && existingItems.length > 0) {
        // Update existing user settings
        const existingItem = existingItems[0];
        await props.context.spHttpClient.post(
          `${props.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('UserSettingsList')/items(${existingItem.Id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': '',
              'IF-MATCH': '*',
              'X-HTTP-Method': 'MERGE'
            },
            body: JSON.stringify(itemData)
          }
        );
        console.log('User settings updated successfully for:', userName);
      } else {
        // Create new user settings (first time or list is empty)
        try {
          const response = await props.context.spHttpClient.post(
            `${props.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('UserSettingsList')/items`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': ''
              },
              body: JSON.stringify(itemData)
            }
          );
          
          if (response.ok) {
            console.log('User settings created successfully for first-time user:', userName);
            console.log('Data saved:', itemData);
          } else {
            const errorText = await response.text();
            console.error('Failed to create user settings. Response:', response.status, errorText);
            throw new Error(`HTTP ${response.status}: ${errorText}`);
          }
        } catch (createError) {
          console.error('Error creating user settings. List might not exist:', createError);
          
          // Try to provide helpful error information
          if (createError.message?.includes('404')) {
            console.error('UserSettingsList does not exist. Please create the list with the following columns:');
            console.error('- Title (Single line of text) - for user display name');
            console.error('- UserID (Single line of text) - for user principal name');
            console.error('- PersonalisedCards (Multiple lines of text) - for JSON data');
          }
          
          // Don't throw the error, just log it so the UI doesn't break
          alert('Settings could not be saved. Please contact your administrator to set up the UserSettingsList with columns: Title, UserID, and PersonalisedCards.');
        }
      }

      console.log('User settings JSON prepared:', userSettingsJSON);
      console.log('User information:', { userName, userPrincipalName });
    } catch (error) {
      console.error('Failed to save user settings to SharePoint:', error);
      
      // Provide user-friendly error message
      alert('There was an error saving your settings. Please try again or contact your administrator.');
    }

    dismissCustomizePanel();
  };

  return (
    <DndProvider backend={HTML5Backend}>
      <section className={styles.modernSharePointDashboard}>
        <div className={styles.headerBar}>
          <CommandButton
            iconProps={{ iconName: 'Customize' }}
            text="Customize"
            className="customizeButton"
            onClick={openCustomizePanel}
          />
        </div>
        
        <div className={styles.dashboardGrid}>
          {cards.filter(card => card.visible).slice(0, 4).map((card) => (
            <CardComponent
              key={card.id}
              cardData={{
                id: card.id,
                title: card.title,
                cardViewJSON: card.cardViewJSON
              }}
            />
          ))}
        </div>

        <Panel
          isOpen={isCustomizePanelOpen}
          onDismiss={dismissCustomizePanel}
          headerText="Customize Dashboard"
          closeButtonAriaLabel="Close"
        >
          <div className={styles.customizePanel}>
            <div style={{ marginBottom: '16px', fontSize: '14px', color: '#605e5c' }}>
              {fixedCount > 0 && (
                <div style={{ fontSize: '12px', color: '#a19f9d', marginTop: '4px' }}>
                  ({fixedCount} fixed card{fixedCount !== 1 ? 's' : ''}, {selectableCount}/{maxSelectableCards} selectable)
                </div>
              )}
            </div>
            
            {cards.map((card, index) => (
              <DraggableCard
                key={card.id}
                card={card}
                index={index}
                moveCard={moveCard}
                onSelectionChange={handleSelectionChange}
                maxSelectionReached={maxSelectionReached}
              />
            ))}
            
            <div style={{ marginTop: '20px', display: 'flex', gap: '12px' }}>
              <PrimaryButton
                text="Save"
                onClick={handleSave}
                disabled={selectedCount !== 4}
              />
              <DefaultButton
                text="Cancel"
                onClick={dismissCustomizePanel}
              />
            </div>
          </div>
        </Panel>
      </section>
    </DndProvider>
  );
};

export default ModernSharePointDashboard;
