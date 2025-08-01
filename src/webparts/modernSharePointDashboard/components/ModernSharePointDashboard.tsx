import * as React from 'react';
import { useState, useEffect, useRef, useCallback } from 'react';
import styles from './ModernSharePointDashboard.module.scss';
import type { IModernSharePointDashboardProps } from './IModernSharePointDashboardProps';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Panel } from '@fluentui/react/lib/Panel';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { Icon } from '@fluentui/react/lib/Icon';
import { useBoolean } from '@fluentui/react-hooks';
import { DndContext, DragEndEvent, useDraggable, useDroppable, DragOverlay, DragStartEvent } from '@dnd-kit/core';
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
  CardTooltip?: string;
}

interface ISelectedCard {
  id: number;
  title: string;
  order: number;
}

interface ISharePointCardItem {
  Id: number;
  Title: string;
  Fixed: boolean; // SharePoint Yes/No (checkbox) column returns boolean
  DefaultOrder: number;
  CardViewJSON: string;
  CardTooltip?: string;
}

// Draggable Card Component for Customize Panel using @dnd-kit/core
const DraggableCard: React.FC<{
  card: ICard;
  index: number;
  onSelectionChange: (cardId: number, selected: boolean) => void;
  maxSelectionReached: boolean;
  requiredUserSelections: number;
  isDragOverlay?: boolean;
}> = ({ card, index, onSelectionChange, maxSelectionReached, requiredUserSelections, isDragOverlay = false }) => {
  
  // Only allow dragging for selected cards (not fixed and not available)
  const canDrag = card.selected && !card.fixed;
  
  const {
    attributes,
    listeners,
    setNodeRef: setDraggableRef,
    transform,
    isDragging,
  } = useDraggable({
    id: `card-${card.id}`,
    data: {
      card,
      index,
    },
    disabled: !canDrag,
  });

  const {
    setNodeRef: setDroppableRef,
    isOver,
  } = useDroppable({
    id: `droppable-${card.id}`,
    data: {
      card,
      index,
    },
    disabled: !card.selected, // Only allow dropping on selected cards
  });

  const handleCheckboxChange = (ev: React.FormEvent<HTMLElement>, checked?: boolean): void => {
    onSelectionChange(card.id, checked || false);
  };

  const style = transform ? {
    transform: `translate3d(${transform.x}px, ${transform.y}px, 0)`,
  } : undefined;

  return (
    <div
      ref={(node) => {
        setDraggableRef(node);
        setDroppableRef(node);
      }}
      style={{
        ...style,
        opacity: isDragOverlay ? 0.8 : isDragging ? 0.4 : 1,
        cursor: canDrag ? 'move' : 'default',
        userSelect: 'none',
        transition: isDragging || isDragOverlay ? 'none' : 'all 0.2s ease',
        transform: isDragging ? 'rotate(2deg) scale(1.05)' : isOver ? 'translateY(-2px)' : style?.transform || 'none',
        boxShadow: isDragging || isDragOverlay ? '0 8px 16px rgba(0,0,0,0.3)' : 
                   isOver ? '0 4px 12px rgba(0,120,212,0.4), inset 0 2px 0 #0078d4' : 
                   '0 1px 3px rgba(0,0,0,0.1)',
        borderColor: isOver ? '#0078d4' : undefined,
        backgroundColor: isOver ? 'rgba(0,120,212,0.05)' : undefined,
      }}
      className={`${styles.draggableCard}${card.selected ? ' selected' : ''}${card.fixed ? ' fixed' : ''}${isDragging ? ' isDragging' : ''}${isOver ? ' isOver' : ''}`}
      title={
        card.fixed 
          ? "This card is fixed and cannot be moved" 
          : !card.selected 
          ? "This card is not selected and cannot be moved"
          : "Click and drag to reorder"
      }
    >
      {card.fixed && <Icon iconName="Lock" className="lockIcon" />}
      {card.selected && <div className="selectionStripe" />}
      {canDrag && (
        <div 
          className="dragHandle" 
          title="Drag handle - click and drag to reorder"
          style={{ 
            cursor: isDragging ? 'grabbing' : 'grab',
            opacity: isDragging ? 0.7 : 1,
            transform: isDragging ? 'scale(1.1)' : 'scale(1)',
            transition: 'all 0.2s ease'
          }}
          {...listeners}
          {...attributes}
        >
          <Icon 
            iconName="GripperDotsVertical" 
            className="gripperIcon" 
            styles={{
              root: {
                cursor: isDragging ? 'grabbing' : 'grab',
                fontSize: '16px',
                color: isDragging ? '#0078d4' : '#605e5c',
                transition: 'color 0.2s ease'
              }
            }}
            style={{ pointerEvents: 'none' }}
          />
        </div>
      )}
      <Checkbox
        checked={card.selected}
        onChange={handleCheckboxChange}
        disabled={card.fixed || (!card.selected && maxSelectionReached)}
        title={
          card.fixed
            ? "This card is fixed and cannot be deselected"
            : !card.selected && maxSelectionReached
            ? `Maximum ${requiredUserSelections} cards can be selected`
            : undefined
        }
      />
      <span className="cardTitle">{card.title}</span>
    </div>
  );
};

const ModernSharePointDashboard: React.FC<IModernSharePointDashboardProps> = (props) => {
  const [cards, setCards] = useState<ICard[]>([]);
  const [originalCards, setOriginalCards] = useState<ICard[]>([]);
  const [isCustomizePanelOpen, { setTrue: openCustomizePanel, setFalse: dismissCustomizePanel }] = useBoolean(false);
  const [searchText, setSearchText] = useState('');
  const [activeCard, setActiveCard] = useState<ICard | null>(null);

  // Handle drag start for @dnd-kit
  const handleDragStart = (event: DragStartEvent) => {
    const cardData = event.active.data.current;
    if (cardData?.card) {
      setActiveCard(cardData.card);
    }
  };

  // Handle drag end for @dnd-kit
  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event;
    setActiveCard(null);

    if (!over || active.id === over.id) {
      return;
    }

    const activeData = active.data.current;
    const overData = over.data.current;

    if (!activeData?.card || !overData?.card) {
      return;
    }

    const activeCard = activeData.card as ICard;
    const overCard = overData.card as ICard;

    // Only allow reordering among selected cards
    if (!activeCard.selected || !overCard.selected) {
      return;
    }

    setCards((prevCards) => {
      const activeIndex = prevCards.findIndex(card => card.id === activeCard.id);
      const overIndex = prevCards.findIndex(card => card.id === overCard.id);

      if (activeIndex === -1 || overIndex === -1) {
        return prevCards;
      }

      const newCards = [...prevCards];
      const [movedCard] = newCards.splice(activeIndex, 1);
      newCards.splice(overIndex, 0, movedCard);

      // Update the order based on new positions
      const updatedCards = newCards.map((card, index) => ({
        ...card,
        order: index + 1
      }));

      console.log(`🔄 Moved card "${activeCard.title}" from position ${activeIndex} to ${overIndex}`);
      return updatedCards;
    });
  };

  const loadUserSettings = useCallback(async (initialCards: ICard[]): Promise<void> => {
    try {
      const currentUser = props.context.pageContext.user;
      const userPrincipalName = currentUser.loginName;

      console.log('Loading user settings for:', userPrincipalName);
      console.log('Initial cards from SharePoint:', initialCards.map(c => ({ id: c.id, title: c.title, fixed: c.fixed })));

      const response = await props.context.spHttpClient.get(
        `${props.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('UserSettingsList')/items?$filter=UserID eq '${userPrincipalName}'&$top=1`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        if (data.value && data.value.length > 0) {
          const userSettings = JSON.parse(data.value[0].PersonalisedCards);
          console.log('Found user settings:', userSettings);
          
          const userCardMap = new Map<number, ISelectedCard>(
            userSettings.selectedCards.map((uc: ISelectedCard) => [uc.id, uc])
          );

          const updatedCards = initialCards.map(card => {
            if (card.fixed) {
              return { ...card, visible: true, selected: true };
            }
            const userSetting = userCardMap.get(card.id);
            return {
              ...card,
              visible: !!userSetting,
              selected: !!userSetting,
              order: userSetting ? userSetting.order : card.order,
            };
          });

          console.log('Final cards after applying user settings:', updatedCards.map(c => ({ 
            id: c.id, 
            title: c.title, 
            fixed: c.fixed, 
            selected: c.selected, 
            visible: c.visible 
          })));

          setCards(updatedCards);
          setOriginalCards([...updatedCards]);
          return;
        }
      }
    } catch (error) {
      console.log('No user settings found or error loading them:', error);
    }

    // Default behavior when no user settings found
    const defaultVisibleCards = initialCards.map((card, index) => ({
      ...card,
      visible: card.fixed || index < 4,
      selected: card.fixed || index < 4,
    }));
    
    console.log('Using default card visibility:', defaultVisibleCards.map(c => ({ 
      id: c.id, 
      title: c.title, 
      fixed: c.fixed, 
      selected: c.selected, 
      visible: c.visible 
    })));
    
    setCards(defaultVisibleCards);
    setOriginalCards([...defaultVisibleCards]);
  }, [props.context]);

  const loadCardsFromSharePoint = useCallback(async (): Promise<void> => {
    try {
      const response: SPHttpClientResponse = await props.context.spHttpClient.get(
        `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('MasterCardList')/items?$select=Id,Title,Fixed,DefaultOrder,CardViewJSON,CardTooltip&$orderby=DefaultOrder`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      const items: ISharePointCardItem[] = data.value;

      const cardsData: ICard[] = items.map((item, index) => {
        // Handle Fixed field - SharePoint Yes/No column returns boolean directly
        const isFixed = item.Fixed === true;
        
        console.log(`Card "${item.Title}": Fixed field value = ${item.Fixed} (type: ${typeof item.Fixed}), isFixed = ${isFixed}`);
        
        return {
          id: item.Id,
          title: item.Title,
          order: item.DefaultOrder || index + 1,
          visible: isFixed,
          selected: isFixed,
          fixed: isFixed,
          defaultOrder: item.DefaultOrder || index + 1,
          cardViewJSON: item.CardViewJSON || '',
          CardTooltip: item.CardTooltip || ''
        };
      });

      await loadUserSettings(cardsData);
    } catch (error) {
      console.error('Failed to load cards from SharePoint:', error);
      setCards([]);
    }
  }, [props.context]);

  useEffect(() => {
    loadCardsFromSharePoint().catch(console.error);
  }, [loadCardsFromSharePoint]);

  const handleSelectionChange = (cardId: number, selected: boolean): void => {
    setCards(currentCards => {
      const newCards = currentCards.map(card => {
        if (card.id === cardId && !card.fixed) {
          return { ...card, selected };
        }
        return card;
      });

      const fixedCount = newCards.filter(c => c.fixed).length;
      const requiredSelections = 4 - fixedCount;
      const newSelectableCount = newCards.filter(c => !c.fixed && c.selected).length;

      if (newSelectableCount > requiredSelections) {
        return currentCards; // Abort change
      }
      return newCards;
    });
  };

  const handleOpenCustomizePanel = (): void => {
    setOriginalCards([...cards]);
    openCustomizePanel();
  };

  const handleCancel = (): void => {
    setCards([...originalCards]);
    dismissCustomizePanel();
  };

  const handleSave = async (): Promise<void> => {
    const updatedCards = cards.map(card => ({
      ...card,
      visible: card.selected,
    }));

    const visibleCount = updatedCards.filter(c => c.visible).length;
    if (visibleCount > 4) {
      alert(`You can only select up to 4 cards.`);
      return;
    }

    setCards(updatedCards);

    try {
      const selectedCardsForSaving = updatedCards
        .filter(c => c.selected)
        .map((c, index) => ({
          id: c.id,
          title: c.title,
          order: index + 1, // Assign new order based on final position
        }));

      const userSettingsJSON = {
        timestamp: new Date().toISOString(),
        selectedCards: selectedCardsForSaving,
      };

      const currentUser = props.context.pageContext.user;
      const userPrincipalName = currentUser.loginName;
      const userName = currentUser.displayName;

      const listName = 'UserSettingsList';
      const siteUrl = props.context.pageContext.site.absoluteUrl;
      const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

      const getResponse = await props.context.spHttpClient.get(
        `${apiUrl}?$filter=UserID eq '${userPrincipalName}'&$top=1`,
        SPHttpClient.configurations.v1
      );

      const itemData = {
        Title: userName,
        UserID: userPrincipalName,
        PersonalisedCards: JSON.stringify(userSettingsJSON),
      };

      if (getResponse.ok) {
        const data = await getResponse.json();
        if (data.value && data.value.length > 0) {
          const existingItemId = data.value[0].Id;
          await props.context.spHttpClient.post(
            `${apiUrl}(${existingItemId})`,
            SPHttpClient.configurations.v1,
            {
              headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
              body: JSON.stringify(itemData),
            }
          );
        } else {
          await props.context.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, {
            body: JSON.stringify(itemData),
          });
        }
      }
      console.log('User settings saved successfully.');
    } catch (error) {
      console.error('Failed to save user settings:', error);
      alert('There was an error saving your settings.');
    }

    dismissCustomizePanel();
  };

  const fixedCount = cards.filter(c => c.fixed).length;
  const selectableCount = cards.filter(c => !c.fixed && c.selected).length;
  const requiredUserSelections = 4 - fixedCount;
  const hasCorrectSelections = selectableCount === requiredUserSelections;
  const maxSelectionReached = selectableCount >= requiredUserSelections;

  const renderDashboardGrid = (): JSX.Element[] => {
    const visibleCards = cards.filter(c => c.visible);
    const sortedCards = visibleCards.sort((a, b) => a.order - b.order);
    const gridSlots: (ICard | null)[] = Array(4).fill(null);

    const fixedCards = sortedCards.filter(c => c.fixed);
    const nonFixedCards = sortedCards.filter(c => !c.fixed);

    fixedCards.forEach(card => {
      if (card.defaultOrder > 0 && card.defaultOrder <= 4) {
        gridSlots[card.defaultOrder - 1] = card;
      }
    });

    let nonFixedIndex = 0;
    for (let i = 0; i < 4 && nonFixedIndex < nonFixedCards.length; i++) {
      if (gridSlots[i] === null) {
        gridSlots[i] = nonFixedCards[nonFixedIndex++];
      }
    }

    return gridSlots
      .filter(card => card !== null)
      .map(card => (
        <CardComponent
          key={card!.id}
          cardData={{
            id: card!.id,
            title: card!.title,
            cardViewJSON: card!.cardViewJSON,
            CardTooltip: card!.CardTooltip
          }}
        />
      ));
  };

  return (
    <DndContext onDragStart={handleDragStart} onDragEnd={handleDragEnd}>
      <section className={styles.modernSharePointDashboard}>
        <div className={styles.headerBar}>
          <PrimaryButton
            iconProps={{ iconName: 'Settings' }}
            text="Customize Dashboard"
            className="customizeButton"
            onClick={handleOpenCustomizePanel}
          />
        </div>
        <div className={styles.dashboardGrid}>{renderDashboardGrid()}</div>
        <Panel
          isOpen={isCustomizePanelOpen}
          onDismiss={handleCancel}
          headerText="Customize Dashboard"
          closeButtonAriaLabel="Close"
        >
          <div className={styles.customizePanel}>
            <div style={{ marginBottom: '16px', fontSize: '14px', color: '#605e5c' }}>
              💡 Tip: Click and drag any card to reorder. Fixed cards (with lock icons) cannot be moved. Cards with blue backgrounds are selected.
            </div>
            
            {/* Search Box */}
            <div style={{ marginBottom: '16px', position: 'relative' }}>
              <input
                type="text"
                placeholder="🔍 Search cards by name..."
                value={searchText}
                onChange={(e) => setSearchText(e.target.value)}
                style={{
                  width: '100%',
                  padding: '8px 12px',
                  paddingRight: searchText ? '40px' : '12px',
                  border: '1px solid #8a8886',
                  borderRadius: '4px',
                  fontSize: '14px',
                  fontFamily: '"Segoe UI", system-ui, sans-serif',
                  boxSizing: 'border-box'
                }}
              />
              {searchText && (
                <button
                  onClick={() => setSearchText('')}
                  style={{
                    position: 'absolute',
                    right: '8px',
                    top: '50%',
                    transform: 'translateY(-50%)',
                    background: 'none',
                    border: 'none',
                    cursor: 'pointer',
                    color: '#605e5c',
                    fontSize: '16px',
                    padding: '4px'
                  }}
                  title="Clear search"
                >
                  ✕
                </button>
              )}
            </div>
            
            {/* Selected Cards Section */}
            {(() => {
              // Filter cards by search text first
              const filteredCards = cards.filter(card => 
                card.title.toLowerCase().includes(searchText.toLowerCase())
              );

              // Separate filtered cards into categories
              const userSelectedCards = filteredCards
                .filter(card => card.selected && !card.fixed)
                .sort((a, b) => a.order - b.order);

              const fixedCards = filteredCards
                .filter(card => card.selected && card.fixed)
                .sort((a, b) => a.defaultOrder - b.defaultOrder);

              const unselectedCards = filteredCards
                .filter(card => !card.selected)
                .sort((a, b) => a.defaultOrder - b.defaultOrder);

              // Combine arrays: user-selected first, then fixed, then unselected
              const sortedCards = [...userSelectedCards, ...fixedCards, ...unselectedCards];

              // Check if we have any results
              const hasResults = sortedCards.length > 0;
              const isSearching = searchText.trim() !== '';

              return (
                <>
                  {isSearching && !hasResults && (
                    <div style={{ 
                      padding: '16px', 
                      textAlign: 'center', 
                      color: '#605e5c',
                      backgroundColor: '#f3f2f1',
                      border: '1px solid #edebe9',
                      borderRadius: '4px',
                      marginBottom: '12px'
                    }}>
                      No cards found matching &quot;{searchText}&quot;
                    </div>
                  )}

                  {userSelectedCards.length > 0 && (
                    <div style={{ marginBottom: '12px' }}>
                      <div style={{ 
                        fontSize: '13px', 
                        fontWeight: '600', 
                        color: '#0078d4', 
                        marginBottom: '8px',
                        borderBottom: '1px solid #edebe9',
                        paddingBottom: '4px'
                      }}>
                        User Selected Cards ({userSelectedCards.length})
                      </div>
                      {userSelectedCards.map((card, index) => (
                        <DraggableCard
                          key={card.id}
                          card={card}
                          index={sortedCards.findIndex(c => c.id === card.id)}
                          onSelectionChange={handleSelectionChange}
                          maxSelectionReached={maxSelectionReached}
                          requiredUserSelections={requiredUserSelections}
                        />
                      ))}
                    </div>
                  )}

                  {fixedCards.length > 0 && (
                    <div style={{ marginBottom: '12px' }}>
                      <div style={{ 
                        fontSize: '13px', 
                        fontWeight: '600', 
                        color: '#107c10', 
                        marginBottom: '8px',
                        borderBottom: '1px solid #edebe9',
                        paddingBottom: '4px'
                      }}>
                        Fixed Cards ({fixedCards.length})
                      </div>
                      {fixedCards.map((card, index) => (
                        <DraggableCard
                          key={card.id}
                          card={card}
                          index={sortedCards.findIndex(c => c.id === card.id)}
                          onSelectionChange={handleSelectionChange}
                          maxSelectionReached={maxSelectionReached}
                          requiredUserSelections={requiredUserSelections}
                        />
                      ))}
                    </div>
                  )}

                  {unselectedCards.length > 0 && (
                    <div style={{ marginBottom: '12px' }}>
                      <div style={{ 
                        fontSize: '13px', 
                        fontWeight: '600', 
                        color: '#605e5c', 
                        marginBottom: '8px',
                        borderBottom: '1px solid #edebe9',
                        paddingBottom: '4px'
                      }}>
                        Available Cards ({unselectedCards.length})
                      </div>
                      {unselectedCards.map((card, index) => (
                        <DraggableCard
                          key={card.id}
                          card={card}
                          index={sortedCards.findIndex(c => c.id === card.id)}
                          onSelectionChange={handleSelectionChange}
                          maxSelectionReached={maxSelectionReached}
                          requiredUserSelections={requiredUserSelections}
                        />
                      ))}
                    </div>
                  )}
                </>
              );
            })()}
            
            <div style={{ marginTop: '16px', padding: '8px 12px', backgroundColor: '#f3f2f1', borderRadius: '4px' }}>
              <div style={{ fontSize: '12px', color: '#605e5c', fontWeight: '600' }}>
                Selection Status: {fixedCount} fixed, {selectableCount}/{requiredUserSelections} selectable selected.
              </div>
            </div>
            <div style={{ marginTop: '20px', display: 'flex', gap: '12px' }}>
              <PrimaryButton
                text="Save"
                onClick={handleSave}
                disabled={!hasCorrectSelections}
                title={!hasCorrectSelections ? `Please select exactly ${requiredUserSelections} more card(s).` : ''}
              />
              <DefaultButton text="Cancel" onClick={handleCancel} />
            </div>
          </div>
        </Panel>
      </section>
    </DndContext>
  );
};

export default ModernSharePointDashboard;
