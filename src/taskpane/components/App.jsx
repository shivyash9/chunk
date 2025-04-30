/* global console */

import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import { makeStyles, tokens, shorthands } from "@fluentui/react-components";
import { 
  Button, Text, List, ListItem, Dialog, DialogSurface, 
  DialogBody, DialogTitle, DialogContent, DialogActions, Textarea, Field, Input,
  Card, CardHeader, CardFooter, Divider, TabList, Tab, Badge, Spinner
} from "@fluentui/react-components";
import { analyzeDocument, deleteContext, selectContentControlById, insertParagraphAfter, deleteContentControlById, replaceParagraphContent, addCommentToParagraph } from "../wordService";
import { 
  AddRegular, 
  DeleteRegular, 
  EditRegular, 
  CommentRegular,
  DocumentSearchRegular,
  DeleteDismissRegular,
  ArrowSyncRegular,
  DocumentBulletListRegular,
  ArrowForwardRegular,
  SearchRegular
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    ...shorthands.padding("0"),
    backgroundColor: tokens.colorNeutralBackground1,
    display: "flex",
    flexDirection: "column",
  },
  contentArea: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap("16px"),
    ...shorthands.padding("16px"),
    flex: 1,
    overflowY: "auto",
  },
  buttonsContainer: {
    display: "flex",
    ...shorthands.gap("8px"),
    marginBottom: "16px",
  },
  actionButton: {
    minWidth: "auto",
    flex: 1,
  },
  card: {
    ...shorthands.margin("0", "0", "16px", "0"),
  },
  paragraph: {
    margin: "0 0 12px 0",
  },
  listContainer: {
    maxHeight: "calc(100vh - 200px)",
    minHeight: "200px",
    overflowY: "auto",
    ...shorthands.borderRadius("4px"),
    scrollBehavior: "smooth",
    "&::-webkit-scrollbar": {
      width: "8px",
    },
    "&::-webkit-scrollbar-track": {
      background: tokens.colorNeutralBackground1,
    },
    "&::-webkit-scrollbar-thumb": {
      background: tokens.colorNeutralStroke1,
      ...shorthands.borderRadius("4px"),
      "&:hover": {
        background: tokens.colorNeutralStroke2,
      },
    },
  },
  listItem: {
    ...shorthands.margin("0", "0", "8px", "0"),
    ...shorthands.padding("12px"),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius("4px"),
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke1),
    wordBreak: "break-word",
    cursor: "pointer",
    transition: "all 0.2s ease",
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground2,
      ...shorthands.borderColor(tokens.colorBrandBackground),
    },
  },
  activeItem: {
    backgroundColor: tokens.colorBrandBackground2,
    ...shorthands.borderColor(tokens.colorBrandBackground),
    "&:hover": {
      backgroundColor: tokens.colorBrandBackground2,
    },
  },
  noElements: {
    fontStyle: "italic",
    color: tokens.colorNeutralForeground3,
    ...shorthands.margin("24px", "0"),
    textAlign: "center",
  },
  listItemContent: {
    width: '100%', 
    display: 'flex', 
    justifyContent: 'space-between', 
    alignItems: 'flex-start',
  },
  listItemActions: {
    display: 'flex', 
    ...shorthands.gap('4px'),
    marginLeft: '8px',
  },
  searchInput: {
    flexGrow: 1,
  },
  searchSection: {
    ...shorthands.padding("16px"),
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.borderRadius("6px"),
    ...shorthands.margin("0", "0", "16px", "0"),
  },
  tabContent: {
    ...shorthands.padding("16px", "0", "0", "0"),
  },
  badge: {
    marginLeft: "8px",
  },
  loadingOverlay: {
    position: "absolute",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: "rgba(255, 255, 255, 0.7)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 1000,
  },
  spinnerContainer: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    ...shorthands.gap("12px"),
  },
  ellipsis: {
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    maxWidth: "180px",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [rawContentControls, setRawContentControls] = React.useState([]);
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [activeItemId, setActiveItemId] = React.useState(null);
  const [insertText, setInsertText] = React.useState("");
  const [showInsertDialog, setShowInsertDialog] = React.useState(false);
  const [currentInsertId, setCurrentInsertId] = React.useState(null);
  const [analysisTime, setAnalysisTime] = React.useState(0);
  
  // State variables for paragraph ID search functionality
  const [searchParaId, setSearchParaId] = React.useState("");
  const [showReplaceDialog, setShowReplaceDialog] = React.useState(false);
  const [showCommentDialog, setShowCommentDialog] = React.useState(false);
  const [replaceText, setReplaceText] = React.useState("");
  const [commentText, setCommentText] = React.useState("");
  
  // Tab state
  const [selectedTab, setSelectedTab] = React.useState("paragraphs");
  const [processingMessage, setProcessingMessage] = React.useState("");

  // Memoize the sorted content controls
  const sortedContentControls = React.useMemo(() => {
    if (!rawContentControls || rawContentControls.length === 0) {
      return [];
    }
    
    // Create a stable sort by using both index and id for tiebreakers
    return [...rawContentControls].sort((a, b) => {
      // First sort by index
      if (a.index !== b.index) {
        return a.index - b.index;
      }
      // If indexes are the same, sort by id to ensure stable order
      return a.id.localeCompare(b.id);
    });
  }, [rawContentControls]);
  
  // Force re-render when tab is changed
  const handleTabChange = (_, data) => {
    setSelectedTab(data.value);
  };

  const handleAnalyzeDocument = async () => {
    try {
      setIsProcessing(true);
      setProcessingMessage("Analyzing document...");
      const startTime = performance.now();
      
      // First, try to delete existing context
      try {
        await deleteContext();
      } catch (deleteError) {
        console.warn("Warning: Could not completely clear previous context:", deleteError);
        // Continue with analysis even if context deletion fails
      }
      
      // Then analyze the document
      const controls = await analyzeDocument();
      console.log("Controls received:", controls);
      
      // Store the raw controls unmodified to ensure we don't accidentally sort twice
      setRawContentControls(controls);
      
      const endTime = performance.now();
      setAnalysisTime((endTime - startTime) / 1000); // Convert to seconds
    } catch (error) {
      console.error("Error in handleAnalyzeDocument:", error);
      // Show an error message to the user
      setProcessingMessage("Analysis failed. Please try again or reload the add-in.");
      setTimeout(() => setProcessingMessage(""), 5000); // Clear message after 5 seconds
    } finally {
      setIsProcessing(false);
      setProcessingMessage("");
    }
  };

  const handleDeleteContext = async () => {
    try {
      setIsProcessing(true);
      setProcessingMessage("Clearing context...");
      await deleteContext();
      setRawContentControls([]);
      setActiveItemId(null);
    } catch (error) {
      console.error("Error in handleDeleteContext:", error);
      // Show an error message to the user
      setProcessingMessage("Clearing failed. Please try again or reload the add-in.");
      setTimeout(() => setProcessingMessage(""), 5000); // Clear message after 5 seconds
    } finally {
      setIsProcessing(false);
      // if (processingMessage === "Clearing context...") {
        setProcessingMessage("");
      // }
    }
  };
  
  const handleItemClick = async (itemId) => {
    try {
      const success = await selectContentControlById(itemId);
      if (success) {
        setActiveItemId(itemId);
        setSearchParaId(itemId);
      }
    } catch (error) {
      console.error("Error selecting content control:", error);
    }
  };

  const handleInsertAfter = (itemId) => {
    setCurrentInsertId(itemId);
    setInsertText("");
    setShowInsertDialog(true);
  };

  const handleInsertConfirm = async () => {
    try {
      setIsProcessing(true);
      setProcessingMessage("Inserting paragraph...");
      const result = await insertParagraphAfter(currentInsertId, insertText);
      
      if (result.success) {
        // Refresh document analysis to show the new paragraph
        await handleAnalyzeDocument();
        // Highlight the newly inserted paragraph
        await selectContentControlById(result.id);
        setActiveItemId(result.id);
      }
      
      setShowInsertDialog(false);
    } catch (error) {
      console.error("Error inserting paragraph:", error);
    } finally {
      setIsProcessing(false);
      setProcessingMessage("");
    }
  };

  const handleDeleteItem = async (itemId) => {
    try {
      const success = await deleteContentControlById(itemId);
      
      if (success) {
        // Update local state by filtering out the deleted item
        setRawContentControls(prevControls => 
          prevControls.filter(control => control.id !== itemId)
        );
        // Clear active item if it was the deleted one
        if (activeItemId === itemId) {
          setActiveItemId(null);
        }
      }
    } catch (error) {
      console.error("Error deleting content control:", error);
    }
  };

  const handleReplaceContent = async () => {
    if (!searchParaId.trim()) return;
    
    try {
      const success = await selectContentControlById(searchParaId);
      if (success) {
        setActiveItemId(searchParaId);
        setReplaceText("");
        setShowReplaceDialog(true);
      }
    } catch (error) {
      console.error("Error finding paragraph:", error);
    }
  };

  const handleAddComment = async () => {
    if (!searchParaId.trim()) return;
    
    try {
      const success = await selectContentControlById(searchParaId);
      if (success) {
        setActiveItemId(searchParaId);
        setCommentText("");
        setShowCommentDialog(true);
      }
    } catch (error) {
      console.error("Error finding paragraph:", error);
    }
  };

  const handleReplaceConfirm = async () => {
    try {
      setIsProcessing(true);
      setProcessingMessage("Replacing content...");
      const success = await replaceParagraphContent(searchParaId, replaceText);
      
      if (success) {
        // Refresh document analysis to show the updated paragraph
        await handleAnalyzeDocument();
        // Highlight the modified paragraph
        await selectContentControlById(searchParaId);
        setActiveItemId(searchParaId);
      }
      
      setShowReplaceDialog(false);
    } catch (error) {
      console.error("Error replacing paragraph content:", error);
    } finally {
      setIsProcessing(false);
      setProcessingMessage("");
    }
  };

  const handleCommentConfirm = async () => {
    try {
      setIsProcessing(true);
      setProcessingMessage("Adding comment...");
      const success = await addCommentToParagraph(searchParaId, commentText);
      
      if (success) {
        // No need to refresh the document here, just close the dialog
        setShowCommentDialog(false);
      }
    } catch (error) {
      console.error("Error adding comment:", error);
    } finally {
      setIsProcessing(false);
      setProcessingMessage("");
    }
  };

  return (
    <div className={styles.root}>
      <Header 
        logo="assets/logo-filled.png" 
        title={title} 
        message={processingMessage || "Pramata Document Analyser"}
        onClear={handleDeleteContext}
        onAnalyse={handleAnalyzeDocument}
        isProcessing={isProcessing}
        analysisTime={analysisTime}
      />
      
      {isProcessing && (
        <div className={styles.loadingOverlay}>
          <div className={styles.spinnerContainer}>
            <Spinner size="medium" />
            <Text weight="semibold">{processingMessage}</Text>
          </div>
        </div>
      )}
      
      <div className={styles.contentArea}>
        <TabList 
          selectedValue={selectedTab}
          onTabSelect={handleTabChange}
          appearance="subtle"
        >
          <Tab value="paragraphs" icon={<DocumentBulletListRegular />}>
            Paragraphs
            {sortedContentControls.length > 0 && (
              <Badge appearance="filled" className={styles.badge} size="small" color="brand">
                {sortedContentControls.length}
              </Badge>
            )}
          </Tab>
          <Tab value="actions" icon={<SearchRegular />}>
            Quick Actions
          </Tab>
        </TabList>

        {selectedTab === "paragraphs" ? (
          <div className={styles.tabContent}>
            {sortedContentControls.length > 0 ? (
              <div className={styles.listContainer} key={`list-container`}>
                <List>
                  {sortedContentControls.map((item, idx) => (
                    <ListItem 
                      key={`${item.id}-${idx}`}
                      className={`${styles.listItem} ${activeItemId === item.id ? styles.activeItem : ''}`}
                      onClick={() => handleItemClick(item.id)}
                    >
                      <div className={styles.listItemContent}>
                        <div>
                          <Text weight="semibold" size={200} className={styles.ellipsis}>
                            {item.title}: {item.id} (idx: {item.index})
                          </Text>
                          <Text block size={300}>{item.text}</Text>
                        </div>
                        <div className={styles.listItemActions}>
                          <Button 
                            icon={<AddRegular />}
                            appearance="subtle"
                            size="small" 
                            onClick={(e) => {
                              e.stopPropagation();
                              handleInsertAfter(item.id);
                            }}
                            disabled={isProcessing}
                            title="Insert after"
                          />
                          <Button 
                            icon={<DeleteRegular />}
                            appearance="subtle"
                            size="small" 
                            onClick={(e) => {
                              e.stopPropagation();
                              handleDeleteItem(item.id);
                            }}
                            disabled={isProcessing}
                            title="Delete paragraph"
                          />
                        </div>
                      </div>
                    </ListItem>
                  ))}
                </List>
              </div>
            ) : (
              <div>
                <Text className={styles.noElements}>
                  No analyzed elements yet. Click 'Analyse' to begin.
                </Text>
              </div>
            )}
          </div>
        ) : (
          <div className={styles.tabContent}>
            <Card className={styles.card}>
              <CardHeader 
                header={<Text weight="semibold">Quick Actions</Text>}
              />
              <CardFooter>
                <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
                  <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                    <Input 
                      className={styles.searchInput}
                      placeholder="Enter paragraph ID" 
                      value={searchParaId}
                      onChange={(e) => setSearchParaId(e.target.value)}
                      contentBefore={<DocumentSearchRegular />}
                    />
                    <Button 
                      icon={<ArrowForwardRegular />}
                      onClick={() => handleItemClick(searchParaId)}
                      disabled={isProcessing || !searchParaId.trim()}
                      title="Find paragraph"
                    />
                  </div>
                  
                  <Divider />
                  
                  <div style={{ display: "flex", gap: "8px" }}>
                    <Button 
                      icon={<EditRegular />}
                      onClick={handleReplaceContent}
                      disabled={isProcessing || !searchParaId.trim()}
                      title="Replace content"
                      className={styles.actionButton}
                    >
                      Replace
                    </Button>
                    <Button 
                      icon={<CommentRegular />}
                      onClick={handleAddComment}
                      disabled={isProcessing || !searchParaId.trim()}
                      title="Add comment"
                      className={styles.actionButton}
                    >
                      Comment
                    </Button>
                  </div>
                </div>
              </CardFooter>
            </Card>
          </div>
        )}
      </div>

      <Dialog open={showInsertDialog}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Insert New Paragraph</DialogTitle>
            <DialogContent>
              <Field label="Enter paragraph text:">
                <Textarea 
                  value={insertText} 
                  onChange={(e) => setInsertText(e.target.value)}
                  style={{ width: '100%', minHeight: '100px' }}
                />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setShowInsertDialog(false)}>Cancel</Button>
              <Button appearance="primary" onClick={handleInsertConfirm} disabled={isProcessing}>
                Insert
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      
      <Dialog open={showReplaceDialog}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Replace Paragraph Content</DialogTitle>
            <DialogContent>
              <Text block style={{ marginBottom: "10px" }}>Paragraph ID: {searchParaId}</Text>
              <Field label="Enter new content:">
                <Textarea 
                  value={replaceText} 
                  onChange={(e) => setReplaceText(e.target.value)}
                  style={{ width: '100%', minHeight: '100px' }}
                />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setShowReplaceDialog(false)}>Cancel</Button>
              <Button appearance="primary" onClick={handleReplaceConfirm} disabled={isProcessing}>
                Replace
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      
      <Dialog open={showCommentDialog}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Add Comment to Paragraph</DialogTitle>
            <DialogContent>
              <Text block style={{ marginBottom: "10px" }}>Paragraph ID: {searchParaId}</Text>
              <Field label="Enter comment:">
                <Textarea 
                  value={commentText} 
                  onChange={(e) => setCommentText(e.target.value)}
                  style={{ width: '100%', minHeight: '100px' }}
                />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setShowCommentDialog(false)}>Cancel</Button>
              <Button appearance="primary" onClick={handleCommentConfirm} disabled={isProcessing}>
                Add Comment
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
