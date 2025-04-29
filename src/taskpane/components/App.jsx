/* global console */

import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import { makeStyles } from "@fluentui/react-components";
import { Button, Text, List, ListItem, Dialog, DialogSurface, 
  DialogBody, DialogTitle, DialogContent, DialogActions, Textarea, Field, Input } from "@fluentui/react-components";
import { analyzeDocument, deleteContext, selectContentControlById, insertParagraphAfter, deleteContentControlById, replaceParagraphContent, addCommentToParagraph } from "../wordService";
import { AddRegular, DeleteRegular, EditRegular, CommentRegular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "1rem",
  },
  buttonsContainer: {
    display: "flex",
    gap: "10px",
    marginBottom: "20px",
  },
  button: {
    minWidth: "120px",
  },
  listContainer: {
    maxHeight: "400px",
    overflowY: "auto",
    border: "1px solid #e0e0e0",
    marginTop: "10px",
    padding: "10px",
  },
  listItem: {
    marginBottom: "5px",
    padding: "5px",
    backgroundColor: "#f5f5f5",
    borderRadius: "4px",
    wordBreak: "break-word",
    cursor: "pointer",
    transition: "background-color 0.2s",
    "&:hover": {
      backgroundColor: "#e0e0e0",
    },
  },
  activeItem: {
    backgroundColor: "#d1e8ff",
    "&:hover": {
      backgroundColor: "#c0e0ff",
    },
  },
  noElements: {
    fontStyle: "italic",
    color: "#666",
    margin: "10px 0",
  }
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [contentControls, setContentControls] = React.useState([]);
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [activeItemId, setActiveItemId] = React.useState(null);
  const [insertText, setInsertText] = React.useState("");
  const [showInsertDialog, setShowInsertDialog] = React.useState(false);
  const [currentInsertId, setCurrentInsertId] = React.useState(null);
  
  // New state variables for paragraph ID search functionality
  const [searchParaId, setSearchParaId] = React.useState("");
  const [showReplaceDialog, setShowReplaceDialog] = React.useState(false);
  const [showCommentDialog, setShowCommentDialog] = React.useState(false);
  const [replaceText, setReplaceText] = React.useState("");
  const [commentText, setCommentText] = React.useState("");

  const handleAnalyzeDocument = async () => {
    try {
      setIsProcessing(true);
      
      // First, delete existing context
      await deleteContext();
      
      // Then analyze the document
      const controls = await analyzeDocument();
      setContentControls(controls);
    } catch (error) {
      console.error("Error in handleAnalyzeDocument:", error);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDeleteContext = async () => {
    try {
      setIsProcessing(true);
      await deleteContext();
      setContentControls([]);
      setActiveItemId(null);
    } catch (error) {
      console.error("Error in handleDeleteContext:", error);
    } finally {
      setIsProcessing(false);
    }
  };
  
  const handleItemClick = async (itemId) => {
    try {
      setIsProcessing(true);
      const success = await selectContentControlById(itemId);
      if (success) {
        setActiveItemId(itemId);
        // Auto-populate the search input field with the selected item's ID
        setSearchParaId(itemId);
      }
    } catch (error) {
      console.error("Error selecting content control:", error);
    } finally {
      setIsProcessing(false);
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
    }
  };

  const handleDeleteItem = async (itemId) => {
    try {
      setIsProcessing(true);
      const success = await deleteContentControlById(itemId);
      
      if (success) {
        // Refresh the document analysis
        await handleAnalyzeDocument();
      }
    } catch (error) {
      console.error("Error deleting content control:", error);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleReplaceContent = async () => {
    if (!searchParaId.trim()) return;
    
    try {
      setIsProcessing(true);
      // First try to select the paragraph to verify it exists
      const success = await selectContentControlById(searchParaId);
      
      if (success) {
        setActiveItemId(searchParaId);
        setReplaceText("");
        setShowReplaceDialog(true);
      } else {
        console.log("Paragraph ID not found");
        // Could add a notification here that the ID was not found
      }
    } catch (error) {
      console.error("Error finding paragraph:", error);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleAddComment = async () => {
    if (!searchParaId.trim()) return;
    
    try {
      setIsProcessing(true);
      // First try to select the paragraph to verify it exists
      const success = await selectContentControlById(searchParaId);
      
      if (success) {
        setActiveItemId(searchParaId);
        setCommentText("");
        setShowCommentDialog(true);
      } else {
        console.log("Paragraph ID not found");
        // Could add a notification here that the ID was not found
      }
    } catch (error) {
      console.error("Error finding paragraph:", error);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleReplaceConfirm = async () => {
    try {
      setIsProcessing(true);
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
    }
  };

  const handleCommentConfirm = async () => {
    try {
      setIsProcessing(true);
      const success = await addCommentToParagraph(searchParaId, commentText);
      
      if (success) {
        // No need to refresh the document here, just close the dialog
        setShowCommentDialog(false);
      }
    } catch (error) {
      console.error("Error adding comment:", error);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={title} message="Document Analysis Tool" />
      
      <div className={styles.buttonsContainer}>
        <Button 
          className={styles.button}
          appearance="primary" 
          onClick={handleAnalyzeDocument}
          disabled={isProcessing}
        >
          Analyse Document
        </Button>
        
        <Button 
          className={styles.button}
          onClick={handleDeleteContext}
          disabled={isProcessing}
        >
          Delete Context
        </Button>
      </div>
      
      {/* New section for paragraph ID search and modification */}
      <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#f5f5f5", borderRadius: "4px" }}>
        <Text size={500} weight="semibold" block style={{ marginBottom: "10px" }}>
          Search and Modify by Paragraph ID
        </Text>
        
        <div style={{ display: "flex", gap: "10px", marginBottom: "10px" }}>
          <Input 
            placeholder="Enter paragraph ID" 
            value={searchParaId}
            onChange={(e) => setSearchParaId(e.target.value)}
            style={{ flexGrow: 1 }}
          />
        </div>
        
        <div style={{ display: "flex", gap: "10px" }}>
          <Button 
            icon={<EditRegular />}
            onClick={handleReplaceContent}
            disabled={isProcessing || !searchParaId.trim()}
          >
            Replace Content
          </Button>
          <Button 
            icon={<CommentRegular />}
            onClick={handleAddComment}
            disabled={isProcessing || !searchParaId.trim()}
          >
            Add Comment
          </Button>
        </div>
      </div>
      
      <Text size={500} weight="semibold">Document Elements ({contentControls.length})</Text>
      
      {contentControls.length > 0 ? (
        <div className={styles.listContainer}>
          <List>
            {contentControls.map((item) => (
              <React.Fragment key={item.id}>
                <ListItem 
                  className={`${styles.listItem} ${activeItemId === item.id ? styles.activeItem : ''}`}
                  onClick={() => handleItemClick(item.id)}
                >
                  <div style={{ width: '100%', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div>
                      <Text weight="semibold">{item.title}: {item.id}</Text>
                      <Text block>{item.text}</Text>
                    </div>
                    <div style={{ display: 'flex', gap: '5px' }}>
                      <Button 
                        icon={<AddRegular />}
                        appearance="transparent"
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
                        appearance="transparent"
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
              </React.Fragment>
            ))}
          </List>
        </div>
      ) : (
        <Text className={styles.noElements}>No analyzed elements yet. Click 'Analyse Document' to begin.</Text>
      )}

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
      
      {/* Add new dialogs for Replace and Comment */}
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
