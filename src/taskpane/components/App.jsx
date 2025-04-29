/* global console */

import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import { makeStyles } from "@fluentui/react-components";
import { Button, Text, List, ListItem, Dialog, DialogSurface, 
  DialogBody, DialogTitle, DialogContent, DialogActions, Textarea, Field } from "@fluentui/react-components";
import { analyzeDocument, deleteContext, selectContentControlById, insertParagraphAfter, deleteContentControlById } from "../wordService";
import { AddRegular, DeleteRegular } from "@fluentui/react-icons";

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
      
      <Text size={500} weight="semibold">Document Elements ({contentControls.length})</Text>
      
      {contentControls.length > 0 ? (
        <div className={styles.listContainer}>
          <List>
            {contentControls.map((item, index) => (
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
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
