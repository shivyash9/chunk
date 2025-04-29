import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import { makeStyles } from "@fluentui/react-components";
import { Button, Text, List, ListItem } from "@fluentui/react-components";
import { analyzeDocument, deleteContext, selectContentControlById } from "../wordService";

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
            {contentControls.map((item) => (
              <ListItem 
                key={item.id} 
                className={`${styles.listItem} ${activeItemId === item.id ? styles.activeItem : ''}`}
                onClick={() => handleItemClick(item.id)}
              >
                <div>
                  <Text weight="semibold">{item.title}: {item.id}</Text>
                  <Text block>{item.text}</Text>
                </div>
              </ListItem>
            ))}
          </List>
        </div>
      ) : (
        <Text className={styles.noElements}>No analyzed elements yet. Click 'Analyse Document' to begin.</Text>
      )}
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
