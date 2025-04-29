/* global Word, console */

/**
 * Analyzes the document and wraps paragraphs and tables with content controls
 * @returns {Promise<Array>} - Array of content controls with their info
 */
export async function analyzeDocument() {
  try {
    const controls = [];

    await Word.run(async (context) => {
      // Get all paragraphs and tables in the document in order
      const body = context.document.body;
      const contentItems = body.contentControls;
      const paragraphs = body.paragraphs;
      const tables = body.tables;
      
      paragraphs.load("text, parentTableOrNullObject");
      tables.load("values");
      
      await context.sync();
      
      let paragraphCounter = 1;
      let tableCounter = 1;
      
      // Process paragraphs (excluding those in tables)
      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        // Skip if paragraph is inside a table
        if (paragraph.parentTableOrNullObject.isNullObject) {
          // Skip empty paragraphs or those with default placeholder text
          if (!paragraph.text || paragraph.text.trim() === "" || paragraph.text === "Click or tap here to enter text.") {
            continue;
          }
          
          // Check if already has a content control
          paragraph.contentControls.load("tag");
          await context.sync();
          
          if (paragraph.contentControls.items.length === 0) {
            const uniqueId = `para-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
            const contentControl = paragraph.insertContentControl();
            contentControl.tag = uniqueId;
            contentControl.title = `paragraph ${paragraphCounter}`;
            
            controls.push({
              id: uniqueId,
              text: paragraph.text,
              type: "paragraph",
              index: i, // Save the original position
              title: `paragraph ${paragraphCounter}`
            });
            
            paragraphCounter++;
          }
        }
      }
      
      // Process tables
      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        // Check if the table already has a content control
        table.contentControls.load("tag");
        await context.sync();
        
        if (table.contentControls.items.length === 0) {
          const uniqueId = `table-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
          const contentControl = table.insertContentControl();
          contentControl.tag = uniqueId;
          contentControl.title = `table ${tableCounter}`;
          
          // Get table content as a string representation
          let tableText = "Table with content: ";
          for (let row = 0; row < table.values.length; row++) {
            for (let col = 0; col < table.values[row].length; col++) {
              tableText += `[${table.values[row][col]}] `;
            }
            tableText += " | ";
          }
          
          controls.push({
            id: uniqueId,
            text: tableText,
            type: "table",
            index: paragraphs.items.length + i, // Position after paragraphs
            title: `table ${tableCounter}`
          });
          
          tableCounter++;
        }
      }
      
      await context.sync();
      
      // Load all content controls after insertion for accurate ordering
      contentItems.load("title, tag");
      await context.sync();
    });
    
    // Sort controls by their index to maintain document reading order
    controls.sort((a, b) => a.index - b.index);
    return controls;
  } catch (error) {
    console.error("Error analyzing document:", error);
    throw error;
  }
}

/**
 * Deletes all content controls from the document
 */
export async function deleteContext() {
  try {
    await Word.run(async (context) => {
      // Get all content controls in the document
      const contentControls = context.document.contentControls;
      contentControls.load("tag");
      
      await context.sync();
      
      // Delete all content controls but preserve their content
      for (let i = 0; i < contentControls.items.length; i++) {
        contentControls.items[i].delete(true); // true = keep content
      }
      
      await context.sync();
    });
  } catch (error) {
    console.error("Error deleting context:", error);
    throw error;
  }
} 