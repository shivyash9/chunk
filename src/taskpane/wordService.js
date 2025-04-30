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
      const document = context.document;
      
      // Store the current change tracking mode
      document.load("changeTrackingMode");
      paragraphs.load("text, parentTableOrNullObject");
      contentItems.load("tag, title, text");
      tables.load("values");
      
      await context.sync();
      
      // Save the current tracking mode
      const originalTrackingMode = document.changeTrackingMode;
      
      // Temporarily turn off change tracking for our content control operations
      document.changeTrackingMode = "Off";
      await context.sync();
      
      console.log(`Total paragraphs found: ${paragraphs.items.length}`);
      console.log(`Total tables found: ${tables.items.length}`);
      console.log(`Total content controls found: ${contentItems.items.length}`);
      
      let paragraphCounter = 1;
      let tableCounter = 1;
      let skippedTableParagraphs = 0;
      let skippedEmptyParagraphs = 0;
      let paragraphsWithControls = 0;
      
      // Check for existing content controls first and preserve them in our control list
      const existingControlTags = new Set();
      for (let i = 0; i < contentItems.items.length; i++) {
        const control = contentItems.items[i];
        existingControlTags.add(control.tag);
        
        // If it's an inserted paragraph (has the right title format), preserve it
        if (control.title && control.title.includes("paragraph (inserted)")) {
          controls.push({
            id: control.tag,
            text: control.text,
            type: "paragraph",
            index: i,
            title: ""
          });
          paragraphsWithControls++;
        }
      }
      
      // Process paragraphs (excluding those in tables)
      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        try {
          // Skip if paragraph is inside a table
          if (!paragraph.parentTableOrNullObject.isNullObject) {
            skippedTableParagraphs++;
            continue;
          }
          
          // Skip empty paragraphs or those with default placeholder text
          if (!paragraph.text || paragraph.text.trim() === "" || paragraph.text === "Click or tap here to enter text.") {
            skippedEmptyParagraphs++;
            continue;
          }
          
          // Insert a content control only if there isn't one already
          const uniqueId = `para-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
          let contentControl = paragraph.insertContentControl();
          contentControl.tag = uniqueId;
          contentControl.appearance = "Hidden";
          
          controls.push({
            id: uniqueId,
            text: paragraph.text,
            type: "paragraph",
            index: i, // Save the original position
            title: "" // Empty title
          });
          
          paragraphCounter++;
          paragraphsWithControls++;
        } catch (error) {
          console.error(`Error processing paragraph ${i+1}:`, error);
        }
      }
      
      console.log(`Paragraphs skipped (in tables): ${skippedTableParagraphs}`);
      console.log(`Paragraphs skipped (empty/placeholder): ${skippedEmptyParagraphs}`);
      console.log(`Paragraphs with content controls: ${paragraphsWithControls}`);
      
      let tablesWithControls = 0;
      
      // Process tables
      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        
        try {
          // Create a content control for the table
          const uniqueId = `table-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
          const contentControl = table.insertContentControl();
          contentControl.tag = uniqueId;
          contentControl.appearance = "Hidden";
          
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
            title: "" // Empty title
          });
          
          tableCounter++;
          tablesWithControls++;
        } catch (error) {
          console.error(`Error processing table ${i+1}:`, error);
        }
      }
      
      console.log(`Tables with content controls: ${tablesWithControls}`);
      
      await context.sync();
      
      // Load all content controls after insertion for accurate ordering
      contentItems.load("title, tag, text");
      await context.sync();
      
      // Update indexes based on the actual order of content controls in the document
      const orderedControls = [];
      for (let i = 0; i < contentItems.items.length; i++) {
        const control = contentItems.items[i];
        // Find this control in our existing controls array
        const existingControl = controls.find(c => c.id === control.tag);
        if (existingControl) {
          existingControl.index = i; // Update the index to its current position
          orderedControls.push(existingControl);
        }
      }
      
      // Replace controls array with the ordered one if we have all controls
      if (orderedControls.length === controls.length) {
        controls.length = 0;
        controls.push(...orderedControls);
      } else {
        // Fall back to sorting by index
        controls.sort((a, b) => a.index - b.index);
      }
      
      // Restore the original tracking mode at the end
      document.changeTrackingMode = originalTrackingMode;
      await context.sync();
    });
    
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
      const document = context.document;
      
      // Load content controls and their tags
      contentControls.load("items");
      contentControls.load("tag");
      
      // Store the current change tracking mode
      document.load("changeTrackingMode");
      
      await context.sync();
      
      console.log(`Deleting ${contentControls.items.length} content controls`);
      
      if (contentControls.items.length === 0) {
        // No content controls to delete
        return;
      }
      
      // Save the current tracking mode
      const originalTrackingMode = document.changeTrackingMode;
      
      try {
        // Temporarily turn off change tracking for content control deletion
        document.changeTrackingMode = "Off";
        await context.sync();
        
        // Delete content controls in batches to avoid timeout issues
        const batchSize = 10;
        for (let i = 0; i < contentControls.items.length; i += batchSize) {
          const batch = contentControls.items.slice(i, Math.min(i + batchSize, contentControls.items.length));
          
          for (let j = 0; j < batch.length; j++) {
            try {
              // Delete control but preserve content
              batch[j].delete(true);
            } catch (controlError) {
              console.warn(`Could not delete content control: ${controlError.message}`);
              // Continue with other controls even if one fails
            }
          }
          
          // Sync after each batch
          await context.sync();
        }
      } finally {
        try {
          // Restore the original tracking mode regardless of success or failure
          document.changeTrackingMode = originalTrackingMode;
          await context.sync();
        } catch (restoreError) {
          console.warn("Could not restore tracking mode:", restoreError);
        }
      }
    });
  } catch (error) {
    console.error("Error deleting context:", error);
    // Don't throw the error - it's ok if we can't delete everything
    // This prevents the error from bubbling up and affecting the UI
  }
}

/**
 * Scrolls to and selects a content control by its tag ID
 * @param {string} tagId - The tag ID of the content control to select
 * @returns {Promise<boolean>} - True if successful, false if not found
 */
export async function selectContentControlById(tagId) {
  try {
    let found = false;
    
    await Word.run(async (context) => {
      // Get all content controls in the document
      const contentControls = context.document.contentControls;
      contentControls.load("tag");
      
      await context.sync();
      
      // Find the content control with the matching tag
      for (let i = 0; i < contentControls.items.length; i++) {
        if (contentControls.items[i].tag === tagId) {
          // Select the content control - this automatically scrolls it into view
          contentControls.items[i].select();
          found = true;
          break;
        }
      }
      
      await context.sync();
    });
    
    return found;
  } catch (error) {
    console.error("Error selecting content control:", error);
    return false;
  }
}

/**
 * Inserts a new paragraph between two content controls or at the end of the document
 * @param {string} adjacentId - The ID of the content control to insert after, or null to insert at the end
 * @param {string} text - The text for the new paragraph
 * @returns {Promise<{id: string, success: boolean}>} - The ID of the new content control and success status
 */
export async function insertParagraphAfter(adjacentId, text = "") {
  try {
    let newId = null;
    let success = false;
    
    await Word.run(async (context) => {
      // First, make sure track changes is enabled
      const document = context.document;
      document.changeTrackingMode = "TrackAll";
      
      // Generate a unique ID for the new paragraph
      const uniqueId = `para-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
      
      if (!adjacentId) {
        // Insert at the end of the document if no adjacentId is provided
        const paragraph = context.document.body.insertParagraph(text, Word.InsertLocation.end);
        const contentControl = paragraph.insertContentControl();
        contentControl.tag = uniqueId;
        contentControl.appearance = "Hidden";
        success = true;
        newId = uniqueId;
      } else {
        // Find the content control with the specified ID
        const contentControls = context.document.contentControls;
        contentControls.load("tag");
        
        await context.sync();
        
        // Find the target content control
        let targetControl = null;
        for (let i = 0; i < contentControls.items.length; i++) {
          if (contentControls.items[i].tag === adjacentId) {
            targetControl = contentControls.items[i];
            break;
          }
        }
        
        if (targetControl) {
          // Insert paragraph after the found content control
          const paragraph = targetControl.insertParagraph(text, Word.InsertLocation.after);
          const contentControl = paragraph.insertContentControl();
          contentControl.tag = uniqueId;
          contentControl.appearance = "Hidden";
          success = true;
          newId = uniqueId;
        }
      }
      
      await context.sync();
    });
    
    return { id: newId, success };
  } catch (error) {
    console.error("Error inserting paragraph:", error);
    return { id: null, success: false };
  }
}

/**
 * Deletes a specific content control by its ID
 * @param {string} tagId - The ID of the content control to delete
 * @returns {Promise<boolean>} - True if successfully deleted, false otherwise
 */
export async function deleteContentControlById(tagId) {
  try {
    let success = false;
    
    await Word.run(async (context) => {
      // First, make sure track changes is enabled
      const document = context.document;
      document.changeTrackingMode = "TrackAll";
      
      // Get all content controls in the document
      const contentControls = context.document.contentControls;
      contentControls.load("tag, type");
      
      await context.sync();
      
      // Find the content control with the matching tag
      for (let i = 0; i < contentControls.items.length; i++) {
        if (contentControls.items[i].tag === tagId) {
          // Get the content within the content control
          const range = contentControls.items[i].getRange();
          range.load("text");
          await context.sync();
          
          // Delete the content control preserving its contents
          contentControls.items[i].delete(true);
          await context.sync();
          
          // Now get the range which was previously in the content control and delete it (with track changes)
          // This will show up as a deletion in review mode
          range.delete();
          success = true;
          break;
        }
      }
      
      await context.sync();
    });
    
    return success;
  } catch (error) {
    console.error("Error deleting content control:", error);
    return false;
  }
}

/**
 * Replaces the content of a paragraph with new content by its ID
 * @param {string} paraId - The ID of the paragraph to replace content for
 * @param {string} newContent - The new content to replace with
 * @returns {Promise<boolean>} - True if successfully replaced, false otherwise
 */
export async function replaceParagraphContent(paraId, newContent) {
  try {
    let success = false;
    
    await Word.run(async (context) => {
      // First, make sure track changes is enabled
      const document = context.document;
      document.changeTrackingMode = "TrackAll";
      
      // Get all content controls in the document
      const contentControls = context.document.contentControls;
      contentControls.load("tag, text");
      
      await context.sync();
      
      // Find the content control with the matching tag
      for (let i = 0; i < contentControls.items.length; i++) {
        if (contentControls.items[i].tag === paraId) {
          // Get the content control's range
          const contentControl = contentControls.items[i];
          
          // Set the text directly in the content control
          contentControl.insertText(newContent, Word.InsertLocation.replace);
          contentControl.appearance = "Hidden";
          
          success = true;
          break;
        }
      }
      
      await context.sync();
    });
    
    return success;
  } catch (error) {
    console.error("Error replacing paragraph content:", error);
    return false;
  }
}

/**
 * Adds a comment to a paragraph by its ID
 * @param {string} paraId - The ID of the paragraph to add a comment to
 * @param {string} commentText - The text of the comment to add
 * @returns {Promise<boolean>} - True if successfully added comment, false otherwise
 */
export async function addCommentToParagraph(paraId, commentText) {
  try {
    let success = false;
    
    await Word.run(async (context) => {
      // Get all content controls in the document
      const contentControls = context.document.contentControls;
      contentControls.load("tag");
      
      await context.sync();
      
      // Find the content control with the matching tag
      for (let i = 0; i < contentControls.items.length; i++) {
        if (contentControls.items[i].tag === paraId) {
          // Get the range to add a comment to
          const range = contentControls.items[i].getRange();
          
          // Add a comment to the range
          range.insertComment(commentText);
          
          success = true;
          break;
        }
      }
      
      await context.sync();
    });
    
    return success;
  } catch (error) {
    console.error("Error adding comment to paragraph:", error);
    return false;
  }
} 