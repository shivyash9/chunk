1. Analyse the document and assign unique cc tag to each para & tables [tag with user_id, regex]
2. Create chunk based on cc tags [Store which cc id belongs to which chunk]
   - To ensure ordering (this cc id will not help in any way) -> rely on order in which para was recieved 
3. Pass these cc tag id to LLM and have conversation between taskpane and Word doc using these cc tag id only


Document Editing Handling:
1. Edited via taskpane:
   
   
2. Edited via document:
   - How to identify [ where ] change happened and exactly which cc id was changed/new cc id added/ old cc id removed?
   - How to identify [ what ] change was made to update the chunk?





