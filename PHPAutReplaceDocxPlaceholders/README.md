# PH Auto Replace Docx Placeholder 
### Features of the code Imp
- Replaces *%varName*% style placeholders with text, images, or entire tables.
- Dynamic Image Insertion Replaces a placeholder with an image using an associative array ['image' => 'path/to/image.jpg', 'width' => 300, 'height' => 300].
- Full Table Replacement: Replaces a single placeholder with an entirely new table, useful for dynamic content that isn't already in a table format.
- Automatic Row Generation adds new rows to a table for each row of data provided
- copies the styling of the template row, including cell borders, paragraph alignment, and text formatting

## Sample input & output 
Input
<img width="1206" height="807" alt="image" src="https://github.com/user-attachments/assets/8dd8834e-85fd-4f4f-bdbc-29f75809b31a" />
Output
<img width="1600" height="727" alt="image" src="https://github.com/user-attachments/assets/9b3376de-b039-4591-b485-245e393a8850" />
