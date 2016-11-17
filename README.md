# Excel Macro MasterList DataList Find Matches
An excel macro for comparing two lists, master and data. For each row of the master list, we will find all matching rows in data list, then add up totals.

### Why?
You need a macro to extract data from one excel file, while using a second excel file as a reference.

### Implemented
* Collections
* Error Handling
* User Defined Functions

### Installation
1. Create a new folder on your desktop named ExcelMacro

2. Download both excel files from repository, MasterList.xlsx and DataList.xlsx, and drop them into ExcelMacro folder.

3. Open DataList.xlsx file and click Developer tab on the top menu Ribbon

4. Inside Developer Tab, click Macros button

5. Type RunMacro in the "Macro name:" textbox and click Create button

6. A Visual Basic Editor Window will pop up. Highlight all text in the Editor Window and delete it.

7. Back in repository, click excelmacro.vbs and highlight all the text in the file

8. In the Editor Window, paste all of the copied text

9. Goto line 15 and specify your local path for the MasterList.xlsx file

9. Click the green play button on top to run the macro


