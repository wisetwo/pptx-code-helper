
![Alt text](img/codepen-square-fill.png?raw=true "pptx-code-helper")

# pptx-code-helper
An helper tool that helps getting object dimensions (size and position) and font attribute (font name and size etc.) in powerpoint files, by which improves the experience of generating powerpoint by [pptxgenjs](https://github.com/gitbrent/PptxGenJS). 

## Features
 Current features include:
 | Group | Feature |
 |-|-|
 | Object | - Object horizon and vertical position (x, y)<br>- Object width and height (w, h)|
 | Font | - fontFace, fontSize, color, bold, italic, align, lineSpacing, paraSpaceBefore, paraSpaceAfter, valign, margin |
 
## Examples

![Alt text](img/shape-attribute.png?raw=true "shape-attribute")

After select the shape and click the button shown above, the following code will be copied to clipboard:

```
x: 5.972,
y: 1.46,
w: 1.389,
h: 1.389,

```

![Alt text](img/font-attribute.png?raw=true "font-attribute")

After select the text and click the button shown above, the following code will be copied to clipboard:

```
fontFace: 'Arial Narrow',
fontSize: 24,
color: '#4472C4'
italic: true,
align: 'center',
lineSpacingMultiple: 0.9,
valign: 'bottom',
margin: [7.2, 7.2, 3.6, 3.6],
```
> Note that `italic`, `align`, `lineSpacingMultiple` (or `lineSpacing`), `valign`, `margin` attributes are for the whole text content in the text box.

# How to install 
Pptx-code-helper is a Visual Basic for Applications (VBA) add-in that can be installed within Powerpoint, requiring no administrative rights on most enterprise systems.

## Windows
You can save the add-in to your PC and then install the add-in by adding it to the Available Add-Ins list:
- Download the add-in file in the latest version (https://github.com/wisetwo/pptx-code-helper/blob/main/bin/PptxCodeHelper.ppam) and save it in a fixed location
- Open Powerpoint, click the File tab, and then click Options
- In the Options dialog box, click Add-Ins.
- In the Manage list at the bottom of the dialog box, click PowerPoint Add-ins, and then click Go.
- In the Add-Ins dialog box, click Add New.
- In the Add New PowerPoint Add-In dialog box, browse for the add-in file, and then click OK.
- A security notice appears. Click Enable Macros, and then click Close.
- There now should be an "PptxCodeHelper" page in the Powerpoint ribbon

## Mac
You can save the add-in to your Mac and then install the add-in by adding it to the Add-Ins list:
- Download the add-in file in the latest version (https://github.com/wisetwo/pptx-code-helper/blob/main/bin/PptxCodeHelper.ppam) and save it in a fixed location
- Open Powerpoint, click Tools in the application menu, and then click Add-ins...
- In the Add-Ins dialog box, click the + button, browse for the add-in file, and then click Open.
- Click Ok to close the Add-ins dialog box
- There now should be an "PptxCodeHelper" page in the Powerpoint ribbon
