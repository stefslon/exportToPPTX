## Overview

exportToPPTX allows user to create PowerPoint 2007+ (PPTX) files without using COM-objects automation (or PowerPoint application itself). Proper XML files are created and packaged into PPTX file that can be read and displayed by PowerPoint.

*Note about PowerPoint 2003 and older:* To open these PPTX files in older PowerPoint software you will need to get a free office compatibility pack from Microsoft: http://www.microsoft.com/en-us/download/details.aspx?id=3

## Usage

**Basic command syntax:**
```matlab
exportToPPTX('command',parameters,...)
```
    
**List of available commands:**

```matlab
exportToPPTX('new',...)
```

Creates new PowerPoint presentation. Actual PowerPoint files are not written until 'save' command is called. No required inputs. This command does not return any values. *Additional options:*
* `Dimensions` Two element vector specifying presentation's width and height in inches. Default size is 10 x 7.5 in.
* `Author` Specify presentation's author. Default is exportToPPTX.
* `Title` Specify presentation's title. Default is "Blank".
* `Subject` Specify presentation's subject line. Default is empty (blank).
* `Comments` Specify presentation's comments. Default is empty (blank).
* `BackgroundColor` Three element vector specifying document's background RGB value in the range from 0 to 1. By default background is white.

```matlab
exportToPPTX('open',filename)
```

Opens existing PowerPoint presentation. Requires file name of the PowerPoint file to be open. This command does not return any values.

```matlab
exportToPPTX('addslide')
```

Adds a slide to the presentation. No additional inputs required. Returns newly created slide number. *Additional options:*
* `Position` Specify position at which to insert new slide. The value must be between 1 and the total number of slides.
* `BackgroundColor` Three element vector specifying slide's background RGB value in the range from 0 to 1. By default background is white.

```matlab
exportToPPTX('addpicture',[figureHandle|axesHandle|imageFilename|CDATA],...)
```

Adds picture to the current slide. Requires figure or axes handle or image filename or CDATA to be supplied. Images supplied as handles or CDATA matricies are saved in PNG format. This command does not return any values. *Additional options:*
* `Scale` Controls how image is placed on the slide:
    * noscale - No scaling (place figure as is in the center of the slide) (default)
    * maxfixed - Max size while preserving aspect ratio
    * max - Max size with no aspect ratio preservation
* `Position` Four element vector: x, y, width, height (in inches) that controls the placement and size of the image. This property overrides Scale.
* `LineWidth` Width of the picture's edge line, a single value (in points). Edge is not drawn by default. Unless either LineWidth or EdgeColor are specified. 
* `EdgeColor` Color of the picture's edge, a three element vector specifying RGB value. Edge is not drawn by default. Unless either LineWidth or EdgeColor are specified. 

```matlab
exportToPPTX('addtext',textboxText,...)
```

Adds textbox to the current slide. Requires text of the box to be added. This command does not return any values. *Additional options:*
* `Position` Four element vector: x, y, width, height (in inches) that controls the placement and size of the textbox. Coordinates x=0, y=0 are in the upper left corner of the slide.
* `Color` Three element vector specifying RGB value in range from 0 to 1. Default text color is black.
* `BackgroundColor` Three element vector specifying RGB value in the range from 0 to 1. By default background is transparent.
* `FontSize` Specifies the font size to use for text. Default font size is 12.
* `FontWeight` Weight of text characters:
    * normal - use regular font (default)
    * bold - use bold font
* `FontAngle` Character slant:
    * normal - no character slant (default)
    * italic - use slanted font
* `Rotation` Determines the orientation of the textbox. Specify values of rotation in degrees (positive angles cause counterclockwise rotation).
* `HorizontalAlignment` Horizontal alignment of text:
    * left - left-aligned text (default)
    * center - centered text
    * right - right-aligned text
* `VerticalAlignment` Vertical alignment of text:
    * top - top-aligned text (default)
    * middle - align to the middle of the textbox
    * bottom - bottom-aligned text
* `LineWidth` Width of the textbox's edge line, a single value (in points). Edge is not drawn by default. Unless either LineWidth or EdgeColor are specified. 
* `EdgeColor` Color of the textbox's edge, a three element vector specifying RGB value. Edge is not drawn by default. Unless either LineWidth or EdgeColor are specified. 

```matlab
exportToPPTX('addnote',noteText)
```

Adds notes information to the current slide. Requires text of the notes to be added. This command does not return any values. Note: repeat calls overwrite previous information. *Additional options:*
* `FontWeight` Weight of text characters:
    * normal - use regular font (default)
    * bold - use bold font
* `FontAngle` Character slant:
    * normal - no character slant (default)
    * italic - use slanted font

```matlab
exportToPPTX('save',filename)
```

Saves current presentation. If PowerPoint was created with 'new' command, then filename to save to is required. If PowerPoint was openned, then by default it will write changes back to the same file. If another filename is provided, then changes will be written to the new file (effectively a 'Save As' operation). Returns full name of the presentation file written.

```matlab
exportToPPTX('close')
```

Cleans temporary files and closes current presentation. No additional inputs required. No outputs.

```matlab
exportToPPTX('saveandclose')
```

Shortcut to save and close at the same time. No additional inputs required. No outputs.

```matlab
exportToPPTX('query')
```

Returns current status either to the command window (if no output arguments) or to the output variable. If no presentation is currently open, returned value is null.

## Markdown

Any textual inputs (addtext, addnote) support basic markdown formatting: 
- Bulleted lists (lines start with "-")
- Numbered lists (lines start with "#")
- Bolded text (enclosed in "\*\*") Ex. this **word** is bolded
- Italicized text (enclosed in "\*") Ex. this *is* italics
- Underlined text (enclosed in "\_") Ex. this text is _underlined_

## Example

Here is a very simple example

```matlab
% Start new presentation
exportToPPTX('new','Dimensions',[6 6]);

% Just an example image
load mandrill; figure('color','w'); image(X); colormap(map); axis off; axis image;

% Add slide, then add image to it, then add box
exportToPPTX('addslide');
exportToPPTX('addpicture',gcf,'Scale','maxfixed');
exportToPPTX('addtext','Mandrill','Position',[0 5 6 1],'FontWeight','bold','HorizontalAlignment','center','VerticalAlignment','bottom');

% Save and close
exportToPPTX('save','example.pptx');
exportToPPTX('close');
```

A more elaborate example is included in the `examples_exportToPPTX.m` file.
