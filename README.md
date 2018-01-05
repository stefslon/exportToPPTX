## Overview

exportToPPTX allows user to create PowerPoint 2007+ (PPTX) files without using COM-objects automation (or PowerPoint application itself). Proper XML files are created and packaged into PPTX file that can be read and displayed by PowerPoint.

*Note about PowerPoint 2003 and older:* To open these PPTX files in older PowerPoint software you will need to get a free office compatibility pack from Microsoft: http://www.microsoft.com/en-us/download/details.aspx?id=3

## Usage

**Basic command syntax:**
```matlab
pptx    = exportToPPTX();
pptx.<command>(...)
```
    
**List of available commands:**

```matlab
pptx    = exportToPPTX(fileName,...);
```

Creates new PowerPoint presentation if first parameter is empty or opens existing presentation. Actual PowerPoint files are not written until 'save' command is called. No required inputs. This command does not return any values. *Additional options, applicable if creating new presentation:*
* `Dimensions` Two element vector specifying presentation's width and height in inches. Default size is 10 x 7.5 in.
* `Author` Specify presentation's author. Default is exportToPPTX.
* `Title` Specify presentation's title. Default is "Blank".
* `Subject` Specify presentation's subject line. Default is empty (blank).
* `Comments` Specify presentation's comments. Default is empty (blank).


```matlab
slideId = h.addSlide();
```

Adds a slide to the presentation. No additional inputs required. Returns newly created slide ID. Note slide IDs are not neccessarily in order.

```matlab
pptx.addPicture([figureHandle|axesHandle|imageFilename|CDATA],...)
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
pptx.addTextbox(textboxText,...)
```

Adds textbox to the current slide. Requires text of the box to be added. This command does not return any values. *Additional options:*
* `Position` Four element vector: x, y, width, height (in inches) that controls the placement and size of the textbox.
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
pptx.addNote(noteText,...)
```

Adds notes information to the current slide. Requires text of the notes to be added. This command does not return any values. Note: repeat calls overwrite previous information. *Additional options:*
* `FontWeight`  Weight of text characters:
    * normal - use regular font (default)
    * bold - use bold font
* `FontAngle` Character slant:
    * normal - no character slant (default)
    * italic - use slanted font

```matlab
fileOutPath = pptx.save(fileName)
```

Saves current presentation. If new PowerPoint was created, then filename to save to is required. If PowerPoint was openned, then by default it will write changes back to the same file. If another filename is provided, then changes will be written to the new file (effectively a 'Save As' operation). Returns full name of the presentation file written.


## Example

Here is a very simple example

```matlab
% Start new presentation
pptx    = exportToPPTX();

% Just an example image
load mandrill; figure('color','w'); image(X); colormap(map); axis off; axis image;

% Add slide, then add image to it, then add box
pptx.addSlide();
pptx.addPicture(gcf,'Scale','maxfixed');
pptx.addTextbox('Mandrill','Position',[0 5 6 1],'FontWeight','bold','HorizontalAlignment','center','VerticalAlignment','bottom');

% Save
pptx.save('example.pptx');
```

A more elaborate example is included in the `examples_exportToPPTX.m` file.
