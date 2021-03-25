## Overview

exportToPPTX allows user to create PowerPoint 2007+ (PPTX) files without using COM-objects automation (or PowerPoint application itself). Proper XML files are created and packaged into PPTX file that can be read and displayed by PowerPoint.

## Usage

**Basic command syntax:**
```matlab
pptx    = exportToPPTX();
pptx.<command>(...)
```
    
#### List of Available Commands
-   [exportToPPTX](#exportToPPTX)
-   [save](#save)
-   [addSlide](#addSlide)
-   [switchSlide](#switchSlide)
-   [addPicture](#addPicture)
-   [addShape](#addShape)
-   [addNote](#addNote)
-   [addTextbox](#addTextbox)
-   [addTable](#addTable)
	
### exportToPPTX

```matlab
pptx    = exportToPPTX(fileName,...);
```

Creates new PowerPoint presentation if first parameter is empty or opens existing presentation. Actual PowerPoint files are not written until 'save' command is called. No required inputs. This command does not return any values. 

#### Additional parameters, applicable if creating new presentation:
* `Dimensions` Two element vector specifying presentation's width and height in inches. Default size is 10 x 7.5 in.
* `Author` Specify presentation's author. Default is exportToPPTX.
* `Title` Specify presentation's title. Default is "Blank".
* `Subject` Specify presentation's subject line. Default is empty (blank).
* `Comments` Specify presentation's comments. Default is empty (blank).

### save

```matlab
fileOutPath = pptx.save(fileName)
```

Saves current presentation. If new PowerPoint was created, then filename to save to is required. If PowerPoint was openned, then by default it will write changes back to the same file. If another filename is provided, then changes will be written to the new file (effectively a 'Save As' operation). Returns full name of the presentation file written.

### addSlide

```matlab
slideId = h.addSlide();
```

Adds a slide to the presentation. No additional inputs required. Returns newly created slide ID (sequential slide number signifying total slides in the deck, not neccessarily slide order).

#### Additional parameters
* `Position` Specify position at which to insert new slide. The value must be between 1 and the total number of slides.
* `BackgroundColor` Three element vector specifying RGB value in the range from 0 to 1. By default background is white.
* `Master` Master layout ID or name. By default first master layout is used.
* `Layout` Slide template layout ID or name. By default first slide template layout is used.

### switchSlide

```matlab
slideId = h.switchSlide(slideNum);
```

Switches current slide to be operated on. Requires slide ID as the second input parameter.

### addPicture

```matlab
pptx.addPicture([figureHandle|axesHandle|imageFilename|CDATA],...)
```

Adds picture to the current slide. Requires figure or axes handle or image filename or CDATA to be supplied. Images supplied as handles or CDATA matricies are saved in PNG format. This command does not return any values.

#### Additional parameters
* `Scale` Controls how image is placed on the slide:
    * `noscale` - No scaling (place figure as is in the center of the slide) (default)
    * `maxfixed` - Max size while preserving aspect ratio
    * `max` - Max size with no aspect ratio preservation
* `Position` Four element vector: x, y, width, height (in inches) that controls the placement and size of the image. This property overrides Scale.
* `LineWidth` Width of the picture's edge line, a single value (in points). Edge is not drawn by default. Unless either LineWidth or EdgeColor are specified. 
* `EdgeColor` Color of the picture's edge, a three element vector specifying RGB value. Edge is not drawn by default. Unless either LineWidth or EdgeColor are specified. 

### addShape

```matlab
pptx.addShape(xData,yData,...)
```

Add lines or closed shapes to the current slide. Requires X and Y data to be supplied. This command does not return any values.

#### Additional parameters
* `ClosedShape` Specifies whether the shape is automatically closed or not. Default value is false.
* `LineWidth` Width of the line, a single value (in points). Default line width is 1 point. Set LineWidth to zero have no edge drawn.
* `LineColor` Color of the drawn line, a three element vector specifying RGB value. Default color is black.
* `LineStyle` Style of the drawn line. Default style is a solid line. The following styles are available:
	* `-` solid
	* `:` dotted
	* `-.` dash dot
	* `--` dashes
* `BackgroundColor` Shape fill color, a three element vector specifying RGB value. By default shapes are drawn transparent.

### addNote

```matlab
pptx.addNote(noteText,...)
```

Adds notes information to the current slide. Requires text of the notes to be added. This command does not return any values. Note: repeat calls overwrite previous information.

#### Additional parameters
* `FontWeight`  Weight of text characters:
    * `normal` - use regular font (default)
    * `bold` - use bold font
* `FontAngle` Character slant:
    * `normal` - no character slant (default)
    * `italic` - use slanted font

### addTextbox

```matlab
pptx.addTextbox(textboxText,...)
```

Adds textbox to the current slide. Requires text of the box to be added. This command does not return any values.

#### Additional parameters
* `Position` Four element vector: x, y, width, height (in inches) that controls the placement and size of the textbox.
* `Color` Three element vector specifying RGB value in range from 0 to 1. Default text color is black.
* `BackgroundColor` Three element vector specifying RGB value in the range from 0 to 1. By default background is transparent.
* `FontSize` Specifies the font size to use for text. Default font size is 12.
* `FontWeight` Weight of text characters:
    * `normal` - use regular font (default)
    * `bold` - use bold font
* `FontAngle` Character slant:
    * `normal` - no character slant (default)
    * `italic` - use slanted font
* `Rotation` Determines the orientation of the textbox. Specify values of rotation in degrees (positive angles cause counterclockwise rotation).
* `HorizontalAlignment` Horizontal alignment of text:
    * `left` - left-aligned text (default)
    * `center` - centered text
    * `right` - right-aligned text
* `VerticalAlignment` Vertical alignment of text:
    * `top` - top-aligned text (default)
    * `middle` - align to the middle of the textbox
    * `bottom` - bottom-aligned text
* `LineWidth` Width of the textbox's edge line, a single value (in points). Edge is not drawn by default. Unless either LineWidth or EdgeColor are specified. 
* `EdgeColor` Color of the textbox's edge, a three element vector specifying RGB value. Edge is not drawn by default. Unless either LineWidth or EdgeColor are specified. 

### addTable

```matlab
pptx.addTable(tableData,...)
```

Adds PowerPoint table to the current slide. Requires table content to be supplied in the form of a cell matrix. This command does not return any values.

All of the 'addTextbox' Additional parameters apply to the table as well.

## Example

Here is a very simple example

```matlab
% Start new presentation
pptx    = exportToPPTX();
pptx.title  = 'Basic example'

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
