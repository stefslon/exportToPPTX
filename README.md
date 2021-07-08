[![View exportToPPTX on File Exchange](https://www.mathworks.com/matlabcentral/images/matlab-file-exchange.svg)](https://www.mathworks.com/matlabcentral/fileexchange/40277-exporttopptx)

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
pptx    = exportToPPTX([fileName],...);
```

Creates new PowerPoint presentation if first parameter is empty or opens existing presentation. Actual PowerPoint files are not written until `save` command is called. No required inputs. This command does not return any values. 

#### Additional parameters, applicable if creating new presentation:
* `Dimensions` Two element vector specifying presentation's width and height in inches. Default size is 10 x 7.5 in.
* `Author` Specify presentation's author. Default is exportToPPTX.
* `Title` Specify presentation's title. Default is "Blank".
* `Subject` Specify presentation's subject line. Default is empty (blank).
* `Comments` Specify presentation's comments. Default is empty (blank).

### save

```matlab
fileOutPath = pptx.save([filename])
```

Saves current presentation. If new PowerPoint was created, then filename to save to is required. If PowerPoint was openned, then by default it will write changes back to the same file. If another filename is provided, then changes will be written to the new file (effectively a 'Save As' operation). Returns full name of the presentation file written.

### addSlide

```matlab
slideId = pptx.addSlide();
```

Adds a slide to the presentation. No additional inputs required. Returns newly created slide ID (sequential slide number signifying total slides in the deck, not neccessarily slide order).

#### Additional parameters
* `Position` Specify position at which to insert new slide. The value must be between 1 and the total number of slides.
* `BackgroundColor` Three element vector specifying RGB value in the range from 0 to 1. By default background is white.
* `Master` Master layout ID or name. By default first master layout is used.
* `Layout` Slide template layout ID or name. By default first slide template layout is used.

### switchSlide

```matlab
slideId = pptx.switchSlide(slideId);
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
    * `maxfixed` - Max size while preserving aspect ratio (default when Position is not set)
    * `max` - Max size with no aspect ratio preservation (default when Position is set)
* `Position` Four element vector: x, y, width, height (in inches) that controls the placement and size of the image. This property overrides `Scale`.
* `LineWidth` Width of the picture's edge line, a single value (in points). Edge is not drawn by default. Unless either `LineWidth` or `EdgeColor` are specified. 
* `EdgeColor` Color of the picture's edge, a three element vector specifying RGB value. Edge is not drawn by default. Unless either `LineWidth` or `EdgeColor` are specified. 
* `OnClick` Links text to another slide (if slide number is given as an integer) or URL or another file

### addShape

```matlab
pptx.addShape(xData,yData,...)
```

Add lines or closed shapes to the current slide. Requires X and Y data to be supplied. This command does not return any values.

#### Additional parameters
* `ClosedShape` Specifies whether the shape is automatically closed or not. Default value is false.
* `LineWidth` Width of the line, a single value (in points). Default line width is 1 point. Set `LineWidth` to zero have no edge drawn.
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
* `LineWidth` Width of the textbox's edge line, a single value (in points). Edge is not drawn by default. Unless either `LineWidth` or `EdgeColor` are specified. 
* `EdgeColor` Color of the textbox's edge, a three element vector specifying RGB value. Edge is not drawn by default. Unless either `LineWidth` or `EdgeColor` are specified. 
* `OnClick` Links text to another slide (if slide number is given as an integer) or URL or another file

### addTable

```matlab
pptx.addTable(tableData,...)
```

Adds PowerPoint table to the current slide. Requires table content to be supplied in the form of a cell matrix. This command does not return any values.

All of the `addTextbox` Additional parameters apply to the table as well.

## Markdown

Any textual inputs (addtext, addnote) support basic markdown formatting: 
- Bulleted lists (lines start with "-" followed by a space)
- Numbered lists (lines start with "#" followed by a space)
- Bolded text (enclosed in "\*\*") Ex. this **word** is bolded
- Italicized text (enclosed in "\*") Ex. this *is* italics
- Underlined text (enclosed in "\_") Ex. this text is _underlined_
- Markdown characters can be escaped with backslash "\\" to be treated literally

## Notes about Templates

In order to use PowerPoint templates with exportToPPTX a basic knowledge of the structure of the template is required. You will need to know master layout name (especially if there are more than one master layout), slide layout names, and placeholder names on each layout slide. There are multiple ways of getting this information. 

The easiest way of getting template structure information is to open the presentation in PowerPoint and to look under Layout drop-down menu on the Home tab. Master name will be given at the top of the list. Layout names will be listed under each slide thumbnail. Placeholder names are not easy (if not impossible) to get to from PowerPoint itself. But typically they are named with obvious names such as Title, Content Placeholder, Text Placeholder, etc. 

Alternative way of getting template structure information is to open presentation template with exportToPPTX and run `disp` which will list out all available master layouts, slide layouts, and placeholders on each slide layout. Here is an example with the included `Parallax.pptx` template:

```matlab
>> pptx    = exportToPPTX('Parallax.pptx');
>> pptx
pptx = 
	File:          D:\Stefan\MATLAB\exportToPPTX\Parallax.pptx
	Total slides:  0
	Current slide: 0
	Author:        Stefan Slonevskiy
	Title:         PowerPoint Presentation
	Subject:       
	Description:   
	Dimensions:    13.33 x 7.50 in

	Master #1: Parallax
		Layout #1: Content with Caption (Title 1, Content Placeholder 2, Text Placeholder 3, Date Placeholder 4, Footer Placeholder 5, Slide Number Placeholder 6)
		Layout #2: Name Card (Title 1, Text Placeholder 2, Date Placeholder 3, Footer Placeholder 4, Slide Number Placeholder 5)
		Layout #3: Section Header (Title 1, Text Placeholder 2, Date Placeholder 3, Footer Placeholder 4, Slide Number Placeholder 5)
		Layout #4: Blank (Date Placeholder 1, Footer Placeholder 2, Slide Number Placeholder 3)
		Layout #5: Quote with Caption (TextBox 13, TextBox 14, Title 1, Text Placeholder 9, Text Placeholder 2, Date Placeholder 3, Footer Placeholder 4, Slide Number Placeholder 5)
		Layout #6: Vertical Title and Text (Vertical Title 1, Vertical Text Placeholder 2, Date Placeholder 3, Footer Placeholder 4, Slide Number Placeholder 5)
		Layout #7: Title and Content (Title 1, Content Placeholder 2, Date Placeholder 3, Footer Placeholder 4, Slide Number Placeholder 5)
		Layout #8: Title and Vertical Text (Title 1, Vertical Text Placeholder 2, Date Placeholder 3, Footer Placeholder 4, Slide Number Placeholder 5)
		Layout #9: Title Slide (Freeform 6, Freeform 7, Freeform 9, Freeform 10, Freeform 11, Freeform 12, Title 1, Subtitle 2, Date Placeholder 3, Footer Placeholder 4, Slide Number Placeholder 5)
		Layout #10: Title Only (Title 1, Date Placeholder 2, Footer Placeholder 3, Slide Number Placeholder 4)
		Layout #11: Title and Caption (Title 1, Text Placeholder 2, Date Placeholder 3, Footer Placeholder 4, Slide Number Placeholder 5)
		Layout #12: Comparison (Title 1, Text Placeholder 2, Content Placeholder 3, Text Placeholder 4, Content Placeholder 5, Date Placeholder 6, Footer Placeholder 7, Slide Number Placeholder 8)
		Layout #13: True or False (Title 1, Text Placeholder 9, Text Placeholder 2, Date Placeholder 3, Footer Placeholder 4, Slide Number Placeholder 5)
		Layout #14: Panoramic Picture with Caption (Title 1, Picture Placeholder 2, Text Placeholder 3, Date Placeholder 4, Footer Placeholder 5, Slide Number Placeholder 6)
		Layout #15: Two Content (Title 1, Content Placeholder 2, Content Placeholder 3, Date Placeholder 4, Footer Placeholder 5, Slide Number Placeholder 6)
		Layout #16: Picture with Caption (Title 1, Picture Placeholder 2, Text Placeholder 3, Date Placeholder 4, Footer Placeholder 5, Slide Number Placeholder 6)
		Layout #17: Quote Name Card (TextBox 13, TextBox 14, Title 1, Text Placeholder 9, Text Placeholder 2, Date Placeholder 3, Footer Placeholder 4, Slide Number Placeholder 5)
```

Once you have all this structure information available it's easy to use templates with exportToPPTX. When creating new slide specify which slide layout you want to use. When adding text, images, or tables to the slide, instead of giving exact position and dimensions of the element, simply pass in the placeholder name. 

## Examples

### Basic Example

Here is a simple example that does not use templates

```matlab
% Start new presentation
pptx    = exportToPPTX();

% Set presentation title
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

### Template Example

Here is another simple example that uses templates

```matlab
% Open presentation template
pptx    = exportToPPTX('Parallax.pptx');

% Add new slide with layout #9 (Title Slide)
pptx.addSlide('Master',1,'Layout','Title Slide');

% Add title text and subtitle text
pptx.addTextbox('Example Presentation','Position','Title');
pptx.addTextbox('Created with exportToPPTX','Position','Subtitle');

% Save as another presentation and close
pptx.save('example2');
```

A more elaborate template example is included in the `examples2_exportToPPTX.m` file.

`Parallax.pptx`, included with this tool, is one of the default PowerPoint templates distributed with Microsoft Office 2013.

## Switching from previous (non-class) version

Changing over to class-based implementation is a major change and breaks all previous usage of the code. However updating to a new implementation should not be too bad. See the following table of the old to new commands mapping.

Old Code | New Code
-------- | --------
`exportToPPTX('new',...)` | `pptx = exportToPPTX(...)`
`exportToPPTX('open',file)` | `pptx = exportToPPTX(file)`
`exportToPPTX('save',file)` | `pptx.save(file)`
`exportToPPTX('addslide',...)` | `pptx.addSlide(...)`
`exportToPPTX('switchslide',slideID)` | `pptx.switchSlide(slideID)`
`exportToPPTX('addpicture',h,...)` | `pptx.addPicture(h,...)`
`exportToPPTX('addtext',txt,...)` | `pptx.addTextbox(txt,...)`
`exportToPPTX('addnote',txt,...)` | `pptx.addNote(txt,...)`
`exportToPPTX('addshape',x,y,...)` | `pptx.addShape(x,y,...)`
`exportToPPTX('addtable',tbl,...`) | `pptx.addTable(tbl,...)`
`exportToPPTX('close')` | There is no need to explicitly call close command. Clearing handle causes temporary files to be cleaned up.
`exportToPPTX('saveandclose')` | `pptx.save(file)` 
`exportToPPTX('query')` | `pptx` prints information about the file


