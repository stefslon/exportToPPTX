exportToPPTX
============

# Overview

exportToPPTX allows user to create PowerPoint 2007 (PPTX) files without using COM-objects automation. Proper XML files are created and packed into PPTX file that can be read and displayed by PowerPoint.

*Note about PowerPoint 2003 and older:*
To open these PPTX files in older PowerPoint software you will need to get a free office compatibility pack from Microsoft: http://www.microsoft.com/en-us/download/details.aspx?id=3

# Usage

**Basic command syntax:**
```
exportToPPTX('command',parameters,...)
```
    
## List of possible commands:

* **new** Creates new PowerPoint presentation. Actual PowerPoint files are not written until 'save' command is called. No required inputs. This command does not return any values. 
    * Additional options:
        * **Dimensions** - two element vector specifying presentation's width and height in inches. Default size is 10 x 7.5 in.

* **open** Opens existing PowerPoint presentation. Requires file name of the PowerPoint file to be open. This command does not return any values.

* **addslide** Adds a slide to the presentation. No additional inputs required. Returns newly created slide number.

* **addpicture** Adds picture to the current slide. Requires figure or axes handle to be supplied. All files are saved in a PNG format. This command does not return any values.
    * Additional options:
        * **Scale** Controls how image is placed on the slide:
            * noscale - No scaling (place figure as is in the center of the slide) (default)
            * maxfixed - Max size while preserving aspect ratio
            * max - Max size with no aspect ratio preservation
        * **Position** Four element vector: x, y, width, height (in inches) that controls the placement and size of the image. This property overrides Scale.

* **addtext** Adds textbox to the current slide. Requires text of the box to be added. This command does not return any values.
    * Additional options:
        * **Position** Four element vector: x, y, width, height (in inches) that controls the placement and size of the textbox.
        * **Color** Three element vector specifying RGB value in range from 0 to 1. Default text color is black.
        * **FontSize** Specifies the font size to use for text. Default font size is 12.
        * **FontWeight** Weight of text characters:
            * normal - use regular font (default)
            * bold - use bold font
        * **FontAngle**   Character slant:
            * normal - no character slant (default)
            * italic - use slanted font
        * **HorizontalAlignment** Horizontal alignment of text:
            * left - left-aligned text (default)
            * center - centered text
            * right - right-aligned text
        * **VerticalAlignment** Vertical alignment of text:
            * top - top-aligned text (default)
            * middle - align to the middle of the textbox
            * bottom - bottom-aligned text
        
* **save** Saves current presentation. If PowerPoint was created with 'new' command, then filename to save to is required. If PowerPoint was openned, then by default it will write changes back to the same file. If another filename is provided, then changes will be written to the new file (effectively a 'Save As' operation). Returns full name of the presentation file written.

* **close** Cleans temporary files and closes current presentation. No additional inputs required. No outputs.

* **saveandclose** Shortcut to save and close at the same time. No additional inputs required. No outputs.

* **query** Returns current status either to the command window (if no output arguments) or to the output variable. If no presentation is currently open, returned value is null.

