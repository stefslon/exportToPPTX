function varargout = exportToPPTX(varargin)
%  exportToPPTX Creates PowerPoint 2007 (PPTX) slides
%
%       exportToPPTX allows user to create PowerPoint 2007 (PPTX) files
%       without using COM-objects automation. Proper XML files are created and
%       packed into PPTX file that can be read and displayed by PowerPoint.
%       
%       Note about PowerPoint 2003 and older:
%       To open these PPTX files in older PowerPoint software you will need
%       to get a free office compatibility pack from Microsoft:
%       http://www.microsoft.com/en-us/download/details.aspx?id=3
%
%       Basic command syntax:
%       exportToPPTX('command',parameters,...)
%           
%       List of available commands:
%
%       new
%           Creates new PowerPoint presentation. Actual PowerPoint files
%           are not written until 'save' command is called. No required 
%           inputs. This command does not return any values. 
%
%           Additional options:
%               Dimensions  two element vector specifying presentation's
%                           width and height in inches. Default size is 
%                           10 x 7.5 in.
%
%       open
%           Opens existing PowerPoint presentation. Requires file name of
%           the PowerPoint file to be open. This command does not return
%           any values.
% 
%       addslide
%           Adds a slide to the presentation. No additional inputs
%           required. Returns newly created slide number.
%   
%       addpicture
%           Adds picture to the current slide. Requires figure or axes
%           handle or image filename or CDATA to be supplied. Images
%           supplied as handles or CDATA matricies are saved in PNG format.
%           This command does not return any values.
%
%           Additional options:
%               Scale       Controls how image is placed on the slide
%                           noscale - No scaling (place figure as is in the 
%                                     center of the slide) (default)
%                           maxfixed - Max size while preserving aspect ratio
%                           max - Max size with no aspect ratio preservation
%               Position    Four element vector: x, y, width, height (in
%                           inches) that controls the placement and size of
%                           the image. This property overrides Scale.
%               LineWidth   Width of the picture's edge line, a single
%                           value (in points). Edge is not drawn by 
%                           default. Unless either LineWidth or EdgeColor 
%                           are specified. 
%               EdgeColor   Color of the picture's edge, a three element
%                           vector specifying RGB value. Edge is not drawn
%                           by default. Unless either LineWidth or
%                           EdgeColor are specified. 
%
%       addtext
%           Adds textbox to the current slide. Requires text of the box to
%           be added. This command does not return any values.
%
%           Additional options:
%               Position    Four element vector: x, y, width, height (in
%                           inches) that controls the placement and size of
%                           the textbox.
%               Color       Three element vector specifying RGB value in the
%                           range from 0 to 1. Default text color is black.
%               BackgroundColor         Three element vector specifying RGB
%                           value in the range from 0 to 1. By default 
%                           background is transparent.
%               FontSize    Specifies the font size to use for text.
%                           Default font size is 12.
%               FontWeight  Weight of text characters:
%                           normal - use regular font (default)
%                           bold - use bold font
%               FontAngle   Character slant:
%                           normal - no character slant (default)
%                           italic - use slanted font
%               Rotation    Determines the orientation of the textbox. 
%                           Specify values of rotation in degrees (positive 
%                           angles cause counterclockwise rotation).
%               HorizontalAlignment     Horizontal alignment of text:
%                           left - left-aligned text (default)
%                           center - centered text
%                           right - right-aligned text
%               VerticalAlignment       Vertical alignment of text:
%                           top - top-aligned text (default)
%                           middle - align to the middle of the textbox
%                           bottom - bottom-aligned text
%               LineWidth   Width of the textbox's edge line, a single
%                           value (in points). Edge is not drawn by 
%                           default. Unless either LineWidth or EdgeColor 
%                           are specified. 
%               EdgeColor   Color of the textbox's edge, a three element
%                           vector specifying RGB value. Edge is not drawn
%                           by default. Unless either LineWidth or
%                           EdgeColor are specified. 
%               
%       save
%           Saves current presentation. If PowerPoint was created with
%           'new' command, then filename to save to is required. If
%           PowerPoint was openned, then by default it will write changes
%           back to the same file. If another filename is provided, then
%           changes will be written to the new file (effectively a 'Save
%           As' operation). Returns full name of the presentation file 
%           written.
%
%       close
%           Cleans temporary files and closes current presentation. No
%           additional inputs required. No outputs.
%
%       saveandclose
%           Shortcut to save and close at the same time. No additional
%           inputs required. No outputs.
%
%       query
%           Returns current status either to the command window (if no
%           output arguments) or to the output variable. If no presentation
%           is currently open, returned value is null.
%       
%

% Author: S.Slonevskiy, 02/11/2013
% File bug reports at: 
%       https://github.com/stefslon/exportToPPTX/issues

% Versions:
%   02/11/2013, Initial version
%   03/12/2013, Throw a more descriptive error if there is one when writing PPTX file
%               Move all constants into PPTXInfo.CONST structure
%               Support additional input types for addPicture
%               Add additional textbox and image formatting options
%   05/01/2013, Fix a bug with multiple images overwritting each other on slide
%               Fix a crash with new PPTX after old one was closed
%



%% Keep track of the current PPTX presentation
persistent PPTXInfo


%% Constants
% - EMUs (English Metric Units) are defined as 1/360,000 of a centimeter 
%   and thus there are 914,400 EMUs per inch, and 12,700 EMUs per point
PPTXInfo.CONST.IN_TO_EMU           = 914400;
PPTXInfo.CONST.PT_TO_EMU           = 12700;
PPTXInfo.CONST.DEG_TO_EMU          = 60000;
PPTXInfo.CONST.FONT_PX_TO_PPTX     = 100;
PPTXInfo.CONST.DEFAULT_DIMENSIONS  = [10 7.5];     % in inches


%% Initialize PPTXInfo (if it's not initialized yet)
if ~isfield(PPTXInfo,'fileOpen'),
    PPTXInfo.fileOpen   = false;
end


%% Parse inputs
if nargin==0,
    action  = 'query';
else
    action  = varargin{1};
end


%% Set blank outputs
if nargout>0,
    varargout   = cell(1,nargout);
end


%% Do what we are told to do...
switch lower(action),
    case 'new',
        %% Check if there is PPTX already in progress
        if PPTXInfo.fileOpen,
            error('exportToPPTX:fileStillOpen','PPTX file %s already in progress. Save and close first, before starting new file.',PPTXInfo.fullName);
        end
        
        %% Check for additional input parameters
        if nargin>1 && any(strcmpi(varargin,'Dimensions')),
            idx                     = find(strcmpi(varargin,'Dimensions'));
            PPTXInfo.dimensions     = round(varargin{idx+1}.*PPTXInfo.CONST.IN_TO_EMU);
            if numel(PPTXInfo.dimensions)~=2,
                error('exportToPPTX:badDimensions','Slide dimensions vector must have two values only: width x height');
            end
        else
            PPTXInfo.dimensions     = round(PPTXInfo.CONST.DEFAULT_DIMENSIONS.*PPTXInfo.CONST.IN_TO_EMU);
        end
        
        %% Obtain temp folder name
        tempName    = tempname;
        while exist(tempName,'dir'),
            tempName    = tempname;
        end
        
        %% Create temp folder to hold all PPTX files
        mkdir(tempName);
        PPTXInfo.tempName   = tempName;

        %% Create new (empty) PPTX files structure
        PPTXInfo.fullName       = [];
        PPTXInfo.createdDate    = datestr(now,'yyyy-mm-ddTHH:MM:SS');
        initBlankPPTX(PPTXInfo);

        %% Load important XML files into memory
        PPTXInfo    = loadAndParseXMLFiles(PPTXInfo);
        
        %% No outputs

        
        
    case 'open',
        %% Check if there is PPTX already in progress
        if PPTXInfo.fileOpen,
            error('exportToPPTX:fileStillOpen','PPTX file %s already in progress. Save and close first, before starting new file.',PPTXInfo.fullName);
        end

        %% Inputs
        if nargin<2,
            error('exportToPPTX:minInput','Second argument required: filename to create or open');
        end
        fileName    = varargin{2};

        %% Check input validity
        [filePath,fileName,fileExt]     = fileparts(fileName);
        if isempty(filePath),
            filePath    = pwd;
        end
        if ~strncmpi(fileExt,'.pptx',5),
            fileExt     = cat(2,fileExt,'.pptx');
        end
        fullName            = fullfile(filePath,cat(2,fileName,fileExt));
        PPTXInfo.fullName   = fullName;
        
        %% Obtain temp folder name
        tempName    = tempname;
        while exist(tempName,'dir'),
            tempName    = tempname;
        end
        
        %% Create temp folder to hold all PPTX files
        mkdir(tempName);
        PPTXInfo.tempName   = tempName;
        
        %% Check destination location/file
        if exist(fullName,'file'),
            % Open
            openExistingPPTX(PPTXInfo);
        else
            error('exportToPPTX:fileNotFound','PPTX to open not found. Start new PPTX using new command');
        end
        
        %% Load important XML files into memory
        PPTXInfo    = loadAndParseXMLFiles(PPTXInfo);
        
        %% No outputs

        
        
    case 'addslide',
        %% Check if there is PPT to add to
        if ~PPTXInfo.fileOpen,
            error('exportToPPTX:addSlideFail','No PPTX in progress. Start new or open PPTX file using new or open commands');
        end
        
        %% Create new blank slide
        PPTXInfo    = addSlide(PPTXInfo);
        
        %% Outputs
        if nargout>0,
            varargout{1}    = PPTXInfo.numSlides;
        end

        
        
    case 'addpicture',
        %% Check if there is PPT to add to
        if ~PPTXInfo.fileOpen,
            error('exportToPPTX:addPictureFail','No PPTX in progress. Start new or open PPTX file using new or open commands');
        end
        
        %% Inputs
        if nargin<2,
            error('exportToPPTX:minInput','Second argument required: figure handle or filename or CDATA');
        end
        imgData     = varargin{2};
        
        %% Add image
        if nargin>2,
            PPTXInfo    = addPicture(PPTXInfo,imgData,varargin{3:end});
        else
            PPTXInfo    = addPicture(PPTXInfo,imgData);
        end
        
        
        
    %case 'addnote',
    %    %% Check if there is PPT to add to
    %    if ~PPTXInfo.fileOpen,
    %        error('No file open. Open file using new or open commands');
    %    end
        
        
        
    case 'addtext',
        %% Check if there is PPT to add to
        if ~PPTXInfo.fileOpen,
            error('exportToPPTX:addTextFail','No PPTX in progress. Start new or open PPTX file using new or open commands');
        end
        
        %% Inputs
        if nargin<2,
            error('exportToPPTX:minInput','Second argument required: textbox contents');
        end
        boxText     = varargin{2};
        
        % Defaults
        % Places textbox the size of the slide
        textPos     = [0 0 PPTXInfo.dimensions];
        
        % Input overrides
        if nargin>2 && any(strcmpi(varargin,'Position')),
            idx         = find(strcmpi(varargin,'Position'));
            textPos     = round(varargin{idx+1}.*PPTXInfo.CONST.IN_TO_EMU);
        end
        
        %% Add textbox
        if nargin>2,
            PPTXInfo    = addTextbox(PPTXInfo,boxText,textPos,varargin{3:end});
        else
            PPTXInfo    = addTextbox(PPTXInfo,boxText,textPos);
        end
        
        
    case 'saveandclose',
        exportToPPTX('save');
        exportToPPTX('close');
        
        
        
    case 'save',
        %% Check if there is PPT to save
        if ~PPTXInfo.fileOpen,
            warning('exportToPPTX:noFileOpen','No PPTX in progress. Nothing to save.');
            return;
        end
        
        %% Inputs
        if nargin<2 && isempty(PPTXInfo.fullName),
            error('exportToPPTX:minInput','For new presentation filename to save to is required');
        end
        if nargin>=2,
            fileName    = varargin{2};
        else
            fileName    = PPTXInfo.fullName;
        end

        %% Check input validity
        [filePath,fileName,fileExt]     = fileparts(fileName);
        if isempty(filePath),
            filePath    = pwd;
        end
        if ~strncmpi(fileExt,'.pptx',5),
            fileExt     = cat(2,fileExt,'.pptx');
        end
        fullName            = fullfile(filePath,cat(2,fileName,fileExt));
        PPTXInfo.fullName   = fullName;

        
        %% Add some more useful information
        PPTXInfo.revNumber      = PPTXInfo.revNumber+1;
        PPTXInfo.updatedDate    = datestr(now,'yyyy-mm-ddTHH:MM:SS');
        
        %% Commit changes to PPTX file
        writePPTX(PPTXInfo);
        
        %% Output
        if nargout>0,
            varargout{1}    = PPTXInfo.fullName;
        end
        
        
    case 'close',
        %% Check if there is PPT to add to
        if ~PPTXInfo.fileOpen,
            warning('exportToPPTX:noFileOpen','No PPTX in progress. Nothing to close.');
            return;
        end
        
        try
            cleanUp(PPTXInfo);
            PPTXInfo.fileOpen   = false;
        catch
        end
        
        % Remove fields from PPTXInfo
        PPTXInfo    = rmfield(PPTXInfo,{'dimensions','tempName','fullName','createdDate','XML','numSlides','revNumber','imageTypes','Slide','lastSlideId','lastRId','updatedDate'});
        
        
    case 'query',
        if nargout==0,
            if ~PPTXInfo.fileOpen,
                fprintf('No file open\n');
            else
                fprintf('File: %s\n',PPTXInfo.fullName);
                fprintf('Dimensions: %.2f x %.2f in\n',PPTXInfo.dimensions./PPTXInfo.CONST.IN_TO_EMU);
                fprintf('Slides: %d\n',PPTXInfo.numSlides);
            end
        else
            if PPTXInfo.fileOpen,
                ret.fullName    = PPTXInfo.fullName;
                ret.dimensions  = PPTXInfo.dimensions./PPTXInfo.CONST.IN_TO_EMU;
                ret.numSlides   = PPTXInfo.numSlides;
                
                varargout{1}    = ret;
            end
        end
        
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function retCode = writeTextFile(fileName,fileContent)
fId     = fopen(fileName,'w');
if fId~=-1,
    fprintf(fId,'%s\n',fileContent{:});
    fclose(fId);
    retCode = 1;
else
    retCode = 0;
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function PPTXInfo = loadAndParseXMLFiles(PPTXInfo)
% Load files into persistent memory
% These will be kept in memory for the duration of the PPTX construction
PPTXInfo.XML.TOC        = xmlread(fullfile(PPTXInfo.tempName,'[Content_Types].xml'));
PPTXInfo.XML.App        = xmlread(fullfile(PPTXInfo.tempName,'docProps','app.xml'));
PPTXInfo.XML.Core       = xmlread(fullfile(PPTXInfo.tempName,'docProps','core.xml'));
PPTXInfo.XML.Pres       = xmlread(fullfile(PPTXInfo.tempName,'ppt','presentation.xml'));
PPTXInfo.XML.PresRel    = xmlread(fullfile(PPTXInfo.tempName,'ppt','_rels','presentation.xml.rels'));

% Parse overall information
PPTXInfo.numSlides      = str2num(char(getNodeValue(PPTXInfo.XML.App,'Slides')));
PPTXInfo.revNumber      = str2num(char(getNodeValue(PPTXInfo.XML.Core,'cp:revision')));
% PPTXInfo.title          = char(getNodeValue(PPTXInfo.XML.Core,'dc:title'));
% PPTXInfo.description    = char(getNodeValue(PPTXInfo.XML.Core,'dc:description'));
PPTXInfo.dimensions(1)  = str2num(char(getNodeAttribute(PPTXInfo.XML.Pres,'p:sldSz','cx')));
PPTXInfo.dimensions(2)  = str2num(char(getNodeAttribute(PPTXInfo.XML.Pres,'p:sldSz','cy')));

% Parse image type information
fileExt     = getNodeAttribute(PPTXInfo.XML.TOC,'Default','Extension');
fileType    = getNodeAttribute(PPTXInfo.XML.TOC,'Default','ContentType');
imgIdx      = find(strncmpi(fileType,'image',5));
PPTXInfo.imageTypes     = fileExt(imgIdx);

% Parse slide information
allIds                  = getNodeAttribute(PPTXInfo.XML.Pres,'p:sldId','id');
allRIds                 = getNodeAttribute(PPTXInfo.XML.Pres,'p:sldId','r:id');
if ~isempty(allIds),
    % Presentation already has some slides
    for islide=1:numel(allIds),
        PPTXInfo.Slide(islide).id   = str2num(allIds{islide});
        PPTXInfo.Slide(islide).rId  = allRIds{islide};
        
        allRels     = getNodeAttribute(PPTXInfo.XML.PresRel,'Relationship','Id');
        allTargs    = getNodeAttribute(PPTXInfo.XML.PresRel,'Relationship','Target');
        relIdx      = (strcmpi(allRels,PPTXInfo.Slide(islide).rId));
        [~,fileName]                = fileparts(allTargs{relIdx});
        PPTXInfo.Slide(islide).file = fileName;
    end
    
    PPTXInfo.lastSlideId    = max([PPTXInfo.Slide.id]);
    allRIds2                = char(allRIds{:});
    PPTXInfo.lastRId        = max(str2num(allRIds2(:,4:end)));
else
    % Completely blank presentation
    PPTXInfo.Slide          = struct([]);
    PPTXInfo.lastSlideId    = 255;  % not starting with 0 because PPTX doesn't seem to do that
    PPTXInfo.lastRId        = 2;    % not starting with 0 because rId1 is for theme, rId2 is for slide master
end

% Safety check
if PPTXInfo.numSlides~=numel(PPTXInfo.Slide),
    warning('Badly formed PPTX. Number of slides is corrupted');
    PPTXInfo.numSlides  = numel(PPTXInfo.Slide);
end

% Flag as ready to be used
PPTXInfo.fileOpen   = true;


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function openExistingPPTX(PPTXInfo)
unzip(PPTXInfo.fullName,PPTXInfo.tempName);


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function success = writePPTX(PPTXInfo)
% Commit changes to current slide, if any
if isfield(PPTXInfo.XML,'Slide') && ~isempty(PPTXInfo.XML.Slide),
    fileName    = PPTXInfo.Slide(PPTXInfo.numSlides).file;
    xmlwrite(fullfile(PPTXInfo.tempName,'ppt','slides',fileName),PPTXInfo.XML.Slide);
    xmlwrite(fullfile(PPTXInfo.tempName,'ppt','slides','_rels',cat(2,fileName,'.rels')),PPTXInfo.XML.SlideRel);
end

% Do last minute updates to the file
% if isfield(PPTXInfo,'author'),
%     setNodeValue(PPTXInfo.XML.Core,'dc:creator',PPTXInfo.author);
% end
% if isfield(PPTXInfo,'updatedBy'),
%     setNodeValue(PPTXInfo.XML.Core,'cp:lastModifiedBy',PPTXInfo.updatedBy);
% end
if isfield(PPTXInfo,'createdDate'),
    setNodeValue(PPTXInfo.XML.Core,'dcterms:created',PPTXInfo.createdDate);
end
if isfield(PPTXInfo,'updatedDate'),
    setNodeValue(PPTXInfo.XML.Core,'dcterms:modified',PPTXInfo.updatedDate);
end

% Update overall PPTX information
setNodeValue(PPTXInfo.XML.Core,'cp:revision',PPTXInfo.revNumber);
setNodeValue(PPTXInfo.XML.App,'Slides',PPTXInfo.numSlides);
% setNodeValue(PPTXInfo.XML.Core,'dc:title',PPTXInfo.title);
% setNodeValue(PPTXInfo.XML.Core,'dc:description',PPTXInfo.description);

% Commit all changes to XML files
xmlwrite(fullfile(PPTXInfo.tempName,'[Content_Types].xml'),PPTXInfo.XML.TOC);
xmlwrite(fullfile(PPTXInfo.tempName,'docProps','app.xml'),PPTXInfo.XML.App);
xmlwrite(fullfile(PPTXInfo.tempName,'docProps','core.xml'),PPTXInfo.XML.Core);
xmlwrite(fullfile(PPTXInfo.tempName,'ppt','presentation.xml'),PPTXInfo.XML.Pres);
xmlwrite(fullfile(PPTXInfo.tempName,'ppt','_rels','presentation.xml.rels'),PPTXInfo.XML.PresRel);

% Write all files to ZIP (aka PPTX) file
% 1) Zip temp directory with all XML files to tempName.zip
% 2) Rename tempName.zip to fullName (this leaves only .pptx extension in the filename)
zip(PPTXInfo.tempName,'*',PPTXInfo.tempName);
[success,msg] = movefile(cat(2,PPTXInfo.tempName,'.zip'),PPTXInfo.fullName);
if ~success,
    error('exportToPPTX:saveFail','PPTX file cannot be written to: %s.',msg);
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function cleanUp(PPTXInfo)
rmdir(PPTXInfo.tempName,'s');


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function PPTXInfo = addSlide(PPTXInfo)
% Adding a new slide to PPTX
%   1. Create slide#.xml file
%   2. Create slide#.xml.rels file
%   3. Update presentation.xml to include new slide
%   4. Update presentation.xml.rels to link new slide to presentation.xml
%   5. Update [Content_Types].xml 

% Before creating new slide, is there a current slide that needs to be
% saved to XML file?
if isfield(PPTXInfo.XML,'Slide') && ~isempty(PPTXInfo.XML.Slide),
    fileName    = PPTXInfo.Slide(PPTXInfo.numSlides).file;
    xmlwrite(fullfile(PPTXInfo.tempName,'ppt','slides',fileName),PPTXInfo.XML.Slide);
    xmlwrite(fullfile(PPTXInfo.tempName,'ppt','slides','_rels',cat(2,fileName,'.rels')),PPTXInfo.XML.SlideRel);
end

% Init useful variables
PPTXInfo.numSlides      = PPTXInfo.numSlides+1;
PPTXInfo.lastSlideId    = PPTXInfo.lastSlideId+1;
PPTXInfo.lastRId        = PPTXInfo.lastRId+1;

% Create new XML file
% file name and path are relative to presentation.xml file
fileName        = sprintf('slide%d.xml',PPTXInfo.numSlides);
fileContent     = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:cSld>'
    '<p:spTree>'
    '<p:nvGrpSpPr>'
    '<p:cNvPr id="1" name=""/>'
    '<p:cNvGrpSpPr/>'
    '<p:nvPr/>'
    '</p:nvGrpSpPr>'
    '<p:grpSpPr>'
    '<a:xfrm>'
    '<a:off x="0" y="0"/>'
    '<a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/>'
    '<a:chExt cx="0" cy="0"/>'
    '</a:xfrm>'
    '</p:grpSpPr>'
    '</p:spTree>'
    '</p:cSld>'
    '<p:clrMapOvr>'
    '<a:masterClrMapping/>'
    '</p:clrMapOvr>'
    '</p:sld>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'ppt','slides',fileName),fileContent);

% Create new XML relationships file
fileRelsPath    = cat(2,fileName,'.rels');
fileContent     = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '</Relationships>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'ppt','slides','_rels',fileRelsPath),fileContent) & retCode;

% Link new slide to presentation.xml
addNode(PPTXInfo.XML.Pres,'p:sldIdLst','p:sldId', ...
    {'id',int2str(PPTXInfo.lastSlideId), ...
    'r:id',cat(2,'rId',int2str(PPTXInfo.lastRId))});

% Link new slide filename to presentation.xml slide ID
addNode(PPTXInfo.XML.PresRel,'Relationships','Relationship', ...
    {'Id',cat(2,'rId',int2str(PPTXInfo.lastRId)), ...
    'Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', ...
    'Target',cat(2,'slides/',fileName)});

% Include new slide in the table of contents
addNode(PPTXInfo.XML.TOC,'Types','Override', ...
    {'PartName',cat(2,'/ppt/slides/',fileName), ...
    'ContentType','application/vnd.openxmlformats-officedocument.presentationml.slide+xml'});

% Load created files
PPTXInfo.XML.Slide      = xmlread(fullfile(PPTXInfo.tempName,'ppt','slides',fileName));
PPTXInfo.XML.SlideRel   = xmlread(fullfile(PPTXInfo.tempName,'ppt','slides','_rels',fileRelsPath));

% Assign data to slide
PPTXInfo.Slide(PPTXInfo.numSlides).id       = PPTXInfo.lastSlideId;
PPTXInfo.Slide(PPTXInfo.numSlides).rId      = PPTXInfo.lastRId;
PPTXInfo.Slide(PPTXInfo.numSlides).file     = fileName;
PPTXInfo.Slide(PPTXInfo.numSlides).objId    = 1;


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function PPTXInfo = addPicture(PPTXInfo,imgData,varargin)
% Adding image to existing slide
%   1. Add <p:pic> node to slide#.xml
%   2. Update slide#.xml.rels to link new pic to an image file
%

% Defaults
showLn  = false;
lnW     = 1;
lnCol   = [0 0 0];

% Set object ID
PPTXInfo.Slide(PPTXInfo.numSlides).objId    = PPTXInfo.Slide(PPTXInfo.numSlides).objId+1;
objId       = PPTXInfo.Slide(PPTXInfo.numSlides).objId;

% Get screen size (used in a few places)
screenSize      = get(0,'ScreenSize');
screenSize      = screenSize(1,3:4);
emusPerPx       = max(PPTXInfo.dimensions./screenSize);

% Based on the input format, decide what to do about it
if isnumeric(imgData) && numel(imgData)==1 && ...
        (ishghandle(imgData,'Figure') || ishghandle(imgData,'Axes'))
    % Either figure or axes handle
    img         = getframe(imgData);
    imageName   = sprintf('image-%d-%d.png',PPTXInfo.numSlides,objId);
    imagePath   = fullfile(PPTXInfo.tempName,'ppt','media',imageName);
    imwrite(img.cdata,imagePath);
    imdims      = size(img.cdata);
elseif isnumeric(imgData) && numel(imgData)>1,
    % Image CDATA
    imageName   = sprintf('image-%d-%d.png',PPTXInfo.numSlides,objId);
    imagePath   = fullfile(PPTXInfo.tempName,'ppt','media',imageName);
    imwrite(imgData,imagePath);
    imdims      = size(imgData);
elseif ischar(imgData),
    if exist(imgData,'file'),
        % Image filename
        [~,~,inImgExt]  = fileparts(imgData);
        inImgExt        = inImgExt(2:end);
        imageName       = sprintf('image-%d-%d.%s',PPTXInfo.numSlides,objId,inImgExt);
        imagePath       = fullfile(PPTXInfo.tempName,'ppt','media',imageName);
        if ~any(strcmpi(PPTXInfo.imageTypes,inImgExt)),
            % If XML file does not support this format yet, then add it
            PPTXInfo    = addImageTypeSupport(PPTXInfo,inImgExt);
        end
        copyfile(imgData,imagePath);
        if any(strcmpi({'emf','wmf','eps'},inImgExt)),
            % Vector images cannot be loaded by MatLab to determine their
            % native sizes, but they can be scaled to any size anyway
            imdims      = PPTXInfo.dimensions([2 1])./emusPerPx;
        else
            imgCdata    = imread(imgData);
            imdims      = size(imgCdata);
        end
    else
        error('exportToPPTX:fileNotFound','Image file requested to be added to the slide was not found');
    end
else
    % Error condition
    error('exportToPPTX:badInput','addPicture command requires a valid figure/axes handle or filename or CDATA');
end

% % Save image -- this code would be MUCH faster, but less supported: requires additional MEX file and uses unsupported hardcopy
% % Obtain a copy of fast PNG writing routine: http://www.mathworks.com/matlabcentral/fileexchange/40384
% imageName   = sprintf('image%d.png',PPTXInfo.numSlides);
% imagePath   = fullfile(PPTXInfo.tempName,'ppt','media',imageName);
% cdata       = hardcopy(figH,'-Dopengl','-r0');
% savepng(cdata,imagePath);

% Figure out picture size in PPTX units (EMUs)
imEMUs          = imdims([2 1]).*emusPerPx;

% Default picture position
picPosition     = round([(PPTXInfo.dimensions-imEMUs)./2 imEMUs]);

if nargin>2,
    if any(strcmpi(varargin,'Scale')),
        idx         = find(strcmpi(varargin,'Scale'));
        scaleOpt    = (varargin{idx+1});
        switch lower(scaleOpt),
            case 'noscale',
                % Do nothing, default option already defined
            case 'maxfixed',
                scaleSize       = min(PPTXInfo.dimensions./imEMUs);
                newImEMUs       = imEMUs.*scaleSize;
                picPosition     = round([(PPTXInfo.dimensions-newImEMUs)./2 newImEMUs]);
            case 'max',
                picPosition     = round([0 0 PPTXInfo.dimensions]);
            otherwise,
                error('exportToPPTX:badProperty','Bad property value found in Scale');
        end
    end
    
    if any(strcmpi(varargin,'Position')),
        idx         = find(strcmpi(varargin,'Position'));
        if ~isnumeric(varargin{idx+1}) || numel(varargin{idx+1})~=4,
            error('exportToPPTX:badProperty','Bad property value found in Position');
        end
        picPosition = round(varargin{idx+1}.*PPTXInfo.CONST.IN_TO_EMU);
    end
    
    if any(strcmpi(varargin,'LineWidth')),
        idx         = find(strcmpi(varargin,'LineWidth'));
        lnW         = varargin{idx+1};
        showLn      = true;
        if ~isnumeric(lnW) || numel(lnW)~=1,
            error('exportToPPTX:badProperty','Bad property value found in LineWidth');
        end
    end
    
    if any(strcmpi(varargin,'EdgeColor')),
        idx         = find(strcmpi(varargin,'EdgeColor'));
        lnCol       = varargin{idx+1};
        showLn      = true;
        if ~isnumeric(lnCol) || numel(lnCol)~=3,
            error('exportToPPTX:badProperty','Bad property value found in EdgeColor');
        end
    end
    
end

% Add image to slide XML file
picNode     = addNode(PPTXInfo.XML.Slide,'p:spTree','p:pic');
nvPicPr     = addNode(PPTXInfo.XML.Slide,picNode,'p:nvPicPr');
              addNode(PPTXInfo.XML.Slide,nvPicPr,'p:cNvPr',{'id',objId,'name','Picture 1','descr',imageName});
cNvPicPr    = addNode(PPTXInfo.XML.Slide,nvPicPr,'p:cNvPicPr');
              addNode(PPTXInfo.XML.Slide,cNvPicPr,'a:picLocks',{'noChangeAspect','1'});
              addNode(PPTXInfo.XML.Slide,nvPicPr,'p:nvPr');

blipFill    = addNode(PPTXInfo.XML.Slide,picNode,'p:blipFill');
              addNode(PPTXInfo.XML.Slide,blipFill,'a:blip',{'r:embed',sprintf('rId%d',objId)','cstate','print'});
stretch     = addNode(PPTXInfo.XML.Slide,blipFill,'a:stretch');
              addNode(PPTXInfo.XML.Slide,stretch,'a:fillRect');

spPr        = addNode(PPTXInfo.XML.Slide,picNode,'p:spPr');
axfm        = addNode(PPTXInfo.XML.Slide,spPr,'a:xfrm');
              addNode(PPTXInfo.XML.Slide,axfm,'a:off',{'x',picPosition(1),'y',picPosition(2)});
              addNode(PPTXInfo.XML.Slide,axfm,'a:ext',{'cx',picPosition(3),'cy',picPosition(4)});
prstGeom    = addNode(PPTXInfo.XML.Slide,spPr,'a:prstGeom',{'prst','rect'});
              addNode(PPTXInfo.XML.Slide,prstGeom,'a:avLst');

if showLn,
    aLn     = addNode(PPTXInfo.XML.Slide,spPr,'a:ln',{'w',lnW*PPTXInfo.CONST.PT_TO_EMU});
    sFil    = addNode(PPTXInfo.XML.Slide,aLn,'a:solidFill');
              addNode(PPTXInfo.XML.Slide,sFil,'a:srgbClr',{'val',sprintf('%s',dec2hex(round(lnCol*255),2).')});
end

% Add image reference to rels file
addNode(PPTXInfo.XML.SlideRel,'Relationships','Relationship', ...
    {'Id',sprintf('rId%d',objId), ...
    'Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', ...
    'Target',cat(2,'../media/',imageName)});


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function PPTXInfo = addTextbox(PPTXInfo,boxText,textPos,varargin)
% Adding image to existing slide
%   1. Add <p:sp> node to slide#.xml

% Defaults
hAlign  = 'l';
vAlign  = 't';
fSize   = 12;
fItal   = 0;
fBold   = 0;
fCol    = [0 0 0];
fRot    = 0;
bCol    = NaN;
showLn  = false;
lnW     = 1;
lnCol   = [0 0 0];

if nargin>3,
    if any(strcmpi(varargin,'HorizontalAlignment')),
        idx         = find(strcmpi(varargin,'HorizontalAlignment'));
        switch lower(varargin{idx+1}),
            case 'left',
                hAlign  = 'l';
            case 'right',
                hAlign  = 'r';
            case 'center',
                hAlign  = 'ctr';
            otherwise,
                error('exportToPPTX:badProperty','Bad property value found in HorizontalAlignment');
        end
    end
    
    if any(strcmpi(varargin,'VerticalAlignment')),
        idx         = find(strcmpi(varargin,'VerticalAlignment'));
        switch lower(varargin{idx+1}),
            case 'top',
                vAlign  = 't';
            case 'bottom',
                vAlign  = 'b';
            case 'middle',
                vAlign  = 'ctr';
            otherwise,
                error('exportToPPTX:badProperty','Bad property value found in VerticalAlignment');
        end
    end
    
    if any(strcmpi(varargin,'FontAngle')),
        idx         = find(strcmpi(varargin,'FontAngle'));
        switch lower(varargin{idx+1}),
            case {'italic','oblique'},
                fItal   = '1';
            case 'normal',
                fItal   = '0';
            otherwise,
                error('exportToPPTX:badProperty','Bad property value found in FontAngle');
        end
    end
    
    if any(strcmpi(varargin,'FontWeight')),
        idx         = find(strcmpi(varargin,'FontWeight'));
        switch lower(varargin{idx+1}),
            case {'bold','demi'},
                fBold   = '1';
            case {'normal','light'},
                fBold   = '0';
            otherwise,
                error('exportToPPTX:badProperty','Bad property value found in FontWeight');
        end
    end
    
    if any(strcmpi(varargin,'Color')),
        idx         = find(strcmpi(varargin,'Color'));
        fCol        = varargin{idx+1};
        if ~isnumeric(fCol) || numel(fCol)~=3,
            error('exportToPPTX:badProperty','Bad property value found in Color');
        end
    end
    
    if any(strcmpi(varargin,'BackgroundColor')),
        idx         = find(strcmpi(varargin,'BackgroundColor'));
        bCol        = varargin{idx+1};
        if ~isnumeric(bCol) || numel(bCol)~=3,
            error('exportToPPTX:badProperty','Bad property value found in BackgroundColor');
        end
    end
    
    if any(strcmpi(varargin,'FontSize')),
        idx         = find(strcmpi(varargin,'FontSize'));
        fSize       = varargin{idx+1};
        if ~isnumeric(fSize),
            error('exportToPPTX:badProperty','Bad property value found in FontSize');
        end
    end
    
    if any(strcmpi(varargin,'Rotation')),
        idx         = find(strcmpi(varargin,'Rotation'));
        fRot       = varargin{idx+1};
        if ~isnumeric(fRot) || numel(fRot)~=1,
            error('exportToPPTX:badProperty','Bad property value found in Rotation');
        end
    end
    
    if any(strcmpi(varargin,'LineWidth')),
        idx         = find(strcmpi(varargin,'LineWidth'));
        lnW         = varargin{idx+1};
        showLn      = true;
        if ~isnumeric(lnW) || numel(lnW)~=1,
            error('exportToPPTX:badProperty','Bad property value found in LineWidth');
        end
    end
    
    if any(strcmpi(varargin,'EdgeColor')),
        idx         = find(strcmpi(varargin,'EdgeColor'));
        lnCol       = varargin{idx+1};
        showLn      = true;
        if ~isnumeric(lnCol) || numel(lnCol)~=3,
            error('exportToPPTX:badProperty','Bad property value found in EdgeColor');
        end
    end
    
end

% Set object ID
PPTXInfo.Slide(PPTXInfo.numSlides).objId    = PPTXInfo.Slide(PPTXInfo.numSlides).objId+1;
objId       = PPTXInfo.Slide(PPTXInfo.numSlides).objId;

% Add image to slide XML file
spNode      = addNode(PPTXInfo.XML.Slide,'p:spTree','p:sp');
nvPicPr     = addNode(PPTXInfo.XML.Slide,spNode,'p:nvSpPr');
              addNode(PPTXInfo.XML.Slide,nvPicPr,'p:cNvPr',{'id',objId,'name','TextBox'});
              addNode(PPTXInfo.XML.Slide,nvPicPr,'p:cNvSpPr',{'txBox','1'});
              addNode(PPTXInfo.XML.Slide,nvPicPr,'p:nvPr');

spPr        = addNode(PPTXInfo.XML.Slide,spNode,'p:spPr');
axfm        = addNode(PPTXInfo.XML.Slide,spPr,'a:xfrm',{'rot',-fRot*PPTXInfo.CONST.DEG_TO_EMU});
              addNode(PPTXInfo.XML.Slide,axfm,'a:off',{'x',textPos(1),'y',textPos(2)});
              addNode(PPTXInfo.XML.Slide,axfm,'a:ext',{'cx',textPos(3),'cy',textPos(4)});
prstGeom    = addNode(PPTXInfo.XML.Slide,spPr,'a:prstGeom',{'prst','rect'});
              addNode(PPTXInfo.XML.Slide,prstGeom,'a:avLst');

if ~isnan(bCol),
    solidFill=addNode(PPTXInfo.XML.Slide,spPr,'a:solidFill');
              addNode(PPTXInfo.XML.Slide,solidFill,'a:srgbClr',{'val',sprintf('%s',dec2hex(round(bCol*255),2).')});
else
              addNode(PPTXInfo.XML.Slide,spPr,'a:noFill');
end

if showLn,
    aLn     = addNode(PPTXInfo.XML.Slide,spPr,'a:ln',{'w',lnW*PPTXInfo.CONST.PT_TO_EMU});
    sFil    = addNode(PPTXInfo.XML.Slide,aLn,'a:solidFill');
              addNode(PPTXInfo.XML.Slide,sFil,'a:srgbClr',{'val',sprintf('%s',dec2hex(round(lnCol*255),2).')});
end
              
txBody      = addNode(PPTXInfo.XML.Slide,spNode,'p:txBody');
bodyPr      = addNode(PPTXInfo.XML.Slide,txBody,'a:bodyPr',{'wrap','square','rtlCol','0','anchor',vAlign});
              addNode(PPTXInfo.XML.Slide,bodyPr,'a:normAutofit'); % autofit flag
              addNode(PPTXInfo.XML.Slide,txBody,'a:lstStyle');
ap          = addNode(PPTXInfo.XML.Slide,txBody,'a:p');
              addNode(PPTXInfo.XML.Slide,ap,'a:pPr',{'algn',hAlign});
ar          = addNode(PPTXInfo.XML.Slide,ap,'a:r');
rPr         = addNode(PPTXInfo.XML.Slide,ar,'a:rPr',{'lang','en-US','dirty','0','smtClean','0','i',fItal,'b',fBold,'sz',fSize*PPTXInfo.CONST.FONT_PX_TO_PPTX});
sf          = addNode(PPTXInfo.XML.Slide,rPr,'a:solidFill');
              addNode(PPTXInfo.XML.Slide,sf,'a:srgbClr',{'val',sprintf('%s',dec2hex(round(fCol*255),2).')});
at          = addNode(PPTXInfo.XML.Slide,ar,'a:t');
              addNodeValue(PPTXInfo.XML.Slide,at,boxText);
erPr        = addNode(PPTXInfo.XML.Slide,ap,'a:endParaRPr',{'lang','en-US','dirty','0','i',fItal,'b',fBold,'sz',fSize*PPTXInfo.CONST.FONT_PX_TO_PPTX});
esf         = addNode(PPTXInfo.XML.Slide,erPr,'a:solidFill');
              addNode(PPTXInfo.XML.Slide,esf,'a:srgbClr',{'val',sprintf('%s',dec2hex(round(fCol*255),2).')});


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function PPTXInfo = addImageTypeSupport(PPTXInfo,imgExt)
addNode(PPTXInfo.XML.TOC,'Types','Default',{'Extension',imgExt,'ContentType',cat(2,'image/',imgExt)});

        
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function foundNode = findNode(theNode,searchNodeName)
foundNodes  = theNode.getElementsByTagName(searchNodeName);
if foundNodes.getLength>1,
    foundNode   = foundNodes;
else
    foundNode   = foundNodes.item(0);
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function nodeValue = getNodeValue(xmlData,searchNodeName)
nodeValue = {};
foundNode = findNode(xmlData,searchNodeName);
if ~isempty(foundNode),
    if foundNode.getLength>1,
        for inode=0:foundNode.getLength-1,
            if ~isempty(foundNode.item(inode).getFirstChild),
                nodeValue{inode+1} = char(foundNode.item(inode).getFirstChild.getNodeValue);
            end
        end
    else
        if isempty(foundNode.getFirstChild),
            nodeValue{1} = '';
        else
            nodeValue{1} = char(foundNode.getFirstChild.getNodeValue);
        end
    end
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function setNodeValue(xmlData,searchNodeName,newValue)
if ~isempty(newValue),
    % two possible inputs: tag name or node handle
    if ischar(searchNodeName),
        foundNode = findNode(xmlData,searchNodeName);
    else
        foundNode = searchNodeName;
    end
    if ~isempty(foundNode),
        if isnumeric(newValue),
            newValue2   = num2str(newValue);
        else
            newValue2   = newValue;
        end
        foundNode.getFirstChild.setNodeValue(newValue2);
    end
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function addNodeValue(xmlData,searchNodeName,newValue)
if ~isempty(newValue),
    % two possible inputs: tag name or node handle
    if ischar(searchNodeName),
        foundNode = findNode(xmlData,searchNodeName);
    else
        foundNode = searchNodeName;
    end
    if ~isempty(foundNode),
        if isnumeric(newValue),
            newValue2   = num2str(newValue);
        else
            newValue2   = newValue;
        end
        foundNode.appendChild(xmlData.createTextNode(newValue2));
    end
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function nodeAttribute = getNodeAttribute(xmlData,searchNodeName,searchAttribute)
nodeAttribute   = {};
foundNode = findNode(xmlData,searchNodeName);
if ~isempty(foundNode),
    if foundNode.getLength>0,
        for inode=0:foundNode.getLength-1,
            nodeAttribute{inode+1} = char(foundNode.item(inode).getAttribute(searchAttribute));
        end
    else
        nodeAttribute{1} = char(foundNode.getAttribute(searchAttribute));
    end
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function setNodeAttribute(xmlData,searchNodeName,searchAttribute,newValue)
foundNode = findNode(xmlData,searchNodeName);
if ~isempty(foundNode),
    foundNode.setAttribute(searchAttribute,newValue);
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function newNode = addNode(xmlData,parentNode,childNode,allAttribs)
% two possible inputs: tag name or node handle
if ischar(parentNode),
    foundNode = findNode(xmlData,parentNode);
else
    foundNode = parentNode;
end

if ~isempty(foundNode),
    newNode = xmlData.createElement(childNode);
    foundNode.appendChild(newNode);
    % Set attributes
    if exist('allAttribs','var') && ~isempty(allAttribs),
        for iatr=1:2:length(allAttribs),
            if isnumeric(allAttribs{iatr+1}),
                attribValue     = num2str(allAttribs{iatr+1});
            else
                attribValue     = allAttribs{iatr+1};
            end
            newNode.setAttribute(allAttribs{iatr},attribValue);
        end
    end
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function initBlankPPTX(PPTXInfo)
% Creates valid PPTX structure with 0 slides

% Folder structure
mkdir(fullfile(PPTXInfo.tempName,'docProps'));
mkdir(fullfile(PPTXInfo.tempName,'ppt'));
mkdir(fullfile(PPTXInfo.tempName,'_rels'));
mkdir(fullfile(PPTXInfo.tempName,'ppt','media'));
mkdir(fullfile(PPTXInfo.tempName,'ppt','slideLayouts'));
mkdir(fullfile(PPTXInfo.tempName,'ppt','slideMasters'));
mkdir(fullfile(PPTXInfo.tempName,'ppt','slides'));
mkdir(fullfile(PPTXInfo.tempName,'ppt','theme'));
mkdir(fullfile(PPTXInfo.tempName,'ppt','_rels'));
mkdir(fullfile(PPTXInfo.tempName,'ppt','slideLayouts','_rels'));
mkdir(fullfile(PPTXInfo.tempName,'ppt','slideMasters','_rels'));
mkdir(fullfile(PPTXInfo.tempName,'ppt','slides','_rels'));

% Absolutely neccessary files
% \[Content_Types].xml
% \_rels\.rels
% \docProps\app.xml
% \docProps\core.xml
% \ppt\presentation.xml
% \ppt\_rels\presentation.xml.rels
% \ppt\slideLayouts\slideLayout1.xml
% \ppt\slideLayouts\_rels\slideLayout1.xml.rels
% \ppt\slideMasters\slideMaster1.xml
% \ppt\slideMasters\_rels\slideMaster1.xml.rels
% \ppt\theme\theme1.xml

retCode     = 1;

% \[Content_Types].xml
fileContent    = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
    '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
    '</Types>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'[Content_Types].xml'),fileContent) & retCode;

% \_rels\.rels
fileContent    = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
    '</Relationships>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'_rels','.rels'),fileContent) & retCode;

% \docProps\app.xml
fileContent    = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '<Template/>'
    '<TotalTime>0</TotalTime>'
    '<Words>0</Words>'
    '<Application>Microsoft Office PowerPoint</Application>'
    '<PresentationFormat>On-screen Show (4:3)</PresentationFormat>'
    '<Paragraphs>0</Paragraphs>'
    '<Slides>0</Slides>'
    '<Notes>0</Notes>'
    '<HiddenSlides>0</HiddenSlides>'
    '<MMClips>0</MMClips>'
    '<ScaleCrop>false</ScaleCrop>'
    '<HeadingPairs>'
    '<vt:vector size="4" baseType="variant">'
    '<vt:variant>'
    '<vt:lpstr>Theme</vt:lpstr>'
    '</vt:variant>'
    '<vt:variant>'
    '<vt:i4>1</vt:i4>'
    '</vt:variant>'
    '<vt:variant>'
    '<vt:lpstr>Slide Titles</vt:lpstr>'
    '</vt:variant>'
    '<vt:variant>'
    '<vt:i4>1</vt:i4>'
    '</vt:variant>'
    '</vt:vector>'
    '</HeadingPairs>'
    '<TitlesOfParts>'
    '<vt:vector size="2" baseType="lpstr">'
    '<vt:lpstr>Default Theme</vt:lpstr>'
    '<vt:lpstr></vt:lpstr>'
    '</vt:vector>'
    '</TitlesOfParts>'
    '<Company></Company>'
    '<LinksUpToDate>false</LinksUpToDate>'
    '<SharedDoc>false</SharedDoc>'
    '<HyperlinksChanged>false</HyperlinksChanged>'
    '<AppVersion>12.0000</AppVersion>'
    '</Properties>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'docProps','app.xml'),fileContent) & retCode;

% \docProps\core.xml
fileContent    = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '<dc:title>Blank</dc:title>'
    '<dc:subject>Blank</dc:subject>'
    '<dc:creator>MatLab</dc:creator>'                % To be filled out
    '<cp:keywords></cp:keywords>'
    '<dc:description></dc:description>'
    '<cp:lastModifiedBy>Blank</cp:lastModifiedBy>'   % To be filled out
    '<cp:revision>0</cp:revision>'
    ['<dcterms:created xsi:type="dcterms:W3CDTF">' PPTXInfo.createdDate '</dcterms:created>']     % To be filled out
    ['<dcterms:modified xsi:type="dcterms:W3CDTF">' PPTXInfo.createdDate '</dcterms:modified>']   % To be filled out
    '</cp:coreProperties>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'docProps','core.xml'),fileContent) & retCode;

% \ppt\presentation.xml
fileContent    = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">'
    '<p:sldMasterIdLst>'
    '<p:sldMasterId id="2147483663" r:id="rId2"/>'
    '</p:sldMasterIdLst>'
    '<p:sldIdLst>'
    '</p:sldIdLst>'
    ['<p:sldSz cx="' int2str(PPTXInfo.dimensions(1)) '" cy="' int2str(PPTXInfo.dimensions(2)) '" type="screen4x3"/>']
    ['<p:notesSz cx="' int2str(PPTXInfo.dimensions(1)) '" cy="' int2str(PPTXInfo.dimensions(2)) '"/>']
    '</p:presentation>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'ppt','presentation.xml'),fileContent) & retCode;

% \ppt\_rels\presentation.xml.rels
fileContent    = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
    '</Relationships>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'ppt','_rels','presentation.xml.rels'),fileContent) & retCode;

% \ppt\slideLayouts\slideLayout1.xml
fileContent    = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1" userDrawn="1">'
    '<p:cSld name="Title Slide">'
    '<p:spTree>'
    '<p:nvGrpSpPr>'
    '<p:cNvPr id="1" name=""/>'
    '<p:cNvGrpSpPr/>'
    '<p:nvPr/>'
    '</p:nvGrpSpPr>'
    '<p:grpSpPr>'
    '<a:xfrm>'
    '<a:off x="0" y="0"/>'
    '<a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/>'
    '<a:chExt cx="0" cy="0"/>'
    '</a:xfrm>'
    '</p:grpSpPr>'
    '</p:spTree>'
    '</p:cSld>'
    '<p:clrMapOvr>'
    '<a:masterClrMapping/>'
    '</p:clrMapOvr>'
    '<p:hf hdr="0" ftr="0"/>'
    '</p:sldLayout>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'ppt','slideLayouts','slideLayout1.xml'),fileContent) & retCode;

% \ppt\slideLayouts\_rels\slideLayout1.xml.rels
fileContent    = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
    '</Relationships>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'ppt','slideLayouts','_rels','slideLayout1.xml.rels'),fileContent) & retCode;

% \ppt\slideMasters\slideMaster1.xml
fileContent    = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:cSld>'
    '<p:bg>'
    '<p:bgRef idx="1001">'
    '<a:schemeClr val="bg1"/>'
    '</p:bgRef>'
    '</p:bg>'
    '<p:spTree>'
    '<p:nvGrpSpPr>'
    '<p:cNvPr id="1" name=""/>'
    '<p:cNvGrpSpPr/>'
    '<p:nvPr/>'
    '</p:nvGrpSpPr>'
    '<p:grpSpPr>'
    '<a:xfrm>'
    '<a:off x="0" y="0"/>'
    '<a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/>'
    '<a:chExt cx="0" cy="0"/>'
    '</a:xfrm>'
    '</p:grpSpPr>'
    '</p:spTree>'
    '</p:cSld>'
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    '<p:sldLayoutIdLst>'
    '<p:sldLayoutId id="2147483664" r:id="rId1"/>'
    '</p:sldLayoutIdLst>'
    '<p:hf hdr="0" ftr="0"/>'
    '</p:sldMaster>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'ppt','slideMasters','slideMaster1.xml'),fileContent) & retCode;

% \ppt\slideMasters\_rels\slideMaster1.xml.rels
fileContent    = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '</Relationships>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'ppt','slideMasters','_rels','slideMaster1.xml.rels'),fileContent) & retCode;

% \ppt\theme\theme1.xml
fileContent    = { ...
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
    '<a:themeElements>'
    '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>'
    '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
    '<a:fmtScheme name="Office">'
    '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"> <a:schemeClr val="phClr"> <a:tint val="37000"/> <a:satMod val="300000"/> </a:schemeClr> </a:gs> <a:gs pos="100000"> <a:schemeClr val="phClr"> <a:tint val="15000"/> <a:satMod val="350000"/> </a:schemeClr> </a:gs> </a:gsLst> <a:lin ang="16200000" scaled="1"/> </a:gradFill> <a:gradFill rotWithShape="1"> <a:gsLst> <a:gs pos="0"> <a:schemeClr val="phClr"> <a:shade val="51000"/> <a:satMod val="130000"/> </a:schemeClr> </a:gs> <a:gs pos="80000"> <a:schemeClr val="phClr"> <a:shade val="93000"/> <a:satMod val="130000"/> </a:schemeClr> </a:gs> <a:gs pos="100000"> <a:schemeClr val="phClr"> <a:shade val="94000"/> <a:satMod val="135000"/> </a:schemeClr> </a:gs> </a:gsLst> <a:lin ang="16200000" scaled="0"/> </a:gradFill></a:fillStyleLst>'
    '<a:lnStyleLst> <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"> <a:solidFill> <a:schemeClr val="phClr"> <a:shade val="95000"/> <a:satMod val="105000"/> </a:schemeClr> </a:solidFill> <a:prstDash val="solid"/> </a:ln> <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"> <a:solidFill> <a:schemeClr val="phClr"/> </a:solidFill> <a:prstDash val="solid"/> </a:ln> <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"> <a:solidFill> <a:schemeClr val="phClr"/> </a:solidFill> <a:prstDash val="solid"/> </a:ln> </a:lnStyleLst>'
    '<a:effectStyleLst> <a:effectStyle> <a:effectLst> <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"> <a:srgbClr val="000000"> <a:alpha val="38000"/> </a:srgbClr> </a:outerShdw> </a:effectLst> </a:effectStyle> <a:effectStyle> <a:effectLst> <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"> <a:srgbClr val="000000"> <a:alpha val="35000"/> </a:srgbClr> </a:outerShdw> </a:effectLst> </a:effectStyle> <a:effectStyle> <a:effectLst> <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"> <a:srgbClr val="000000"> <a:alpha val="35000"/> </a:srgbClr> </a:outerShdw> </a:effectLst> <a:scene3d> <a:camera prst="orthographicFront"> <a:rot lat="0" lon="0" rev="0"/> </a:camera> <a:lightRig rig="threePt" dir="t"> <a:rot lat="0" lon="0" rev="1200000"/> </a:lightRig> </a:scene3d> <a:sp3d> <a:bevelT w="63500" h="25400"/> </a:sp3d> </a:effectStyle> </a:effectStyleLst>'
    '<a:bgFillStyleLst> <a:solidFill> <a:schemeClr val="phClr"/> </a:solidFill> <a:gradFill rotWithShape="1"> <a:gsLst> <a:gs pos="0"> <a:schemeClr val="phClr"> <a:tint val="40000"/> <a:satMod val="350000"/> </a:schemeClr> </a:gs> <a:gs pos="40000"> <a:schemeClr val="phClr"> <a:tint val="45000"/> <a:shade val="99000"/> <a:satMod val="350000"/> </a:schemeClr> </a:gs> <a:gs pos="100000"> <a:schemeClr val="phClr"> <a:shade val="20000"/> <a:satMod val="255000"/> </a:schemeClr> </a:gs> </a:gsLst> <a:path path="circle"> <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/> </a:path> </a:gradFill> <a:gradFill rotWithShape="1"> <a:gsLst> <a:gs pos="0"> <a:schemeClr val="phClr"> <a:tint val="80000"/> <a:satMod val="300000"/> </a:schemeClr> </a:gs> <a:gs pos="100000"> <a:schemeClr val="phClr"> <a:shade val="30000"/> <a:satMod val="200000"/> </a:schemeClr> </a:gs> </a:gsLst> <a:path path="circle"> <a:fillToRect l="50000" t="50000" r="50000" b="50000"/> </a:path> </a:gradFill> </a:bgFillStyleLst>'
    '</a:fmtScheme>'
    '</a:themeElements>'
    '<a:objectDefaults/>'
    '<a:extraClrSchemeLst/>'
    '</a:theme>'};
retCode     = writeTextFile(fullfile(PPTXInfo.tempName,'ppt','theme','theme1.xml'),fileContent) & retCode;

if ~retCode,
	error('Error creating temporary PPTX files');
end


