classdef exportToPPTX < handle
    %  exportToPPTX Creates PowerPoint 2007+ (PPTX) slides
    %
    %       exportToPPTX allows user to create PowerPoint 2007+ (PPTX) files
    %       without using COM-objects automation. Proper XML files are created and
    %       packed into PPTX file that can be read and displayed by PowerPoint.
    %       
    %       exportToPPTX methods:
    %           exportToPPTX    - Starts new presentation or opens an existing presentation
    %           save            - Saves current presentation
    %           addSlide        - Adds a slide to the presentation
    %           switchSlide     - Switches current slide to a given slide ID
    %           addPicture      - Adds picture to the current slide
    %           addShape        - Adds lines or closed shapes to the current slide
    %           addNote         - Adds notes information to the current slide
    %           addTextbox      - Adds textbox to the current slide
    %           addTable        - Adds PowerPoint table to the current slide
    %
    %       exportToPPTX properties:
    %           author          - Presentation's author, default is 'exportToPPTX'
    %           title           - Presentation's title, default is 'Blank'
    %           subject         - Presentation's subject, default is blank
    %           description     - Presentation's description, default is blank
    %
    %       Basic usage example:
    %
    %           % Start new presentation
    %           pptx        = exportToPPTX();
    %
    %           % Set presentation title
    %           pptx.title  = 'Basic example'
    %       
    %           % Just an example image
    %           load mandrill; figure('color','w'); image(X); colormap(map); axis off; axis image;
    %
    %           % Add slide, then add image to it, then add box
    %           pptx.addSlide();
    %           pptx.addPicture(gcf,'Scale','maxfixed');
    %           pptx.addTextbox('Mandrill','Position',[0 5 6 1],'FontWeight','bold','HorizontalAlignment','center','VerticalAlignment','bottom');
    %
    %           % Save
    %           pptx.save('example.pptx');
    %
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%   PROPERTIES
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

    %% Define constants
    properties (Constant,Access=private)
        CONST_IN_TO_EMU           = 914400;
        CONST_PT_TO_EMU           = 12700;
        CONST_DEG_TO_EMU          = 60000;
        CONST_FONT_PX_TO_PPTX     = 100;
    end
    
    %% Define run-time public (settable) properties
    properties (Access=public)
        % author    Presentation's author, default is 'exportToPPTX'
        %
        %   Example:
        %       pptx.author = 'exportToPPTX Example';
        author          = 'exportToPPTX';   
        
        % title     Presentation's title, default is 'Blank'
        % 
        %   Example:
        %       pptx.title = 'Demonstration Presentation';
        title           = 'Blank';   
        
        % subject   Presentation's subject, default is blank
        %
        %   Example:
        %       pptx.subject = 'Demonstration of various exportToPPTX commands';
        subject         = '';  
        
        % description   Presentation's description, default is blank
        %
        %   Example:
        %       pptx.description = 'Description goes in here';
        description     = '';
    end
    
    %% Define run-time gettable properties (not settable)
    properties (GetAccess=public,SetAccess=private)
        dimensions      = [10 7.5];     % Presentation's dimensions in inches (read only)
        
        fullName        = '';           % Full filename of the openned presentation (empty if new)
        numSlides       = 0;            % Total number of slides
        currentSlide                    % Current slide
    end

    %% Define run-time private properties (for internal use)
    properties (Access=private)
        revNumber       = 0;
        numMasters
        
        tempName
        createdDate
        updatedDate
        imageTypes
        videoTypes
        bgColor
        lastSlideId
        lastRId
        themeNum        = 0;
        Slide
        SlideMaster
        NotesMaster
    end
    
    %% Define internal XML structures that are kept in memory for modifications
    properties (Access=private)
        XML
    end
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%   METHODS
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %% Public methods
    methods
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function PPTX = exportToPPTX(fileName,varargin)
            % exportToPPTX([fileName],...)
            %   
            %   Start new presentation or open an existing presentation
            %   Actual PowerPoint files are not written until 'save' command is called. 
            %   No required inputs. This command does not return any values.
            %
            %   Additional parameters:
            %       Dimensions  two element vector specifying presentation's width and 
            %                   height in inches. Default size is 10 x 7.5 in.
            %       Author      specify presentation's author. Default is "exportToPPTX".
            %       Title       specify presentation's title. Default is "Blank".
            %       Subject     specify presentation's subject line. Default is empty (blank).
            %       Comments    specify presentation's comments. Default is empty (blank).
            %       BackgroundColor     Three element vector specifying RGB value in the 
            %                           range from 0 to 1. By default background is white.
            %   
            %   Examples:
            %       % Start new presentation
            %       pptx    = exportToPPTX();
            %
            %       % Open existing presentation
            %       pptx    = exportToPPTX('\\path\to\some\existing\presentation.pptx');

            %% Obtain temp folder name
            tempName    = tempname;
            while exist(tempName,'dir'),
                tempName    = tempname;
            end
            
            %% Create temp folder to hold all PPTX files
            mkdir(tempName);
            PPTX.tempName   = tempName;
            
            if nargin>0 && ~isempty(fileName)
                % Open file
                
                %% Check input validity
                [filePath,fileName,fileExt]     = fileparts(fileName);
                if isempty(filePath),
                    filePath        = pwd;
                end
                if ~strncmpi(fileExt,'.pptx',5),
                    fileExt         = cat(2,fileExt,'.pptx');
                end
                fullName            = fullfile(filePath,cat(2,fileName,fileExt));
                PPTX.fullName       = fullName;
                
                if exist(PPTX.fullName,'file'),
                    PPTX.openExistingPPTX();
                    
                    %% Load important XML files into memory
                    PPTX.loadAndParseXMLFiles();
                else
                    error('exportToPPTX:fileNotFound','PPTX to open (%s) not found.',PPTX.fullName);
                end
            else
                % New file
                
                mi                      = false(size(varargin));
                [PPTX.dimensions,mi]    = exportToPPTX.getPVPair(varargin,'Dimensions',PPTX.dimensions,mi);
                [PPTX.author,mi]        = exportToPPTX.getPVPair(varargin,'Author',PPTX.author,mi);
                [PPTX.title,mi]         = exportToPPTX.getPVPair(varargin,'Title',PPTX.title,mi);
                [PPTX.subject,mi]       = exportToPPTX.getPVPair(varargin,'Subject',PPTX.subject,mi);
                [PPTX.description,mi]   = exportToPPTX.getPVPair(varargin,'Comments',PPTX.description,mi);
                [PPTX.bgColor,mi]       = exportToPPTX.getPVPair(varargin,'BackgroundColor',[],mi);
                
                if any(~mi)
                    error('exportToPPTX:badProperty','Unrecognized property %s',varargin{find(~mi,1)});
                end
                
                if numel(PPTX.dimensions)~=2,
                    error('exportToPPTX:badDimensions','Slide dimensions vector must have two values only: width x height');
                end
                
                if ~isempty(PPTX.bgColor),
                    if ~isnumeric(PPTX.bgColor) || numel(PPTX.bgColor)~=3,
                        error('exportToPPTX:badProperty','Bad property value found in BackgroundColor');
                    end
                end
                
                %% Create new (empty) PPTX files structure
                PPTX.fullName       = [];
                PPTX.createdDate    = datestr(now,'yyyy-mm-ddTHH:MM:SS');
                
                PPTX.initBlankPPTX();
                PPTX.loadAndParseXMLFiles();
            end
        end
        
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function delete(PPTX)
            % Destructor: clears out temporary storage

            % TODO: check if recently saved before destroying changes
            % Remove temporary directory and all its contents
            rmdir(PPTX.tempName,'s');
        end
        
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function disp(PPTX)
            % Command window display
            
            showFileName    = PPTX.fullName;
            if isempty(showFileName),
                showFileName    = '<new>';
            end
            fprintf('\tFile:          %s\n',showFileName);
            fprintf('\tTotal slides:  %d\n',PPTX.numSlides);
            fprintf('\tCurrent slide: %d\n',PPTX.currentSlide);
            fprintf('\tAuthor:        %s\n',PPTX.author);
            fprintf('\tTitle:         %s\n',PPTX.title);
            fprintf('\tSubject:       %s\n',PPTX.subject);
            fprintf('\tDescription:   %s\n',PPTX.description);
            fprintf('\tDimensions:    %.2f x %.2f in\n',PPTX.dimensions);
            fprintf('\n');
        
            if ~isempty(PPTX.SlideMaster),
                for imast=1:numel(PPTX.SlideMaster),
                    fprintf('\tMaster #%d: %s\n',imast,PPTX.SlideMaster(imast).name);
                    if ~isempty(PPTX.SlideMaster(imast).Layout),
                        for ilay=1:numel(PPTX.SlideMaster(imast).Layout),
                            placeHolders    = sprintf('%s, ',PPTX.SlideMaster(imast).Layout(ilay).place{:});
                            if ~isempty(placeHolders),
                                placeHolders    = sprintf('(%s)',placeHolders(1:end-2));
                            end
                            fprintf('\t\tLayout #%d: %s %s\n',ilay,PPTX.SlideMaster(imast).Layout(ilay).name,placeHolders);
                        end
                    else
                        fprintf('\t\tNo layouts defined.\n');
                    end
                end
            else
                fprintf('\tNo master layouts defined.\n');
            end
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function currentSlideId = addSlide(PPTX,varargin)
            % addSlide(...)
            %
            %   Adds a slide to the presentation. No additional inputs required. Returns 
            %   newly created slide ID (sequential slide number signifying total slides in 
            %   the deck, not neccessarily slide order).
            % 
            %   Additional parameters:
            %       Position    Specify position at which to insert new slide. The value 
            %                   must be between 1 and the total number of slides.
            %       BackgroundColor     Three element vector specifying RGB value in the 
            %                           range from 0 to 1. By default background is white.
            %       Master      Master layout ID or name. By default first master layout is used.
            %       Layout      Slide template layout ID or name. By default first slide 
            %                   template layout is used.
            %
            %   Examples:
            %       % Add new slide
            %       pptx.addSlide();
            %
            %       % Add new slide with green background at position 1
            %       pptx.addSlide('BackgroundColor','g','Position',1);
            
            % Adding a new slide to PPTX
            %   1. Create slide#.xml file
            %   2. Create slide#.xml.rels file
            %   3. Update presentation.xml to include new slide
            %   4. Update presentation.xml.rels to link new slide to presentation.xml
            %   5. Update [Content_Types].xml
            
            % Parse optional inputs
            mi              = false(size(varargin));
            [bgCol,mi]      = exportToPPTX.getPVPair(varargin,'BackgroundColor',[],mi);
            [insPos,mi]     = exportToPPTX.getPVPair(varargin,'Position',[],mi);
            [masterID,mi]   = exportToPPTX.getPVPair(varargin,'Master',1,mi);
            [layoutID,mi]   = exportToPPTX.getPVPair(varargin,'Layout',1,mi);
            
            if any(~mi)
                error('exportToPPTX:badProperty','Unrecognized property %s',varargin{find(~mi,1)});
            end
            
            bgCol       = exportToPPTX.validateColor(bgCol);
            if ~isempty(bgCol),
                bgContent   = { ...
                    '<p:bg>'
                    '<p:bgPr>'
                    '<a:solidFill>'
                    ['<a:srgbClr val="' bgCol '"/>']
                    '</a:solidFill>'
                    '<a:effectLst/>'
                    '</p:bgPr>'
                    '</p:bg>'
                    };
            else
                bgContent   = {};
            end

            if ~isempty(insPos) && (insPos<1 || insPos>PPTX.numSlides+1 || numel(insPos)>1),
                % Error condition
                error('exportToPPTX:badProperty','addSlide position must be between 1 and the total number of slides');
            end
            
            % Parse master number and layout number
            % Use first master slide by default
            if isnumeric(masterID),
                masterNum       = masterID;
            elseif ischar(masterID),
                masterNum       = find(strncmpi({PPTX.SlideMaster.name},masterID,length(masterID)));
                if isempty(masterNum),
                    warning('exportToPPTX:badName','Master layout "%s" does not exist in the current presentation',masterID);
                    masterNum   = 1;
                end
                if numel(masterNum)>1,
                    warning('exportToPPTX:badName','There are multiple matches for master layout "%s"',masterID);
                    masterNum   = masterNum(1);
                end
            end
            if ~isempty(masterNum) && (masterNum<1 || masterNum>PPTX.numMasters),
                error('exportToPPTX:badProperty','addSlide cannot proceed because requested slide master number does not exist');
            end
            
            % Use first layout slide by default
            if isnumeric(layoutID),
                layoutNum       = layoutID;
            elseif ischar(layoutID)
                layoutNum       = find(strcmp({PPTX.SlideMaster(masterNum).Layout.name},layoutID));
                if isempty(layoutNum)
                    layoutNum   = find(strncmpi({PPTX.SlideMaster(masterNum).Layout.name},layoutID,length(layoutID)));
                end
                if isempty(layoutNum)
                    warning('exportToPPTX:badName','Layout "%s" does not exist in the current master',layoutID);
                    layoutNum   = 1;
                end
                if numel(layoutNum)>1
                    warning('exportToPPTX:badName','There are multiple matches for layout "%s"',layoutID);
                    layoutNum   = layoutNum(1);
                end
            end
            if ~isempty(layoutNum) && (layoutNum<1 || layoutNum>PPTX.SlideMaster(masterNum).numLayout),
                error('exportToPPTX:badProperty','addSlide cannot proceed because requested slide layout number does not exist');
            end

            % Check if slides folder exists
            if ~exist(fullfile(PPTX.tempName,'ppt','slides'),'dir'),
                mkdir(fullfile(PPTX.tempName,'ppt','slides'));
                mkdir(fullfile(PPTX.tempName,'ppt','slides','_rels'));
            end
            
            % Before creating new slide, is there a current slide that needs to be
            % saved to XML file?
            if isfield(PPTX.XML,'Slide') && ~isempty(PPTX.XML.Slide),
                fileName    = PPTX.Slide(PPTX.currentSlide).file;
                xmlwrite(fullfile(PPTX.tempName,'ppt','slides',fileName),PPTX.XML.Slide);
                xmlwrite(fullfile(PPTX.tempName,'ppt','slides','_rels',cat(2,fileName,'.rels')),PPTX.XML.SlideRel);
            end
            
            % Update useful variables
            PPTX.numSlides      = PPTX.numSlides+1;
            PPTX.lastSlideId    = PPTX.lastSlideId+1;
            PPTX.lastRId        = PPTX.lastRId+1;
            
            % Create new XML file
            % file name and path are relative to presentation.xml file
            fileName        = sprintf('slide%d.xml',PPTX.numSlides);
            fileContent     = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
                '<p:cSld>'
                [bgContent{:}]
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
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','slides',fileName),fileContent);
            
            % Create new XML relationships file
            fileRelsPath    = cat(2,fileName,'.rels');
            fileContent     = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/' PPTX.SlideMaster(masterNum).Layout(layoutNum).file '"/>']
                '</Relationships>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','slides','_rels',fileRelsPath),fileContent) & retCode;
            
            
            % Link new slide to presentation.xml
            % Check if list of slides node exists
            sldIdLstNode = exportToPPTX.findNode(PPTX.XML.Pres,'p:sldIdLst');
            if isempty(sldIdLstNode),
                % Note: order of tags within XML structure is important, slide list
                % must show up after p:sldMasterIdLst and before p:sldSz
                exportToPPTX.addNodeBefore(PPTX.XML.Pres,'p:presentation','p:sldIdLst','p:sldSz');
            end
            
            % Order of items in the sldIdLst node determines the order of the slides
            if ~isempty(insPos),
                exportToPPTX.addNodeAtPosition(PPTX.XML.Pres,'p:sldIdLst','p:sldId',insPos, ...
                    {'id',int2str(PPTX.lastSlideId), ...
                    'r:id',sprintf('rId%d',PPTX.lastRId)});
            else
                exportToPPTX.addNode(PPTX.XML.Pres,'p:sldIdLst','p:sldId', ...
                    {'id',int2str(PPTX.lastSlideId), ...
                    'r:id',sprintf('rId%d',PPTX.lastRId)});
            end
            
            
            % Link new slide filename to presentation.xml slide ID
            exportToPPTX.addNode(PPTX.XML.PresRel,'Relationships','Relationship', ...
                {'Id',cat(2,'rId',int2str(PPTX.lastRId)), ...
                'Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', ...
                'Target',cat(2,'slides/',fileName)});
            
            % Include new slide in the table of contents
            exportToPPTX.addNode(PPTX.XML.TOC,'Types','Override', ...
                {'PartName',cat(2,'/ppt/slides/',fileName), ...
                'ContentType','application/vnd.openxmlformats-officedocument.presentationml.slide+xml'});
            
            % Load created files
            PPTX.XML.Slide      = xmlread(fullfile(PPTX.tempName,'ppt','slides',fileName));
            PPTX.XML.SlideRel   = xmlread(fullfile(PPTX.tempName,'ppt','slides','_rels',fileRelsPath));
            
            % Even though masterNum and layoutNum variables are available in this
            % function we still want to get actual numbers based on generated XML file
            [mNum,lNum]         = PPTX.parseMasterLayoutNumber();
            
            % Assign data to slide
            PPTX.Slide(PPTX.numSlides).id           = PPTX.lastSlideId;
            PPTX.Slide(PPTX.numSlides).rId          = sprintf('rId%d',PPTX.lastRId);
            PPTX.Slide(PPTX.numSlides).file         = fileName;
            PPTX.Slide(PPTX.numSlides).objId        = 1;
            PPTX.Slide(PPTX.numSlides).masterNum    = mNum;
            PPTX.Slide(PPTX.numSlides).layoutNum    = lNum;
            PPTX.currentSlide                       = PPTX.numSlides;
            
            currentSlideId  = PPTX.currentSlide;
        end
        
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function currentSlide = switchSlide(PPTX,slideId)
            % switchSlide(slideId)
            %
            %   Switches current slide to be operated on. Requires slide ID as the second 
            %   input parameter.
            % 
            %   Examples:
            %       % Switch current slide to slide number 2
            %       pptx.switchSlide(2);
            
            %% Inputs
            if nargin<2,
                error('exportToPPTX:minInput','Second argument required: slide ID to switch to');
            end
        
            % Check slide numbers
            if slideId<1 || slideId>PPTX.numSlides || numel(slideId)>1,
                % Error condition
                error('exportToPPTX:badProperty','switchSlide position must be between 1 and the total number of slides');
            end
            
            % Before creating new slide, is there a current slide that needs to be
            % saved to XML file?
            if isfield(PPTX.XML,'Slide') && ~isempty(PPTX.XML.Slide),
                fileName    = PPTX.Slide(PPTX.currentSlide).file;
                xmlwrite(fullfile(PPTX.tempName,'ppt','slides',fileName),PPTX.XML.Slide);
                xmlwrite(fullfile(PPTX.tempName,'ppt','slides','_rels',cat(2,fileName,'.rels')),PPTX.XML.SlideRel);
            end
            
            fileRelsPath        = cat(2,PPTX.Slide(slideId).file,'.rels');
            PPTX.XML.Slide      = xmlread(fullfile(PPTX.tempName,'ppt','slides',PPTX.Slide(slideId).file));
            PPTX.XML.SlideRel   = xmlread(fullfile(PPTX.tempName,'ppt','slides','_rels',fileRelsPath));
            
            allIDs                      = exportToPPTX.getAllAttribute(PPTX.XML.Slide,'id');
            allIDNums                   = str2num(char(reshape(allIDs,[],1)));
            PPTX.Slide(slideId).objId   = max(allIDNums);
            
            [mNum,lNum]                     = PPTX.parseMasterLayoutNumber();
            PPTX.Slide(slideId).masterNum   = mNum;
            PPTX.Slide(slideId).layoutNum   = lNum;
            PPTX.currentSlide               = slideId;
            
            currentSlide    = PPTX.currentSlide;
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function addPicture(PPTX,imgData,varargin)
            % addPicture([figureHandle|axesHandle|imageFilename|CDATA],...)
            %
            %   Adds picture to the current slide. Requires figure or axes handle or image 
            %   filename or CDATA to be supplied. Images supplied as handles or CDATA 
            %   matricies are saved in PNG format. This command does not return any values.
            % 
            %   Additional parameters:
            %       Scale       Controls how image is placed on the slide
            %                   noscale  - No scaling (place figure as is in the center of 
            %                              the slide)
            %                   maxfixed - Max size while preserving aspect ratio (default 
            %                              when Position is not set)
            %                   max      - Max size with no aspect ratio preservation
            %                              (default when Position is set)
            %       Position    Four element vector: x, y, width, height (in inches) or 
            %                   template placeholder ID or name.
            %                   Coordinates x=0, y=0 are in the upper left corner of the slide.
            %       LineWidth   Width of the picture's edge line, a single value (in 
            %                   points). Edge is not drawn by default. Unless either 
            %                   LineWidth or EdgeColor are specified. 
            %       EdgeColor   Color of the picture's edge, a three element vector 
            %                   specifying RGB value. Edge is not drawn by default. Unless 
            %                   either LineWidth or EdgeColor are specified. 
            %       OnClick     Links text to another slide (if slide number is given as 
            %                   an integer) or URL or another file
            %
            %   Examples:
            %       % Add current figure
            %       figure,plot(rand(10));
            %       pptx.addPicture(gcf);
            %
            %       % Lower left corner picture inserted via image CDATA (height x width x 3)
            %       rgb = imread('ngc6543a.jpg');
            %       pptx.addPicture(rgb,'Position',[1 3.5 3 2]);
            
            % Adding image to existing slide
            %   1. Add <p:pic> node to slide#.xml
            %   2. Update slide#.xml.rels to link new pic to an image file
            %
            
            % Inputs
            if nargin<2,
                error('exportToPPTX:minInput','Second argument required: figure handle or filename or CDATA');
            end
            
            mi                  = false(size(varargin));
            [picPositionNew,mi] = exportToPPTX.getPVPair(varargin,'Position',[],mi);
            if isempty(picPositionNew)
                [scaleOpt,mi]   = exportToPPTX.getPVPair(varargin,'Scale','maxfixed',mi);
            else
                [scaleOpt,mi]   = exportToPPTX.getPVPair(varargin,'Scale','max',mi);
            end
            [lnWNew,mi]         = exportToPPTX.getPVPair(varargin,'LineWidth',[],mi);
            [lnColNew,mi]       = exportToPPTX.getPVPair(varargin,'EdgeColor',[],mi);
            [onClick,mi]        = exportToPPTX.getPVPair(varargin,'OnClick',[],mi);
            if any(~mi)
                error('exportToPPTX:badProperty','Unrecognized property %s',varargin{find(~mi,1)});
            end
            
            % Defaults
            showLn  = false;
            lnW     = 1;
            lnCol   = [0 0 0];
            
            % Check if media folder exists
            if ~exist(fullfile(PPTX.tempName,'ppt','media'),'dir'),
                mkdir(fullfile(PPTX.tempName,'ppt','media'));
            end
            
            % Set object ID
            PPTX.Slide(PPTX.currentSlide).objId    = PPTX.Slide(PPTX.currentSlide).objId+1;
            objId       = PPTX.Slide(PPTX.currentSlide).objId;
            
            % Get screen size (used in a few places)
            screenSize      = get(0,'ScreenSize');
            screenSize      = screenSize(1,3:4);
            emusPerPx       = max(round(PPTX.dimensions.*PPTX.CONST_IN_TO_EMU)./screenSize);
            isVideoClass    = false;
            
            % Based on the input format, decide what to do about it
            if ischar(imgData),
                if exist(imgData,'file'),
                    % Image or video filename
                    [d,d,inImgExt]  = fileparts(imgData);
                    inImgExt        = inImgExt(2:end);
                    imageName       = sprintf('media-%d-%d.%s',PPTX.currentSlide,objId,inImgExt);
                    imagePath       = fullfile(PPTX.tempName,'ppt','media',imageName);
                    copyfile(imgData,imagePath);
                    if any(strcmpi({'emf','wmf','eps'},inImgExt)),
                        % Vector images cannot be loaded by MatLab to determine their
                        % native sizes, but they can be scaled to any size anyway
                        imdims      = PPTX.dimensions([2 1])./emusPerPx.*PPTX.CONST_IN_TO_EMU;
                        
                    elseif exist('VideoReader','class') && any(strcmpi(get(VideoReader.getFileFormats,'Extension'),inImgExt)),
                        % This is a new MatLab style access to video information
                        % Only format types supported on this system are allowed
                        % because video has to be read to create a thumbnail image
                        videoName       = imageName;
                        videoPath       = imagePath;
                        vidinfo         = VideoReader(videoPath);
                        imdims          = [vidinfo.Height vidinfo.Width];
                        
                        % Video gets its own relationship ID
                        PPTX.Slide(PPTX.currentSlide).objId    = PPTX.Slide(PPTX.currentSlide).objId+1;
                        vidRId          = PPTX.Slide(PPTX.currentSlide).objId;
                        PPTX.Slide(PPTX.currentSlide).objId    = PPTX.Slide(PPTX.currentSlide).objId+1;
                        vidR2Id         = PPTX.Slide(PPTX.currentSlide).objId;
                        
                        % Prepare a thumbnail image
                        thumbData       = read(vidinfo,1);
                        imageName       = sprintf('media-thumb-%d-%d.png',PPTX.currentSlide,objId);
                        imagePath       = fullfile(PPTX.tempName,'ppt','media',imageName);
                        imwrite(thumbData,imagePath);
                        
                        clear vidinfo;
                        isVideoClass    = true;
                        
                    elseif any(strcmpi({'avi'},inImgExt)),
                        % This is an older MatLab style, which supports only AVI
                        videoName       = imageName;
                        videoPath       = imagePath;
                        vidinfo         = aviinfo(videoPath);
                        imdims          = [vidinfo.Height vidinfo.Width];
                        
                        % Video gets its own relationship ID
                        PPTX.Slide(PPTX.currentSlide).objId    = PPTX.Slide(PPTX.currentSlide).objId+1;
                        vidRId          = PPTX.Slide(PPTX.currentSlide).objId;
                        PPTX.Slide(PPTX.currentSlide).objId    = PPTX.Slide(PPTX.currentSlide).objId+1;
                        vidR2Id         = PPTX.Slide(PPTX.currentSlide).objId;
                        
                        % Prepare a thumbnail image
                        thumbData       = aviread(imgData,1);
                        imageName       = sprintf('media-thumb-%d-%d.png',PPTX.currentSlide,objId);
                        imagePath       = fullfile(PPTX.tempName,'ppt','media',imageName);
                        imwrite(thumbData.cdata,imagePath);
                        
                        clear vidinfo;
                        isVideoClass    = true;
                        
                    else
                        imgCdata    = imread(imgData);
                        imdims      = size(imgCdata);
                        
                    end
                    
                else
                    error('exportToPPTX:fileNotFound','Image file requested to be added to the slide was not found');
                end
                
            elseif isnumeric(imgData) && numel(imgData)>1,
                % Image CDATA
                imageName   = sprintf('image-%d-%d.png',PPTX.currentSlide,objId);
                imagePath   = fullfile(PPTX.tempName,'ppt','media',imageName);
                imwrite(imgData,imagePath);
                imdims      = size(imgData);
                inImgExt    = 'png';
                
            elseif ishghandle(imgData,'Figure') || ishghandle(imgData,'Axes'),
                % Either figure or axes handle
                img         = getframe(imgData);
                imageName   = sprintf('image-%d-%d.png',PPTX.currentSlide,objId);
                imagePath   = fullfile(PPTX.tempName,'ppt','media',imageName);
                imwrite(img.cdata,imagePath);
                imdims      = size(img.cdata);
                inImgExt    = 'png';
                
            else
                % Error condition
                error('exportToPPTX:badInput','addPicture command requires a valid figure/axes handle or filename or CDATA');
            end
            
            
            % If XML file does not support this format yet, then add it
            if isVideoClass,
                if ~any(strcmpi(PPTX.videoTypes,inImgExt)),
                    PPTX.addVideoTypeSupport(inImgExt);
                end
                % Make sure PNG (format for video thumbnail image) is supported
                if ~any(strcmpi(PPTX.imageTypes,'png')),
                    PPTX.addImageTypeSupport('png');
                end
            else
                if ~any(strcmpi(PPTX.imageTypes,inImgExt)),
                    PPTX.addImageTypeSupport(inImgExt);
                end
            end

            % % Save image -- this code would be MUCH faster, but less supported: requires additional MEX file and uses unsupported hardcopy
            % % Obtain a copy of fast PNG writing routine: http://www.mathworks.com/matlabcentral/fileexchange/40384
            % imageName   = sprintf('image%d.png',PPTX.currentSlide);
            % imagePath   = fullfile(PPTX.tempName,'ppt','media',imageName);
            % cdata       = hardcopy(figH,'-Dopengl','-r0');
            % savepng(cdata,imagePath);
            
            % Figure out picture size in PPTX units (EMUs)
            imEMUs          = imdims([2 1]).*emusPerPx;

            % Check picture position (absolute page position, or placeholder position)
            % Adding a picture to a placeholder follows a different procedure than
            % textbox. Picture is added into a placeholder, but its postion attributes
            % are filled out anyway, but matched to the underlying placeholder size.
            picPlaceholder  = [];
            picPosition     = [];
            mNum            = PPTX.Slide(PPTX.currentSlide).masterNum;
            lNum            = PPTX.Slide(PPTX.currentSlide).layoutNum;
            if ~isempty(picPositionNew)
                if ischar(picPositionNew)
                    % Change placeholder name into placeholder ID
                    posID   = find(strcmp(PPTX.SlideMaster(mNum).Layout(lNum).place,picPositionNew));
                    if isempty(posID)
                        posID   = find(strncmpi(PPTX.SlideMaster(mNum).Layout(lNum).place,picPositionNew,length(picPositionNew)));
                    end
                    if isempty(posID)
                        % Try search in ph attribute for the name match
                        posID   = find(strcmp(PPTX.SlideMaster(mNum).Layout(lNum).ph,picPositionNew));
                    end
                    if isempty(posID)
                        posID  = find(strncmpi(PPTX.SlideMaster(mNum).Layout(lNum).ph,picPositionNew,length(picPositionNew)));
                    end
                    if isempty(posID)
                        warning('exportToPPTX:badName','Placeholder "%s" does not exist in the current layout',picPositionNew);
                        posID   = 1;
                    end
                    if numel(posID)>1
                        warning('exportToPPTX:badName','There are multiple matches for placeholder "%s"',picPositionNew);
                        posID   = posID(1);
                    end
                    picPlaceholder  = posID;
                elseif isnumeric(picPositionNew) && numel(picPositionNew)==1,
                    if (picPositionNew<1 || picPositionNew>numel(PPTX.SlideMaster(mNum).Layout(lNum).ph)),
                        error('exportToPPTX:badProperty','Invalid placeholder index');
                    end
                    picPlaceholder  = picPositionNew;
                elseif isnumeric(picPositionNew) && numel(picPositionNew)==4,
                    picPosition     = round(picPositionNew.*PPTX.CONST_IN_TO_EMU);
                else
                    error('exportToPPTX:badProperty','Bad property value found in Position');
                end
            end
            if ~isempty(picPlaceholder),
                frameDimensions     = PPTX.SlideMaster(mNum).Layout(lNum).position{picPlaceholder};
            elseif ~isempty(picPosition)
                frameDimensions     = picPosition;
            else
                frameDimensions     = [0 0 round(PPTX.dimensions.*PPTX.CONST_IN_TO_EMU)];
            end
            
            aspRatioAttrib   = {};
            switch lower(scaleOpt),
                case 'noscale',
                    picPosition     = round([frameDimensions([1 2])+(frameDimensions([3 4])-imEMUs)./2 imEMUs]);
                    aspRatioAttrib  = [aspRatioAttrib {'noChangeAspect','1'}];
                case 'maxfixed',
                    scaleSize       = min(frameDimensions([3 4])./imEMUs);
                    newImEMUs       = imEMUs.*scaleSize;
                    picPosition     = round([frameDimensions([1 2])+(frameDimensions([3 4])-newImEMUs)./2 newImEMUs]);
                    aspRatioAttrib  = [aspRatioAttrib {'noChangeAspect','1'}];
                case 'max',
                    picPosition     = round(frameDimensions);
                otherwise,
                    error('exportToPPTX:badProperty','Bad property value found in Scale');
            end

            if ~isempty(lnWNew),
                showLn      = true;
                lnW         = lnWNew;
                if ~isnumeric(lnW) || numel(lnW)~=1,
                    error('exportToPPTX:badProperty','Bad property value found in LineWidth');
                end
            end

            if ~isempty(lnColNew),
                lnCol       = lnColNew;
                showLn      = true;
                if ~isnumeric(lnCol) || numel(lnCol)~=3,
                    error('exportToPPTX:badProperty','Bad property value found in EdgeColor');
                end
            end
            
            % Set object name
            objName     = 'Media File';
            if ~isempty(picPlaceholder),
                objName = PPTX.SlideMaster(mNum).Layout(lNum).place{picPlaceholder};
            end
            
            % Add image/video to slide XML file
            picNode     = exportToPPTX.addNode(PPTX.XML.Slide,'p:spTree','p:pic');
            nvPicPr     = exportToPPTX.addNode(PPTX.XML.Slide,picNode,'p:nvPicPr');
            cNvPr       = exportToPPTX.addNode(PPTX.XML.Slide,nvPicPr,'p:cNvPr',{'id',objId,'name',objName,'descr',imageName});
            if isVideoClass,
                exportToPPTX.addNode(PPTX.XML.Slide,cNvPr,'a:hlinkClick',{'r:id','','action','ppaction://media'});
            end
            cNvPicPr    = exportToPPTX.addNode(PPTX.XML.Slide,nvPicPr,'p:cNvPicPr');
                          exportToPPTX.addNode(PPTX.XML.Slide,cNvPicPr,'a:picLocks',aspRatioAttrib);
            nvPr        = exportToPPTX.addNode(PPTX.XML.Slide,nvPicPr,'p:nvPr');
            
            if ~isempty(picPlaceholder),
                allAttribs  = {};
                if ~isempty(PPTX.SlideMaster(mNum).Layout(lNum).idx{picPlaceholder}),
                    allAttribs  = [allAttribs {'idx',PPTX.SlideMaster(mNum).Layout(lNum).idx{picPlaceholder}}];
                end
                if ~isempty(PPTX.SlideMaster(mNum).Layout(lNum).ph{picPlaceholder}),
                    allAttribs  = [allAttribs {'type',PPTX.SlideMaster(mNum).Layout(lNum).ph{picPlaceholder}}];
                end
                exportToPPTX.addNode(PPTX.XML.Slide,nvPr,'p:ph',allAttribs);
            end
            
            if isVideoClass,
                exportToPPTX.addNode(PPTX.XML.Slide,nvPr,'a:videoFile',{'r:link',sprintf('rId%d',vidRId)});
                extLst = exportToPPTX.addNode(PPTX.XML.Slide,nvPr,'p:extLst');
                pExt   = exportToPPTX.addNode(PPTX.XML.Slide,extLst,'p:ext',{'uri','{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}'});
                exportToPPTX.addNode(PPTX.XML.Slide,pExt,'p14:media',{'xmlns:p14','http://schemas.microsoft.com/office/powerpoint/2010/main','r:embed',sprintf('rId%d',vidR2Id)});
            end
            
            blipFill    = exportToPPTX.addNode(PPTX.XML.Slide,picNode,'p:blipFill');
                          exportToPPTX.addNode(PPTX.XML.Slide,blipFill,'a:blip',{'r:embed',sprintf('rId%d',objId)','cstate','print'});
            stretch     = exportToPPTX.addNode(PPTX.XML.Slide,blipFill,'a:stretch');
                          exportToPPTX.addNode(PPTX.XML.Slide,stretch,'a:fillRect');
            
            spPr        = exportToPPTX.addNode(PPTX.XML.Slide,picNode,'p:spPr');
            
            axfm        = exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:xfrm');
                          exportToPPTX.addNode(PPTX.XML.Slide,axfm,'a:off',{'x',picPosition(1),'y',picPosition(2)});
                          exportToPPTX.addNode(PPTX.XML.Slide,axfm,'a:ext',{'cx',picPosition(3),'cy',picPosition(4)});
            prstGeom    = exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:prstGeom',{'prst','rect'});
                          exportToPPTX.addNode(PPTX.XML.Slide,prstGeom,'a:avLst');
            
            if showLn,
                aLn     = exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:ln',{'w',lnW*PPTX.CONST_PT_TO_EMU});
                sFil    = exportToPPTX.addNode(PPTX.XML.Slide,aLn,'a:solidFill');
                          exportToPPTX.addNode(PPTX.XML.Slide,sFil,'a:srgbClr',{'val',sprintf('%s',dec2hex(round(lnCol*255),2).')});
            end
            
            % Add slide timing info, node <p:timing>
            if isVideoClass,
                % TODO: this needs to be added at sometime later...
                pTiming     = exportToPPTX.addNode(PPTX.XML.Slide,'p:sld','p:timing');
                tnLst       = exportToPPTX.addNode(PPTX.XML.Slide,pTiming,'p:tnLst');
                pPar        = exportToPPTX.addNode(PPTX.XML.Slide,tnLst,'p:par');
                              exportToPPTX.addNode(PPTX.XML.Slide,pPar,'p:cTn',{'id',1,'dur','indefinite','restart','never','nodeType','tmRoot'});
            end
            
            % Add image reference to rels file
            exportToPPTX.addNode(PPTX.XML.SlideRel,'Relationships','Relationship', ...
                {'Id',sprintf('rId%d',objId), ...
                'Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', ...
                'Target',cat(2,'../media/',imageName)});
            if isVideoClass,
                exportToPPTX.addNode(PPTX.XML.SlideRel,'Relationships','Relationship', ...
                    {'Id',sprintf('rId%d',vidRId), ...
                    'Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/video', ...
                    'Target',cat(2,'../media/',videoName)});
                exportToPPTX.addNode(PPTX.XML.SlideRel,'Relationships','Relationship', ...
                    {'Id',sprintf('rId%d',vidR2Id), ...
                    'Type','http://schemas.microsoft.com/office/2007/relationships/media', ...
                    'Target',cat(2,'../media/',videoName)});
            end
            
            PPTX.addLink(onClick,cNvPr);
            
        end
        
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function addShape(PPTX,xData,yData,varargin)
            % addShape(xData,yData,...)
            %
            %   Add lines or closed shapes to the current slide. Requires X and Y data to 
            %   be supplied. This command does not return any values.
            % 
            %   Additional parameters:
            %       ClosedShape Specifies whether the shape is automatically closed or not. 
            %                   Default value is false.
            %       LineWidth   Width of the line, a single value (in points). Default line 
            %                   width is 1 point. Set LineWidth to zero have no edge drawn.
            %       LineColor   Color of the drawn line, a three element vector specifying 
            %                   RGB value. Default color is black.
            %       LineStyle   Style of the drawn line. Default style is a solid line. 
            %                   The following styles are available:
            %                       - (solid), : (dotted), -. (dash dot), -- (dashes)
            %       BackgroundColor     Shape fill color, a three element vector specifying 
            %                           RGB value. By default shapes are drawn transparent.
            %
            %   Examples:
            %       % Add lines
            %       xData   = linspace(0,12,101);
            %       yData  = 3+sin(xData*2)*0.3;
            %       pptx.addShape(xData,yData,'LineWidth',2,'LineStyle',':');
            %
            %       % Add filled circle
            %       theta   = linspace(0,2*pi,101); xData = sin(theta); yData = cos(theta);
            %       pptx.addShape(xData+4,yData+1,'LineWidth',2,'LineColor','r','LineStyle','--','BackgroundColor','g','ClosedShape',true);
            
            % Adding a line/patch segment (custGeom) to PPTX
            %   1. Add <p:sp> node to slide#.xml

            %% Inputs
            if nargin<3,
                error('exportToPPTX:minInput','Two input argument required: X and Y data');
            end
            
            mi              = false(size(varargin));
            [bCol,mi]       = exportToPPTX.getPVPair(varargin,'BackgroundColor',[],mi);
            [isClosed,mi]   = exportToPPTX.getPVPair(varargin,'ClosedShape',false,mi);
            [lnWVal,mi]     = exportToPPTX.getPVPair(varargin,'LineWidth',[],mi);
            [lnColVal,mi]   = exportToPPTX.getPVPair(varargin,'LineColor',[],mi);
            [lnStyleVal,mi] = exportToPPTX.getPVPair(varargin,'LineStyle',[],mi);
            
            if any(~mi)
                error('exportToPPTX:badProperty','Unrecognized property %s',varargin{find(~mi,1)});
            end
            
            % Input error checking
            if isempty(xData) || isempty(yData),
                % Error condition
                error('exportToPPTX:badInput','addShape command requires non-empty X and Y data');
            end
            
            if size(xData)~=size(yData),
                % Error condition
                error('exportToPPTX:badInput','addShape command requires X and Y data sizes to match');
            end
            
            if numel(xData)==1 || numel(yData)==1,
                % Error condition
                error('exportToPPTX:badInput','addShape command requires at least two X and Y data point');
            end
            
            % If needed, reshape data (dim 1 = segments of a single line, dim 2 = different lines)
            if size(xData,1)==1,
                xData   = reshape(xData,[],1);
                yData   = reshape(yData,[],1);
            end
            
            % Convert to PPTX coordinates
            xData       = round(xData*PPTX.CONST_IN_TO_EMU);
            yData       = round(yData*PPTX.CONST_IN_TO_EMU);
            
            % Defaults
            showLn      = true;
            lnW         = 1;
            lnCol       = exportToPPTX.validateColor([0 0 0]);
            lnStyle     = '';
            
            bCol        = exportToPPTX.validateColor(bCol);

            if ~isempty(lnWVal),
                lnW         = lnWVal;
                if ~isnumeric(lnW) || numel(lnW)~=1,
                    error('exportToPPTX:badProperty','Bad property value found in LineWidth');
                end
                if lnW==0,
                    showLn  = false;
                end
            end

            lnColVal    = exportToPPTX.validateColor(lnColVal);
            if ~isempty(lnColVal),
                lnCol   = lnColVal;
            end

            if ~isempty(lnStyleVal),
                switch (lnStyleVal),
                    case '-',
                        lnStyle     = 'solid';
                    case ':',
                        lnStyle     = 'sysDot';
                    case '-.',
                        lnStyle     = 'sysDashDot';
                    case '--',
                        lnStyle     = 'sysDash';
                    otherwise,
                        error('exportToPPTX:badProperty','Bad property value found in LineStyle');
                end
                % Other possible values supported by PowerPoint
                %     dash
                %     dashDot
                %     dot
                %     lgDash (large dash)
                %     lgDashDot
                %     lgDashDotDot
                %     solid
                %     sysDash (system dash)
                %     sysDashDot
                %     sysDashDotDot
                %     sysDot
            end
            
            numShapes   = size(xData,2);
            numSeg      = size(xData,1);
            
            % Iterate over all shapes
            for ishape=1:numShapes,
                
                % Set object ID
                PPTX.Slide(PPTX.currentSlide).objId    = PPTX.Slide(PPTX.currentSlide).objId+1;
                objId       = PPTX.Slide(PPTX.currentSlide).objId;
                
                % Determine line boundaries
                xLim        = [min(xData(:,ishape)) max(xData(:,ishape))];
                yLim        = [min(yData(:,ishape)) max(yData(:,ishape))];
                
                % Add line to slide XML file
                spNode      = exportToPPTX.addNode(PPTX.XML.Slide,'p:spTree','p:sp');
                nvPicPr     = exportToPPTX.addNode(PPTX.XML.Slide,spNode,'p:nvSpPr');
                              exportToPPTX.addNode(PPTX.XML.Slide,nvPicPr,'p:cNvPr',{'id',objId,'name',sprintf('Freeform %d',objId)});
                              exportToPPTX.addNode(PPTX.XML.Slide,nvPicPr,'p:cNvSpPr');
                              exportToPPTX.addNode(PPTX.XML.Slide,nvPicPr,'p:nvPr');
                
                spPr        = exportToPPTX.addNode(PPTX.XML.Slide,spNode,'p:spPr');
                axfm        = exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:xfrm');
                              exportToPPTX.addNode(PPTX.XML.Slide,axfm,'a:off',{'x',xLim(1),'y',yLim(1)});
                              exportToPPTX.addNode(PPTX.XML.Slide,axfm,'a:ext',{'cx',xLim(2)-xLim(1),'cy',yLim(2)-yLim(1)});
                custGeom    = exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:custGeom');
                              exportToPPTX.addNode(PPTX.XML.Slide,custGeom,'a:avLst');
                              exportToPPTX.addNode(PPTX.XML.Slide,custGeom,'a:gdLst');
                              exportToPPTX.addNode(PPTX.XML.Slide,custGeom,'a:ahLst');
                              exportToPPTX.addNode(PPTX.XML.Slide,custGeom,'a:cxnLst');
                pathLst     = exportToPPTX.addNode(PPTX.XML.Slide,custGeom,'a:pathLst');
                aPath       = exportToPPTX.addNode(PPTX.XML.Slide,pathLst,'a:path',{'w',xLim(2)-xLim(1),'h',yLim(2)-yLim(1)});
                
                moveTo      = exportToPPTX.addNode(PPTX.XML.Slide,aPath,'a:moveTo');
                              exportToPPTX.addNode(PPTX.XML.Slide,moveTo,'a:pt',{'x',xData(1,ishape)-xLim(1),'y',yData(1,ishape)-yLim(1)});
                for iseg=2:numSeg,
                    lineTo  = exportToPPTX.addNode(PPTX.XML.Slide,aPath,'a:lnTo');
                              exportToPPTX.addNode(PPTX.XML.Slide,lineTo,'a:pt',{'x',xData(iseg,ishape)-xLim(1),'y',yData(iseg,ishape)-yLim(1)});
                end
                
                if isClosed,
                              exportToPPTX.addNode(PPTX.XML.Slide,aPath,'a:close');
                end
                
                if ~isempty(bCol),
                    solidFill=exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:solidFill');
                              exportToPPTX.addNode(PPTX.XML.Slide,solidFill,'a:srgbClr',{'val',bCol});
                else
                              exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:noFill');
                end
                
                if showLn,
                    aLn     = exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:ln',{'w',lnW*PPTX.CONST_PT_TO_EMU});
                    sFil    = exportToPPTX.addNode(PPTX.XML.Slide,aLn,'a:solidFill');
                              exportToPPTX.addNode(PPTX.XML.Slide,sFil,'a:srgbClr',{'val',lnCol});
                    if ~isempty(lnStyle),
                              exportToPPTX.addNode(PPTX.XML.Slide,aLn,'a:prstDash',{'val',lnStyle});
                    end
                end
                
            end
        end
        
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function addNote(PPTX,notesText,varargin)
            % addNote(noteText,...)
            %   
            %   Add notes information to the current slide. Requires text of the notes to 
            %   be added. This command does not return any values. 
            %   Note: repeat calls overwrite previous information.
            % 
            %   Additional parameters:
            %       FontWeight  Weight of text characters:
            %                       normal - use regular font (default)
            %                       bold - use bold font
            %       FontAngle   Character slant:
            %                       normal - no character slant (default)
            %                       italic - use slanted font
            %
            %   Examples:
            %       % Adding a simple note
            %       pptx.addNote('This is a simple note');
            %
            %       % Adding a multiline formatted note
            %       pptx.addNote( ...
            %         {'FontName property can also be used in the notes field.', ...
            %         'It is useful for adding code type output.', ...
            %         '', ...
            %         '**To view text formatting changes that you make in the Notes pane, do the following**', ...
            %         '# On the View tab, in the Presentation Views group, click Outline View.', ...
            %         '# Right-click the Outline pane, and then select Show Text Formatting in the popped up menu.'}, ...
            %         'FontName','FixedWidth','FontSize',10);
            
            % Adding a new notes slide to PPTX
            %   1. Create notesSlide#.xml file
            %   2. Create notesSlide#.xml.rels file
            %   3. Update [Content_Types].xml
            %   4. Add link to notesSlide#.xml file in slide#.xml.rels
            
            %% Inputs
            if nargin<2,
                error('exportToPPTX:minInput','Second argument required: text for the notes field');
            end
            
            % Check if notes have already been added to this slide and if so, then overwrite existing content
            fileName    = sprintf('notesSlide%d.xml',PPTX.Slide(PPTX.currentSlide).id);
            
            % Check if notes master exists
            if isempty(PPTX.NotesMaster)
                PPTX.themeNum   = PPTX.themeNum+1;
                PPTX.lastRId    = PPTX.lastRId+1;
                retCode         = PPTX.initNotesMaster(PPTX.themeNum);
                
                % Update [Content_Types].xml
                exportToPPTX.addNode(PPTX.XML.TOC,'Types','Override', ...
                    {'PartName',cat(2,'/ppt/theme/',sprintf('theme%d.xml',PPTX.themeNum)), ...
                    'ContentType','application/vnd.openxmlformats-officedocument.theme+xml'});
                
                exportToPPTX.addNode(PPTX.XML.TOC,'Types','Override', ...
                    {'PartName','/ppt/notesMasters/notesMaster1.xml', ...
                    'ContentType','application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml'});

                % Link new notes master to presentation.xml slide ID
                exportToPPTX.addNode(PPTX.XML.PresRel,'Relationships','Relationship', ...
                    {'Id',cat(2,'rId',int2str(PPTX.lastRId)), ...
                    'Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster', ...
                    'Target','notesMasters/notesMaster1.xml'});
                
                notesMasterIdLstNode    = exportToPPTX.findNode(PPTX.XML.Pres,'p:notesMasterIdLst');
                if isempty(notesMasterIdLstNode),
                    % Note: order of tags within XML structure is important, notes master list
                    % must show up before slide id list
                    exportToPPTX.addNodeBefore(PPTX.XML.Pres,'p:presentation','p:notesMasterIdLst','p:sldIdLst');
                end

                exportToPPTX.addNode(PPTX.XML.Pres,'p:notesMasterIdLst','p:notesMasterId', ...
                    {'r:id',sprintf('rId%d',PPTX.lastRId)});
                
                PPTX.NotesMaster.rId    = PPTX.lastRId;
                PPTX.NotesMaster.file   = 'notesMaster1.xml';
            end
            
            if ~exist(fullfile(PPTX.tempName,'ppt','notesSlides',fileName),'file'),
                
                % Create notes folder if it doesn't exist
                if ~exist(fullfile(PPTX.tempName,'ppt','notesSlides'),'dir'),
                    mkdir(fullfile(PPTX.tempName,'ppt','notesSlides'));
                    mkdir(fullfile(PPTX.tempName,'ppt','notesSlides','_rels'));
                end
                
                % Create new XML file
                % file name and path are relative to presentation.xml file
                fileContent     = { ...
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
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
                    '<p:sp>'
                    '<p:nvSpPr>'
                    '<p:cNvPr id="3" name="Notes Placeholder 2"/>'
                    '<p:cNvSpPr>'
                    '<a:spLocks noGrp="1"/>'
                    '</p:cNvSpPr>'
                    '<p:nvPr>'
                    '<p:ph type="body" idx="1"/>'
                    '</p:nvPr>'
                    '</p:nvSpPr>'
                    '<p:spPr/>'     % p:txBody is skipped on purposes, added later via XML manipulations
                    '</p:sp>'
                    '</p:spTree>'
                    '</p:cSld>'
                    '<p:clrMapOvr>'
                    '<a:masterClrMapping/>'
                    '</p:clrMapOvr>'
                    '</p:notes>'};
                retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','notesSlides',fileName),fileContent);
                
                % Create new XML relationships file
                fileRelsPath    = cat(2,fileName,'.rels');
                fileContent     = { ...
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    ['<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/' PPTX.Slide(PPTX.currentSlide).file '"/>']
                    ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/' PPTX.NotesMaster.file '"/>']
                    '</Relationships>'};
                retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','notesSlides','_rels',fileRelsPath),fileContent) & retCode;
                
                % Update [Content_Types].xml
                exportToPPTX.addNode(PPTX.XML.TOC,'Types','Override', ...
                    {'PartName',cat(2,'/ppt/notesSlides/',fileName), ...
                    'ContentType','application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'});
                
                % Add link to notesSlide#.xml file in slide#.xml.rels
                PPTX.Slide(PPTX.currentSlide).objId    = PPTX.Slide(PPTX.currentSlide).objId+1;
                objId       = PPTX.Slide(PPTX.currentSlide).objId;
                exportToPPTX.addNode(PPTX.XML.SlideRel,'Relationships','Relationship', ...
                    {'Id',sprintf('rId%d',objId), ...
                    'Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide', ...
                    'Target',cat(2,'../notesSlides/',fileName)});
                
                % Load notes slide
                notesXMLSlide   = xmlread(fullfile(PPTX.tempName,'ppt','notesSlides',fileName));
                spNode          = exportToPPTX.findNode(notesXMLSlide,'p:sp');
                
            else
                % Slide file already exists, so replace txBody node with new content
                notesXMLSlide   = xmlread(fullfile(PPTX.tempName,'ppt','notesSlides',fileName));
                exportToPPTX.removeNode(notesXMLSlide,'p:txBody');
                spNode          = exportToPPTX.findNode(notesXMLSlide,'p:sp');
                
            end
            
            % check if node shape node was found
            if isempty(spNode),
                error('Shape node not found.');
            end
            
            % and add formatted text to it
            txBody      = exportToPPTX.addNode(notesXMLSlide,spNode,'p:txBody');
            PPTX.addTxBodyNode(notesXMLSlide,txBody,notesText,varargin{:});
            
            % Save XML file back
            xmlwrite(fullfile(PPTX.tempName,'ppt','notesSlides',fileName),notesXMLSlide);
        end
        
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function addTextbox(PPTX,boxText,varargin)
            % addTextbox(textboxText,...)
            %
            %   Adds textbox to the current slide. Requires text of the box to be added. 
            %   This command does not return any values.
            %
            %   Additional parameters:
            %       Position    Four element vector: x, y, width, height (in inches) or 
            %                   template placeholder ID or name. Coordinates x=0, y=0 are 
            %                   in the upper left corner of the slide.
            %       Color       Three element vector specifying RGB value in the range from 
            %                   0 to 1. Default text color is black.
            %       BackgroundColor     Three element vector specifying RGB value in the 
            %                           range from 0 to 1. By default background is transparent.
            %       FontSize    Specifies the font size to use for text. Default font size is 12.
            %       FontName    Specifies font name to be used. Default is whatever is 
            %                   template defined font is. Specifying FixedWidth for font 
            %                   name will use monospaced font defined on the system.
            %       FontWeight  Weight of text characters:
            %                       normal - use regular font (default)
            %                       bold - use bold font
            %       FontAngle   Character slant:
            %                       normal - no character slant (default)
            %                       italic - use slanted font
            %       Rotation    Determines the orientation of the textbox. Specify values 
            %                   of rotation in degrees (positive angles cause counterclockwise 
            %                   rotation).
            %       HorizontalAlignment     Horizontal alignment of text:
            %                                   left - left-aligned text (default)
            %                                   center - centered text
            %                                   right - right-aligned text
            %       VerticalAlignment       Vertical alignment of text:
            %                                   top - top-aligned text (default)
            %                                   middle - align to the middle of the textbox
            %                                   bottom - bottom-aligned text
            %       LineWidth   Width of the textbox's edge line, a single value (in points). 
            %                   Edge is not drawn by default. Unless either LineWidth or 
            %                   EdgeColor are specified.
            %       EdgeColor   Color of the textbox's edge, a three element vector 
            %                   specifying RGB value. Edge is not drawn by default. Unless 
            %                   either LineWidth or EdgeColor are specified.
            %       OnClick     Links text to another slide (if slide number is given as 
            %                   an integer) or URL or another file
            %
            %   Examples:
            %       % Add simple textbox
            %       pptx.addTextbox('Simple textbox text');
            %
            %       % Right-top aligned, bold and italicized
            %       pptx.addTextbox('Right-top and bold italics','Position',[8 0 4 2], ...
            %           'HorizontalAlignment','right','FontWeight','bold','FontAngle','italic');
            
            % Adding image to existing slide
            %   1. Add <p:sp> node to slide#.xml
            
            %% Inputs
            if nargin<2,
                error('exportToPPTX:minInput','Second argument required: textbox contents');
            end
            
            mi              = false(size(varargin));
            [textPos,mi]    = exportToPPTX.getPVPair(varargin,'Position',[0 0 PPTX.dimensions],mi);
            [bCol,mi]       = exportToPPTX.getPVPair(varargin,'BackgroundColor',[],mi);
            [fRot,mi]       = exportToPPTX.getPVPair(varargin,'Rotation',0,mi);
            [lnWVal,mi]     = exportToPPTX.getPVPair(varargin,'LineWidth',[],mi);
            [lnColVal,mi]   = exportToPPTX.getPVPair(varargin,'EdgeColor',[],mi);
            
            % Not checking the state of mi because more options are passed into
            % addTxBodyNode

            % Defaults
            showLn      = false;
            lnW         = 1;
            lnCol       = exportToPPTX.validateColor([0 0 0]);
            
            % Get position (if specified)
            % By default places textbox the size of the slide
            if isnumeric(textPos) && numel(textPos)>1,
                textPos = round(textPos.*PPTX.CONST_IN_TO_EMU);
            end
            
            % Check textbox position (absolute page position, or placeholder position)
            mNum        = PPTX.Slide(PPTX.currentSlide).masterNum;
            lNum        = PPTX.Slide(PPTX.currentSlide).layoutNum;
            if ischar(textPos)
                % Change textbox placeholder name into placeholder ID
                % Fix by hjw.
                posID       = find(strcmp(PPTX.SlideMaster(mNum).Layout(lNum).place,textPos));
                if isempty(posID)
                    posID   = find(strncmpi(PPTX.SlideMaster(mNum).Layout(lNum).place,textPos,length(textPos)));
                end
                if isempty(posID)
                    % Try search in ph attribute for the name match
                    posID   = find(strcmp(PPTX.SlideMaster(mNum).Layout(lNum).ph,textPos));
                end
                if isempty(posID)
                    posID   = find(strncmpi(PPTX.SlideMaster(mNum).Layout(lNum).ph,textPos,length(textPos)));
                end
                if isempty(posID)
                    warning('exportToPPTX:badName','Placeholder "%s" does not exist in the current layout',textPos);
                    posID   = 1;
                end
                if numel(posID)>1
                    warning('exportToPPTX:badName','There are multiple matches for placeholder "%s"',textPos);
                    posID   = posID(1);
                end
                textPos     = posID;
            end
            if numel(textPos)==1 && (textPos<1 || textPos>numel(PPTX.SlideMaster(mNum).Layout(lNum).ph)),
                error('exportToPPTX:badProperty','Invalid placeholder index');
            end

            bCol        = exportToPPTX.validateColor(bCol);

            if ~isnumeric(fRot) || numel(fRot)~=1,
                error('exportToPPTX:badProperty','Bad property value found in Rotation');
            end

            if ~isempty(lnWVal),
                lnW         = lnWVal;
                showLn      = true;
                if ~isnumeric(lnW) || numel(lnW)~=1,
                    error('exportToPPTX:badProperty','Bad property value found in LineWidth');
                end
            end
            
            lnColVal    = exportToPPTX.validateColor(lnColVal);
            if ~isempty(lnColVal),
                lnCol   = lnColVal;
                showLn  = true;
            end
            
            
            % Set object ID
            PPTX.Slide(PPTX.currentSlide).objId    = PPTX.Slide(PPTX.currentSlide).objId+1;
            objId       = PPTX.Slide(PPTX.currentSlide).objId;
            
            objName     = 'TextBox';
            if numel(textPos)==1,
                objName = PPTX.SlideMaster(mNum).Layout(lNum).place{textPos};
            end
            
            % Add textbox to slide XML file
            spNode      = exportToPPTX.addNode(PPTX.XML.Slide,'p:spTree','p:sp');
            nvPicPr     = exportToPPTX.addNode(PPTX.XML.Slide,spNode,'p:nvSpPr');
                          exportToPPTX.addNode(PPTX.XML.Slide,nvPicPr,'p:cNvPr',{'id',objId,'name',objName});
                          exportToPPTX.addNode(PPTX.XML.Slide,nvPicPr,'p:cNvSpPr',{'txBox','1'});
            nvPr        = exportToPPTX.addNode(PPTX.XML.Slide,nvPicPr,'p:nvPr');
            
            % Insert link to placeholder (if requested)
            if numel(textPos)==1,
                allAttribs  = {};
                if ~isempty(PPTX.SlideMaster(mNum).Layout(lNum).idx{textPos}),
                    allAttribs  = [allAttribs {'idx',PPTX.SlideMaster(mNum).Layout(lNum).idx{textPos}}];
                end
                if ~isempty(PPTX.SlideMaster(mNum).Layout(lNum).ph{textPos}),
                    allAttribs  = [allAttribs {'type',PPTX.SlideMaster(mNum).Layout(lNum).ph{textPos}}];
                end
                exportToPPTX.addNode(PPTX.XML.Slide,nvPr,'p:ph',allAttribs);
            end
            
            spPr        = exportToPPTX.addNode(PPTX.XML.Slide,spNode,'p:spPr');
            
            % If not a placeholder, then set position on the page
            if numel(textPos)==4,
                axfm    = exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:xfrm',{'rot',-fRot*PPTX.CONST_DEG_TO_EMU});
                          exportToPPTX.addNode(PPTX.XML.Slide,axfm,'a:off',{'x',textPos(1),'y',textPos(2)});
                          exportToPPTX.addNode(PPTX.XML.Slide,axfm,'a:ext',{'cx',textPos(3),'cy',textPos(4)});
                prstGeom= exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:prstGeom',{'prst','rect'});
                          exportToPPTX.addNode(PPTX.XML.Slide,prstGeom,'a:avLst');
            end
            
            if ~isempty(bCol),
                solidFill=exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:solidFill');
                          exportToPPTX.addNode(PPTX.XML.Slide,solidFill,'a:srgbClr',{'val',bCol});
            else
                          exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:noFill');
            end
            
            if showLn,
                aLn     = exportToPPTX.addNode(PPTX.XML.Slide,spPr,'a:ln',{'w',lnW*PPTX.CONST_PT_TO_EMU});
                sFil    = exportToPPTX.addNode(PPTX.XML.Slide,aLn,'a:solidFill');
                          exportToPPTX.addNode(PPTX.XML.Slide,sFil,'a:srgbClr',{'val',lnCol});
            end
            
            txBody      = exportToPPTX.addNode(PPTX.XML.Slide,spNode,'p:txBody');
            PPTX.addTxBodyNode(PPTX.XML.Slide,txBody,boxText,varargin{~mi});
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function addTable(PPTX,tableData,varargin)
            % addTable(tableData,...)
            %
            %   Adds PowerPoint table to the current slide. Requires table content to be 
            %   supplied in the form of a cell matrix. This command does not return any values.
            % 
            %   All of the 'addTextbox' Additional parameters apply to the table as well.
            %
            %   Examples:
            %       % Basic table insert
            %       tableData   = { ...
            %             'Header 1','Header 2','Header 3'; ...
            %             'Row 1','Data A','Data B'; ...
            %             'Row 2',10,20; ...
            %             'Row 3',NaN,'N/A'; };
            %       pptx.addTable(tableData,'Position',[2 1 8 4]);
            
            % Adding a table element to PPTX
            %   1. Add <p:graphicFrame>
            %   2. Add a set of <p:nvGraphicFramePr> properties
            %   3. Add <p:xfrm> for location
            %   4. Add graphic root element <a:graphic>, <a:graphicData> and <a:tbl>
            %   5. Set <a:tblPr> and <a:tblGrid>
            %   6. Loop over rows adding <a:tr>
            %   7. Loop over columns within each row adding <a:tc>
            
            %% Inputs
            if nargin<2,
                error('exportToPPTX:minInput','Second argument required: table contents');
            end
            
            mi                  = false(size(varargin));
            [tablePos,mi]       = exportToPPTX.getPVPair(varargin,'Position',[0 0 PPTX.dimensions],mi);
            [colWidthVal,mi]    = exportToPPTX.getPVPair(varargin,'ColumnW',[],mi);
            
            % Not checking the state of mi because more options are passed into addTableCell
            
            % Check input
            numCols     = size(tableData,2);
            numRows     = size(tableData,1);
            
            % Get position (if specified)
            % By default places table the size of the slide
            
            if isnumeric(tablePos) && numel(tablePos)>1,
                tablePos    = round(tablePos.*PPTX.CONST_IN_TO_EMU);
            end
            
            % Check textbox position (absolute page position, or placeholder position)
            usePlaceholder  = [];
            mNum            = PPTX.Slide(PPTX.currentSlide).masterNum;
            lNum            = PPTX.Slide(PPTX.currentSlide).layoutNum;
            if ischar(tablePos)
                % Change textbox placeholder name into placeholder ID
                posID   = find(strcmp(PPTX.SlideMaster(mNum).Layout(lNum).place,tablePos));
                if isempty(posID)
                    posID   = find(strncmpi(PPTX.SlideMaster(mNum).Layout(lNum).place,tablePos,length(tablePos)));
                end
                if isempty(posID)
                    % Try search in ph attribute for the name match
                    posID   = find(strcmp(PPTX.SlideMaster(mNum).Layout(lNum).ph,tablePos));
                end
                if isempty(posID)
                    posID   = find(strncmpi(PPTX.SlideMaster(mNum).Layout(lNum).ph,tablePos,length(tablePos)));
                end
                if isempty(posID)
                    warning('exportToPPTX:badName','Placeholder "%s" does not exist in the current layout',tablePos);
                    posID   = 1;
                end
                if numel(posID)>1
                    warning('exportToPPTX:badName','There are multiple matches for placeholder "%s"',tablePos);
                    posID   = posID(1);
                end
                usePlaceholder  = posID;
            elseif isnumeric(tablePos) && numel(tablePos)==1,
                usePlaceholder  = tablePos(1);
            end
            if numel(tablePos)==1 && (tablePos<1 || tablePos>numel(PPTX.SlideMaster(mNum).Layout(lNum).ph)),
                error('exportToPPTX:badProperty','Invalid placeholder index');
            end
            if ~isempty(usePlaceholder),
                tablePos    = PPTX.SlideMaster(mNum).Layout(lNum).position{usePlaceholder};
            end

            if ~isempty(colWidthVal),
                if numel(colWidthVal)~=numCols,
                    error('exportToPPTX:badProperty','Number of elements in ColumnWidth must equal the number of columns');
                end
                % if sum(round(colWidthVal*100))~=100,
                if sum(colWidthVal*100)~=100,
                    error('exportToPPTX:badProperty','Sum of all elements of ColumnWidth must add up to 1');
                end
                colWidth    = round(colWidthVal*tablePos(3));
            else
                colWidth    = repmat(round(tablePos(3)/numCols),1,numCols);
            end
            
            % Set object ID
            PPTX.Slide(PPTX.currentSlide).objId    = PPTX.Slide(PPTX.currentSlide).objId+1;
            objId       = PPTX.Slide(PPTX.currentSlide).objId;
            
            objName     = sprintf('Table %d',objId);
            if ~isempty(usePlaceholder),
                objName = PPTX.SlideMaster(mNum).Layout(lNum).place{usePlaceholder};
            end
            
            % Add all basic table elements
            pGf         = exportToPPTX.addNode(PPTX.XML.Slide,'p:spTree','p:graphicFrame');
            pNvGf       = exportToPPTX.addNode(PPTX.XML.Slide,pGf,'p:nvGraphicFramePr');
                          exportToPPTX.addNode(PPTX.XML.Slide,pNvGf,'p:cNvPr',{'id',objId,'name',objName});
                          exportToPPTX.addNode(PPTX.XML.Slide,pNvGf,'p:cNvGraphicFramePr');
            nvPr        = exportToPPTX.addNode(PPTX.XML.Slide,pNvGf,'p:nvPr');
            
            % Insert link to placeholder (if requested)
            if ~isempty(usePlaceholder),
                allAttribs  = {};
                if ~isempty(PPTX.SlideMaster(mNum).Layout(lNum).idx{usePlaceholder}),
                    allAttribs  = [allAttribs {'idx',PPTX.SlideMaster(mNum).Layout(lNum).idx{usePlaceholder}}];
                end
                if ~isempty(PPTX.SlideMaster(mNum).Layout(lNum).ph{usePlaceholder}),
                    allAttribs  = [allAttribs {'type',PPTX.SlideMaster(mNum).Layout(lNum).ph{usePlaceholder}}];
                end
                exportToPPTX.addNode(PPTX.XML.Slide,nvPr,'p:ph',allAttribs);
            end
            
            % If not a placeholder, then set position on the page
            axfm        = exportToPPTX.addNode(PPTX.XML.Slide,pGf,'p:xfrm');
                          exportToPPTX.addNode(PPTX.XML.Slide,axfm,'a:off',{'x',tablePos(1),'y',tablePos(2)});
                          exportToPPTX.addNode(PPTX.XML.Slide,axfm,'a:ext',{'cx',tablePos(3),'cy',tablePos(4)});
            
            aG          = exportToPPTX.addNode(PPTX.XML.Slide,pGf,'a:graphic');
            aGd         = exportToPPTX.addNode(PPTX.XML.Slide,aG,'a:graphicData',{'uri','http://schemas.openxmlformats.org/drawingml/2006/table'});
            
            aTbl        = exportToPPTX.addNode(PPTX.XML.Slide,aGd,'a:tbl');
                          exportToPPTX.addNode(PPTX.XML.Slide,aTbl,'a:tblPr');
            aGrid       = exportToPPTX.addNode(PPTX.XML.Slide,aTbl,'a:tblGrid');
            
            % loop over columns
            for icol=1:numCols,
                exportToPPTX.addNode(PPTX.XML.Slide,aGrid,'a:gridCol',{'w',colWidth(icol)});
            end
            
            % loop over rows
            for irow=1:numRows,
                aTr     = exportToPPTX.addNode(PPTX.XML.Slide,aTbl,'a:tr',{'h',round(tablePos(4)/numRows)});
                
                % loop over columns within each row
                for icol=1:numCols,
                    aTc         = exportToPPTX.addNode(PPTX.XML.Slide,aTr,'a:tc');
                    cellData    = tableData{irow,icol};
                    
                    if iscell(cellData),
                        combArgs    = [cellData(2:end) varargin(~mi)];
                        PPTX.addTableCell(aTc,cellData{1},combArgs{:});
                    else
                        PPTX.addTableCell(aTc,cellData,varargin{~mi});
                    end
                end
            end
        end
        
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function fullName = save(PPTX,fileName)
            % save([filename])
            %
            %   Saves current presentation. If new PowerPoint file was created, then 
            %   filename to save to is required. If PowerPoint was opened, then by default 
            %   it will write changes back to the same file. If another filename is 
            %   provided, then changes will be written to the new file (effectively a 
            %   'Save As' operation). 
            %
            %   Returns full path and name of the presentation file written.
            %
            %   Examples:
            %       % Save everything to example.pptx
            %       pptx.save('example.pptx');
            %
            %       % Save back to the openned file
            %       pptx    = exportToPPTX('example.pptx');
            %       pptx.save();
            
            %% Inputs
            if nargin<2 && isempty(PPTX.fullName),
                error('exportToPPTX:minInput','For new presentation filename to save to is required');
            end
            if nargin<2,
                fileName    = PPTX.fullName;
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
            PPTX.fullName       = fullName;
            
            %% Add some more useful information
            PPTX.revNumber      = PPTX.revNumber+1;
            PPTX.updatedDate    = datestr(now,'yyyy-mm-ddTHH:MM:SS');
            
            %% Commit changes to PPTX file
            PPTX.writePPTX();
            
            %% Output
            if nargout>0,
                fullName    = PPTX.fullName;
            end
            
        end
    end
    
    
    
    %% Supporting private methods
    methods (Access=private)
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function openExistingPPTX(PPTX)
            % Extract PPTX file to a temporary location
            
            unzip(PPTX.fullName,PPTX.tempName);
        end

        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function addTableCell(PPTX,aTc,cellData,varargin)
            % Add individual cell
            
            mi              = false(size(varargin));
            [vertAlign,mi]  = exportToPPTX.getPVPair(varargin,'Vert','',mi);
            [bCol,mi]       = exportToPPTX.getPVPair(varargin,'BackgroundColor',[],mi);
            [lnColVal,mi]   = exportToPPTX.getPVPair(varargin,'EdgeColor',[],mi);
            [lnWVal,mi]     = exportToPPTX.getPVPair(varargin,'LineWidth',[],mi);
            
            % Not checking unmached mi because there is another check in addTxBodyNode
            
            % Default
            showLn      = false;
            lnCol       = [0 0 0];
            lnStyle     = 'solid';
            lnW         = 1;
            
            switch lower(vertAlign),
                case 'top',
                    vAlign  = {'anchor','t'};
                case 'bottom',
                    vAlign  = {'anchor','b'};
                case 'middle',
                    vAlign  = {'anchor','ctr'}';
                case '',
                    vAlign  = {};
                otherwise,
                    error('exportToPPTX:badProperty','Bad property value found in VerticalAlignment');
            end
            
            bCol        = exportToPPTX.validateColor(bCol);
            
            lnColVal    = exportToPPTX.validateColor(lnColVal);
            if ~isempty(lnColVal),
                lnCol   = lnColVal;
                showLn  = true;
            end

            if ~isempty(lnWVal),
                lnW     = lnWVal;
                showLn  = true;
                if ~isnumeric(lnW) || numel(lnW)~=1,
                    error('exportToPPTX:badProperty','Bad property value found in LineWidth');
                end
            end
            
            if ~ischar(cellData),
                cellData        = num2str(cellData);
            end
            
            txBody  = exportToPPTX.addNode(PPTX.XML.Slide,aTc,'a:txBody');
            PPTX.addTxBodyNode(PPTX.XML.Slide,txBody,cellData,varargin{~mi});
            aTcPr   = exportToPPTX.addNode(PPTX.XML.Slide,aTc,'a:tcPr',vAlign);
            
            if showLn,
                % Loop over all four sides
                sidesTag    = {'a:lnL','a:lnR','a:lnT','a:lnB'};
                for iside=1:4,
                    aLn     = exportToPPTX.addNode(PPTX.XML.Slide,aTcPr,sidesTag{iside}',{'w',lnW*PPTX.CONST_PT_TO_EMU});
                    sFil    = exportToPPTX.addNode(PPTX.XML.Slide,aLn,'a:solidFill');
                    exportToPPTX.addNode(PPTX.XML.Slide,sFil,'a:srgbClr',{'val',lnCol});
                    if ~isempty(lnStyle),
                        exportToPPTX.addNode(PPTX.XML.Slide,aLn,'a:prstDash',{'val',lnStyle});
                    end
                end
            end
            
            if ~isempty(bCol),
                solidFill=exportToPPTX.addNode(PPTX.XML.Slide,aTcPr,'a:solidFill');
                exportToPPTX.addNode(PPTX.XML.Slide,solidFill,'a:srgbClr',{'val',bCol});
            end
        end
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function loadAndParseXMLFiles(PPTX)
            % Load files into persistent memory
            % These will be kept in memory for the duration of the PPTX construction
            PPTX.XML.TOC        = xmlread(fullfile(PPTX.tempName,'[Content_Types].xml'));
            PPTX.XML.App        = xmlread(fullfile(PPTX.tempName,'docProps','app.xml'));
            PPTX.XML.Core       = xmlread(fullfile(PPTX.tempName,'docProps','core.xml'));
            PPTX.XML.Pres       = xmlread(fullfile(PPTX.tempName,'ppt','presentation.xml'));
            PPTX.XML.PresRel    = xmlread(fullfile(PPTX.tempName,'ppt','_rels','presentation.xml.rels'));
            
            % Parse overall information
            PPTX.numSlides      = str2num(char(exportToPPTX.getNodeValue(PPTX.XML.App,'Slides')));
            PPTX.revNumber      = str2num(char(exportToPPTX.getNodeValue(PPTX.XML.Core,'cp:revision')));
            PPTX.title          = char(exportToPPTX.getNodeValue(PPTX.XML.Core,'dc:title'));
            PPTX.subject        = char(exportToPPTX.getNodeValue(PPTX.XML.Core,'dc:subject'));
            PPTX.author         = char(exportToPPTX.getNodeValue(PPTX.XML.Core,'dc:creator'));
            PPTX.description    = char(exportToPPTX.getNodeValue(PPTX.XML.Core,'dc:description'));
            pptDims(1)          = str2num(char(exportToPPTX.getNodeAttribute(PPTX.XML.Pres,'p:sldSz','cx')));
            pptDims(2)          = str2num(char(exportToPPTX.getNodeAttribute(PPTX.XML.Pres,'p:sldSz','cy')));
            PPTX.dimensions     = pptDims./PPTX.CONST_IN_TO_EMU;
            
            PPTX.createdDate    = char(exportToPPTX.getNodeValue(PPTX.XML.Core,'dcterms:created'));
            PPTX.updatedDate    = char(exportToPPTX.getNodeValue(PPTX.XML.Core,'dcterms:modified'));
            
            % Parse image/video type information
            fileExt             = exportToPPTX.getNodeAttribute(PPTX.XML.TOC,'Default','Extension');
            fileType            = exportToPPTX.getNodeAttribute(PPTX.XML.TOC,'Default','ContentType');
            imgIdx              = (strncmpi(fileType,'image',5));
            PPTX.imageTypes     = fileExt(imgIdx);
            vidIdx              = (strncmpi(fileType,'video',5));
            PPTX.videoTypes     = fileExt(vidIdx);
            
            % Parse slide information
            % Parsing is based on relationship XML file PresRel, which will contain
            % more IDs than presentation XML file itself
            allRIds         = exportToPPTX.getNodeAttribute(PPTX.XML.PresRel,'Relationship','Id');
            allTargs        = exportToPPTX.getNodeAttribute(PPTX.XML.PresRel,'Relationship','Target');
            allSlideIds     = exportToPPTX.getNodeAttribute(PPTX.XML.Pres,'p:sldId','id');
            allSlideRIds    = exportToPPTX.getNodeAttribute(PPTX.XML.Pres,'p:sldId','r:id');
            if ~isempty(allSlideRIds),
                % Presentation possibly contains some slides
                for irid=1:numel(allRIds),
                    % Find slide ID based on relationship ID
                    slideIdx    = find(strcmpi(allSlideRIds,allRIds{irid}));
                    
                    if ~isempty(slideIdx),
                        % There is a slide with a given relationship ID, save its information
                        PPTX.Slide(slideIdx).id     = str2num(allSlideIds{slideIdx});
                        PPTX.Slide(slideIdx).rId    = allSlideRIds{slideIdx};
                        
                        % Note: objId field is populated on demand when switchSlide is
                        % invoked.
                        
                        [d,fileName,fileExt]        = fileparts(allTargs{irid});
                        fileName                    = cat(2,fileName,fileExt);
                        PPTX.Slide(slideIdx).file   = fileName;
                    end
                end
                
                PPTX.lastSlideId    = max([PPTX.Slide.id]);
                allRIds2            = char(allRIds{:});
                PPTX.lastRId        = max(str2num(allRIds2(:,4:end)));
            else
                % Completely blank presentation
                PPTX.Slide          = struct([]);
                PPTX.lastSlideId    = 255;  % not starting with 0 because PPTX doesn't seem to do that
                allRIds2            = char(allRIds{:});
                PPTX.lastRId        = max(str2num(allRIds2(:,4:end)));
            end
            
            % Parse template information
            % Once again, parsing is based on relationship XML file PresRel, which
            % contains links and filenames of theme and master XML files
            mCnt        = 0;
            allTypes    = exportToPPTX.getNodeAttribute(PPTX.XML.PresRel,'Relationship','Type');
            if ~isempty(allRIds),
                % Check if presentation contains slideMaster and theme type slides
                for irid=1:numel(allRIds),
                    if ~isempty(strfind(allTypes{irid},'slideMaster')),
                        % We have at least one slide master
                        mCnt                        = mCnt+1;
                        PPTX.SlideMaster(mCnt).rId  = allRIds{irid};
                        
                        [d,fileName,fileExt]        = fileparts(allTargs{irid});
                        fileName                    = cat(2,fileName,fileExt);
                        PPTX.SlideMaster(mCnt).file = fileName;
                        
                        % Open slide master relationship file to find all slide layouts
                        fileRelsPath                = cat(2,PPTX.SlideMaster(mCnt).file,'.rels');
                        PPTX.XML.SlideMaster        = xmlread(fullfile(PPTX.tempName,'ppt','slideMasters','_rels',fileRelsPath));
                        
                        masterRIds          = exportToPPTX.getNodeAttribute(PPTX.XML.SlideMaster,'Relationship','Id');
                        masterTargs         = exportToPPTX.getNodeAttribute(PPTX.XML.SlideMaster,'Relationship','Target');
                        masterTypes         = exportToPPTX.getNodeAttribute(PPTX.XML.SlideMaster,'Relationship','Type');
                        layoutIdx           = find(~cellfun('isempty',strfind(masterTypes,'slideLayout')));
                        if ~isempty(layoutIdx),
                            % Loop over all available layouts
                            for ilay=1:numel(layoutIdx),
                                PPTX.SlideMaster(mCnt).Layout(ilay).rId     = masterRIds{layoutIdx(ilay)};
                                
                                [d,fileName,fileExt]                        = fileparts(masterTargs{layoutIdx(ilay)});
                                fileName                                    = cat(2,fileName,fileExt);
                                PPTX.SlideMaster(mCnt).Layout(ilay).file    = fileName;
                                
                                % Open slide layout file and get its name
                                PPTX.XML.SlideLayout    = xmlread(fullfile(PPTX.tempName,'ppt','slideLayouts',PPTX.SlideMaster(mCnt).Layout(ilay).file));
                                %layoutName             = exportToPPTX.getNodeAttribute(PPTX.XML.SlideLayout,'p:cSld','name'); % ISSUE: does not work when there is only one node found
                                cSldNode                = exportToPPTX.findNode(PPTX.XML.SlideLayout,'p:cSld');
                                if cSldNode.length>0,
                                    PPTX.SlideMaster(mCnt).Layout(ilay).name   = char(cSldNode.getAttribute('name'));
                                else
                                    PPTX.SlideMaster(mCnt).Layout(ilay).name   = 'Error Getting Layout Name';
                                end
                                
                                % Get all placeholders on a given layout slide
                                nvSpNodes     = exportToPPTX.findNode(PPTX.XML.SlideLayout,'p:sp');
                                if ~isempty(nvSpNodes),
                                    if any(strfind(class(nvSpNodes),'DeferredElementImpl')),
                                        numNodes        = 1;
                                    else
                                        numNodes        = nvSpNodes.getLength;
                                    end
                                    
                                    for ielem=0:numNodes-1,
                                        if any(strfind(class(nvSpNodes),'DeferredElementImpl')),
                                            processNode     = nvSpNodes;
                                        else
                                            processNode     = nvSpNodes.item(ielem);
                                        end
                                        
                                        retType     = exportToPPTX.getNodeAttribute(processNode,'p:ph','type');
                                        retIdx      = exportToPPTX.getNodeAttribute(processNode,'p:ph','idx');
                                        retName     = exportToPPTX.getNodeAttribute(processNode,'p:cNvPr','name');
                                        % For position, search within a:xfrm node
                                        xfrmNode    = exportToPPTX.findNode(processNode,'a:xfrm');
                                        if ~isempty(xfrmNode)
                                            posX    = exportToPPTX.getNodeAttribute(xfrmNode,'a:off','x');
                                            posY    = exportToPPTX.getNodeAttribute(xfrmNode,'a:off','y');
                                            posW    = exportToPPTX.getNodeAttribute(xfrmNode,'a:ext','cx');
                                            posH    = exportToPPTX.getNodeAttribute(xfrmNode,'a:ext','cy');
                                        else
                                            posX    = []; posY = []; posW = []; posH = [];
                                        end
                                        if isempty(retType), retType = {''}; end
                                        if isempty(retIdx), retIdx = {''}; end
                                        if isempty(retName), retName = {''}; end
                                        if isempty(posX), posX = NaN; else, posX = str2double(posX); end
                                        if isempty(posY), posY = NaN; else, posY = str2double(posY); end
                                        if isempty(posW), posW = NaN; else, posW = str2double(posW); end
                                        if isempty(posH), posH = NaN; else, posH = str2double(posH); end
                                        
                                        PPTX.SlideMaster(mCnt).Layout(ilay).ph(ielem+1)        = retType;
                                        PPTX.SlideMaster(mCnt).Layout(ilay).idx(ielem+1)       = retIdx;
                                        PPTX.SlideMaster(mCnt).Layout(ilay).place(ielem+1)     = retName;
                                        PPTX.SlideMaster(mCnt).Layout(ilay).position(ielem+1)  = {[posX posY posW posH]};
                                    end
                                else
                                    PPTX.SlideMaster(mCnt).Layout(ilay).ph          = {};
                                    PPTX.SlideMaster(mCnt).Layout(ilay).idx         = {};
                                    PPTX.SlideMaster(mCnt).Layout(ilay).place       = {};
                                    PPTX.SlideMaster(mCnt).Layout(ilay).position    = {};
                                end
                                
                            end
                            PPTX.SlideMaster(mCnt).numLayout   = numel(layoutIdx);
                        else
                            error('exportToPPTX:badPPTX','Invalid PowerPoint presentation: missing at least one slide layout');
                        end
                        
                        % Open theme file and get slide master name
                        themeIdx            = find(~cellfun('isempty',strfind(masterTypes,'theme')));
                        if ~isempty(themeIdx),
                            PPTX.SlideMaster(mCnt).Theme.rId    = masterRIds{themeIdx(1)};
                            
                            [d,fileName,fileExt]                = fileparts(masterTargs{themeIdx(1)});
                            fileName                            = cat(2,fileName,fileExt);
                            PPTX.SlideMaster(mCnt).Theme.file   = fileName;
                            
                            % Open slide layout file and get its name
                            PPTX.XML.SlideTheme     = xmlread(fullfile(PPTX.tempName,'ppt','theme',PPTX.SlideMaster(mCnt).Theme.file));
                            aThemeNode              = exportToPPTX.findNode(PPTX.XML.SlideTheme,'a:theme');
                            if aThemeNode.length>0,
                                PPTX.SlideMaster(mCnt).name     = char(aThemeNode.getAttribute('name'));
                            else
                                PPTX.SlideMaster(mCnt).name     = 'Error Getting Master Name';
                            end
                        else
                            PPTX.SlideMaster(mCnt).Theme   = [];
                        end
                        
                    elseif ~isempty(strfind(allTypes{irid},'theme')),
                        % Nothing to do here, theme files associated with each master
                        % layout have been processed inside master layout if clause
                        [d,themeName]   = fileparts(allTargs{irid});
                        PPTX.themeNum   = max(sscanf(themeName,'theme%d'),PPTX.themeNum);
                        
                    elseif ~isempty(strfind(allTypes{irid},'notesMaster')),
                        % We have at least one notes master, 
                        PPTX.NotesMaster.rId    = allRIds{irid};
                        
                        [d,fileName,fileExt]    = fileparts(allTargs{irid});
                        fileName                = cat(2,fileName,fileExt);
                        PPTX.NotesMaster.file   = fileName;
                        
                    end
                end
            end
            
            % No templates defined
            if mCnt==0,
                error('exportToPPTX:badPPTX','Invalid PowerPoint presentation: missing at least one slide master');
            else
                PPTX.numMasters     = mCnt;
            end

            % Simple file integrity check
            if PPTX.numSlides~=numel(PPTX.Slide),
                warning('exportToPPTX:badPPTX','Badly formed PPTX. Number of slides is corrupted');
                PPTX.numSlides      = numel(PPTX.Slide);
            end
            
            % If openning existing PPTX with slides, load last one
            if PPTX.numSlides>0,
                PPTX                = PPTX.switchSlide(PPTX.numSlides);
            else
                PPTX.currentSlide   = 0;
            end
        end
        
 
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function [mNum,lNum] = parseMasterLayoutNumber(PPTX)
            % Parse out master and layout slides numbers from the currently loaded
            % slide XML files
            mNum        = 0;
            lNum        = 0;
            
            masterTargs = exportToPPTX.getNodeAttribute(PPTX.XML.SlideRel,'Relationship','Target');
            masterTypes = exportToPPTX.getNodeAttribute(PPTX.XML.SlideRel,'Relationship','Type');
            layoutIdx   = find(~cellfun('isempty',strfind(masterTypes,'slideLayout')));
            if ~isempty(layoutIdx),
                [d,fileName,fileExt]    = fileparts(masterTargs{layoutIdx(1)});
                fileName                = cat(2,fileName,fileExt);
                for imast=1:PPTX.numMasters,
                    layMatch    = find(strcmpi({PPTX.SlideMaster(imast).Layout.file},fileName));
                    if ~isempty(layMatch),
                        mNum    = imast;
                        lNum    = layMatch;
                    end
                end
            else
                error('exportToPPTX:badPPTX','Invalid PowerPoint presentation: missing at least one slide layout');
            end
            if mNum==0 || lNum==0,
                error('exportToPPTX:badPPTX','Invalid PowerPoint presentation: master layout(s) and slide layout(s) are not properly linked');
            end
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function success = writePPTX(PPTX)
            % Commit changes to current slide, if any
            if isfield(PPTX.XML,'Slide') && ~isempty(PPTX.XML.Slide),
                fileName    = PPTX.Slide(PPTX.currentSlide).file;
                xmlwrite(fullfile(PPTX.tempName,'ppt','slides',fileName),PPTX.XML.Slide);
                xmlwrite(fullfile(PPTX.tempName,'ppt','slides','_rels',cat(2,fileName,'.rels')),PPTX.XML.SlideRel);
            end
            
            % Do last minute updates to the file
            exportToPPTX.setNodeValue(PPTX.XML.Core,'dcterms:created',PPTX.createdDate);
            exportToPPTX.setNodeValue(PPTX.XML.Core,'dcterms:modified',PPTX.updatedDate);
            
            % Update overall PPTX information
            exportToPPTX.setNodeValue(PPTX.XML.Core,'cp:revision',PPTX.revNumber);
            exportToPPTX.setNodeValue(PPTX.XML.App,'Slides',PPTX.numSlides);
            exportToPPTX.setNodeValue(PPTX.XML.Core,'dc:title',PPTX.title);
            exportToPPTX.setNodeValue(PPTX.XML.Core,'dc:creator',PPTX.author);
            exportToPPTX.setNodeValue(PPTX.XML.Core,'dc:subject',PPTX.subject);
            exportToPPTX.setNodeValue(PPTX.XML.Core,'dc:description',PPTX.description);
            exportToPPTX.setNodeValue(PPTX.XML.Core,'cp:lastModifiedBy',PPTX.author);
            
            % Commit all changes to XML files
            xmlwrite(fullfile(PPTX.tempName,'[Content_Types].xml'),PPTX.XML.TOC);
            xmlwrite(fullfile(PPTX.tempName,'docProps','app.xml'),PPTX.XML.App);
            xmlwrite(fullfile(PPTX.tempName,'docProps','core.xml'),PPTX.XML.Core);
            xmlwrite(fullfile(PPTX.tempName,'ppt','presentation.xml'),PPTX.XML.Pres);
            xmlwrite(fullfile(PPTX.tempName,'ppt','_rels','presentation.xml.rels'),PPTX.XML.PresRel);
            
            % Write all files to ZIP (aka PPTX) file
            % 1) Zip temp directory with all XML files to tempName.zip
            % 2) Rename tempName.zip to fullName (this leaves only .pptx extension in the filename)
            zip(PPTX.tempName,'*',PPTX.tempName);
            [success,msg] = movefile(cat(2,PPTX.tempName,'.zip'),PPTX.fullName);
            if ~success,
                error('exportToPPTX:saveFail','PPTX file cannot be written to: %s.',msg);
            end
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function addImageTypeSupport(PPTX,imgExt)
            exportToPPTX.addNode(PPTX.XML.TOC,'Types','Default',{'Extension',imgExt,'ContentType',cat(2,'image/',imgExt)});
            PPTX.imageTypes     = cat(2,PPTX.imageTypes,{imgExt});
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function addVideoTypeSupport(PPTX,vidExt)
            exportToPPTX.addNode(PPTX.XML.TOC,'Types','Default',{'Extension',vidExt,'ContentType',cat(2,'video/',vidExt)});
            PPTX.videoTypes     = cat(2,PPTX.videoTypes,{vidExt});
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function addTxBodyNode(PPTX,fileXML,txBody,inputText,varargin)
            % Input: PPTX (for constants), fileXML to modify, rootNode XML node to attach text to, regular text string
            % Output: modified XML file
            
            mi                  = false(size(varargin));
            [horizAlign,mi]     = exportToPPTX.getPVPair(varargin,'Horiz','',mi);
            [vertAlign,mi]      = exportToPPTX.getPVPair(varargin,'Vert','',mi);
            [fontAngle,mi]      = exportToPPTX.getPVPair(varargin,'FontAngle','',mi);
            [fontWeight,mi]     = exportToPPTX.getPVPair(varargin,'FontWeight','',mi);
            [fCol,mi]           = exportToPPTX.getPVPair(varargin,'Color',[],mi);
            [fSizeVal,mi]       = exportToPPTX.getPVPair(varargin,'FontSize',[],mi);
            [fNameVal,mi]       = exportToPPTX.getPVPair(varargin,'FontName',[],mi);
            [onClick,mi]        = exportToPPTX.getPVPair(varargin,'OnClick',[],mi);
            [allowMarkdown,mi]  = exportToPPTX.getPVPair(varargin,'markdown',true,mi);

            if any(~mi)
                error('exportToPPTX:badProperty','Unrecognized property %s',varargin{find(~mi,1)});
            end

            switch lower(horizAlign),
                case 'left',
                    hAlign  = {'algn','l'};
                case 'right',
                    hAlign  = {'algn','r'};
                case 'center',
                    hAlign  = {'algn','ctr'};
                case '',
                    hAlign  = {};
                otherwise,
                    error('exportToPPTX:badProperty','Bad property value found in HorizontalAlignment');
            end
            
            switch lower(vertAlign),
                case 'top',
                    vAlign  = {'anchor','t'};
                case 'bottom',
                    vAlign  = {'anchor','b'};
                case 'middle',
                    vAlign  = {'anchor','ctr'}';
                case '',
                    vAlign  = {};
                otherwise,
                    error('exportToPPTX:badProperty','Bad property value found in VerticalAlignment');
            end
            
            switch lower(fontAngle),
                case {'italic','oblique'},
                    fItal   = {'i','1'};
                case 'normal',
                    fItal   = {'i','0'};
                case '',
                    fItal   = {};
                otherwise,
                    error('exportToPPTX:badProperty','Bad property value found in FontAngle');
            end
            
            switch lower(fontWeight),
                case {'bold','demi'},
                    fBold   = {'b','1'};
                case {'normal','light'},
                    fBold   = {'b','0'};
                case '',
                    fBold   = {};
                otherwise,
                    error('exportToPPTX:badProperty','Bad property value found in FontWeight');
            end
            
            fCol        = exportToPPTX.validateColor(fCol);

            if ~isnumeric(fSizeVal),
                error('exportToPPTX:badProperty','Bad property value found in FontSize');
            elseif ~isempty(fSizeVal),
                fSize   = {'sz',fSizeVal*PPTX.CONST_FONT_PX_TO_PPTX};
            else
                fSize   = {};
            end

            if isempty(fNameVal),
                fName   = '';
            elseif ~ischar(fNameVal),
                error('exportToPPTX:badProperty','Bad property value found in FontName');
            elseif strncmpi(fNameVal,'FixedWidth',10),
                fName   = get(0,'FixedWidthFontName');
            else
                fName   = fNameVal;
            end
            
            
            % Formatting notes:
            %   p:sp                                input to this function
            %       p:nvSpPr                        non-visual shape props (handled outside)
            %       p:spPr                          visual shape props (also handled outside)
            %       p:txBody                  <---- created by this subfunction
            %           a:bodyPr
            %           a:lstStyle
            %           a:p                         paragraph node (one in each ipara loop)
            %               a:pPr                   paragraph properties (text alignment, bullet (a:buChar), number (a:buAutoNum) )
            %               a:r                     run node (one in each irun loop)
            %                   a:rPr               run formatting (bold, italics, font size, etc.)
            %                       a:solidFill     text coloring
            %                           a:srfbClr
            %                   a:t                 text node
            
            % Bad news for text auto-fit: font sizes and line spacing are determined by
            % PowerPoint based on fonts measurements. If fontScale and lnSpcReduction
            % are filled out in the OpenXML file then PowerPoint doesn't re-render
            % those textboxes and simply uses OpenXML values. This means that in order
            % for auto-fit to work in this tool font measurements have to be done in
            % here.
            bodyPr      = exportToPPTX.addNode(fileXML,txBody,'a:bodyPr',{'wrap','square','rtlCol','0',vAlign{:}});
            %             exportToPPTX.addNode(fileXML,bodyPr,'a:normAutofit');%,{'fontScale','50.000%','lnSpcReduction','50.000%'}); % autofit flag
            exportToPPTX.addNode(fileXML,txBody,'a:lstStyle');
            
            % Break text into paragraphs and add each paragraph as a separate a:p node
            if ischar(inputText),
                paraText    = regexp(inputText,'\n','split');
            else
                paraText    = inputText;
            end
            numParas        = numel(paraText);
            
            defMargin   = 0.3;  % inches
            
            for ipara=1:numParas,
                ap          = exportToPPTX.addNode(fileXML,txBody,'a:p');
                pPr         = exportToPPTX.addNode(fileXML,ap,'a:pPr',hAlign);
                
                % Check for paragraph level markdown: bulletted lists, numbered lists
                
                if ~isempty(paraText{ipara}),
                    
                    % Make a copy of paragraph text for modification
                    addParaText = paraText{ipara};
                    
                    % Re-init here and modify if further down than one level
                    useMargin   = defMargin;
                    
                    % Check for tabs before removing them
                    if double(paraText{ipara}(1))==9,   % 9=tab characeter
                        firstNonTabIdx  = find(double(paraText{ipara})~=9,1);
                        numTabs         = firstNonTabIdx-1;
                        useMargin       = useMargin*(numTabs+1);
                        exportToPPTX.setNodeAttribute(pPr,{'lvl',numTabs});
                    end
                    
                    % Clean up (remove extra whitespaces)
                    addParaText     = strtrim(addParaText);
                    
                    % Bullet and/or number are allowed only if there is a space between
                    % dash and text that follows (this guards against negative numbers
                    % showing up as bullets)
                    if allowMarkdown && length(paraText{ipara})>=2 && isequal(paraText{ipara}(1:2),'- '),
                        % if paraText{ipara}(1)=='-' && numParas>1 && ...
                        %         (   (~isempty(paraText{max(ipara-1,1)}) && ipara-1>=1 && paraText{max(ipara-1,1)}(1)=='-') || ...
                        %             (~isempty(paraText{min(ipara+1,numParas)}) && ipara+1<=numParas && paraText{min(ipara+1,numParas)}(1)=='-') ),
                        addParaText(1)      = [];   % remove the actual character
                        exportToPPTX.setNodeAttribute(pPr,{'marL',useMargin*PPTX.CONST_IN_TO_EMU,'indent',-defMargin*PPTX.CONST_IN_TO_EMU});
                        exportToPPTX.addNode(fileXML,pPr,'a:buChar',{'char',char(8226)});   % TODO: add character control here
                    end
                    
                    if allowMarkdown && length(paraText{ipara})>=2 && isequal(paraText{ipara}(1:2),'# '),
                        % if paraText{ipara}(1)=='#' && numParas>1 && ...
                        %         (   (~isempty(paraText{max(ipara-1,1)}) && paraText{max(ipara-1,1)}(1)=='#') || ...
                        %             (~isempty(paraText{min(ipara+1,numParas)}) && paraText{min(ipara+1,numParas)}(1)=='#') ),
                        addParaText(1)      = [];   % remove the actual character
                        exportToPPTX.setNodeAttribute(pPr,{'marL',useMargin*PPTX.CONST_IN_TO_EMU,'indent',-defMargin*PPTX.CONST_IN_TO_EMU});
                        exportToPPTX.addNode(fileXML,pPr,'a:buAutoNum',{'type','arabicPeriod'}); % TODO: add numeral control here
                    end
                    
                    % Clean up again in case there were more spaces between markdown
                    % character and beginning of the text
                    addParaText     = strtrim(addParaText);
                    
                    % Check for run level markdown: bold, italics, underline
                    if allowMarkdown,
                        [fmtStart,fmtStop,tokStr,splitStr] = regexpi(addParaText,'\<(*{2}|*{1}|_{1})([a-z0-9]+[()\- ,:.\w]*)(?=\1)\1\>','start','end','tokens','split');
                    else
                        fmtStart    = [];
                        fmtStop     = [];
                        tokStr      = [];
                        splitStr    = {addParaText};
                    end
                    
                    allRuns     = cat(2,1,reshape(cat(1,fmtStart,fmtStop),1,[]));
                    runTypes    = cat(2,-1,reshape(cat(1,1:numel(fmtStart),-(2:numel(splitStr))),1,[]));    % setup resets (negative numbers)
                    numRuns     = numel(allRuns);
                    
                    for irun=1:numRuns,
                        fItalR      = fItal;
                        fBoldR      = fBold;
                        fUnderR     = {};       % u="sng"
                        
                        if runTypes(irun)<0,
                            runText     = splitStr{-runTypes(irun)};
                        else
                            runText     = tokStr{runTypes(irun)}{2};
                            runFormat   = tokStr{runTypes(irun)}{1};
                            if strcmpi(runFormat,'*'), fItalR = {'i','1'}; end
                            if strcmpi(runFormat,'**'), fBoldR = {'b','1'}; end
                            if strcmpi(runFormat,'_'), fUnderR = {'u','sng'}; end
                        end
                        
                        if allowMarkdown
                            runText     = strrep(runText,'\**','**');
                            runText     = strrep(runText,'\*','*');
                            runText     = strrep(runText,'\_','_');
                            runText     = strrep(runText,'\\','\');
                        end
                        
                        ar          = exportToPPTX.addNode(fileXML,ap,'a:r');    % run node
                        rPr         = exportToPPTX.addNode(fileXML,ar,'a:rPr',{'lang','en-US','dirty','0','smtClean','0',fItalR{:},fBoldR{:},fUnderR{:},fSize{:}});
                        if ~isempty(fCol),
                            sf      = exportToPPTX.addNode(fileXML,rPr,'a:solidFill');
                            exportToPPTX.addNode(fileXML,sf,'a:srgbClr',{'val',fCol});
                        end
                        if ~isempty(fName),
                            exportToPPTX.addNode(fileXML,rPr,'a:latin',{'typeface',fName});
                            exportToPPTX.addNode(fileXML,rPr,'a:cs',{'typeface',fName});
                        end
                        
                        PPTX.addLink(onClick,rPr);
                        
                        at          = exportToPPTX.addNode(fileXML,ar,'a:t');
                        exportToPPTX.addNodeValue(fileXML,at,runText);
                    end
                    
                    
                else
                    % Empty paragraph, do nothing
                    
                end
                
                % Add end paragraph
                endParaRPr      = exportToPPTX.addNode(fileXML,ap,'a:endParaRPr',{'lang','en-US','dirty','0'});
                if ~isempty(fName),
                    exportToPPTX.addNode(fileXML,endParaRPr,'a:latin',{'typeface',fName});
                    exportToPPTX.addNode(fileXML,endParaRPr,'a:cs',{'typeface',fName});
                end
                
                
            end
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function addLink(PPTX,onClick,cNvPr)
            % Add XML to link an object
            
            if isempty(onClick)
                return
            end
            
            if ~isempty(onClick) && (isnumeric(onClick) && (onClick<1 || onClick>PPTX.numSlides || numel(onClick)>1))
                % Error condition
                error('exportToPPTX:badInput','OnClick slide number must be between 1 and the total number of slides');
            end
            
            if ~isempty(onClick)
                PPTX.Slide(PPTX.currentSlide).objId    = PPTX.Slide(PPTX.currentSlide).objId+1;
                rId     = sprintf('rId%d',PPTX.Slide(PPTX.currentSlide).objId);
                if isnumeric(onClick) % numeric link signifies a jump to slide number
                    actionTags  = {'action','ppaction://hlinksldjump'};
                    relTags     = {'Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', ...
                        'Target',PPTX.Slide(onClick).file};
                else % non-numeric link is an external (URL or file) link
                    actionTags  = {'action','ppaction://hlinkpres'};
                    relTags     = {'Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', ...
                        'TargetMode','External', ...
                        'Target',onClick};
                end
                exportToPPTX.addNode(PPTX.XML.Slide,cNvPr,'a:hlinkClick',{'r:id',rId,actionTags{:}});
                nodeAttribute   = exportToPPTX.getNodeAttribute(PPTX.XML.SlideRel,'Relationship','Id');
                if ~any(strcmpi(nodeAttribute,rId))
                    exportToPPTX.addNode(PPTX.XML.SlideRel,'Relationships','Relationship', ...
                        {'Id',rId,relTags{:}});
                end
            end
        end

        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function initBlankPPTX(PPTX)
            % Creates valid PPTX structure with 0 slides
            
            % Folder structure
            mkdir(fullfile(PPTX.tempName,'_rels'));
            mkdir(fullfile(PPTX.tempName,'docProps'));
            mkdir(fullfile(PPTX.tempName,'ppt'));
            mkdir(fullfile(PPTX.tempName,'ppt','_rels'));
            mkdir(fullfile(PPTX.tempName,'ppt','media'));
            mkdir(fullfile(PPTX.tempName,'ppt','notesSlides'));
            mkdir(fullfile(PPTX.tempName,'ppt','notesSlides','_rels'));
            mkdir(fullfile(PPTX.tempName,'ppt','slideLayouts'));
            mkdir(fullfile(PPTX.tempName,'ppt','slideLayouts','_rels'));
            mkdir(fullfile(PPTX.tempName,'ppt','slideMasters'));
            mkdir(fullfile(PPTX.tempName,'ppt','slideMasters','_rels'));
            mkdir(fullfile(PPTX.tempName,'ppt','slides'));
            mkdir(fullfile(PPTX.tempName,'ppt','slides','_rels'));
            mkdir(fullfile(PPTX.tempName,'ppt','theme'));

            % Absolutely neccessary files (by trial and error, because PresentationML
            % documentation says less files are required than what seems to be the
            % case)
            %
            % \[Content_Types].xml
            % \_rels\.rels
            % \docProps\app.xml
            % \docProps\core.xml
            % \ppt\presentation.xml
            % \ppt\presProps.xml
            % \ppt\tableStyles.xml
            % \ppt\viewProps.xml
            % \ppt\_rels\presentation.xml.rels
            % \ppt\notesMasters\notesMaster1.xml
            % \ppt\notesMasters\_rels\notesMaster1.xml.rels
            % \ppt\slideLayouts\slideLayout1.xml
            % \ppt\slideLayouts\_rels\slideLayout1.xml.rels
            % \ppt\slideMasters\slideMaster1.xml
            % \ppt\slideMasters\_rels\slideMaster1.xml.rels
            % \ppt\theme\theme1.xml
            % \ppt\theme\theme2.xml
            
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
                '<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
                '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
                '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
                '<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>'
                '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
                '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'
                '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>'
                '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>'
                '</Types>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'[Content_Types].xml'),fileContent) & retCode;
            
            % \_rels\.rels
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
                '</Relationships>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'_rels','.rels'),fileContent) & retCode;
            
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
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'docProps','app.xml'),fileContent) & retCode;
            
            % \docProps\core.xml
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                ['<dc:title>' PPTX.title '</dc:title>']
                ['<dc:subject>' PPTX.subject '</dc:subject>']
                ['<dc:creator>' PPTX.author '</dc:creator>']
                '<cp:keywords></cp:keywords>'
                ['<dc:description>' PPTX.description '</dc:description>']
                ['<cp:lastModifiedBy>' PPTX.author '</cp:lastModifiedBy>']
                '<cp:revision>0</cp:revision>'
                ['<dcterms:created xsi:type="dcterms:W3CDTF">' PPTX.createdDate '</dcterms:created>']
                ['<dcterms:modified xsi:type="dcterms:W3CDTF">' PPTX.createdDate '</dcterms:modified>']
                '</cp:coreProperties>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'docProps','core.xml'),fileContent) & retCode;
            
            % \ppt\presentation.xml
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">'
                '<p:sldMasterIdLst>'
                '<p:sldMasterId id="2147483663" r:id="rId2"/>'
                '</p:sldMasterIdLst>'
                '<p:notesMasterIdLst>'
                '<p:notesMasterId r:id="rId3"/>'
                '</p:notesMasterIdLst>'
                '<p:sldIdLst>'
                '</p:sldIdLst>'
                ['<p:sldSz cx="' int2str(round(PPTX.dimensions(1).*PPTX.CONST_IN_TO_EMU)) '" cy="' int2str(round(PPTX.dimensions(2).*PPTX.CONST_IN_TO_EMU)) '" type="screen4x3"/>']
                ['<p:notesSz cx="' int2str(round(PPTX.dimensions(1).*PPTX.CONST_IN_TO_EMU)) '" cy="' int2str(round(PPTX.dimensions(2).*PPTX.CONST_IN_TO_EMU)) '"/>']
                '</p:presentation>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','presentation.xml'),fileContent) & retCode;
            
            % \ppt\_rels\presentation.xml.rels
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>'
                '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>'
                '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>'
                '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>'
                '</Relationships>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','_rels','presentation.xml.rels'),fileContent) & retCode;
            
            % \ppt\presProps.xml
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','presProps.xml'),fileContent) & retCode;
            
            % \ppt\tableStyles.xml
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','tableStyles.xml'),fileContent) & retCode;
            
            % \ppt\viewProps.xml
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
                '<p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>'
                '<p:slideViewPr>'
                '<p:cSldViewPr showGuides="1">'
                '<p:cViewPr varScale="1">'
                '<p:scale><a:sx n="65" d="100"/><a:sy n="65" d="100"/></p:scale>'
                '<p:origin x="-1452" y="-114"/>'
                '</p:cViewPr>'
                '<p:guideLst/>'
                '</p:cSldViewPr>'
                '</p:slideViewPr>'
                '<p:notesTextViewPr>'
                '<p:cViewPr>'
                '<p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale>'
                '<p:origin x="0" y="0"/>'
                '</p:cViewPr>'
                '</p:notesTextViewPr>'
                '<p:gridSpacing cx="78028800" cy="78028800"/>'
                '</p:viewPr>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','viewProps.xml'),fileContent) & retCode;
                        
            % \ppt\slideLayouts\slideLayout1.xml
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1" userDrawn="1">'
                '<p:cSld name="Blank Slide">'
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
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','slideLayouts','slideLayout1.xml'),fileContent) & retCode;
            
            % \ppt\slideLayouts\_rels\slideLayout1.xml.rels
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
                '</Relationships>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','slideLayouts','_rels','slideLayout1.xml.rels'),fileContent) & retCode;
            
            % \ppt\slideMasters\slideMaster1.xml
            if ~isempty(PPTX.bgColor),
                bgContent   = { ...
                    '<p:bg>'
                    '<p:bgPr>'
                    '<a:solidFill>'
                    ['<a:srgbClr val="' sprintf('%s',dec2hex(round(PPTX.bgColor*255),2).') '"/>']
                    '</a:solidFill>'
                    '<a:effectLst/>'
                    '</p:bgPr>'
                    '</p:bg>'
                    };
            else
                bgContent   = { ...
                    '<p:bg>'
                    '<p:bgRef idx="1001">'
                    '<a:schemeClr val="bg1"/>'
                    '</p:bgRef>'
                    '</p:bg>'
                    };
            end
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
                '<p:cSld>'
                [bgContent{:}]
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
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','slideMasters','slideMaster1.xml'),fileContent) & retCode;
            
            % \ppt\slideMasters\_rels\slideMaster1.xml.rels
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
                '</Relationships>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','slideMasters','_rels','slideMaster1.xml.rels'),fileContent) & retCode;
            
            % \ppt\theme\theme1.xml
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="exportToPPTX Blank Theme">'
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
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','theme','theme1.xml'),fileContent) & retCode;
            
            % \ppt\theme\theme2.xml
            % \ppt\notesMasters\notesMaster1.xml
            % \ppt\notesMasters\_rels\notesMaster1.xml.rels
            retCode     = PPTX.initNotesMaster(2) & retCode;
            
            if ~retCode,
                error('Error creating temporary PPTX files');
            end
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function retCode = initNotesMaster(PPTX,themeNum)
            % Need a special function for notes because there is no gurantee that opened
            % PPTX will have notes master. In fact, it appears that default themes do not
            % have notes master.
            
            themeName   = sprintf('theme%d.xml',themeNum);
            
            mkdir(fullfile(PPTX.tempName,'ppt','notesMasters'));
            mkdir(fullfile(PPTX.tempName,'ppt','notesMasters','_rels'));
            
            % \ppt\theme\theme[themeNum].xml
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="exportToPPTX Blank Theme">'
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
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','theme',themeName),fileContent);
            
            % \ppt\notesMasters\notesMaster1.xml
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
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
                '<p:notesStyle>'
                '<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
                '<a:defRPr sz="1200" kern="1200">'
                '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
                '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>'
                '</a:defRPr>'
                '</a:lvl1pPr>'
                '</p:notesStyle>'
                '</p:notesMaster>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','notesMasters','notesMaster1.xml'),fileContent) & retCode;
            
            % \ppt\notesMasters\_rels\notesMaster1.xml.rels
            fileContent    = { ...
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/' themeName '"/>']
                '</Relationships>'};
            retCode     = exportToPPTX.writeTextFile(fullfile(PPTX.tempName,'ppt','notesMasters','_rels','notesMaster1.xml.rels'),fileContent) & retCode;
            
        end
    end
    
    
    
    %% Supporting static private methods
    methods (Static,Access=private)
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function [propValue,mi] = getPVPair(inputArr,propName,defValue,mi)
            % Get Property-Value Pair from varargin cell-array
            % Inputs: input cell array (varargin), property name, default value
            % Output: value specified on the input, otherwise default value
            
            if numel(inputArr)>=2 && any(strncmpi(inputArr,propName,length(propName))),
                % If multiple identical properties present, only value for the first one is returned
                % This provides the ability to override the list of properties by prepending varargin
                idx         = find(strncmpi(inputArr,propName,length(propName)));
                propValue   = inputArr{idx(1)+1};
                mi(idx)     = true;
                mi(idx+1)   = true;
            else
                propValue   = defValue;
            end
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function colorStr = validateColor(bCol)
            if ~isempty(bCol)
                if isnumeric(bCol) && numel(bCol)==3,
                    rgbCol  = bCol;
                elseif ischar(bCol) && numel(bCol)==1,
                    switch bCol,
                        case 'b', % blue
                            rgbCol  = [0 0 1];
                        case 'g', % green
                            rgbCol  = [0 1 0];
                        case 'r', % red
                            rgbCol  = [1 0 0];
                        case 'c', % cyan
                            rgbCol  = [0 1 1];
                        case 'm', % magenta
                            rgbCol  = [1 0 1];
                        case 'y', % yellow
                            rgbCol  = [1 1 0];
                        case 'k', % black
                            rgbCol  = [0 0 0];
                        case 'w', % white
                            rgbCol  = [1 1 1];
                        otherwise,
                            error('exportToPPTX:badProperty','Unknown color code');
                    end
                else
                    error('exportToPPTX:badProperty','Bad color property value found. Color must be either a three element RGB value or a single color character');
                end
                
                colorStr    = sprintf('%s',dec2hex(round(rgbCol*255),2).');
            else
                colorStr    = '';
            end
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function retCode = writeTextFile(fileName,fileContent)
            % Inputs: name of the file, file contents (cell array)
            % Outputs: 1 if file was successfully written, 0 otherwise
            
            fId     = fopen(fileName,'w','n','UTF-8');
            if fId~=-1,
                fprintf(fId,'%s\n',fileContent{:});
                fclose(fId);
                retCode = 1;
            else
                retCode = 0;
            end
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function foundNode = findNode(theNode,searchNodeName)
            foundNodes  = theNode.getElementsByTagName(searchNodeName);
            if foundNodes.getLength>1,
                foundNode   = foundNodes;
            else
                foundNode   = foundNodes.item(0);
            end
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function nodeAttribute = getAllAttribute(theNode,attribName)
            nodeAttribute   = {};
            
            if strcmpi(class(theNode),'org.apache.xerces.dom.DeferredElementImpl'),
                attribValue     = char(theNode.getAttribute(attribName));
                if ~isempty(attribValue),
                    nodeAttribute   = cat(2,nodeAttribute,{attribValue});
                end
            end
            
            for inode=0:theNode.getLength,
                if ~isempty(theNode.item(inode)),
                    if strcmpi(class(theNode),'org.apache.xerces.dom.DeferredElementImpl') || ...
                            strcmpi(class(theNode),'org.apache.xerces.dom.DeferredDocumentImpl'),
                        nodeAttribute   = cat(2,nodeAttribute,exportToPPTX.getAllAttribute(theNode.item(inode),attribName));
                    end
                end
            end
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function nodeValue = getNodeValue(xmlData,searchNodeName)
            nodeValue = {};
            foundNode = exportToPPTX.findNode(xmlData,searchNodeName);
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
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function setNodeValue(xmlData,searchNodeName,newValue)
            if ~isempty(newValue),
                % two possible inputs: tag name or node handle
                if ischar(searchNodeName),
                    foundNode = exportToPPTX.findNode(xmlData,searchNodeName);
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
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function addNodeValue(xmlData,searchNodeName,newValue)
            if ~isempty(newValue),
                % two possible inputs: tag name or node handle
                if ischar(searchNodeName),
                    foundNode = exportToPPTX.findNode(xmlData,searchNodeName);
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
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function nodeAttribute = getNodeAttribute(xmlData,searchNodeName,searchAttribute)
            nodeAttribute   = {};
            foundNode = exportToPPTX.findNode(xmlData,searchNodeName);
            if ~isempty(foundNode),
                if foundNode.getLength>1,
                    for inode=0:foundNode.getLength-1,
                        nodeAttribute{inode+1} = char(foundNode.item(inode).getAttribute(searchAttribute));
                    end
                else
                    nodeAttribute{1} = char(foundNode.getAttribute(searchAttribute));
                end
            end
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function setNodeAttribute(newNode,allAttribs)
            % foundNode = exportToPPTX.findNode(xmlData,searchNodeName);
            % if ~isempty(foundNode),
            %     foundNode.setAttribute(searchAttribute,newValue);
            % end
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
        function newNode = addNode(xmlData,parentNode,childNode,allAttribs)
            % two possible inputs: tag name or node handle
            if ischar(parentNode),
                foundNode = exportToPPTX.findNode(xmlData,parentNode);
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
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function newNode = addNodeBefore(xmlData,parentNode,childNode,insertBefore,allAttribs)
            % two possible inputs: tag name or node handle
            % insertBefore is a node name before which to insert new node
            if ischar(parentNode),
                foundNode   = exportToPPTX.findNode(xmlData,parentNode);
            else
                foundNode   = parentNode;
            end
            if ischar(insertBefore),
                insertNode  = exportToPPTX.findNode(xmlData,insertBefore);
            else
                insertNode  = insertBefore;
            end
            
            
            if ~isempty(foundNode) && ~isempty(insertNode),
                newNode = xmlData.createElement(childNode);
                foundNode.insertBefore(newNode,insertNode);
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
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function newNode = addNodeAtPosition(xmlData,parentNode,childNode,insertPosition,allAttribs)
            % two possible inputs: tag name or node handle
            % insertPosition is an integer number where to insert new node
            if ischar(parentNode),
                foundNode = exportToPPTX.findNode(xmlData,parentNode);
            else
                foundNode = parentNode;
            end
            
            if ~isempty(foundNode),
                existingNode    = exportToPPTX.findNode(foundNode,childNode);
                newNode         = xmlData.createElement(childNode);
                if ~isempty(existingNode) && ...
                        existingNode.getLength > 1 && ...
                        insertPosition <= existingNode.getLength,
                    foundNode.insertBefore(newNode,existingNode.item(insertPosition-1));
                else
                    foundNode.appendChild(newNode);
                end
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
        end
        
        
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        function parentNode = removeNode(xmlData,remNode)
            if ischar(remNode),
                foundNode = exportToPPTX.findNode(xmlData,remNode);
            else
                foundNode = remNode;
            end
            parentNode  = foundNode.getParentNode;
            parentNode.removeChild(foundNode);
        end
        
        
        
    end
end