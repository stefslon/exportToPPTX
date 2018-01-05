%
%   Example 2: usage of exportToPPTX with a template
%

%% Open existing presentation
pptx    = exportToPPTX('Parallax.pptx');


%% See all available masters and layout templates
% You can programmatically access master/layout information
% You can also run exportToPPTX by itself to have this information printed
% to command window
pptx


%% Add title slide
pptx.addSlide('Master',1,'Layout','Title Slide');
pptx.addTextbox('Example Presentation','Position','Title');
pptx.addTextbox('Created with exportToPPTX','Position','Subtitle');


%% Add content slide
% By default Master 1 is assumed
sldID   = pptx.addSlide('Layout','Title and Content');
pptx.addTextbox('Content Formatting','Position','Title');
pptx.addTextbox( ...
    {'The following in-text markdown formatting options are supported:', ...
    '- Words can be **bolded** by surrounding them with "**" symbols', ...
    '- Words can be *italicized* using "*"', ...
    '- Words can also be _underlined_ by using "_" symbol', ...
    '- **Various** types of formatting *can be* combined _together_', ...
    '- Variable NAMES_WITH_UNDERSCORES are not underlined', ...
    '- _Markdown_ formatting *can span multiple* words', ...
    sprintf('\t- It can even be tabbed (if using printable \\t sequence)'), ...
    sprintf('\t\t- Further subtabbing is allowed as well')}, ...
    'Position','Content');

pptx.addNote('Notes **can** be added to the slides as well');


%% Adding more markup content
pptx.addSlide('Layout','Title and Content');
pptx.addTextbox('Content Formatting','Position','Title');
pptx.addTextbox( ...
    {'# First observation', ...
    '# Second observation', ...
    '# Third observation (*the last one*)'}, ...
    'Position','Content');

pptx.addNote( ...
    {'The following in-text markdown formatting options are supported on notes fields:', ...
    '- Words can be **bolded** by surrounding them with "**" symbols', ...
    '- Words can be *italicized* using "*"', ...
    '- Words can also be _underlined_ by using "_" symbol', ...
    '- **Various** types of formatting *can be* combined _together_', ...
    '- Variable NAMES_WITH_UNDERSCORES are not underlined', ...
    '- _Markdown_ formatting *can span multiple* words', ...
    sprintf('\t- It can even be tabbed (if using printable \\t sequence)'), ...
    sprintf('\t\t- Further subtabbing is allowed as well')});


%% Add content slide with two images
pptx.addSlide('Layout','Two Content');
pptx.addTextbox('Adding Images','Position','Title');

% Add image to the left placeholder. Set scaling to fill up the whole
% placeholder.
load earth; figure('Renderer','zbuffer'); image(X); colormap(map); axis off;
pptx.addPicture(gcf,'Position','Content Placeholder 2','Scale','max');
close(gcf);

% Add image to the right placeholder. By default, scaling is set to 
% fill up the whole placeholder while preserving aspect ratio.
load mandrill; figure('Renderer','zbuffer','Color','w'); image(X); colormap(map); axis off;
pptx.addPicture(gca,'Position','Content Placeholder 3');
close(gcf);


%% Add table
pptx.addSlide('Layout','Content with Caption');
pptx.addTextbox('Adding Tables','Position','Title');

pptx.addTextbox('The following is an example of a table added via exportToPPTX with a single cell customized','Position','Text Placeholder 3');

tableData   = { ...
    'Header 1','Header 2','Header 3'; ...
    'Row 1','Data A','Data B'; ...
    'Row 2',10,20; ...
    'Row 3','Data C',{'Click here to jump to formatting slide', 'OnClick', sldID, 'FontSize',18,'BackgroundColor','r','Vert','bottom','Horiz','right'} };

pptx.addTable(tableData,'Position','Content Placeholder 2','Vert','middle','Horiz','center');


%% Adding some more content variety
% Note how only a portion of the layout name is used, exportToPPTX 
% will match starting portion of the layout name with an available name. If
% unique match cannot be found a warning will be thrown.
pptx.addSlide('Layout','Panoramic');

% Add image using CDATA directly, image is set to fill the entire
% placeholder rectangle
rgb = imread('ngc6543a.jpg');
pptx.addPicture(rgb,'Position','Picture','Scale','max');
pptx.addTextbox('NGC 6543','Position','Title');

% Placehold content can be adjusted with all the standard textbox
% properties
pptx.addTextbox('Some description text with a custom background color', ...
    'Position','Text','FontSize',25,'BackgroundColor','k','Color',[0.9 0.9 0.9]);


%% Update slide out of order
pptx.switchSlide(sldID);
pptx.addTextbox('Added using ''switchslide'' command.','Position',[pptx.dimensions(1)/2-2 pptx.dimensions(2)-0.5 4 0.5],'FontSize',9,'Horiz','center');
pptx.addNote('Added using ''switchslide'' command.');


%% Save As this presentation
newFile = pptx.save('example2');


fprintf('New file has been saved: <a href="matlab:open(''%s'')">%s</a>\n',newFile,newFile);

