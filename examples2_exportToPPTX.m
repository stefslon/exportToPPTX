%
%   Example 2: usage of exportToPPTX with a template
%

%% Open existing presentation
isOpen  = exportToPPTX();
if ~isempty(isOpen),
    % If PowerPoint already started, then close first and then open a new one
    exportToPPTX('close');
end

exportToPPTX('open','Parallax.pptx');


%% See all available masters and layout templates
% You can programmatically access master/layout information
% You can also run exportToPPTX by itself to have this information printed
% to command window
pptxInfo    = exportToPPTX;
fprintf('All available layout templates:\n');
for ilayout=1:numel(pptxInfo.master(1).layout),
    fprintf('\t%d. %s\n',ilayout,pptxInfo.master(1).layout(ilayout).name);
end


%% Add title slide
exportToPPTX('addslide','Master',1,'Layout','Title Slide');
exportToPPTX('addtext','Example Presentation','Position','Title');
exportToPPTX('addtext','Created with exportToPPTX','Position','Subtitle');


%% Add content slide
% By default Master 1 is assumed
sldID   = exportToPPTX('addslide','Layout','Title and Content');
exportToPPTX('addtext','Content Formatting','Position','Title');
exportToPPTX('addtext', ...
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

exportToPPTX('addnote','Notes **can** be added to the slides as well');


%% Adding more markup content
exportToPPTX('addslide','Layout','Title and Content');
exportToPPTX('addtext','Content Formatting','Position','Title');
exportToPPTX('addtext', ...
    {'# First observation', ...
    '# Second observation', ...
    '# Third observation (*the last one*)'}, ...
    'Position','Content');

exportToPPTX('addnote', ...
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
exportToPPTX('addslide','Layout','Two Content');
exportToPPTX('addtext','Adding Images','Position','Title');

% Add image to the left placeholder. Set scaling to fill up the whole
% placeholder.
load earth; figure('Renderer','zbuffer'); image(X); colormap(map); axis off;
exportToPPTX('addpicture',gcf,'Position','Content Placeholder 2','Scale','max');
close(gcf);

% Add image to the right placeholder. By default, scaling is set to 
% fill up the whole placeholder while preserving aspect ratio.
load mandrill; figure('Renderer','zbuffer','Color','w'); image(X); colormap(map); axis off;
exportToPPTX('addpicture',gca,'Position','Content Placeholder 3');
close(gcf);


%% Add table
exportToPPTX('addslide','Layout','Content with Caption');
exportToPPTX('addtext','Adding Tables','Position','Title');

exportToPPTX('addtext','The following is an example of a table added via exportToPPTX with a single cell customized','Position','Text Placeholder 3');

tableData   = { ...
    'Header 1','Header 2','Header 3'; ...
    'Row 1','Data A','Data B'; ...
    'Row 2',10,20; ...
    'Row 3','Data C',{'Click here to jump to formatting slide', 'OnClick', sldID, 'FontSize',18,'BackgroundColor','r','Vert','bottom','Horiz','right'} };

exportToPPTX('addtable',tableData,'Position','Content Placeholder 2','Vert','middle','Horiz','center');


%% Adding some more content variety
% Note how only a portion of the layout name is used, exportToPPTX 
% will match starting portion of the layout name with an available name. If
% unique match cannot be found a warning will be thrown.
exportToPPTX('addslide','Layout','Panoramic');

% Add image using CDATA directly, image is set to fill the entire
% placeholder rectangle
rgb = imread('ngc6543a.jpg');
exportToPPTX('addpicture',rgb,'Position','Picture','Scale','max');
exportToPPTX('addtext','NGC 6543','Position','Title');

% Placehold content can be adjusted with all the standard textbox
% properties
exportToPPTX('addtext','Some description text with a custom background color', ...
    'Position','Text','FontSize',25,'BackgroundColor','k','Color',[0.9 0.9 0.9]);


%% Update slide out of order
exportToPPTX('switchslide',sldID);
exportToPPTX('addtext','Added using ''switchslide'' command.','Position',[pptxInfo.dimensions(1)/2-2 pptxInfo.dimensions(2)-0.5 4 0.5],'FontSize',9,'Horiz','center');
exportToPPTX('addnote','Added using ''switchslide'' command.');


%% Save As this presentation
newFile = exportToPPTX('save','example2');


%% Close presentation (and clear all temporary files)
exportToPPTX('close');

fprintf('New file has been saved: <a href="matlab:open(''%s'')">%s</a>\n',newFile,newFile);

