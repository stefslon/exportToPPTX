%
%   Example usage of exportToPPTX
%


%% Start new presentation
try
    exportToPPTX('new','Dimensions',[12 6]);
catch
    % If PowerPoint already started, then close first and then open a new
    % one
    exportToPPTX('close');
    exportToPPTX('new','Dimensions',[12 6]);
end


%% Add some slides
figH = figure('Renderer','zbuffer'); mesh(peaks); view(0,0);

for islide=1:5,
    slideNum = exportToPPTX('addslide');
    fprintf('Added slide %d\n',slideNum);
    exportToPPTX('addpicture',figH);
    exportToPPTX('addtext',sprintf('Slide Number %d',slideNum));
    
    % Rotate mesh on each slide
    view(18*islide,18*islide);
end
close(figH);


%% Check current presentation
fileStats   = exportToPPTX('query');

if ~isempty(fileStats),
    fprintf('Presentation size: %f x %f\n',fileStats.dimensions);
    fprintf('Number of slides: %d\n',fileStats.numSlides);
end


%% Save presentation -- overwrite file if it already exists
% Filename automatically checked for proper extension
newFile = exportToPPTX('save','example');


%% Add multiple images with custom sizes
load earth; image(X); colormap(map); axis off;

exportToPPTX('addslide');
exportToPPTX('addpicture',gcf,'Position',[1 1 3 2]);
exportToPPTX('addpicture',gcf,'Position',[6 1 3 2]);
exportToPPTX('addpicture',gcf,'Position',[1 3.5 3 2]);
exportToPPTX('addpicture',gcf,'Position',[6 3.5 3 2]);
close;


%% Add image that takes up as much of the slide as possible without losing its aspect ratio
load mandrill; figure('color','w'); image(X); colormap(map); axis off; axis image;

exportToPPTX('addslide');
exportToPPTX('addpicture',gcf,'Scale','maxfixed');
close;


%% Add multiple text boxes with custom sizes
exportToPPTX('addslide');
exportToPPTX('addtext','Red Left-top','Position',[0 0 4 2],'Color',[1 0 0]);
exportToPPTX('addtext','Green Center-top','Position',[4 0 4 2],'HorizontalAlignment','center','Color',[0 1 0]);
exportToPPTX('addtext','Right-top and bold italics','Position',[8 0 4 2],'HorizontalAlignment','right','FontWeight','bold','FontAngle','italic');
exportToPPTX('addtext','Soft Blue Left-middle','Position',[0 2 4 2],'VerticalAlignment','middle','Color',[0.58 0.70 0.84]);
exportToPPTX('addtext','Center-middle and bold','Position',[4 2 4 2],'VerticalAlignment','middle','HorizontalAlignment','center','FontWeight','bold');
exportToPPTX('addtext','Right-middle','Position',[8 2 4 2],'VerticalAlignment','middle','HorizontalAlignment','right');
exportToPPTX('addtext','Left-bottom and italics','Position',[0 4 4 2],'VerticalAlignment','bottom','HorizontalAlignment','left','FontAngle','italic');
exportToPPTX('addtext','Center-bottom size 10','Position',[4 4 4 2],'VerticalAlignment','bottom','HorizontalAlignment','center','FontSize',10);
exportToPPTX('addtext','Right-bottom size 30','Position',[8 4 4 2],'VerticalAlignment','bottom','HorizontalAlignment','right','FontSize',30);


%% Save again
exportToPPTX('save');


%% Close presentation (and clear all temporary files)
exportToPPTX('close');

fprintf('New file has been saved: %s\n',newFile);

