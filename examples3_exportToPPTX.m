%
%   Example 3: working with multiple PowerPoint files
%

clear all;

%% Start first new presentation
pptx(1)     = exportToPPTX();

pptx(1).addSlide();
pptx(1).addTextbox('First presentation');

%% Start second presentation
pptx(2)     = exportToPPTX();

pptx(2).addSlide();
pptx(2).addTextbox('Seconds presentation');

%% Add another slide to first presentation
pptx(1).addSlide();
pptx(1).addTextbox('First presentation');

%% Add another slide to second presentation
pptx(2).addSlide();
pptx(2).addTextbox('Seconds presentation');

%% Finish
pptx(2).save('example3_2');
pptx(1).save('example3_1');

fprintf('Two presentations has been saved:\n');
fprintf('\t<a href="matlab:open(''%s'')">%s</a>\n',pptx(1).fullName,pptx(1).fullName);
fprintf('\t<a href="matlab:open(''%s'')">%s</a>\n',pptx(2).fullName,pptx(2).fullName);
