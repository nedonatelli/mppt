addpath(genpath(fullfile(pwd, '..')))

ppt = objPPT();
ppt.openPPT();
template = 'G:\Documents\Custom Office Templates\testTemplate2.pptx';
ppt.applyTemplate(template);

disp(ppt.AvailableLayouts);

ppt.addSlide('Title');
ppt.addSlide('Blank');
ppt.addSlide('Title + 1 column');

ppt.saveasPPT(fullfile(pwd, 'test.pptx'));
ppt.savePPT;

% ppt.closePPT()


%   You can move but you need to select first!!
% ppt.SlideDeck.slide1.object.Select
% ppt.SlideDeck.slide1.object.Copy
% ppt.SlideDeck.slide1.object.MoveTo(2)