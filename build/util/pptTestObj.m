addpath(genpath(fullfile(pwd, '..')))

ppt = objPPT();
ppt.openPPT();
template = 'G:\Documents\Custom Office Templates\testTemplate2.pptx';
ppt.applyTemplate(template);

disp(ppt.AvailableLayouts);

ppt.addSlide('Title');
ppt.addSlide('Blank', 2);
ppt.addSlide('Title + 1 column', 2);
ppt.SlideDeck.slide1.writeText('Title_1', 'test text \n with 2 lines');

ppt.saveasPPT(fullfile(pwd, 'test.pptx'));
ppt.savePPT;

% ppt.closePPT()


%   You can move but you need to select first!!
% ppt.SlideDeck.slide1.object.Select
% ppt.SlideDeck.slide1.object.Copy
% ppt.SlideDeck.slide1.object.MoveTo(2)