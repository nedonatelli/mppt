addpath(genpath(fullfile(pwd, '..')))

ppt = objPPT();
ppt.openPPT();
template = 'G:\Documents\Custom Office Templates\testTemplate2.pptx';
ppt.applyTemplate(template);

%disp(ppt.AvailableLayouts);

ppt.addSlide('Title');
ppt.addSlide('Blank', 2);
ppt.addSlide('Title + 1 column', 2);
ppt.SlideDeck.slide1.writeText('Title_1', 'test text \n with 2 lines');

fig = figure; imagesc(randn(100, 100));
print(fig, '-clipboard', '-dmeta')

ppt.SlideDeck.slide3.addFigure(fig, 100, 200, 200);
ppt.SlideDeck.slide3.addShape('Rectangle', 300, 200, 100, 200);
ppt.SlideDeck.slide3.fillShape('Rectangle_4', 'no fill');
ppt.SlideDeck.slide3.lineColor('Rectangle_4', [000000000]); % 2147483648 is white

ppt.saveasPPT(fullfile(pwd, 'test.pptx'));
ppt.savePPT;

% ppt.closePPT()


%   You can move but you need to select first!!
% ppt.SlideDeck.slide1.object.Select
% ppt.SlideDeck.slide1.object.Copy
% ppt.SlideDeck.slide1.object.MoveTo(2)


