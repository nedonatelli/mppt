classdef objSlide < handle & dynamicAllocator
% OBJSLIDE Class that dynamically assigns the ActiveX PowerPoint slide
% connection properties and methods to a MATLAB class
    properties
        children
    end
    methods
        function obj = objSlide(slide)
            obj.object = slide;
            obj.dynamicFill();
            obj.findChildren();
        end
        
        function findChildren(obj)
            if obj.Shapes.Count > 0
                for iChild = 1:obj.Shapes.Count
                    name = obj.Shapes.Item(iChild).Name;
                    name = regexprep(name, ' ', '_');
                    obj.children.(name) =...
                        obj.Shapes.Item(iChild);
                end
            else%If there aren't any shapes then return empty
                obj.children = [];
                return
            end   
        end
        
        %TODO: Design Methods to be used for slide design
        
        function writeText(obj, childname, text)
            % WRITETEXT adds text to selected child
            % Supports newline '\n' characters by replacing them with
            % MATLAB's newline
            text = join(split(text, '\n'), newline);
            obj.children.(childname).TextFrame.TextRange.Text = text{1};
        end
        
        function addFigure(obj, fig, left, top, height)
            % ADDFIGURE adds a figure from a figure handle in MATLAB to the
            % selected slide.
            % Can also move the figure to location defined by top & left as
            % well as resize the image by height.
            % If any of those three parameters are set to NaN the image
            % will use the default value
            print(fig, '-clipboard', '-dmeta'); %Capture figure to clipboard
            obj.Select();%Select the slide
            obj.Shapes.PasteSpecial(2); %Paste onto the slide
            close(fig); %Close the figure
            obj.findChildren(); %Update the children for the slide
            iChild = obj.Shapes.count; %Get the new number of children
            childname = obj.Shapes.Item(iChild).Name; %Get the name of new child
            childname = regexprep(childname, ' ', '_'); %Swap spaces for underscores
            
            obj.moveChild(childname, top, left); %Move the new image to the location top, left
            obj.resizeChild(childname, height); %Resize the 
        end
        
        function addShape(obj, shape, left, top, width, height)
            obj.Select();
            obj.Shapes.AddShape(sprintf('msoShape%s', shape), left, top, width, height);
            obj.findChildren();
        end
        
        function fillShape(obj, childname, fill_color)
            obj.Select();
            if ischar(fill_color)
                if strcmp(fill_color, 'no fill')
                    obj.children.(childname).Fill.Visible = false;
                else
                    warning('Unexpected String Argument')
                end
                obj.findChildren();
                return
            end
            
            obj.children.(childname).Fill.ForeColor.RGB = fill_color;
            obj.findChildren();
        end
        
        function lineColor(obj, childname, line_color)
            obj.Select();
            if ischar(line_color)
                if strcmp(fill_color, 'no fill')
                    obj.children.(childname).Line.Visible = false;
                else
                    warning('Unexpected String Argument')
                end
                obj.findChildren();
                return
            end
            
            obj.children.(childname).Line.ForeColor.RGB = line_color;
            obj.findChildren();
        end
        
        function moveChild(obj, childname, top, left)
            obj.Select();
            obj.children.(childname).Select();
            if ~isnan(top)
                obj.children.(childname).Top = top;
            end
            if ~isnan(left)
                obj.children.(childname).Left = left;
            end
        end
        
        function resizeChild(obj, childname, height)
            if ~isnan(height)
                obj.Select();
                obj.children.(childname).Select();
                obj.children.(childname).Height = height;
            end
        end
        
    end
end