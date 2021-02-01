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
            text = join(split(text, '\n'), newline);
            obj.children.(childname).TextFrame.TextRange.Text = text{1};
        end
    end
end