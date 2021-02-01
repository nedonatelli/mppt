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
        end
        %TODO: Design Methods to be used for slide design
    end
end