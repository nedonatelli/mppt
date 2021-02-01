classdef objPPT < handle & dynamicAllocator
% OBJPPT Class that dynamically assigns the ActiveX PowerPoint connection
%   properties and methods to a MATLAB class
%   
%       objPPT - Constructor  for ActiveX PowerPoint connection object 
    
    methods 
        function obj = objPPT()
            obj.object = actxserver('powerpoint.application');
            obj.dynamicFill();
        end
        
        function openPPT(obj)
            % openPPT - Method to create a PowerPoint presentation and add
            % ActivePresentation property to MATLAB ActiveX PowerPoint connection
            % object
            obj.Presentations.Add();
            obj.dynamicAssignSingle('ActivePresentation')
            obj.ActivePresentation = obj.object.ActivePresentation;
            obj.getLayouts;
        end
        
        function applyTemplate(obj, templatename)
            % applyTemplate - Method to apply a PowerPoint template to the
            % PowerPoint presentation. 
            %
            %   templatename: This can be a template file or anotehr
            %       presentation with a theme. (.potx, .ppt, .pptx, etc.)
            obj.ActivePresentation.ApplyTemplate(templatename);
            obj.getLayouts;
        end
        
        function saveasPPT(obj, filename)
            % saveasPPT - Method to save PowerPoint as supplied filename.     
            obj.ActivePresentation.SaveAs(filename);
        end
        
        function savePPT(obj)
            % savePPT - Method to save PowerPoint.
            obj.ActivePresentation.Save;
        end
        
        function getLayouts(obj)
            % getLayouts - Method to identify available layouts that can
            % be used in PowerPoint under the current selected theme.
            obj.dynamicAssignSingle('AvailableLayouts');
            obj.AvailableLayouts = {};
            numLayouts = ...
                obj.ActivePresentation.SlideMaster.CustomLayouts.count;
            for iLayout = 1:numLayouts
                obj.AvailableLayouts{iLayout} = obj.ActivePresentation.SlideMaster.CustomLayouts.Item(iLayout).Name;
            end
        end
        
        function addSlide(obj, layout, index)
            % addSlide - Method to add a slide to the PowerPoint object. 
            %   Adds the a SlideDeck struct if it doesn't already exist
            %   and stores slides in the struct as they are generated.
            if ~exist('index', 'var')
                index = obj.ActivePresentation.Slides.Count+1;
            end
            
            
            obj.dynamicAssignSingle('SlideDeck');
            
            slideregex =... %This is done to combat regex reserved chars
                join(['([', join(split(layout, ' '), ']+) ([') , ']+)']);
            slidename = sprintf('^%s$', slideregex{1});
            
            searchslide = regexp(obj.AvailableLayouts, slidename);
            iLayout = find(cellfun(@(x) ~isempty(x) , searchslide));
            
            if isempty(iLayout)
                warning('There are no matching layouts for your request');
                return
            end
            
            layout = ...
                obj.ActivePresentation.SlideMaster.CustomLayouts.Item(iLayout);
            slidelabel = sprintf('slide%d',...
                obj.ActivePresentation.Slides.Count+1);
            
            slide = objSlide(...
                obj.ActivePresentation.Slides.AddSlide(...
                    index, layout));
            
            obj.SlideDeck.(slidelabel) = slide;
                
        end
        % TODO: Add methods to handle duplicating/selecting/copying/moving
        
        function closePPT(obj)
            obj.ActivePresentation.Close;
            obj.Quit;
        end
    end
end
