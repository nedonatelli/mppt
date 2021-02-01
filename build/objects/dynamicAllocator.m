classdef dynamicAllocator < dynamicprops
% DYNAMICALLOCATOR Class that handles the dynamic allocation of properties
% and methods to the ActiveX PowerPoint-related MATLAB classes
    properties
        object
        invoke
    end
    methods
        function obj = dynamicAllocator()
        end
        
        function dynamicFill(obj)
            fields = fieldnames(obj.object);
            obj.dynamicAssignProperties(fields);
            
            obj.invoke = obj.object.invoke;
            fields = fieldnames(obj.invoke);
            obj.dynamicAssignMethods(fields);
        end
        
        function dynamicAssignProperties(obj, fields)
            for iField = 1:numel(fields)
                try
                    ref = obj.object.(fields{iField});
                    addprop(obj, fields{iField});
                    obj.(fields{iField}) = ref;
                catch
                    warning('The field "%s" was not created due to an error', fields{iField});
                end
            end
        end
        
        function dynamicAssignMethods(obj, fields)
            for iField = 1:numel(fields)
                 try
                    ref = @(x) obj.object.(fields{iField});
                    addprop(obj, fields{iField});
                    obj.(fields{iField}) = ref;
                catch
                    warning('The field "%s" was not created due to an error', fields{iField});
                end
            end
        end
        
        function dynamicAssignSingle(obj, field)
            if ~isprop(obj, field)
                addprop(obj, field);
            end
        end
        
    end
end