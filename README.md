# mppt
This project aims to make operating with PowerPoint from MATLAB through use of ActiveX easier to get done what you want.

## What's going on under the hood
By leveraging some of the [available information](https://docs.microsoft.com/en-us/office/vba/api/overview/powerpoint) about ActiveX controls and utilizing and re-mapping the properties/methods of the 
COM ActiveX connection interface into a more usable MATLAB class.
The classes themselves leverage the [dynamicprops](https://www.mathworks.com/help/matlab/ref/dynamicprops-class.html) and [handle](https://www.mathworks.com/help/matlab/ref/handle-class.html?searchHighlight=handle&s_tid=srchtitle) superclasses in MATLAB in order to make it easier to re-map the COM interface. By using the [addprop](https://www.mathworks.com/help/matlab/ref/dynamicprops.addprop.html?searchHighlight=addprop&s_tid=srchtitle) method along with
MATLAB's [anonymous function handle](https://www.mathworks.com/help/matlab/matlab_prog/creating-a-function-handle.html) capabilities the task gets even easier to work with.




## Very useful information including a guide can be found here:
[OfficeMatlabCookbook](https://github.com/zglin/OfficeMatlabCookbook)

[PowerPoint Visual Basic for Applications Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/powerpoint)
