# xll12

This library makes it easy to create Excel add-ins for
versions of Excel 12 or greater. See
[xlltemplate](https://github.com/keithalewis/xlltemplate) to get started.

In order to hook up Excel [Help on this function]() to your Sandcastle Help File Builder documentation you must run `shfb.bat` in the `xll12` folder.

## Debugging

To debug an add-in you must tell Visual Studio the full path to the Excel executable and what add-in to load.  
In the project `Properties` in the `Debugging` tab the `Command` should be

> `$(registry:HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe)`

To have Excel open your add-in when debugging and specify the default directory set `Command Arguments` to be

> `"$(TargetPath)" /p "$(ProjectDir)"`


Compilation of xll12 in Visual Studio 2019:
Step 1: 	Project>Properties>Configuration Properties>Configuration
		Configuration: Active(Debug)
		Platform:x64
Step 2:	Go back to the main program window. In Solution Explorer, right click on “sample”. Select “Set as StartUp Project”
Step 3: 	Go back to the main program window. In the tool bar, select Debug>Configuration Properties>C/C++>Command Line
		Make sure to delete everything under Additional Options
Step 4: Go to Properties>Configuration Properties>Debugging>Command, and replace everything with
$(registry:HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe)
Step 4:	Go back to the main program window. In Solution Explorer, right click “sample”. Select “Build”


Compilation of Quantlib in Visual Studio 2019:
See documentation online

Compilation of xll12 with Quantlib in Visual Studio 2019:
Step 1: Copy and paste the quantlib usage test code into sample.cpp in xll12 and include the ql/quantlib.hpp in the header.
Step 2: Right click sample project, Properties>Configuration Properties>VC++ Directories>Include Directories. Add boost directory C:/Program Files/boost_1_73_0/ and quantlib directories C:/Program Files/Quantlib1.18/
Step 3: Under the same options, under Library Directories, add C:/Program Files/boost_1_73_0/lib64-msvc-14.2/ and C:/Program Files/Quantlib1.18/lib/
Step 4: Under Properties>Configuration Properties>Linker>General>Additional Library Directories, add C:/Program Files/boost_1_73_0/lib64-msvc-14.2/ and C:/Program Files/Quantlib1.18/lib/
Step 5: Under Properties>Configuration Properties>Linker>Input>Additional Dependencies, add libboost_unit_test_framework-vc142-mt-gd-x64-1_73.lib and QuantLib-x64-mt-gd.lib
Step 6: Under Properties>Configuration Properties>C/C++>Code Generation>Run Time Library, make sure this option in every xll12 project is the same as the option in quantlib.
Step 7: Under Properties>Configuration Properties>C/C++>Preprocessor>Preprocessor Definitions, enter _SILENCE_ALL_CXX17_DEPRECATION_WARNINGS;_HAS_AUTO_PTR_ETC=1;

Step 8: Right Click sample project and select build to build the project.

Add-in installation:
Step 1: Open Excel > File, on the left bottom click Options
Step 2: Select Add-ins in the left column, and then under “Manage:” select “Excel Add-ins” and click “Go”.
Step 3: In the pop up window, select “Browse”
Step 4: In the browsing window, select the address of the add in
	For xll12, the add-in can be in xll12-masters/sample/x64/Release/
Step 5: check the box in front of the add-ins needed. Click “OK”.
