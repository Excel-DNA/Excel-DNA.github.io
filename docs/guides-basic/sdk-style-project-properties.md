---
title: "Excel-DNA properties in SDK-style project files"
---

The following project properties are recognized by the Excel-DNA build task, and can be used to customize the build output.
By setting the appropriate properties, the old-style project .dna files are no longer needed when using SDK-style project files.
As part of the build, if no .dna file is present in the project, the required .dna files will be created as build outputs, using the project properties.

```xml
  <PropertyGroup>
    
    <!-- Base path to ExcelDna.AddIn.Tasks.dll, ExcelDnaPack.exe and ExcelDna.xll. -->
    <!-- Default value: ..\tools -->
    <ExcelDnaToolsPath></ExcelDnaToolsPath>

    <!-- Path to ExcelDnaPack.exe. -->
    <!-- Default value: $(ExcelDnaTasksPath)ExcelDnaPack.exe -->
    <ExcelDnaPackExePath></ExcelDnaPackExePath>

    <!-- Base path for .props location. -->
    <!-- Default value: $(MSBuildProjectDirectory) -->
    <ExcelDnaProjectPath></ExcelDnaProjectPath>

    <!-- Path for configuration properties file location.  -->
    <!-- Default value: $(ExcelDnaProjectPath)\Properties\ExcelDna.Build.props -->
    <ExcelDnaPropsFilePath></ExcelDnaPropsFilePath>

    <!-- Enables creating executable Excel profile in launchSettings.json.  -->
    <!-- Default value: true -->
    <RunExcelDnaSetDebuggerOptions></RunExcelDnaSetDebuggerOptions>

    <!-- Enables creating executable Excel profile in launchSettings.json when building from a command line or alternative IDE.  -->
    <!-- Default value: false -->
    <RunExcelDnaSetDebuggerOptionsOutsideVisualStudio></RunExcelDnaSetDebuggerOptionsOutsideVisualStudio>

    <!-- Enables removing .dna, .xll from the build output folder on Build Clean. -->
    <!-- Default value: true -->
    <RunExcelDnaClean></RunExcelDnaClean>

    <!-- Enables copying .dna, .xll to the build output folder on Build.  -->
    <!-- Default value: true -->
    <RunExcelDnaBuild></RunExcelDnaBuild>

    <!-- Enables creating packed add-in on Build. -->
    <!-- Default value: true -->
    <RunExcelDnaPack></RunExcelDnaPack>

    <!-- Enables to have an .xll file with no packed assemblies in it. -->
    <!-- Default value: false -->
    <ExcelDnaUnpack></ExcelDnaUnpack>

    <!-- The output directory for the 'published' add-in. Use %none% to put the files in the same output directory. -->
    <!-- Default value: publish for SDK-style projects, %none% for old-style projects -->
    <ExcelDnaPublishPath></ExcelDnaPublishPath>

    <!-- Enables creating 32bit add-in. -->
    <!-- Default value: true -->
    <ExcelDnaCreate32BitAddIn></ExcelDnaCreate32BitAddIn>

    <!-- Enables creating 64bit add-in. -->
    <!-- Default value: true -->
    <ExcelDnaCreate64BitAddIn></ExcelDnaCreate64BitAddIn>

    <!-- 32bit add-in name suffix. -->
    <!-- Default value: none -->
    <ExcelDna32BitAddInSuffix></ExcelDna32BitAddInSuffix>

    <!-- 64bit add-in name suffix. Use %none% for no suffix. -->
    <!-- Default value: 64 -->
    <ExcelDna64BitAddInSuffix></ExcelDna64BitAddInSuffix>

    <!-- Packed add-in name suffix.  Use %none% to make the name of the packed output file be the same as the unpacked name. -->
    <!-- Default value: -packed -->
    <ExcelDnaPackXllSuffix></ExcelDnaPackXllSuffix>

    <!-- Explicit 32bit output file name -->
    <!-- Default value: empty -->
    <ExcelDnaPack32BitXllName></ExcelDnaPack32BitXllName>

    <!-- Explicit 64bit output file name -->
    <!-- Default value: empty -->
    <ExcelDnaPack64BitXllName></ExcelDnaPack64BitXllName>

    <!-- Enables packed add-in compression. -->
    <!-- Default value: true -->
    <ExcelDnaPackCompressResources></ExcelDnaPackCompressResources>

    <!-- Enables multithreaded add-in packing. -->
    <!-- Default value: true -->
    <ExcelDnaPackRunMultithreaded></ExcelDnaPackRunMultithreaded>

    <!-- Enables cross-platform resource packing implementation when executing on Windows. -->
    <!-- Default value: false -->
    <ExcelDnaPackManagedResourcePackingOnWindows></ExcelDnaPackManagedResourcePackingOnWindows>

    <!-- Enables packing native libraries from .deps.json. -->
    <!-- Default value: true -->
    <ExcelDnaPackNativeLibraryDependencies></ExcelDnaPackNativeLibraryDependencies>

    <!-- Enables packing managed assemblies from .deps.json. -->
    <!-- Default value: true -->
    <ExcelDnaPackManagedDependencies></ExcelDnaPackManagedDependencies>

    <!-- Semicolon separated file names list to not pack from .deps.json. -->
    <!-- Default value: empty -->
    <ExcelDnaPackExcludeDependencies></ExcelDnaPackExcludeDependencies>

    <!-- EXCEL.EXE path for debugging. -->
    <!-- Default value: auto detect -->
    <ExcelDnaExcelExePath></ExcelDnaExcelExePath>

    <!-- Add-in file name for debugging. -->
    <!-- Default value: auto detect -->
    <ExcelDnaAddInForDebugging></ExcelDnaAddInForDebugging>

    <!-- Add-in name for output files. -->
    <!-- Default value: $(ProjectName)-AddIn -->
    <ExcelAddInFileName></ExcelAddInFileName>

    <!-- DnaLibrary Name in .dna. -->
    <!-- Default value: $(ProjectName) Add-In -->
    <ExcelAddInName></ExcelAddInName>

    <!-- Semicolon separated references list to include in .dna. -->
    <!-- Default value: empty -->
    <ExcelAddInInclude></ExcelAddInInclude>

    <!-- Semicolon separated external libraries to include in .dna. -->
    <!-- Default value: empty -->
    <ExcelAddInExports></ExcelAddInExports>

    <!-- ExternalLibrary Path in .dna. -->
    <!-- Default value: $(TargetFileName) -->
    <ExcelAddInExternalLibraryPath></ExcelAddInExternalLibraryPath>

    <!-- Enable/disable including pdb files in packed add-in. -->
    <!-- Default value: false -->
    <ExcelAddInIncludePdb></ExcelAddInIncludePdb>

    <!-- Control whether the add-in's assemblies are loaded directly from byte arrays under .NET Framework. -->
    <!-- Default value: true -->
    <ExcelAddInLoadFromBytes></ExcelAddInLoadFromBytes>

    <!-- Enable/disable collectible AssemblyLoadContext for .NET 6. -->
    <!-- Default value: false -->
    <ExcelAddInDisableAssemblyContextUnload></ExcelAddInDisableAssemblyContextUnload>

    <!-- Path to TlbExp.exe. E.g. "c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\TlbExp.exe" or $(MSBuildProjectDirectory)\TlbExp.exe.-->
    <!-- Default value: empty -->
    <ExcelAddInTlbExp></ExcelAddInTlbExp>

    <!-- Enable/disable .tlb file creation. -->
    <!-- Default value: false -->
    <ExcelAddInTlbCreate></ExcelAddInTlbCreate>

    <!-- Path to signtool.exe. E.g. "c:\Program Files\Microsoft SDKs\Windows\8.1\bin\x64\signtool.exe" or $(MSBuildProjectDirectory)\signtool.exe -->
    <!-- Default value: empty -->
    <ExcelAddInSignTool></ExcelAddInSignTool>

    <!-- Options for signtool.exe. E.g. /f "$(MSBuildProjectDirectory)\Contoso.pfx" /p 12345678 -->
    <!-- Default value: empty -->
    <ExcelAddInSignOptions></ExcelAddInSignOptions>

    <!-- Replace XLL version information with data read from ExternalLibrary assembly. -->
    <!-- Default value: false -->
    <ExcelAddInUseVersionAsOutputVersion></ExcelAddInUseVersionAsOutputVersion>

    <!-- Prevents every static public function from becomming a UDF, they will need an explicit [ExcelFunction] annotation. -->
    <!-- Default value: false -->
    <ExcelAddInExplicitExports></ExcelAddInExplicitExports>

    <!-- Prevents automatic registration of functions and commands. -->
    <!-- Default value: false -->
    <ExcelAddInExplicitRegistration></ExcelAddInExplicitRegistration>

    <!-- Enable/disable COM Server support. -->
    <!-- Default value: false -->
    <ExcelAddInComServer></ExcelAddInComServer>

    <!-- We don't need the extra 'ref' directory and reference assemblies for the Excel add-in -->
    <ProduceReferenceAssembly>false</ProduceReferenceAssembly>

    <!-- We need all dependencies to be copied to the output directory, as-if we are an 'application' and not a 'library'. This property also sets the CopyLockFileAssemblies property to true. -->
    <EnableDynamicLoading>true</EnableDynamicLoading>
    
  </PropertyGroup>

```
