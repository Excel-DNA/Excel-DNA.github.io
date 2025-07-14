# Excel-DNA Release Notes - v1.9.0

**Released:** _Soon..._ [v1.9.0-rc3 is now available on NuGet](https://www.nuget.org/packages/ExcelDna.AddIn).

This document provides an overview of the new features and enhancements in Excel-DNA version 1.9.0. This release significantly refactors and extends function registration, including simplified support for optional / default parameters and asynchronous functions, and introduces built-in support for object handles.

## Support and Sponsors

Public support for Excel-DNA is provided via the [Excel-DNA Google group](https://groups.google.com/g/exceldna). If you run into any questions or issues with the update, or have suggestions about future improvement, please feel free to start a discussion.

I also offer more formal corporate maintenance and support agreements for those using the library in a mission-critical setting - please contact me directly for more details.

The easiest way to encourage ongoing development and support of Excel-DNA is to sign up as a [GitHub Sponsor](https://github.com/sponsors/Excel-DNA). Higher tiers also provide online support sessions, where I am happy to advise or help you get unstuck with your add-in development. Thank you very much to all the existing sponsors!

## Thanks

Thank you very much to Sergey Vlasov for doing an awesome job in producing this update.
Please have a look at his great set of Visual Studio extensions here: https://vlasovstudio.com/

Thank you also to the other contributors who submitted changes for this version, or reported bugs or problems along the way.

And to all the Excel-DNA add-in creators around the world - I hope you continue to find the library useful.

## Release Highlights

- **[Extended Registration](#registration-refactoring)**  
  The functionality of [`ExcelDna.Registration`](https://github.com/Excel-DNA/Registration) is now built into the core library. `[ExcelFunction]` supports more parameter types and modifiers by default.
- **[Built-in Asynchronous Support](#asynchronous-and-streaming-functions)**  
  Async (`Task<T>`) and streaming (`IObservable<T>`) functions are now supported directly.
- **[Built-in Object Handles](#object-handle-support-with-excelhandle)**  
  Use [`[ExcelHandle]`](https://excel-dna.net/) to pass complex .NET objects between Excel and your add-in. Handles keep objects alive safely across the Excel and .NET boundary.
- **[RTD Server Improvements](#rtd-server-improvements)**  
  New tracking methods in [`ExcelRtdServer`](https://excel-dna.net/) let you monitor RTD interactions and add watchdogs.
- **[Debug and Help File Packing](#other-changes-and-enhancements)**  
  `.chm` help files and `.pdb` files can now be packed in your `.xll` for better deployment.
  
## Registration Refactoring

Excel-DNA's _**extended registration**_ mechanism allows add-ins to define user-defined functions (UDFs) with signatures beyond basic Excel primitives (e.g. to allow optional parameters, `params` arrays, asynchronous methods), add custom function wrappers (for logging, caching, etc.), and register functions created at runtime.

### Changes from Earlier Versions: Integration of `ExcelDna.Registration`
In v1.9, the functionality previously exposed in the separate `ExcelDna.Registration` library and package has been incorporated directly into the main `ExcelDna.Integration` library. This integration streamlines development for add-ins requiring extended registration features. While older versions required explicit steps and inclusion of the `ExcelDna.Registration` package to use these features, v1.9 offers these capabilities out-of-the-box or through more accessible APIs within the core library.

Key changes include:
* **Expanded Default Support:** A significantly broader range of parameter and return types are now supported directly for methods marked with `[ExcelFunction]`.
* **Integrated Async & Object Handles:** Support for asynchronous (`Task<T>`) and streaming (`IObservable<T>`) functions, along with object handles (`[ExcelHandle]`, `[return: ExcelHandle]`), is now built-in.
* **Migrated Extension Points:** Customization hooks like `FunctionExecutionHandler` are now part of the main library, generally found within the `ExcelDna.Registration` namespace.
* **Deprecation of `Excel.Registration` Package:** The `ExcelDna.Registration.*` packages are no longer required, and have been marked as deprecated.

### Function Registration in v1.9

#### Basic Registration for `public static` Functions
For compatibility with previous Excel-DNA versions, all `public static` functions with parameters and return types only having 'primitive' types are registered as Excel UDFs, even if not marked with an `[ExcelFunction]` attribute. The primitive types are: `double`, `string`, `DateTime`, `double[]`, `double[,]`, `bool`, `int`, `short`, `ushort`, `decimal`, `long`, `void`, `object`, `object[]`, `object[,]`. If an add-in project is marked as `<ExcelAddInExplicitExports>true</<ExcelAddInExplicitExports>` then registration is only done for those functions marked with `[ExcelFunction]`, preventing other `public static` functions from being registered by mistake.

#### Extended Default Registration with `[ExcelFunction]`
From this version, the standard `[ExcelFunction]` attribute implicitly supports a wider array of method signatures without requiring additional custom registration configurations:

* **Optional and Default Parameters:** Handles C# optional parameters with default values and VB.NET `Optional` parameters, including for `DateTime` and nullable types.

    ```csharp
    // C# example with optional DateTime
    [ExcelFunction]
    public static string MyOptionalDateTime(DateTime dt = default) // Or e.g., DateTime.MinValue
    {
        if (dt == default(DateTime))
            return "DateTime was default";
        return $"DateTime is {dt:yyyy-MM-dd HH:mm:ss}";
    }

    // C# example with nullable double
    [ExcelFunction]
    public static string DnaParameterConvertTest(double? optTest)
    {
        if (!optTest.HasValue) return "NULL!!!";
        return optTest.Value.ToString("F1");
    }
    ```

    ```vbnet
    ' VB.NET example with optional parameter
    <ExcelFunction>
    Public Function DnaOptionalAnswer(Optional num As Double = 42) As String
        Return "The answer is " & num
    End Function
    ```

* **`params` Arrays:** Supports functions using `params` arrays in C# or `ParamArray` in VB.NET to accept a variable number of arguments.

    ```csharp
    // C# example with params array
    [ExcelFunction]
    public static string DnaParamsFunc(object input, string other, params object[] args)
    {
        return $"Input: {input}, Other: {other}, Args count: {args?.Length ?? 0}";
    }
    ```

    ```vbnet
    ' VB.NET example with ParamArray
    <ExcelFunction>
    Function DnaAddValues(val1 As Double, ParamArray vals As Double()) As Double
        Dim sum As Double = val1
        If vals IsNot Nothing Then
            For Each val As Double In vals
                sum += val
            Next
        End If
        Return sum
    End Function
    ```

#### **Asynchronous and Streaming Functions:** 
Functions that have return types `Task<T>` or `IObservable<T>` will automatically be registered as an RTD-based async or streaming function. If an async function has a final parameter of type [`CancellationToken`](https://learn.microsoft.com/en-us/dotnet/api/system.threading.cancellationtoken), the token will be signaled if the formula is deleted while the async call is outstanding.

##### Directly Usable Helper Methods

In order to customize the async or streaming functions beyond the default wrappers, several utility classes and methods for asynchronous operations and object handling are directly accessible. These allow easy async and streaming function implementations by writing an explicit wrapper.

* **Async Utilities:**
    
    * `ExcelAsyncUtil.RunTask`: Helper for running `Task`-based operations and integrating them with Excel's async model.
        ```csharp
        // Example usage of ExcelAsyncUtil.RunTask
        public static object RunMyTask<TResult>(string callerFunctionName, object callerParameters, Func<Task<TResult>> taskSource)
        {
            return ExcelDna.Integration.ExcelAsyncUtil.RunTask(callerFunctionName, callerParameters, taskSource);
        }
        ```
    * `ExcelDna.Integration.AsyncTaskUtil`: Utility class for managing `Task`-based functions.
    * `ExcelDna.Integration.NativeAsyncTaskUtil`: For deeper integration with native async capabilities.
        ```csharp
        // Conceptual: NativeAsyncTaskUtil.RunTask(() => MyTaskFunction(), asyncHandle);
        ```

    *  `ObservableRtdUtil` Utility
        Introduced `ObservableRtdUtil` for improved handling of real-time data updates using observables. This utility simplifies the integration of `IObservable<T>` sources with Excel's RTD mechanism.
        ```csharp
        // Example usage of ObservableRtdUtil
        // ObservableRtdUtil.Observe("MyRtdFunction", new object[] { param1, param2 }, () => myObservableSource);
        // More specific example from fragments:
        // var rtdResult = ObservableRtdUtil.Observe("FunctionName", new object[] { param1, param2 }, () => myObservable);
        ```
    * `ExcelAsyncUtil.Observe`: Has a new overload that accepts `ExcelObservableOptions` for more control over observable UDFs.
        ```csharp
        // Conceptual usage of new ExcelAsyncUtil.Observe overload
        // var options = new ExcelDna.Integration.ExcelObservableOptions { ... };
        // var result = ExcelDna.Integration.ExcelAsyncUtil.Observe("MyObservableFunc", parametersArray, options, () => myObservableSource);
        ```

#### Registration Customization

Explicit registration by retrieving and processing the `ExcelFunction` list was previously done by extensions in the `ExcelDna.Registration` library. The relevant types from that library are now included in the `ExcelDna.Integration` assembly, but the namespaces and type names are not changed. Thus code previously doing explicit registration should still work after the update. (The `ExcelDna.Registration` NuGet package just becomes obsolete.)

We also provide a simpler approach for including registration extensions in an add-in.

* **Function Execution Handlers:** For common cross-cutting concerns like caching and timing, it can be convenient to add aspect-oriented style function handlers that get woven around the registered functions. We support various ways of registering such handlers:
  * The `FunctionExecutionHandler` base class defines a set or methods that can be overridden to intercept steps in the function call.
  * The `[ExcelFunctionExecutionHandlerSelector]` attribute can be used to declare static selector functions that interact with the registration process.

*  Language Specific Support

When the Excel-DNA Registration extensions were separate libraries, there were some language-specific extensions for VB.NET and F#. 

  * **F#:**

    For F#, the support for asynchronous workflows (`async {}`) should now be incorporated in the add-in project by using the `FsAsyncUtil` helper

    * **Async Functions:** Support for F# asynchronous workflows (`async {}`) integrated with Excel's async UDFs via `FsAsyncUtil`.
        ```fsharp
        // F# Async Example
        open ExcelDna.Integration // For ExcelFunction and FsAsyncUtil (assuming namespace)

        [<ExcelFunction>]
        let dnaFsHelloAsync name (msToWait: int) =
            async {
                do! Async.Sleep msToWait
                return "Hello " + name
            }
            |> FsAsyncUtil.observeAsync // Or other FsAsyncUtil methods as appropriate
        ```
    * **Optional Parameters:** Support for F# optional parameters in UDFs.
        ```fsharp
        // F# Optional Parameter Example
        open ExcelDna.Integration // For ExcelFunction
        open System // For Sprintf

        [<ExcelFunction>]
        let dnaFSharpOptional(?value : double, ?str : string, ?bl : bool) =
            let theValue = defaultArg value 12.3
            let theString = defaultArg str "qwerty"
            let theBool = defaultArg bl true
            sprintf "Value: %f, String: %s, Bool: %b" theValue theString theBool
        ```
    * **New Sample Project:** Added `ExcelDna.AddIn.RegistrationSampleFS` to demonstrate these F# features.

  * **VB.NET:**
    * **Range Parameter Conversion:** Support for automatically converting Excel cell/range references passed as `Object` to `Microsoft.Office.Interop.Excel.Range` objects.
        ```vbnet
        Imports Microsoft.Office.Interop.Excel
        Imports ExcelDna.Integration ' For ExcelFunction
        Imports ExcelDna.Registration.ParameterConversion ' For RangeParameterConversion

        Public Module VbRangeSupport
            <ExcelFunction>
            Public Function GetAddressFromRange(inputRange As Object) As String
                Dim excelRange As Range = RangeParameterConversion.ReferenceToRange(inputRange)
                If excelRange IsNot Nothing Then
                    Return excelRange.Address(False, False)
                Else
                    Return "Not a valid range"
                End If
            End Function
        End Module
        ```
    * **Command Registration:** Simplified registration of VB.NET methods as Excel macros/commands.
        ```vbnet
        Imports ExcelDna.Integration ' For ExcelCommand
        Imports Microsoft.Office.Interop.Excel ' For Application (Excel object model)

        Public Module VbCommands
            <ExcelCommand(MenuName:="VB Samples", MenuText:="Write to A7", ShortCut:="^%D")>
            Sub DnaDumpDataToA7()
                Dim xlApp As Application = ExcelDnaUtil.Application
                xlApp.Range("A7").Value = "Hello from VB.NET Command"
            End Sub
        End Module
        ```
    * **Optional Parameters & `ParamArray`:** Full support similar to C#. (See examples in "Extended Default Registration" section).

* Registration Samples

Various examples of extended registration with parameter conversion and function handlers can be found in the [Excel-DNA samples Repository](https://github.com/Excel-DNA/Samples/blob/master/Registration).

  * AsyncReturnHandler - Customize initial return value of async function, e.g. use #GETTING_DATA instead of #N/A
  * CacheFunctionExecutionHandler - easily cache function results for better performance, especially for slow or async functions
  * FunctionLogging - implement logging for functions
  * InstanceMemberRegistration - register non-static methods of classes
  * ParametersConversions - various additional parameter conversions
  * SuppressInDialogHandler - add an attribute to suppress function calls in the 'Insert Function' dialog.
  * TimingFunctionExecutionHandler - add timing for function calls

#### Project Properties Affecting Registration

Two MSBuild project properties in your `.csproj` file influence how Excel-DNA discovers and registers your functions and commands:

* **`ExcelAddInExplicitExports`**:
    * When set to `true`, only methods explicitly marked with `[ExcelFunction]` (for UDFs) or `[ExcelCommand]` (for macros) attributes will be registered.
    * When `false` (the default), Excel-DNA attempts to register all `public static` methods as UDFs and `public` methods in classes marked with `[ExcelCommandClass]` as commands, unless they are explicitly excluded (e.g., with `[ExcelIgnore]`).
    * **Recommendation:** It is highly recommended to set this property to `true` to avoid accidental exposure of unintended methods and to make your add-in's Excel interface explicit and clear. In a future version of Excel-DNA we might set the default value for this property to `true`.

    ```xml
    <PropertyGroup>
        <ExcelAddInExplicitExports>true</ExcelAddInExplicitExports>
    </PropertyGroup>
    ```

* **`ExcelAddInExplicitRegistration`**:
    * When set to `true`, this property prevents *all* automatic registration of functions and commands by Excel-DNA. If you set this to `true`, you are responsible for providing a complete registration mechanism, typically by implementing `IExcelAddIn.AutoOpen()` and using `ExcelRegistration` helpers, or by providing a full `ExcelRibbon` class that handles all registrations.
    * The default value is `false`.
    * This is an advanced setting and should generally remain `false` unless you have specific requirements to take full control over the entire registration lifecycle.

    ```xml
    <PropertyGroup>
        <ExcelAddInExplicitRegistration>false</ExcelAddInExplicitRegistration>
    </PropertyGroup>
    ```
    
### Object Handle Support with `[ExcelHandle]`
The `[ExcelHandle]` attribute provides a robust mechanism for managing .NET object references passed between your add-in and Excel. This allows complex objects to live in the .NET runtime while Excel manipulates them via handles (e.g., `#MyObject!123`).

* **Usage:** Apply `[ExcelHandle]` to:
    * **Return Values:** `[return: ExcelHandle]` on a function marks its return value to be passed to Excel as an object handle.
    * **Parameters:** `[ExcelHandle]` on a function parameter indicates that an incoming Excel value should be resolved from an object handle back to the .NET object.
    * **Classes and Structs:** Marking a class or struct with `[ExcelHandle]` means instances of this type will be treated as object handles when used as UDF parameters or return values.
    * **Primitive Types:** Can also be used with primitive types if you need to manage them as handles.

    ```csharp
    public class Calc // Assume this is a user-defined class
    {
        public double Value1 { get; set; }
        public double Value2 { get; set; }
        public Calc(double v1, double v2) { Value1 = v1; Value2 = v2; }
        public double Sum() { return Value1 + Value2; }
    }

    [ExcelHandle] // Marking the class itself
    public class CalcAsHandle
    {
        public double Value1 { get; set; }
        // ...
    }


    public static class HandleExamples
    {
        [ExcelFunction(Description = "Creates a Calc object and returns it as a handle.")]
        [return: ExcelHandle] // Mark return value as a handle
        public static Calc MyCreateCalc(double d1, double d2)
        {
            return new Calc(d1, d2);
        }

        [ExcelFunction(Description = "Takes a Calc handle and returns its sum.")]
        public static double MyCalcSum([ExcelHandle] Calc c) // Mark parameter as a handle
        {
            if (c == null) return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            return c.Sum();
        }

        [ExcelFunction(Description = "Creates an integer and returns it as a handle.")]
        [return: ExcelHandle]
        public static int MyCreateSquareIntObject(int i)
        {
            return i * i;
        }

        [ExcelFunction(Description = "Takes an integer handle and prints its info.")]
        public static string MyPrintIntObject([ExcelHandle] int i)
        {
            // Note: The value 'i' here is the actual int, resolved from the handle by Excel-DNA
            return $"IntObject (from handle) value = {i}";
        }
    }
    ```

* **Requirement for `Task<T>`/`IObservable<T>`:** Directly supports methods returning `Task<T>` (for asynchronous UDFs) or `IObservable<T>` (for streaming UDFs).
    * For user-defined types `T` returned within `Task<T>` or `IObservable<T>`, the `[return: ExcelHandle]` attribute is now generally required to ensure the object is correctly managed as an object handle.
    * Excel-DNA internally managed handle types (e.g., a class named `CalcExcelHandle` designed for this purpose) or primitive types do not require `[return: ExcelHandle]`.

    ```csharp
    // C# Task<T> example with a user-defined class 'Calc'
    public class Calc { /* ... content of Calc class ... */ public double Sum() { return 0; } }

    [ExcelFunction]
    [return: ExcelHandle] // Required for user-defined class 'Calc'
    public static async Task<Calc> MyTaskCreateCalc(int msDelay, double d1, double d2)
    {
        await Task.Delay(msDelay);
        return new Calc(); // Replace with actual Calc instantiation
    }

    [ExcelFunction]
    public static double GetCalcSum([ExcelHandle] Calc calc) => calc.Sum();

    // C# Task<T> with an Excel-DNA managed handle type
    [ExcelHandle]
    public class CalcExcelHandle { /* ... */ } // Placeholder for a handle type

    [ExcelFunction]
    // No [return: ExcelHandle] needed if CalcExcelHandle is inherently a handle type recognized by Excel-DNA
    public static async Task<CalcExcelHandle> MyTaskCreateCalcExcelHandle(int msDelay, double d1, double d2)
    {
        await Task.Delay(msDelay);
        return new CalcExcelHandle();
    }

    // C# IObservable<T> example with a user-defined class 'Calc'
    [ExcelFunction]
    [return: ExcelHandle] // Required for user-defined class 'Calc'
    public static IObservable<Calc> MyCalcObservable(double d1, double d2)
    {
        // return Observable.Return(new Calc()); // Replace with actual observable source
        return System.Reactive.Linq.Observable.Timer(TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(1))
                                   .Select(i => new Calc()); // Placeholder
    }
    ```


* **Assembly-Level `ExcelHandleExternal` Attribute:** For types defined in other assemblies that you want Excel-DNA to treat as handles, you can use the `[assembly: ExcelHandleExternal(typeof(YourExternalType))]` attribute.
    ```csharp
    // In AssemblyInfo.cs or a similar file:
    [assembly: ExcelDna.Integration.ExcelHandleExternal(typeof(MyCompany.SharedLibrary.SomeDataType))]
    ```


* **ObjectHandler Helper Class:**
    * The `ExcelDna.Integration.ObjectHandler` class provides explicit control over object handle lifecycle and management if needed.


## RTD Server Improvements

### Virtual Tracking Methods
The RTD server's tracking methods have been made virtual, allowing developers to override them in derived classes for custom tracking of notifications and refresh calls.
The affected methods are:
* `OnUpdateNotifyPostedInsideLock`
* `OnUpdateNotifyInvokedInsideLock`
* `OnRefreshDataProcessedInsideLock`

The sample project RtdClock-Watchdog (in the Excel-DNA\Samples repository) shows how these methods can be used to add a watchdog to an RTD server that monitors whether the RTD data is being fetched timeously.

## Other Changes and Enhancements

### Support for newer .NET Versions
* With this release we extend support to .NET 9 (and preliminary support for .NET 10)

### Help File Packaging
* Added support for packing compiled HTML Help (`.chm`) files into the packed add-in (`.xll`).

### RuntimeFrameworkVersion Support

We now support specifying an exact runtime version (for .NET 5+) by honoring the `RuntimeFrameworkVersion` project property.

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net8.0-windows</TargetFramework>
    <RuntimeFrameworkVersion>8.0.6</RuntimeFrameworkVersion>
    <RollForward>Disable</RollForward>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="ExcelDna.AddIn" Version="1.9.0-*" />
  </ItemGroup>  
</Project>
```

You'll see the version and RollForward option reflected in the generated .runtimeconfig.json file

```json
  "runtimeOptions": {
    "tfm": "net8.0",
    "rollForward": "Disable",
    "framework": {
      "name": "Microsoft.NETCore.App",
      "version": "8.0.6"
    },
    "configProperties": {
      "System.Runtime.Serialization.EnableUnsafeBinaryFormatterSerialization": false
    }
  }
```

This is implemented by adding an attribute in the generated .dna file (generated by the build when you don't have a .dna file in your project directory)

<DnaLibrary RuntimeFrameworkVersion="8.0.6" RollForward="Disable" Name="TestExactVersion Add-In" RuntimeVersion="v8.0" xmlns="http://schemas.excel-dna.net/addin/2020/07/dnalibrary">
  <ExternalLibrary Path="TestExactVersion.dll" ExplicitExports="false" LoadFromBytes="true" Pack="true" IncludePdb="false" />
</DnaLibrary>

Then when loading the add-in, if the matching runtime version cannot be loaded, you should see a clear error message.

You can also use other RollForward options like "LatestPatch" as expected.

-Govert (govert@dnakode.com)
