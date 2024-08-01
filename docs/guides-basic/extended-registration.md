---
title: "Extended Registration"
---

## Nullable parameter

```csharp
[ExcelFunction]
public static string NullableDouble(double? d)
{
    return "Nullable VAL: " + (d.HasValue ? d : "NULL");
}
```

| Cell  | Formula              | Result 
| ----- | -------------------- | ------ 
| A1    | =NullableDouble(1.2) | Nullable VAL: 1.2       
| A2    | =NullableDouble()    | Nullable VAL: NULL     

## Optional parameter

```csharp
[ExcelFunction]
public static string OptionalDouble(double d = 1.23)
{
    return "Optional VAL: " + d.ToString();
}
```

| Cell  | Formula              | Result 
| ----- | -------------------- | ------ 
| A1    | =OptionalDouble(2.3) | Optional VAL: 2.3       
| A2    | =OptionalDouble()    | Optional VAL: 1.23  

## Range parameter

```csharp
[ExcelFunction]
public static string Range(Microsoft.Office.Interop.Excel.Range r)
{
    return r.Address;
}
```

| Cell  | Formula            | Result 
| ----- | ------------------ | ------ 
| A1    | =Range(B2)         | $B$2       
| A2    | =Range(B2:C4)      | $B$2:$C$4  
| A3    | =Range((B2,D5:E6)) | $B$2,$D$5:$E$6  

## Enums parameter and return value

```csharp
[ExcelFunction]
public static string Enum(DateTimeKind e)
{
    return "Enum VAL: " + e.ToString();
}

[ExcelFunction]
public static DateTimeKind EnumReturn(string s)
{
    return Enum.Parse<DateTimeKind>(s);
}
```

| Cell  | Formula                    | Result 
| ----- | -------------------------- | ------ 
| A1    | =Enum("Unspecified")       | Enum VAL: Unspecified       
| A2    | =Enum("Local")             | Enum VAL: Local  
| A3    | =Enum(1)                   | Enum VAL: Utc
| A4    | =EnumReturn("Unspecified") | Unspecified 
| A5    | =EnumReturn("Local")       | Local 

## String array parameter

```csharp
[ExcelFunction]
public static string StringArray(string[] s)
{
    return "StringArray VALS: " + string.Concat(s);
}
```

| Cell  | Formula             | Result 
| ----- | ------------------- | ------ 
| A1    | 01                  |  
| A2    | 2.30                | 
| A3    | World               | 
| B1    | =StringArray(A1:A3) | StringArray VALS: 12.3World

## String array 2D parameter

```csharp
[ExcelFunction]
public static string StringArray2D(string[,] s)
{
    string result = "";
    for (int i = 0; i < s.GetLength(0); i++)
    {
        for (int j = 0; j < s.GetLength(1); j++)
        {
            result += s[i, j];
        }

        result += " ";
    }

    return $"StringArray2D VALS: {result}";
}
```

| Cell  | Formula               | Result 
| ----- | --------------------- | ------ 
| A1    | 01                    |  
| A2    | 2.30                  | 
| A3    | Hello                 | 
| B1    | 5                     |  
| B2    | 6.7                   | 
| B3    | World                 | 
| C1    | =StringArray2D(A1:B3) | StringArray2D VALS: 15 2.36.7 HelloWorld 

## params parameter

```csharp
[ExcelFunction]
public static string ParamsFunc1(
    [ExcelArgument(Name = "first.Input", Description = "is a useful start")]
    object input,
    [ExcelArgument(Description = "is another param start")]
    string QtherInpEt,
    [ExcelArgument(Name = "Value", Description = "gives the Rest")]
    params object[] args)
{
    return input + "," + QtherInpEt + ", : " + args.Length;
}

[ExcelFunction]
public static string ParamsFunc2(
    [ExcelArgument(Name = "first.Input", Description = "is a useful start")]
    object input,
    [ExcelArgument(Name = "second.Input", Description = "is some more stuff")]
    string input2,
    [ExcelArgument(Description = "is another param ")]
    string QtherInpEt,
    [ExcelArgument(Name = "Value", Description = "gives the Rest")]
    params object[] args)
{
    var content = string.Join(",", args.Select(ValueType => ValueType.ToString()));
    return input + "," + input2 + "," + QtherInpEt + ", " + $"[{args.Length}: {content}]";
}

[ExcelFunction]
public static string ParamsJoinString(string separator, params string[] values)
{
    return String.Join(separator, values);
}
```

| Cell  | Formula                                     | Result 
| ----- | ------------------------------------------- | ------ 
| A1    | =ParamsFunc1(1,\"2\",4,5)                   | 1,2, : 2
| A2    | =ParamsFunc2(\"a\",,\"c\",\"d\",,\"f\")     | a,,c, [3: d,ExcelDna.Integration.ExcelMissing,f]
| A3    | =ParamsJoinString(\"//\",\"5\",\"4\",\"3\") | 5//4//3

## Async functions and tasks

```csharp
[ExcelAsyncFunction]
public static string AsyncHello(string name, int msToSleep)
{
    return $"Hello async {name}";
}

[ExcelAsyncFunction]
public static async Task<string> AsyncTaskHello(string name, int msDelay)
{
    await Task.Delay(msDelay);
    return $"Hello async task {name}";
}

[ExcelFunction]
public static Task<string> TaskHello(string name)
{
    return Task.FromResult($"Hello task {name}");
}
```

## Object handles

Create and reuse .NET objects:

```csharp
public class Calc
{
    private double d1, d2;

    public Calc(double d1, double d2)
    {
        this.d1 = d1;
        this.d2 = d2;
    }

    public double Sum()
    {
        return d1 + d2;
    }
}

[ExcelFunction]
public static Calc CreateCalc(double d1, double d2)
{
    return new Calc(d1, d2);
}

[ExcelFunction]
public static double CalcSum(Calc c)
{
    return c.Sum();
}
```

| Cell  | Formula               | Result |
| ----- | --------------------- | ------ |
| A1    | =CreateCalc(1.2, 3.4) |        |
| A2    | =CalcSum(A1)          | 4.6    |


Thread safe creation and use is supported:

```csharp
[ExcelFunction(IsThreadSafe = true)]
public static Calc CreateCalcTS(double d1, double d2)
{
    return new Calc(d1, d2);
}

[ExcelFunction(IsThreadSafe = true)]
public static double CalcSumTS(Calc c)
{
    return c.Sum();
}
```

Object resources are automatically disposed when no longer used:

```csharp
public class DisposableObject : IDisposable
{
    public static int ObjectsCount { get; private set; } = 0;
    private bool disposedValue;

    public DisposableObject()
    {
        ++ObjectsCount;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {
                --ObjectsCount;
            }

            disposedValue = true;
        }
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}

[ExcelFunction]
public static DisposableObject CreateDisposableObject(int x)
{
    return new DisposableObject();
}
```

## User defined parameter conversions

```csharp
public class TestType1
{
    public string Value;

    public TestType1(string value)
    {
        Value = value;
    }
}

public class TestType2
{
    public string Value;

    public TestType2(string value)
    {
        Value = value;
    }
}

[ExcelParameterConversion]
public static TestType2 Order1ToTestType2FromTestType1(TestType1 value)
{
    return new TestType2("From TestType1 " + value.Value);
}

[ExcelParameterConversion]
public static TestType1 Order2ToTestType1(string value)
{
    return new TestType1(value);
}

[ExcelParameterConversion]
public static TestType1 Order3ToTestType1Also(string value)
{
    return new TestType1("Also " + value);
}

[ExcelParameterConversion]
public static Version ToVersion(string s)
{
    return new Version(s);
}

[ExcelFunction]
public static string TestType1(TestType1 tt)
{
    return "The TestType1 value is " + tt.Value;
}

[ExcelFunction]
public static string TestType2(TestType2 tt)
{
    return "The TestType2 value is " + tt.Value;
}

[ExcelFunction]
public static string Version2(Version v)
{
    return "The Version value with field count 2 is " + v.ToString(2);
}
```

| Cell  | Formula              | Result 
| ----- | -------------------- | ------ 
| A1    | =Version2("4.3.2.1") |  The Version value with field count 2 is 4.3
| A2    | =TestType1("world")  | The TestType1 value is world
| A3    | =TestType2("world2") | The TestType2 value is From TestType1 world2

User defined parameter conversions are sorted alphabetically by function name. 

More complex type conversions (like `TestType2 Order1ToTestType2FromTestType1(TestType1 value)`) should be ordered before simpler type conversions they dependent on (like `TestType1 Order2ToTestType1(string value)`). 

Subsequent multiple conversions for the same type (like `TestType1 Order3ToTestType1Also(string value)`) are ignored and the first one (like `TestType1 Order2ToTestType1(string value)`) is used.

## Function execution handler

Monitor Excel functions execution with a custom handler, marked with `ExcelFunctionExecutionHandlerSelector` attribute:

```csharp
internal class FunctionLoggingHandler : FunctionExecutionHandler
{
    public int? ID { get; set; }

    public override void OnEntry(FunctionExecutionArgs args)
    {
        // FunctionExecutionArgs gives access to the function name and parameters,
        // and gives some options for flow redirection.

        // Tag will flow through the whole handler
        if (ID.HasValue)
            args.Tag = $"ID={ID.Value} ";
        else
            args.Tag = "";
        args.Tag += args.FunctionName;

        Logger.Log($"{args.Tag} - OnEntry - Args: {args.Arguments.Select(arg => arg.ToString())}");
    }

    public override void OnSuccess(FunctionExecutionArgs args)
    {
        Logger.Log($"{args.Tag} - OnSuccess - Result: {args.ReturnValue}");
    }

    public override void OnException(FunctionExecutionArgs args)
    {
        Logger.Log($"{args.Tag} - OnException - Message: {args.Exception}");
    }

    public override void OnExit(FunctionExecutionArgs args)
    {
        Logger.Log($"{args.Tag} - OnExit");
    }

    [ExcelFunctionExecutionHandlerSelector]
    public static IFunctionExecutionHandler LoggingHandlerSelector(IExcelFunctionInfo functionInfo)
    {
        if (functionInfo.CustomAttributes.OfType<LoggingAttribute>().Any())
        {
            var loggingAtt = functionInfo.CustomAttributes.OfType<LoggingAttribute>().First();
            return new FunctionLoggingHandler { ID = loggingAtt.ID };
        }

        return new FunctionLoggingHandler();
    }
}
```

The default return value for async functions that are in process is #N/A. You can, for example, return the newer #GETTING_DATA error code creating the following function execution handler:

```csharp
internal class AsyncReturnHandler : FunctionExecutionHandler
{
    public override void OnSuccess(FunctionExecutionArgs args)
    {
        if (args.ReturnValue.Equals(ExcelError.ExcelErrorNA))
            args.ReturnValue = ExcelError.ExcelErrorGettingData;
    }

    [ExcelFunctionExecutionHandlerSelector]
    public static IFunctionExecutionHandler AsyncReturnHandlerSelector(IExcelFunctionInfo functionInfo)
    {
        return new AsyncReturnHandler();
    }
}
```

## Function registration processing

You can implement custom function wrappers during registration using `ExcelFunctionProcessor` attribute:

```csharp
public interface IExcelFunctionInfo
{
    ExcelFunctionAttribute FunctionAttribute { get; }
    List<IExcelFunctionParameter> Parameters { get; }
    IExcelFunctionReturn Return { get; }
    List<object> CustomAttributes { get; }

    LambdaExpression FunctionLambda { get; set; }
}

[ExcelFunctionProcessor]
public static IEnumerable<IExcelFunctionInfo> ProcessFunctions(IEnumerable<IExcelFunctionInfo> registrations, IExcelFunctionRegistrationConfiguration config)
```