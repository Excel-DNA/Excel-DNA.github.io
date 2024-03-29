---
title: "Integrating with VBA"
---
Excel-DNA can make it easy to call between .NET and VBA. This means existing VBA code need not be rewritten. And end users are likely to find VBA much easier to develop in.

To do this, create an Excel-DNA project, and register the a class that will be the entry point from VBA as follows:

```csharp
public class AddInRoot : IExcelAddIn
{
    public void AutoOpen()
    {
        try
        {
            var com_addin = new AddInComRoot();
            com_addin.GetType().InvokeMember("DnaLibrary", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, com_addin, new object[]() { DnaLibrary.CurrentLibrary });

            ExcelComAddInHelper.LoadComAddIn(com_addin);
        }
        catch (Exception e)
        {
            MessageBox.Show("Error loading COM AddIn: " + e.ToString());
        }
    }

    public void AutoClose()
    {
    }
}


[ComVisible(true)]
public class AddInComRoot : ExcelDna.Integration.CustomUI.ExcelComAddIn
           // : IDTExtensibility2, i.e. COM "AddIn".ExcelDNA finds this by magic.
{
    MyAddinObject _helper;

    public AddInComRoot()
    {
    }

    public override void OnConnection(object Application,
        ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
    {
        _helper = new MyAddinObject();

        AddInInst.GetType().InvokeMember("Object",
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
            null,
            AddInInst,
            new object[]() { _helper });
    }

    public override void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
    {
    }

    public override void OnAddInsUpdate(ref Array custom)
    {
    }

    public override void OnStartupComplete(ref Array custom)
    {
    }

    public override void OnBeginShutdown(ref Array custom)
    {
    }
}

// This becomes the VBA addin.Object
[ComVisible(true)]
public class MyAddinObject
{
    public string SayHello()
    {
        return "Hello from the future!";
    }

    public string ActiveCell3()
    {
        var app = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Range r = app.ActiveCell;
        return "ActiveCell3: " + r.Value;
    }
}
```

We then need to get a handle to the Excel DNA file and call these methods. We need to search through Descriptions because we cannot set the ProgId directly. The object is nothing test is also required as one can easily end up with dead entries in the Addins list. CustomUI ribbon objects may also appear in this list, so test for the entire Description.

```csharp
' In VBA:
Sub TestDnaComAddIn()
    Dim cai As COMAddIn
    Dim obj As Object
    For Each cai In Application.COMAddIns
        ' Could check cai.Connect to see if it is loaded.
        Debug.Print cai.Description, cai.GUID
        If InStr(cai.Description, "MyTitle (COM Add-in Helper)") Then
            Set obj = cai.Object
            If obj Is Nothing Then
              Debug.Print "ObjNothing"
            Else
              Debug.Print obj.SayHello(), obj.ActiveCell3
            End If
        End If
    Next
End Sub
```

To call from .NET to VBA it is probably easiest to just use `Application.Run`.

However, be careful about asynchronous calls, see [Performing Asynchronous Work](./performing-asynchronous-work).
