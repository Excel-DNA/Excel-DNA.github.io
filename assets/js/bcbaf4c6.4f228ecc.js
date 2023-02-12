"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[1606],{3905:(e,n,t)=>{t.d(n,{Zo:()=>p,kt:()=>m});var r=t(7294);function a(e,n,t){return n in e?Object.defineProperty(e,n,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[n]=t,e}function i(e,n){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);n&&(r=r.filter((function(n){return Object.getOwnPropertyDescriptor(e,n).enumerable}))),t.push.apply(t,r)}return t}function l(e){for(var n=1;n<arguments.length;n++){var t=null!=arguments[n]?arguments[n]:{};n%2?i(Object(t),!0).forEach((function(n){a(e,n,t[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):i(Object(t)).forEach((function(n){Object.defineProperty(e,n,Object.getOwnPropertyDescriptor(t,n))}))}return e}function o(e,n){if(null==e)return{};var t,r,a=function(e,n){if(null==e)return{};var t,r,a={},i=Object.keys(e);for(r=0;r<i.length;r++)t=i[r],n.indexOf(t)>=0||(a[t]=e[t]);return a}(e,n);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(r=0;r<i.length;r++)t=i[r],n.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(a[t]=e[t])}return a}var s=r.createContext({}),c=function(e){var n=r.useContext(s),t=n;return e&&(t="function"==typeof e?e(n):l(l({},n),e)),t},p=function(e){var n=c(e.components);return r.createElement(s.Provider,{value:n},e.children)},d={inlineCode:"code",wrapper:function(e){var n=e.children;return r.createElement(r.Fragment,{},n)}},u=r.forwardRef((function(e,n){var t=e.components,a=e.mdxType,i=e.originalType,s=e.parentName,p=o(e,["components","mdxType","originalType","parentName"]),u=c(t),m=a,b=u["".concat(s,".").concat(m)]||u[m]||d[m]||i;return t?r.createElement(b,l(l({ref:n},p),{},{components:t})):r.createElement(b,l({ref:n},p))}));function m(e,n){var t=arguments,a=n&&n.mdxType;if("string"==typeof e||a){var i=t.length,l=new Array(i);l[0]=u;var o={};for(var s in n)hasOwnProperty.call(n,s)&&(o[s]=n[s]);o.originalType=e,o.mdxType="string"==typeof e?e:a,l[1]=o;for(var c=2;c<i;c++)l[c]=t[c];return r.createElement.apply(null,l)}return r.createElement.apply(null,t)}u.displayName="MDXCreateElement"},3844:(e,n,t)=>{t.r(n),t.d(n,{assets:()=>s,contentTitle:()=>l,default:()=>d,frontMatter:()=>i,metadata:()=>o,toc:()=>c});var r=t(7462),a=(t(7294),t(3905));const i={title:"COM Server Support"},l=void 0,o={unversionedId:"archive/guides/com-server-support",id:"archive/guides/com-server-support",title:"COM Server Support",description:"Functions and macros created in an Excel-DNA add-in can be called directly from Excel VBA by using Application.Run(...). However, .NET also supports creating rich object models that are exported as COM libraries, which can be Tools->Referenced in VBA. Excel-DNA has some advanced support to host COM-exported objects from Excel-DNA add-ins, giving some advantages over the regular .NET \u2018Register for COM interop\u2019 hosting approach:",source:"@site/docs/archive/guides/com-server-support.md",sourceDirName:"archive/guides",slug:"/archive/guides/com-server-support",permalink:"/docs/archive/guides/com-server-support",draft:!1,tags:[],version:"current",frontMatter:{title:"COM Server Support"},sidebar:"tutorialSidebar",previous:{title:"Checking and Downloading Updates in .NET",permalink:"/docs/archive/guides/checking-and-downloading-updates-in-dotnet"},next:{title:"Configuring NLog Logging",permalink:"/docs/archive/guides/configuring-nlog-logging"}},s={},c=[],p={toc:c};function d(e){let{components:n,...t}=e;return(0,a.kt)("wrapper",(0,r.Z)({},p,t,{components:n,mdxType:"MDXLayout"}),(0,a.kt)("p",null,"Functions and macros created in an Excel-DNA add-in can be called directly from Excel VBA by using Application.Run(...). However, .NET also supports creating rich object models that are exported as COM libraries, which can be Tools->Referenced in VBA. Excel-DNA has some advanced support to host COM-exported objects from Excel-DNA add-ins, giving some advantages over the regular .NET \u2018Register for COM interop\u2019 hosting approach:"),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"COM objects that are created via the Excel-DNA COM server support will be active in the same AppDomain as the rest of the add-in, allowing direct shared access to static variables, internal caches etc.")),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"COM registration for classes hosted by Excel-DNA does not require administrative access (even when registered via RegSvr32.exe).")),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"Everything needed for the COM server can be packed in a single-file .xll add-in, including the type library used for IntelliSense support in VBA.\nMikael Katajam\xe4ki has written some detailed tutorial posts on his ",(0,a.kt)("a",{parentName:"p",href:"http://mikejuniperhill.blogspot.com/"},"Excel in Finance")," blog that explore this Excel-DNA feature, with detailed explanation, step-by-step instructions, screen shots and further links. See:")),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("a",{parentName:"p",href:"http://mikejuniperhill.blogspot.com/2014/03/interfacing-c-and-vba-with-exceldna-no.html"},"Interfacing C# and VBA with Excel-DNA (no intellisense support)"))),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("a",{parentName:"p",href:"http://mikejuniperhill.blogspot.com/2014/03/interfacing-c-and-vba-with-exceldna_16.html"},"Interfacing C# and VBA with Excel-DNA (with intellisense support)")))),(0,a.kt)("p",null,"Note that these techniques would works equally well with code written in VB.NET, allowing you to port VB/VBA libraries to VB.NET with Excel-DNA and then use these from VBA."),(0,a.kt)("hr",null),(0,a.kt)("p",null,"COM visible classes in ",(0,a.kt)("inlineCode",{parentName:"p"},"ExternalLibrary")," tags marked ",(0,a.kt)("inlineCode",{parentName:"p"},'ComServer="true"'),", and COM visible classes that implement IRtdServer can be activated through the .xll directly. Even if the add-in is not loaded in Excel, such objects can be created in VBA."),(0,a.kt)("p",null,"These classes are (persistently) registered by calling ",(0,a.kt)("inlineCode",{parentName:"p"},"regsvr32 <MyAddin>.xll")," or dynamically by the add-in (for example in an ",(0,a.kt)("inlineCode",{parentName:"p"},"AutoOpen")," method) by calling ",(0,a.kt)("inlineCode",{parentName:"p"},"ComServer.DllRegisterServer()"),", and\nunregistered by ",(0,a.kt)("inlineCode",{parentName:"p"},"regsvr32 /u <MyAddin>.xll")," or by ",(0,a.kt)("inlineCode",{parentName:"p"},"ComServer.DllUnregisterServer()"),"."),(0,a.kt)("p",null,"Following are short examples both in C-Sharp and VB, these only demonstrate the unreferenced (late-bound) technique:"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-csharp"},'using ExcelDna.Integration;\nusing ExcelDna.ComInterop;\nusing System.Runtime.InteropServices;\n\n[ComVisible(true)]\n[ClassInterface(ClassInterfaceType.AutoDispatch)]\n[ProgId("ComAddin.FunctionLibrary")]\npublic class AccessibleFunctions\n{\n    public double add(double x, double y)\n    {\n        return x + y;\n    }\n}\n\n[ComVisible(false)]\nclass ExcelAddin : IExcelAddIn\n{\n    public void AutoOpen()\n    {\n        ComServer.DllRegisterServer();\n    }\n    public void AutoClose()\n    {\n        ComServer.DllUnregisterServer();\n    }\n}\n')),(0,a.kt)("p",null,"The same in VB.NET"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb"},'Imports ExcelDna.Integration\nImports ExcelDna.ComInterop\nImports System.Runtime.InteropServices\n\n<ClassInterface(ClassInterfaceType.AutoDispatch)>\n<ProgId("ComAddin.FunctionLibrary")>\n<ComVisible(True)>\nPublic Class AccessibleFunctions\n\n    Public Function add(x As Double, y As Double)\n        Return x + y\n    End Function\n\nEnd Class\n\n<ComVisible(false)>\nPublic Class AddInEvents\n    Implements IExcelAddIn\n\n    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen\n        ComServer.DllRegisterServer()\n    End Sub\n\n    Public Sub AutoClose() Implements IExcelAddIn.AutoClose\n        ComServer.DllUnregisterServer()\n    End Sub\nEnd Class\n')),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-xml"},'<DnaLibrary Name="ComAddin" RuntimeVersion="v4.0">\n  <ExternalLibrary Path="ComAddin.dll" ComServer="true" />\n</DnaLibrary>\n')),(0,a.kt)("p",null,"Usage in VBA:"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb"},'Option Explicit\n\nSub tester()\n    Dim lib As Object: Set lib = CreateObject("ComAddin.FunctionLibrary")\n    Debug.Print lib.Add(12, 13)\n    Set lib = Nothing\nEnd Sub\n')),(0,a.kt)("p",null,"The VB.Net example is also available in the Samples repository ",(0,a.kt)("a",{parentName:"p",href:"https://github.com/Excel-DNA/Samples/tree/master/ComServerVB"},"ComServerVB"),' after building it, start Excel by opening bin/release/ComAddin.xll or ComAddin64.xll (depending on your bitness) and enter the code under "Usage in VBA:" somewhere.'),(0,a.kt)("p",null,"Such classes can be accessed directly as RTD servers or from VBA using ",(0,a.kt)("inlineCode",{parentName:"p"},'CreateObject("MyServer.ItsProgId")'),", and will be loaded in the add-in's AppDomain.\n(The add-in need not be loaded for registered classes to be accessed through COM.)"),(0,a.kt)("p",null,"A type library (.tlb) can be created for the assembly using tlbexp.exe, and will be registered if available (if the ",(0,a.kt)("inlineCode",{parentName:"p"},".tlb")," is found next to the ",(0,a.kt)("inlineCode",{parentName:"p"},".dll"),"). If the assembly is packed in the .xll, the type library will be packed too.\nAn example for this can be found in ",(0,a.kt)("a",{parentName:"p",href:"https://github.com/Excel-DNA/Samples/tree/master/DnaComServer"},"DnaComServer"),"."))}d.isMDXComponent=!0}}]);