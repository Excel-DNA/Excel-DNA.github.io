"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[4580],{3905:(e,t,n)=>{n.d(t,{Zo:()=>u,kt:()=>d});var a=n(7294);function i(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function r(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function s(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?r(Object(n),!0).forEach((function(t){i(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):r(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function o(e,t){if(null==e)return{};var n,a,i=function(e,t){if(null==e)return{};var n,a,i={},r=Object.keys(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||(i[n]=e[n]);return i}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(i[n]=e[n])}return i}var l=a.createContext({}),c=function(e){var t=a.useContext(l),n=t;return e&&(n="function"==typeof e?e(t):s(s({},t),e)),n},u=function(e){var t=c(e.components);return a.createElement(l.Provider,{value:t},e.children)},p={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},m=a.forwardRef((function(e,t){var n=e.components,i=e.mdxType,r=e.originalType,l=e.parentName,u=o(e,["components","mdxType","originalType","parentName"]),m=c(n),d=i,f=m["".concat(l,".").concat(d)]||m[d]||p[d]||r;return n?a.createElement(f,s(s({ref:t},u),{},{components:n})):a.createElement(f,s({ref:t},u))}));function d(e,t){var n=arguments,i=t&&t.mdxType;if("string"==typeof e||i){var r=n.length,s=new Array(r);s[0]=m;var o={};for(var l in t)hasOwnProperty.call(t,l)&&(o[l]=t[l]);o.originalType=e,o.mdxType="string"==typeof e?e:i,s[1]=o;for(var c=2;c<r;c++)s[c]=n[c];return a.createElement.apply(null,s)}return a.createElement.apply(null,n)}m.displayName="MDXCreateElement"},9608:(e,t,n)=>{n.r(t),n.d(t,{assets:()=>l,contentTitle:()=>s,default:()=>p,frontMatter:()=>r,metadata:()=>o,toc:()=>c});var a=n(7462),i=(n(7294),n(3905));const r={title:"Asynchronous Functions"},s=void 0,o={unversionedId:"guides-basic/asynchronous-functions",id:"guides-basic/asynchronous-functions",title:"Asynchronous Functions",description:"Excel-DNA has a core implementation to support asynchronous functions. Two primary ways this could be implemented is through:",source:"@site/docs/guides-basic/asynchronous-functions.md",sourceDirName:"guides-basic",slug:"/guides-basic/asynchronous-functions",permalink:"/docs/guides-basic/asynchronous-functions",draft:!1,tags:[],version:"current",frontMatter:{title:"Asynchronous Functions"},sidebar:"tutorialSidebar",previous:{title:"Accepting Range Parameters in UDFs",permalink:"/docs/guides-basic/accepting-range-parameters-in-udfs"},next:{title:"Creating a Help File",permalink:"/docs/guides-basic/creating-a-help-file"}},l={},c=[{value:"Task-based Async Functions",id:"task-based-async-functions",level:2},{value:"Usage",id:"usage",level:3},{value:"Remarks",id:"remarks",level:3},{value:"Additional Example",id:"additional-example",level:3},{value:"RTD-based Async Functions",id:"rtd-based-async-functions",level:2},{value:"Usage",id:"usage-1",level:3},{value:"Remarks",id:"remarks-1",level:3},{value:"Additional Example",id:"additional-example-1",level:3}],u={toc:c};function p(e){let{components:t,...n}=e;return(0,i.kt)("wrapper",(0,a.Z)({},u,n,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("p",null,"Excel-DNA has a core implementation to support asynchronous functions. Two primary ways this could be implemented is through:"),(0,i.kt)("ol",null,(0,i.kt)("li",{parentName:"ol"},"Task-based async functions (preferred) "),(0,i.kt)("li",{parentName:"ol"},"RTD-based async functions")),(0,i.kt)("p",null,"It is worth noting that RTD-based functions use the same underlying mechanism as Task-based functions. However, it is easier to use Task-based functions as the asynchronous concept is abstracted."),(0,i.kt)("h2",{id:"task-based-async-functions"},"Task-based Async Functions"),(0,i.kt)("p",null,"Task-based functions are the preferred way of async implementation. "),(0,i.kt)("p",null,"Both ",(0,i.kt)("a",{parentName:"p",href:"https://github.com/Excel-DNA/Registration/blob/master/Source/ExcelDna.Registration/Utils/AsyncTaskUtil.cs"},"AsyncTaskUtil.cs")," and ",(0,i.kt)("a",{parentName:"p",href:"https://github.com/Excel-DNA/Registration/blob/master/Source/ExcelDna.Registration/Utils/Disposables.cs"},"Disposables.cs")," from ",(0,i.kt)("a",{parentName:"p",href:"https://github.com/Excel-DNA/Registration"},"Excel-DNA Registration")," library must be included in the project's solution. Once included, the following line must be added at the top of source code file: "),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-csharp"},"using ExcelDna.Registration.Utils;\n")),(0,i.kt)("p",null,(0,i.kt)("strong",{parentName:"p"},"NOTE:")," The ",(0,i.kt)("a",{parentName:"p",href:"https://github.com/Excel-DNA/Registration"},"Excel-DNA Registration")," helper is an extension library that is used to simplify (and modify) the function registration process at runtime. The helper includes conversions to assist in registering task-based async functions. For this example, the helper extensions library is not referenced but the required utility code is imported directly into the project. "),(0,i.kt)("h3",{id:"usage"},"Usage"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"The following example, accepts a target URL (string) and returns a string (object) of characters that was downloaded from the given target URL. The asynchronous UDF should call ",(0,i.kt)("inlineCode",{parentName:"li"},"AsyncTaskUtil.RunTask")," as follows:")),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-csharp"},"//The main function that is exposed to Excel.\npublic static object DownloadStringFromURL(string url)\n{\n    var functionName = nameof(DownloadStringFromURL);\n    var parameters = new object[] { url }; \n    HttpClient myHttpClient = new HttpClient();\n    \n    return AsyncTaskUtil.RunTask(functionName, parameters, async () =>\n    {\n        //The actual asyncronous block of code to execute.\n        return await myHttpClient.GetStringAsync(url);\n    })\n}\n")),(0,i.kt)("h3",{id:"remarks"},"Remarks"),(0,i.kt)("p",null,"The parameters of ",(0,i.kt)("inlineCode",{parentName:"p"},"AsyncTaskUtil.RunTask")," are:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},(0,i.kt)("inlineCode",{parentName:"p"},"string functionName")," - the name of the async function. Used in combination with the ",(0,i.kt)("inlineCode",{parentName:"p"},"parameters")," value to identify this async function by the .NET framework for its internal threading operations. "),(0,i.kt)("p",{parentName:"li"},(0,i.kt)("strong",{parentName:"p"},"NOTE:")," Ensure to enclose the function name within the ",(0,i.kt)("inlineCode",{parentName:"p"},"nameof()")," expression.")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},(0,i.kt)("inlineCode",{parentName:"p"},"object parameters")," - the set of parameters the function is being called with. Although it can be a single object (e.g. a string) it is preferred to enclose the parameter/s in an object","[","]"," array. "),(0,i.kt)("p",{parentName:"li"},(0,i.kt)("strong",{parentName:"p"},"NOTE:")," Ensure to include all the parameters to the UDF as it is used in combination with the ",(0,i.kt)("inlineCode",{parentName:"p"},"functionName")," value to identify this async function by the .NET framework for its internal threading operations. ")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},(0,i.kt)("inlineCode",{parentName:"p"},"Func<Task<T>>")," - a delegate function (can be anonymous) that will be executed asynchronously."))),(0,i.kt)("h3",{id:"additional-example"},"Additional Example"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"The following example, accepts a list of target IPs/hostnames (object[]) and the number of times to ping each target (int). The function returns an array of boolean values for each target indicating ",(0,i.kt)("inlineCode",{parentName:"li"},"true")," if the target is reachable for all ping attempts otherwise ",(0,i.kt)("inlineCode",{parentName:"li"},"false"),". ")),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-csharp"},"//The main function that is exposed to Excel.\npublic static object TaskedPingTargets(object[] targets, int pingCount)\n{\n    var functionName = nameof(TaskedPingTargets);\n    var parameters = new object[] { targets, pingCount };\n    \n    //The task to run is an anonymous async function. \n    //It calls an async task per target and waits for all tasks to complete.\n    //Once all tasks are complete, it returns an array of boolean values stating\n    //each target reachability status.\n    return AsyncTaskUtil.RunTask(functionName, parameters, async () =>\n    {\n        //create an empty list ot tasks.\n        List<Task<PingReply[]>> tasks = new List<Task<PingReply[]>>();\n        \n        //add a PingTargetAsync task per target to the task list and execute it.\n        foreach (string target in targets)\n        {\n            tasks.Add(PingTargetAsync(target, pingCount));\n        }\n\n        //wait for all results to arrive from the tasks.\n        PingReply[][] results = await Task.WhenAll<PingReply[]>(tasks);\n        object[] toReturn = new object[targets.Length];\n\n        //format output to return.\n        for (int i = 0; i<targets.Length; i++)\n        {\n            toReturn[i] = true;\n            for (int j = 0; j<pingCount; j++)\n            {\n                if (results[i][j].Status != IPStatus.Success)\n                {\n                    toReturn[i] = false;\n                    break;\n                }\n            }\n        }\n        return toReturn;\n    });\n}\n//The actual asyncronous payload to execute. This function is not exposed to Excel. \n//NOTE: unlike the previous example, this task returns PingReply[] not double[].\nprivate static async Task<PingReply[]> PingTargetAsync(string target, int pingCount)\n{\n    Ping ping = new Ping();\n    PingReply[] replies = new PingReply[pingCount];\n\n    for (int i = 0; i < pingCount; i++)\n    {\n        replies[i] = await ping.SendPingAsync(target);\n    }\n\n    return replies;\n}\n")),(0,i.kt)("h2",{id:"rtd-based-async-functions"},"RTD-based Async Functions"),(0,i.kt)("p",null,"The RTD-based functions can also be used for async functionality. However, they are a less preferred method of async implementation."),(0,i.kt)("h3",{id:"usage-1"},"Usage"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"The following example function accepts a string value in milliseconds (which is parsed to an int, later on), and sleeps for that duration. The asynchronous UDF should call ",(0,i.kt)("inlineCode",{parentName:"li"},"AsyncUtil.Run")," as follows:")),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-csharp"},'//The main function that is exposed to Excel.\npublic static object SleepAsync(string ms)\n{\n    var functionName = nameof(SleepAsync);\n    var parameters = new object[] { ms };\n    \n    //The task to run is an anonymous function. All it does is sleep for a certain amount of milliseconds.\n    return ExcelAsyncUtil.Run(nameof(functionName), parameters, () =>\n    {\n        Debug.Print("{1:HH:mm:ss.fff} Sleeping for {0} ms", ms, DateTime.Now);\n        Thread.Sleep(int.Parse(ms));\n\n        Debug.Print("{1:HH:mm:ss.fff} Done sleeping {0} ms", ms, DateTime.Now);\n        return "Woke Up at " + DateTime.Now.ToString("1:HH:mm:ss.fff");\n    });\n}\n')),(0,i.kt)("h3",{id:"remarks-1"},"Remarks"),(0,i.kt)("p",null,"The parameters of ",(0,i.kt)("inlineCode",{parentName:"p"},"ExcelAsyncUtil.Run")," are:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},(0,i.kt)("inlineCode",{parentName:"p"},"string functionName")," - the name of the async function. Used in combination with the ",(0,i.kt)("inlineCode",{parentName:"p"},"parameters")," value to identify this async function by the .NET framework for its internal threading operations."),(0,i.kt)("p",{parentName:"li"},(0,i.kt)("strong",{parentName:"p"},"NOTE:")," Ensure to enclose the function name within the ",(0,i.kt)("inlineCode",{parentName:"p"},"nameof()")," expression.")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},(0,i.kt)("inlineCode",{parentName:"p"},"object parameters")," - the set of parameters the function is being called with. Can be a single object (e.g. a string) or an object","[","]"," array of parameters. It should include all the parameters to the UDF as it is used in combination with the ",(0,i.kt)("inlineCode",{parentName:"p"},"functionName")," value to identify this async function by the .NET framework for its internal threading operations."),(0,i.kt)("p",{parentName:"li"},(0,i.kt)("strong",{parentName:"p"},"NOTE:")," Ensure to include all the parameters to the UDF as it is used in combination with the ",(0,i.kt)("inlineCode",{parentName:"p"},"functionName")," value to identify this async function by the .NET framework for its internal threading operations.  ")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},(0,i.kt)("inlineCode",{parentName:"p"}," ExcelFunc function")," - a delegate function (can be anonymous) that will be executed asynchronously."))),(0,i.kt)("h3",{id:"additional-example-1"},"Additional Example"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"The following example function accepts a target IP/hostname (string) to asynchronously send a ping to.")),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-csharp"},"//Sends an ICMP packet to the target and returns the results asynchronously.\npublic static object PingAsync(string target)\n        {\n            return ExcelAsyncUtil.Run(nameof(PingAsync), new object[] { target }, () => Ping(target));\n        }\n\n//This function's payload will be executed by Excel's asynchronous engine.\nprivate static object Ping(string target)\n{\n    return new Ping().Send(target).Status.ToString();\n}\n")))}p.isMDXComponent=!0}}]);