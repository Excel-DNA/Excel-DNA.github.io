"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[5175],{3905:(e,n,t)=>{t.d(n,{Zo:()=>u,kt:()=>v});var r=t(7294);function a(e,n,t){return n in e?Object.defineProperty(e,n,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[n]=t,e}function i(e,n){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);n&&(r=r.filter((function(n){return Object.getOwnPropertyDescriptor(e,n).enumerable}))),t.push.apply(t,r)}return t}function l(e){for(var n=1;n<arguments.length;n++){var t=null!=arguments[n]?arguments[n]:{};n%2?i(Object(t),!0).forEach((function(n){a(e,n,t[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):i(Object(t)).forEach((function(n){Object.defineProperty(e,n,Object.getOwnPropertyDescriptor(t,n))}))}return e}function s(e,n){if(null==e)return{};var t,r,a=function(e,n){if(null==e)return{};var t,r,a={},i=Object.keys(e);for(r=0;r<i.length;r++)t=i[r],n.indexOf(t)>=0||(a[t]=e[t]);return a}(e,n);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(r=0;r<i.length;r++)t=i[r],n.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(a[t]=e[t])}return a}var o=r.createContext({}),c=function(e){var n=r.useContext(o),t=n;return e&&(t="function"==typeof e?e(n):l(l({},n),e)),t},u=function(e){var n=c(e.components);return r.createElement(o.Provider,{value:n},e.children)},b={inlineCode:"code",wrapper:function(e){var n=e.children;return r.createElement(r.Fragment,{},n)}},p=r.forwardRef((function(e,n){var t=e.components,a=e.mdxType,i=e.originalType,o=e.parentName,u=s(e,["components","mdxType","originalType","parentName"]),p=c(t),v=a,d=p["".concat(o,".").concat(v)]||p[v]||b[v]||i;return t?r.createElement(d,l(l({ref:n},u),{},{components:t})):r.createElement(d,l({ref:n},u))}));function v(e,n){var t=arguments,a=n&&n.mdxType;if("string"==typeof e||a){var i=t.length,l=new Array(i);l[0]=p;var s={};for(var o in n)hasOwnProperty.call(n,o)&&(s[o]=n[o]);s.originalType=e,s.mdxType="string"==typeof e?e:a,l[1]=s;for(var c=2;c<i;c++)l[c]=t[c];return r.createElement.apply(null,l)}return r.createElement.apply(null,t)}p.displayName="MDXCreateElement"},2773:(e,n,t)=>{t.r(n),t.d(n,{assets:()=>o,contentTitle:()=>l,default:()=>b,frontMatter:()=>i,metadata:()=>s,toc:()=>c});var r=t(7462),a=(t(7294),t(3905));const i={title:"Reactive Extensions for Excel"},l=void 0,s={unversionedId:"archive/guides/reactive-extensions-for-excel",id:"archive/guides/reactive-extensions-for-excel",title:"Reactive Extensions for Excel",description:"Excel-DNA has support for integrating the Reactive Extensions library (Rx) with Excel via RTD.",source:"@site/docs/archive/guides/reactive-extensions-for-excel.md",sourceDirName:"archive/guides",slug:"/archive/guides/reactive-extensions-for-excel",permalink:"/docs/archive/guides/reactive-extensions-for-excel",draft:!1,tags:[],version:"current",frontMatter:{title:"Reactive Extensions for Excel"},sidebar:"tutorialSidebar",previous:{title:"Reactive Extensions for Excel - VB.NET",permalink:"/docs/archive/guides/reactive-extensions-for-excel-vbnet"},next:{title:"User Settings and the .xll.config File",permalink:"/docs/archive/guides/user-settings-and-the-xllconfig-file"}},o={},c=[{value:"Additional Examples",id:"additional-examples",level:2}],u={toc:c};function b(e){let{components:n,...t}=e;return(0,a.kt)("wrapper",(0,r.Z)({},u,t,{components:n,mdxType:"MDXLayout"}),(0,a.kt)("p",null,"Excel-DNA has support for integrating the ",(0,a.kt)("a",{parentName:"p",href:"http://msdn.microsoft.com/en-us/data/gg577609.aspx"},"Reactive Extensions")," library (Rx) with Excel via RTD."),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},"You have to call ",(0,a.kt)("inlineCode",{parentName:"li"},"ExcelAsyncUtil.Initialize()")," in your ",(0,a.kt)("inlineCode",{parentName:"li"},"AutoOpen")," for any of the Rx stuff to work.")),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-csharp"},'public class AsyncTestAddIn : IExcelAddIn\n{\n    public void AutoOpen()\n    {\n        // This call is required for the async function and Rx support.\n        ExcelAsyncUtil.Initialize();\n\n        // This is optional - allows a custom return value for Exceptions\n        // By default exceptions just return #VALUE\n        ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! EXCEPTION: " + ex.ToString());\n    }\n\n    public void AutoClose()\n    {\n    }\n}\n')),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},"I created an adapter class \u2013 called ",(0,a.kt)("inlineCode",{parentName:"li"},"RxExcel")," but I see it is still in the ",(0,a.kt)("inlineCode",{parentName:"li"},"RxAdapter.cs")," file - to map the .NET 4 Rx types to the Excel-DNA fake types. The idea would be that any add-in doing Rx stuff would just include the RxExcel file.")),(0,a.kt)("p",null,"I need this because I still want to target .NET 2.0 with Excel-DNA, so the ",(0,a.kt)("inlineCode",{parentName:"p"},"System.IObservable")," is not available."),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-csharp"},"using System;\n\nnamespace ExcelDna.Integration.RxExcel\n{\n    public static class RxExcel\n    {\n        public static IExcelObservable ToExcelObservable<T>(this IObservable<T> observable)\n        {\n            return new ExcelObservable<T>(observable);\n        }\n\n        public static object Observe<T>(string functionName, object parameters, Func<IObservable<T>> observableSource)\n        {\n            return ExcelAsyncUtil.Observe(functionName, parameters, () => observableSource().ToExcelObservable());\n        }\n    }\n\n    public class ExcelObservable<T> : IExcelObservable\n    {\n        readonly IObservable<T> _observable;\n\n        public ExcelObservable(IObservable<T> observable)\n        {\n            _observable = observable;\n        }\n\n        public IDisposable Subscribe(IExcelObserver observer)\n        {\n            return _observable.Subscribe(value => observer.OnNext(value), observer.OnError, observer.OnCompleted);\n        }\n    }\n}\n")),(0,a.kt)("p",null,"(VB.NET version here:  ",(0,a.kt)("a",{parentName:"p",href:"reactive-extensions-for-excel-vbnet"},"Reactive Extensions for Excel - VB.NET"),")."),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},"UDFs that hook up Observables then look like this:")),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-csharp"},'// Publishes a single value after the interval elapses.\npublic static object rxTimerWaitInterval(int intervalSeconds)\n{\n    return RxExcel.Observe("rxTimerWaitInterval", intervalSeconds, () =>\n        Observable.Timer(TimeSpan.FromSeconds(intervalSeconds)));\n}\n')),(0,a.kt)("p",null,"The main entry point is then the RxExcel.Observe function, which takes the function name, an object or array of objects representing the \u2018parameters\u2019, and then a delegate that will return the IObservable.\nThe combination of function name and parameters is used to identify the topic. You could have different cells (callers) with their own Observables even though they are calling the same function with the same parameters by adding the caller - XlCall.Excel(XlCall.xlfCaller) - to the array of parameters."),(0,a.kt)("h2",{id:"additional-examples"},"Additional Examples"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-csharp"},'using System;\nusing System.Collections.Generic;\nusing System.Linq;\nusing System.Reactive.Linq;\nusing ExcelDna.Integration;\nusing ExcelDna.Integration.RxExcel;\n\nnamespace AsyncFunctions\n{\n    public class RxTest\n    {\n        // Just returns a single value and completes the sequence.\n        public static object rxReturn(object value)\n        {\n            return RxExcel.Observe("rxReturn", value, () =>\n              Observable.Return(value));\n        }\n\n        // We don\'t currently distinguish between Empty and Never.\n        // Empty is a sequence that immediately completes without pushing a value.\n        // So we return #N/A (the pre-Value \'Not Available\' return state),\n        // and then never have anything else to return when the sequence completes.\n        // CONSIDER: Should we rather transition to an empty string if we comlete without seeing a value?\n        public static object rxEmpty()\n        {\n            return RxExcel.Observe("rxEmpty", null, () =>\n                Observable.Empty<string>());\n        }\n\n        // Never just doesn\'t return anything, so our functions stays in the #N/A pre-value return state.\n        // This seems fine.\n        public static object rxNever()\n        {\n            return RxExcel.Observe("rxNever", null, () =>\n                Observable.Never<string>());\n        }\n\n        // By default, all exceptions are just returned as #VALUE, consistent with the rest of Excel-DNA.\n        // If an UnhandledExceptionHandler is registered via Integration.RegisterUnhandledExceptionHandler,\n        // then the result of that handler will be returned by this function.\n        public static object rxThrow()\n        {\n            return RxExcel.Observe("rxThrow", null, () =>\n                Observable.Throw<string>(new Exception()));\n        }\n\n        // Note that the System.Timers.Timer used here will raise it\'s Elapsed events from a ThreadPool thread.\n        // This is fine - the RxExcel RTD server does all the cross-thread marshaling.\n        public static object rxCreateTimer(int intervalSeconds)\n        {\n            return RxExcel.Observe("rxCreateTimer", intervalSeconds, () =>\n                Observable.Create<string>(observer =>\n                {\n                    var timer = new System.Timers.Timer();\n                    timer.Interval = intervalSeconds * 1000;\n                    timer.Elapsed += (s, e) => observer.OnNext("Tick at" + DateTime.Now.ToString("HH:mm:ss.fff"));\n                    timer.Start();\n                    return timer;\n                }));\n        }\n\n        // Excel will not update for every value in the sequence - just as often as the ThrottleInreval allows.\n        // Observable.Interval might generate many values we ignore.\n        public static object rxInterval(int intervalSeconds)\n        {\n            return RxExcel.Observe("rxInterval", intervalSeconds, () =>\n                Observable.Interval(TimeSpan.FromSeconds(intervalSeconds)));\n        }\n\n        // Publishes a single value after the interval elapses.\n        public static object rxTimerWaitInterval(int intervalSeconds)\n        {\n            return RxExcel.Observe("rxTimerWaitInterval", intervalSeconds, () =>\n                Observable.Timer(TimeSpan.FromSeconds(intervalSeconds)));\n        }\n\n        // Publishes a single value at the given time.\n        public static object rxTimerWaitUntil(DateTime timeUntil)\n        {\n            return RxExcel.Observe("rxTimerWaitUntil", timeUntil, () =>\n                Observable.Timer(timeUntil));\n        }\n\n        // A custom sequence returning squares every 5 seconds, up to 20 * 20.\n        // Not Observing \'Per Caller\' ensures we share a sequnce if using the function in different cells\n        public static object rxCreateValues()\n        {\n            return RxExcel.Observe("rxCreateValuesShared", null, () =>\n                Observable.Generate(\n                    1,\n                    i => i <= 20,\n                    i => i + 1,\n                    i => i * i,\n                    i => TimeSpan.FromSeconds(5)));\n        }\n\n        // A custom sequence returning squares every intervalSeconds seconds, up to 10 * 10.\n        // Observe \'Per Caller\' by sending the caller is one of the \'parameters\' into RxExcel.Observe.\n        // This ensures we get different sequences if using the function in different cells\n        public static object rxCreateValuesPerCaller(int intervalSeconds)\n        {\n            object caller = XlCall.Excel(XlCall.xlfCaller);\n\n            return RxExcel.Observe("rxCreateValues", new[] {intervalSeconds, caller}, () =>\n                Observable.Generate(\n                    1,\n                    i => i <= 10,\n                    i => i + 1,\n                    i => i * i,\n                    i => TimeSpan.FromSeconds(5)));\n        }\n\n        // Some experiments returning arrays ... not really useful yet.\n        public static object rxCreateArrays()\n        {\n            return RxExcel.Observe("rxCreateArrays", null, () =>\n                Observable.Generate(\n                    new List<object> {1,2,3},\n                    lst => true,\n                    lst => { lst.Add((int)lst[lst.Count-1] + 1); return lst;},\n                    lst => Transpose(lst.ToArray()),\n                    lst => TimeSpan.FromSeconds(2)));\n        }\n\n        static object[,] Transpose(object[] array)\n        {\n            object[,] result = new object[array.Length, 1];\n            for (int i = 0; i < array.Length; i++)\n            {\n                result[i,0] = array[i];\n            }\n\n            return result;\n        }\n    }\n}\n')))}b.isMDXComponent=!0}}]);