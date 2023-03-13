"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[680],{3905:(e,n,t)=>{t.d(n,{Zo:()=>u,kt:()=>d});var r=t(7294);function a(e,n,t){return n in e?Object.defineProperty(e,n,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[n]=t,e}function l(e,n){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);n&&(r=r.filter((function(n){return Object.getOwnPropertyDescriptor(e,n).enumerable}))),t.push.apply(t,r)}return t}function s(e){for(var n=1;n<arguments.length;n++){var t=null!=arguments[n]?arguments[n]:{};n%2?l(Object(t),!0).forEach((function(n){a(e,n,t[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):l(Object(t)).forEach((function(n){Object.defineProperty(e,n,Object.getOwnPropertyDescriptor(t,n))}))}return e}function i(e,n){if(null==e)return{};var t,r,a=function(e,n){if(null==e)return{};var t,r,a={},l=Object.keys(e);for(r=0;r<l.length;r++)t=l[r],n.indexOf(t)>=0||(a[t]=e[t]);return a}(e,n);if(Object.getOwnPropertySymbols){var l=Object.getOwnPropertySymbols(e);for(r=0;r<l.length;r++)t=l[r],n.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(a[t]=e[t])}return a}var o=r.createContext({}),c=function(e){var n=r.useContext(o),t=n;return e&&(t="function"==typeof e?e(n):s(s({},n),e)),t},u=function(e){var n=c(e.components);return r.createElement(o.Provider,{value:n},e.children)},b={inlineCode:"code",wrapper:function(e){var n=e.children;return r.createElement(r.Fragment,{},n)}},v=r.forwardRef((function(e,n){var t=e.components,a=e.mdxType,l=e.originalType,o=e.parentName,u=i(e,["components","mdxType","originalType","parentName"]),v=c(t),d=a,p=v["".concat(o,".").concat(d)]||v[d]||b[d]||l;return t?r.createElement(p,s(s({ref:n},u),{},{components:t})):r.createElement(p,s({ref:n},u))}));function d(e,n){var t=arguments,a=n&&n.mdxType;if("string"==typeof e||a){var l=t.length,s=new Array(l);s[0]=v;var i={};for(var o in n)hasOwnProperty.call(n,o)&&(i[o]=n[o]);i.originalType=e,i.mdxType="string"==typeof e?e:a,s[1]=i;for(var c=2;c<l;c++)s[c]=t[c];return r.createElement.apply(null,s)}return r.createElement.apply(null,t)}v.displayName="MDXCreateElement"},5162:(e,n,t)=>{t.d(n,{Z:()=>s});var r=t(7294),a=t(6010);const l="tabItem_Ymn6";function s(e){let{children:n,hidden:t,className:s}=e;return r.createElement("div",{role:"tabpanel",className:(0,a.Z)(l,s),hidden:t},n)}},4866:(e,n,t)=>{t.d(n,{Z:()=>O});var r=t(7462),a=t(7294),l=t(6010),s=t(2466),i=t(6550),o=t(1980),c=t(7392),u=t(12);function b(e){return function(e){return a.Children.map(e,(e=>{if((0,a.isValidElement)(e)&&"value"in e.props)return e;throw new Error(`Docusaurus error: Bad <Tabs> child <${"string"==typeof e.type?e.type:e.type.name}>: all children of the <Tabs> component should be <TabItem>, and every <TabItem> should have a unique "value" prop.`)}))}(e).map((e=>{let{props:{value:n,label:t,attributes:r,default:a}}=e;return{value:n,label:t,attributes:r,default:a}}))}function v(e){const{values:n,children:t}=e;return(0,a.useMemo)((()=>{const e=n??b(t);return function(e){const n=(0,c.l)(e,((e,n)=>e.value===n.value));if(n.length>0)throw new Error(`Docusaurus error: Duplicate values "${n.map((e=>e.value)).join(", ")}" found in <Tabs>. Every value needs to be unique.`)}(e),e}),[n,t])}function d(e){let{value:n,tabValues:t}=e;return t.some((e=>e.value===n))}function p(e){let{queryString:n=!1,groupId:t}=e;const r=(0,i.k6)(),l=function(e){let{queryString:n=!1,groupId:t}=e;if("string"==typeof n)return n;if(!1===n)return null;if(!0===n&&!t)throw new Error('Docusaurus error: The <Tabs> component groupId prop is required if queryString=true, because this value is used as the search param name. You can also provide an explicit value such as queryString="my-search-param".');return t??null}({queryString:n,groupId:t});return[(0,o._X)(l),(0,a.useCallback)((e=>{if(!l)return;const n=new URLSearchParams(r.location.search);n.set(l,e),r.replace({...r.location,search:n.toString()})}),[l,r])]}function m(e){const{defaultValue:n,queryString:t=!1,groupId:r}=e,l=v(e),[s,i]=(0,a.useState)((()=>function(e){let{defaultValue:n,tabValues:t}=e;if(0===t.length)throw new Error("Docusaurus error: the <Tabs> component requires at least one <TabItem> children component");if(n){if(!d({value:n,tabValues:t}))throw new Error(`Docusaurus error: The <Tabs> has a defaultValue "${n}" but none of its children has the corresponding value. Available values are: ${t.map((e=>e.value)).join(", ")}. If you intend to show no default tab, use defaultValue={null} instead.`);return n}const r=t.find((e=>e.default))??t[0];if(!r)throw new Error("Unexpected error: 0 tabValues");return r.value}({defaultValue:n,tabValues:l}))),[o,c]=p({queryString:t,groupId:r}),[b,m]=function(e){let{groupId:n}=e;const t=function(e){return e?`docusaurus.tab.${e}`:null}(n),[r,l]=(0,u.Nk)(t);return[r,(0,a.useCallback)((e=>{t&&l.set(e)}),[t,l])]}({groupId:r}),f=(()=>{const e=o??b;return d({value:e,tabValues:l})?e:null})();(0,a.useLayoutEffect)((()=>{f&&i(f)}),[f]);return{selectedValue:s,selectValue:(0,a.useCallback)((e=>{if(!d({value:e,tabValues:l}))throw new Error(`Can't select invalid tab value=${e}`);i(e),c(e),m(e)}),[c,m,l]),tabValues:l}}var f=t(2389);const h="tabList__CuJ",x="tabItem_LNqP";function g(e){let{className:n,block:t,selectedValue:i,selectValue:o,tabValues:c}=e;const u=[],{blockElementScrollPositionUntilNextRender:b}=(0,s.o5)(),v=e=>{const n=e.currentTarget,t=u.indexOf(n),r=c[t].value;r!==i&&(b(n),o(r))},d=e=>{var n;let t=null;switch(e.key){case"Enter":v(e);break;case"ArrowRight":{const n=u.indexOf(e.currentTarget)+1;t=u[n]??u[0];break}case"ArrowLeft":{const n=u.indexOf(e.currentTarget)-1;t=u[n]??u[u.length-1];break}}null==(n=t)||n.focus()};return a.createElement("ul",{role:"tablist","aria-orientation":"horizontal",className:(0,l.Z)("tabs",{"tabs--block":t},n)},c.map((e=>{let{value:n,label:t,attributes:s}=e;return a.createElement("li",(0,r.Z)({role:"tab",tabIndex:i===n?0:-1,"aria-selected":i===n,key:n,ref:e=>u.push(e),onKeyDown:d,onClick:v},s,{className:(0,l.Z)("tabs__item",x,null==s?void 0:s.className,{"tabs__item--active":i===n})}),t??n)})))}function y(e){let{lazy:n,children:t,selectedValue:r}=e;if(t=Array.isArray(t)?t:[t],n){const e=t.find((e=>e.props.value===r));return e?(0,a.cloneElement)(e,{className:"margin-top--md"}):null}return a.createElement("div",{className:"margin-top--md"},t.map(((e,n)=>(0,a.cloneElement)(e,{key:n,hidden:e.props.value!==r}))))}function E(e){const n=m(e);return a.createElement("div",{className:(0,l.Z)("tabs-container",h)},a.createElement(g,(0,r.Z)({},e,n)),a.createElement(y,(0,r.Z)({},e,n)))}function O(e){const n=(0,f.Z)();return a.createElement(E,(0,r.Z)({key:String(n)},e))}},9139:(e,n,t)=>{t.r(n),t.d(n,{assets:()=>u,contentTitle:()=>o,default:()=>d,frontMatter:()=>i,metadata:()=>c,toc:()=>b});var r=t(7462),a=(t(7294),t(3905)),l=t(4866),s=t(5162);const i={title:"Reactive Extensions for Excel"},o=void 0,c={unversionedId:"guides-advanced/reactive-extensions-for-excel",id:"guides-advanced/reactive-extensions-for-excel",title:"Reactive Extensions for Excel",description:"Excel-DNA has support for integrating the Reactive Extensions library (Rx) with Excel via RTD.",source:"@site/docs/guides-advanced/reactive-extensions-for-excel.md",sourceDirName:"guides-advanced",slug:"/guides-advanced/reactive-extensions-for-excel",permalink:"/docs/guides-advanced/reactive-extensions-for-excel",draft:!1,tags:[],version:"current",frontMatter:{title:"Reactive Extensions for Excel"},sidebar:"tutorialSidebar",previous:{title:"Performing Asynchronous Work",permalink:"/docs/guides-advanced/performing-asynchronous-work"},next:{title:"User Settings and the .xll.config File",permalink:"/docs/guides-advanced/user-settings-and-the-xllconfig-file"}},u={},b=[{value:"Additional Examples",id:"additional-examples",level:2}],v={toc:b};function d(e){let{components:n,...t}=e;return(0,a.kt)("wrapper",(0,r.Z)({},v,t,{components:n,mdxType:"MDXLayout"}),(0,a.kt)("p",null,"Excel-DNA has support for integrating the ",(0,a.kt)("a",{parentName:"p",href:"https://github.com/dotnet/reactive"},"Reactive Extensions")," library (Rx) with Excel via RTD."),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},"To map the .NET Rx types to the Excel-DNA RTD-based mechanism, it is possible to use the following code:")),(0,a.kt)(l.Z,{mdxType:"Tabs"},(0,a.kt)(s.Z,{value:"csharp",label:"C#",mdxType:"TabItem"},(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-csharp"},"using System;\n\nnamespace ExcelDna.Integration.RxExcel\n{\n    public static class RxExcel\n    {\n        public static IExcelObservable ToExcelObservable<T>(this IObservable<T> observable)\n        {\n            return new ExcelObservable<T>(observable);\n        }\n\n        public static object Observe<T>(string functionName, object parameters, Func<IObservable<T>> observableSource)\n        {\n            return ExcelAsyncUtil.Observe(functionName, parameters, () => observableSource().ToExcelObservable());\n        }\n    }\n\n    public class ExcelObservable<T> : IExcelObservable\n    {\n        readonly IObservable<T> _observable;\n\n        public ExcelObservable(IObservable<T> observable)\n        {\n            _observable = observable;\n        }\n\n        public IDisposable Subscribe(IExcelObserver observer)\n        {\n            return _observable.Subscribe(value => observer.OnNext(value), observer.OnError, observer.OnCompleted);\n        }\n    }\n}\n"))),(0,a.kt)(s.Z,{value:"vbnet",label:"VB.Net",mdxType:"TabItem"},(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vbnet"},"Imports System.Runtime.CompilerServices\nImports ExcelDna.Integration\n\nPublic Module RxExcel\n\n    <Extension()>\n    Public Function ToExcelObservable(Of T)(observable As IObservable(Of T)) As IExcelObservable\n        Return New ExcelObservable(Of T)(observable)\n    End Function\n\n    Public Function Observe(Of T)(functionName As String, parameters As Object, _\n                           observableSource As Func(Of IObservable(Of T))) As Object\n        Return ExcelAsyncUtil.Observe(functionName, parameters,\n                                     Function() observableSource().ToExcelObservable())\n    End Function\nEnd Module\n\nPublic Class ExcelObservable(Of T)\n    Implements IExcelObservable\n\n    ReadOnly _observable As IObservable(Of T)\n\n    Public Sub New(observable As IObservable(Of T))\n        _observable = observable\n    End Sub\n\n    Public Function Subscribe(observer As IExcelObserver) As IDisposable _\n        Implements IExcelObservable.Subscribe\n        Return _observable.Subscribe(Sub(value) observer.OnNext(value),\n            Sub(ex) observer.OnError(ex), Sub() observer.OnCompleted())\n    End Function\nEnd Class\n")))),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},"UDFs that hook up Observables then look like this:")),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-csharp"},'// Publishes a single value after the interval elapses.\npublic static object rxTimerWaitInterval(int intervalSeconds)\n{\n    return RxExcel.Observe("rxTimerWaitInterval", intervalSeconds, () =>\n        Observable.Timer(TimeSpan.FromSeconds(intervalSeconds)));\n}\n')),(0,a.kt)("p",null,"The main entry point is then the RxExcel.Observe function, which takes the function name, an object or array of objects representing the \u2018parameters\u2019, and then a delegate that will return the IObservable.\nThe combination of function name and parameters is used to identify the topic. You could have different cells (callers) with their own Observables even though they are calling the same function with the same parameters by adding the caller - XlCall.Excel(XlCall.xlfCaller) - to the array of parameters."),(0,a.kt)("h2",{id:"additional-examples"},"Additional Examples"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-csharp"},'using System;\nusing System.Collections.Generic;\nusing System.Linq;\nusing System.Reactive.Linq;\nusing ExcelDna.Integration;\nusing ExcelDna.Integration.RxExcel;\n\nnamespace AsyncFunctions\n{\n    public class RxTest\n    {\n        // Just returns a single value and completes the sequence.\n        public static object rxReturn(object value)\n        {\n            return RxExcel.Observe("rxReturn", value, () =>\n              Observable.Return(value));\n        }\n\n        // We don\'t currently distinguish between Empty and Never.\n        // Empty is a sequence that immediately completes without pushing a value.\n        // So we return #N/A (the pre-Value \'Not Available\' return state),\n        // and then never have anything else to return when the sequence completes.\n        // CONSIDER: Should we rather transition to an empty string if we comlete without seeing a value?\n        public static object rxEmpty()\n        {\n            return RxExcel.Observe("rxEmpty", null, () =>\n                Observable.Empty<string>());\n        }\n\n        // Never just doesn\'t return anything, so our functions stays in the #N/A pre-value return state.\n        // This seems fine.\n        public static object rxNever()\n        {\n            return RxExcel.Observe("rxNever", null, () =>\n                Observable.Never<string>());\n        }\n\n        // By default, all exceptions are just returned as #VALUE, consistent with the rest of Excel-DNA.\n        // If an UnhandledExceptionHandler is registered via Integration.RegisterUnhandledExceptionHandler,\n        // then the result of that handler will be returned by this function.\n        public static object rxThrow()\n        {\n            return RxExcel.Observe("rxThrow", null, () =>\n                Observable.Throw<string>(new Exception()));\n        }\n\n        // Note that the System.Timers.Timer used here will raise it\'s Elapsed events from a ThreadPool thread.\n        // This is fine - the RxExcel RTD server does all the cross-thread marshaling.\n        public static object rxCreateTimer(int intervalSeconds)\n        {\n            return RxExcel.Observe("rxCreateTimer", intervalSeconds, () =>\n                Observable.Create<string>(observer =>\n                {\n                    var timer = new System.Timers.Timer();\n                    timer.Interval = intervalSeconds * 1000;\n                    timer.Elapsed += (s, e) => observer.OnNext("Tick at" + DateTime.Now.ToString("HH:mm:ss.fff"));\n                    timer.Start();\n                    return timer;\n                }));\n        }\n\n        // Excel will not update for every value in the sequence - just as often as the ThrottleInreval allows.\n        // Observable.Interval might generate many values we ignore.\n        public static object rxInterval(int intervalSeconds)\n        {\n            return RxExcel.Observe("rxInterval", intervalSeconds, () =>\n                Observable.Interval(TimeSpan.FromSeconds(intervalSeconds)));\n        }\n\n        // Publishes a single value after the interval elapses.\n        public static object rxTimerWaitInterval(int intervalSeconds)\n        {\n            return RxExcel.Observe("rxTimerWaitInterval", intervalSeconds, () =>\n                Observable.Timer(TimeSpan.FromSeconds(intervalSeconds)));\n        }\n\n        // Publishes a single value at the given time.\n        public static object rxTimerWaitUntil(DateTime timeUntil)\n        {\n            return RxExcel.Observe("rxTimerWaitUntil", timeUntil, () =>\n                Observable.Timer(timeUntil));\n        }\n\n        // A custom sequence returning squares every 5 seconds, up to 20 * 20.\n        // Not Observing \'Per Caller\' ensures we share a sequnce if using the function in different cells\n        public static object rxCreateValues()\n        {\n            return RxExcel.Observe("rxCreateValuesShared", null, () =>\n                Observable.Generate(\n                    1,\n                    i => i <= 20,\n                    i => i + 1,\n                    i => i * i,\n                    i => TimeSpan.FromSeconds(5)));\n        }\n\n        // A custom sequence returning squares every intervalSeconds seconds, up to 10 * 10.\n        // Observe \'Per Caller\' by sending the caller is one of the \'parameters\' into RxExcel.Observe.\n        // This ensures we get different sequences if using the function in different cells\n        public static object rxCreateValuesPerCaller(int intervalSeconds)\n        {\n            object caller = XlCall.Excel(XlCall.xlfCaller);\n\n            return RxExcel.Observe("rxCreateValues", new[] {intervalSeconds, caller}, () =>\n                Observable.Generate(\n                    1,\n                    i => i <= 10,\n                    i => i + 1,\n                    i => i * i,\n                    i => TimeSpan.FromSeconds(5)));\n        }\n\n        // Some experiments returning arrays ... not really useful yet.\n        public static object rxCreateArrays()\n        {\n            return RxExcel.Observe("rxCreateArrays", null, () =>\n                Observable.Generate(\n                    new List<object> {1,2,3},\n                    lst => true,\n                    lst => { lst.Add((int)lst[lst.Count-1] + 1); return lst;},\n                    lst => Transpose(lst.ToArray()),\n                    lst => TimeSpan.FromSeconds(2)));\n        }\n\n        static object[,] Transpose(object[] array)\n        {\n            object[,] result = new object[array.Length, 1];\n            for (int i = 0; i < array.Length; i++)\n            {\n                result[i,0] = array[i];\n            }\n\n            return result;\n        }\n    }\n}\n')))}d.isMDXComponent=!0}}]);