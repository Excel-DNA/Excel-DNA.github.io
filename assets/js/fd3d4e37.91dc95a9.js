"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[3254],{3905:(e,n,t)=>{t.d(n,{Zo:()=>c,kt:()=>h});var o=t(7294);function i(e,n,t){return n in e?Object.defineProperty(e,n,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[n]=t,e}function a(e,n){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);n&&(o=o.filter((function(n){return Object.getOwnPropertyDescriptor(e,n).enumerable}))),t.push.apply(t,o)}return t}function l(e){for(var n=1;n<arguments.length;n++){var t=null!=arguments[n]?arguments[n]:{};n%2?a(Object(t),!0).forEach((function(n){i(e,n,t[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):a(Object(t)).forEach((function(n){Object.defineProperty(e,n,Object.getOwnPropertyDescriptor(t,n))}))}return e}function r(e,n){if(null==e)return{};var t,o,i=function(e,n){if(null==e)return{};var t,o,i={},a=Object.keys(e);for(o=0;o<a.length;o++)t=a[o],n.indexOf(t)>=0||(i[t]=e[t]);return i}(e,n);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(o=0;o<a.length;o++)t=a[o],n.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(i[t]=e[t])}return i}var d=o.createContext({}),s=function(e){var n=o.useContext(d),t=n;return e&&(t="function"==typeof e?e(n):l(l({},n),e)),t},c=function(e){var n=s(e.components);return o.createElement(d.Provider,{value:n},e.children)},u={inlineCode:"code",wrapper:function(e){var n=e.children;return o.createElement(o.Fragment,{},n)}},p=o.forwardRef((function(e,n){var t=e.components,i=e.mdxType,a=e.originalType,d=e.parentName,c=r(e,["components","mdxType","originalType","parentName"]),p=s(t),h=i,m=p["".concat(d,".").concat(h)]||p[h]||u[h]||a;return t?o.createElement(m,l(l({ref:n},c),{},{components:t})):o.createElement(m,l({ref:n},c))}));function h(e,n){var t=arguments,i=n&&n.mdxType;if("string"==typeof e||i){var a=t.length,l=new Array(a);l[0]=p;var r={};for(var d in n)hasOwnProperty.call(n,d)&&(r[d]=n[d]);r.originalType=e,r.mdxType="string"==typeof e?e:i,l[1]=r;for(var s=2;s<a;s++)l[s]=t[s];return o.createElement.apply(null,l)}return o.createElement.apply(null,t)}p.displayName="MDXCreateElement"},6754:(e,n,t)=>{t.r(n),t.d(n,{assets:()=>d,contentTitle:()=>l,default:()=>u,frontMatter:()=>a,metadata:()=>r,toc:()=>s});var o=t(7462),i=(t(7294),t(3905));const a={title:"Detecting Excel Shutdown and AutoClose"},l=void 0,r={unversionedId:"archive/guides/detecting-excel-shutdown-and-autoclose",id:"archive/guides/detecting-excel-shutdown-and-autoclose",title:"Detecting Excel Shutdown and AutoClose",description:"This is a short note on the IExcelAddIn.AutoClose() callback, noting that it is not called when Excel is shut down, and explaining how the implementation came about.",source:"@site/docs/archive/guides/detecting-excel-shutdown-and-autoclose.md",sourceDirName:"archive/guides",slug:"/archive/guides/detecting-excel-shutdown-and-autoclose",permalink:"/Website/docs/archive/guides/detecting-excel-shutdown-and-autoclose",draft:!1,tags:[],version:"current",frontMatter:{title:"Detecting Excel Shutdown and AutoClose"},sidebar:"tutorialSidebar",previous:{title:"Debugging Addins and Excel-DNA",permalink:"/Website/docs/archive/guides/debugging-addins-and-exceldna"},next:{title:"Dynamic Delegate Registration",permalink:"/Website/docs/archive/guides/dynamic-delegate-registration"}},d={},s=[{value:"Background",id:"background",level:2},{value:"Example Add-In",id:"example-add-in",level:2}],c={toc:s};function u(e){let{components:n,...t}=e;return(0,i.kt)("wrapper",(0,o.Z)({},c,t,{components:n,mdxType:"MDXLayout"}),(0,i.kt)("p",null,"This is a short note on the ",(0,i.kt)("inlineCode",{parentName:"p"},"IExcelAddIn.AutoClose()")," callback, noting that it is not called when Excel is shut down, and explaining how the implementation came about."),(0,i.kt)("p",null,"Excel-DNA will call the ",(0,i.kt)("inlineCode",{parentName:"p"},"IExcelAddIn.AutoClose()")," method when the add-in is removed from the add-ins dialog (",(0,i.kt)("inlineCode",{parentName:"p"},"Alt+t"),", ",(0,i.kt)("inlineCode",{parentName:"p"},"i"),") by the user, or if the add-in is reloaded. In this case you can properly clean up your add-in - remove menus etc. Mostly when Excel shuts down you would not want to do a lot of clean-up - no need to remove menus, deregister functions etc."),(0,i.kt)("p",null,"If you need to be notified of the Excel shutdown:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"If you are running in Excel 2007+ and have an ",(0,i.kt)("inlineCode",{parentName:"li"},"ExcelRibbon"),"-derived\nclass, just override the ",(0,i.kt)("inlineCode",{parentName:"li"},"OnDisconnection")," or ",(0,i.kt)("inlineCode",{parentName:"li"},"OnBeginShutdown"),"."),(0,i.kt)("li",{parentName:"ul"},"To target any Excel version, add a new class that derives from ",(0,i.kt)("inlineCode",{parentName:"li"},"ExcelComAddIn"),", load it in your AutoOpen with ",(0,i.kt)("inlineCode",{parentName:"li"},"ExcelComAddInHelper.LoadComAddIn(...)"),", and override the ",(0,i.kt)("inlineCode",{parentName:"li"},"OnDisconnection")," or ",(0,i.kt)("inlineCode",{parentName:"li"},"OnBeginShutdown"),".")),(0,i.kt)("h2",{id:"background"},"Background"),(0,i.kt)("p",null,"Excel .xll add-in export a few functions that are relevant to the discussion:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"xlAutoOpen")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"xlAutoClose")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"xlAutoAdd")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"xlAutoRemove"))),(0,i.kt)("p",null,"When an add-in is opened, ",(0,i.kt)("inlineCode",{parentName:"p"},"xlAutoOpen")," is called, and Excel-DNA passes that trough to the ",(0,i.kt)("inlineCode",{parentName:"p"},"IExcelAddIn.AutoOpen()"),"."),(0,i.kt)("p",null,"Excel calls ",(0,i.kt)("inlineCode",{parentName:"p"},"xlAutoRemove")," when the add-in is removed from the Add-Ins dialog (thus if the user has explicitly chosen to remove the add-in from a running session). The problem is with ",(0,i.kt)("inlineCode",{parentName:"p"},"xlAutoClose"),". If you have an add-in loaded in Excel with some workbook open and 'dirty', and then press Alt+F4, Excel will call the ",(0,i.kt)("inlineCode",{parentName:"p"},"xlAutoClose"),", and then display a dialog to the user to ask whether to 'Save', 'Don't Save' or 'Cancel'. If the user selects 'Cancel' the session will continue. However, if the add-in has responded to the earlier ",(0,i.kt)("inlineCode",{parentName:"p"},"xlAutoClose"),", it might now be removed although the session still continues, causing functions to fail and the add-in's ribbon to be missing. I didn't like this, so in Excel-DNA I only call ",(0,i.kt)("inlineCode",{parentName:"p"},"ExcelAddIn.AutoClose")," when I have received an ",(0,i.kt)("inlineCode",{parentName:"p"},"xlAutoRemove")," before the ",(0,i.kt)("inlineCode",{parentName:"p"},"xlAutoClose"),"."),(0,i.kt)("p",null,"The resulting behaviour with Excel-DNA is that your add-in's AutoClose is only called when the add-in is actually removed by the user, and not when Excel exits. This allows ",(0,i.kt)("inlineCode",{parentName:"p"},"AutoClose")," to do clean-up work that should be done when an add-in is removed. When Excel is shutting down, the whole process will be shut down, so your add-in should probably not do any clean-up. The operating system will close all handles, and recover all memory. Doing clean-up at this stage will just delay the closing of the Excel process. So I'm happy that this is a reasonable approach."),(0,i.kt)("p",null,"In some cases the add-in might like to be notifies and do additional work when Excel is shutting down. Clearly ",(0,i.kt)("inlineCode",{parentName:"p"},"xlAutoClose")," is not the right place for this, so the Excel C API does not give us an obvious way to implement such behaviour. We need some other mechanism to get a proper notification from Excel, and I suggest using the COM add-in approach, which works well. The COM add-in support in Excel-DNA which allows this is a much more recent addition, only implemented when I added support for the Excel 2007 ribbons, so was not an option when I initially decided what to do with ",(0,i.kt)("inlineCode",{parentName:"p"},"AutoClose()"),"."),(0,i.kt)("p",null,"I have not added the COM add-in and it's shutdown event handler as a standard part of the add-in implementation, so that minimal Excel add-ins exposing only UDFs have no dependency on the COM support and so can claim to be 'pure .xll add-ins' using only the supported C API documented in the Excel SDK. In a sense, doing any COM stuff from the Excel-DNA add-in is making a hydrid with some hacks behind the scenes, and I think it is important to keep the COM part optional."),(0,i.kt)("p",null,"Other events on the Excel Application object or the Workbook object might also be useful and hooked up from Excel-DNA, but there is no special support for these, apart from the ",(0,i.kt)("inlineCode",{parentName:"p"},"ExcelDnaUtil.Application")," call which must be used to get hold of the correct Application root object."),(0,i.kt)("h2",{id:"example-add-in"},"Example Add-In"),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-xml"},'<DnaLibrary RuntimeVersion="v4.0" Language="C#">\n<Reference Name="System.Windows.Forms" />\n<![CDATA[\nusing System;\nusing System.Reflection;\nusing System.Runtime.InteropServices;\nusing SWF = System.Windows.Forms;\nusing ExcelDna.Integration;\nusing ExcelDna.Integration.Extensibility;\nusing ExcelDna.Integration.CustomUI;\n\n    [ComVisible(true)](ComVisible(true))\n    public class MyComAddIn : ExcelComAddIn\n    {\n        public MyComAddIn()\n        {\n        }\n\n        public override void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)\n        {\n            SWF.MessageBox.Show("OnConnection");\n        }\n\n        public override void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)\n        {\n            SWF.MessageBox.Show("OnDisconnection");\n        }\n\n        public override void OnAddInsUpdate(ref Array custom)\n        {\n            SWF.MessageBox.Show("OnAddInsUpdate");\n        }\n\n        public override void OnStartupComplete(ref Array custom)\n        {\n            SWF.MessageBox.Show("OnStartupComplete");\n        }\n\n        public override void OnBeginShutdown(ref Array custom)\n        {\n            SWF.MessageBox.Show("OnBeginShutDown");\n        }\n    }\n\n    public class MyAddIn : IExcelAddIn\n    {\n        ExcelDna.Integration.CustomUI.ExcelComAddIn _comAddIn;\n\n        public void AutoOpen()\n        {\n            try\n            {\n                _comAddIn = new MyComAddIn();\n                ExcelComAddInHelper.LoadComAddIn(_comAddIn);\n            }\n            catch (Exception e)\n            {\n                SWF.MessageBox.Show("Error loading COM AddIn: " + e.ToString());\n            }\n        }\n\n        public void AutoClose()\n        {\n        }\n    }\n]]>\n</DnaLibrary>\n')))}u.isMDXComponent=!0}}]);