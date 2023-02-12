"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[5367],{3905:(e,t,n)=>{n.d(t,{Zo:()=>u,kt:()=>p});var r=n(7294);function i(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function a(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);t&&(r=r.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,r)}return n}function o(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?a(Object(n),!0).forEach((function(t){i(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):a(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,r,i=function(e,t){if(null==e)return{};var n,r,i={},a=Object.keys(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||(i[n]=e[n]);return i}(e,t);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(i[n]=e[n])}return i}var l=r.createContext({}),c=function(e){var t=r.useContext(l),n=t;return e&&(n="function"==typeof e?e(t):o(o({},t),e)),n},u=function(e){var t=c(e.components);return r.createElement(l.Provider,{value:t},e.children)},d={inlineCode:"code",wrapper:function(e){var t=e.children;return r.createElement(r.Fragment,{},t)}},m=r.forwardRef((function(e,t){var n=e.components,i=e.mdxType,a=e.originalType,l=e.parentName,u=s(e,["components","mdxType","originalType","parentName"]),m=c(n),p=i,g=m["".concat(l,".").concat(p)]||m[p]||d[p]||a;return n?r.createElement(g,o(o({ref:t},u),{},{components:n})):r.createElement(g,o({ref:t},u))}));function p(e,t){var n=arguments,i=t&&t.mdxType;if("string"==typeof e||i){var a=n.length,o=new Array(a);o[0]=m;var s={};for(var l in t)hasOwnProperty.call(t,l)&&(s[l]=t[l]);s.originalType=e,s.mdxType="string"==typeof e?e:i,o[1]=s;for(var c=2;c<a;c++)o[c]=n[c];return r.createElement.apply(null,o)}return r.createElement.apply(null,n)}m.displayName="MDXCreateElement"},4934:(e,t,n)=>{n.r(t),n.d(t,{assets:()=>l,contentTitle:()=>o,default:()=>d,frontMatter:()=>a,metadata:()=>s,toc:()=>c});var r=n(7462),i=(n(7294),n(3905));const a={title:"Dynamic Delegate Registration"},o=void 0,s={unversionedId:"archive/guides/dynamic-delegate-registration",id:"archive/guides/dynamic-delegate-registration",title:"Dynamic Delegate Registration",description:"In come cases one might want to implement some kind of function wrapper or transformation at runtime. E.g. automatically wrapping and registering async Task / Rx functions.",source:"@site/docs/archive/guides/dynamic-delegate-registration.md",sourceDirName:"archive/guides",slug:"/archive/guides/dynamic-delegate-registration",permalink:"/docs/archive/guides/dynamic-delegate-registration",draft:!1,tags:[],version:"current",frontMatter:{title:"Dynamic Delegate Registration"},sidebar:"tutorialSidebar",previous:{title:"Detecting Excel Shutdown and AutoClose",permalink:"/docs/archive/guides/detecting-excel-shutdown-and-autoclose"},next:{title:"Enumerating Excel COM Automation Collections",permalink:"/docs/archive/guides/enumerating-excel-com-automation-collections"}},l={},c=[],u={toc:c};function d(e){let{components:t,...n}=e;return(0,i.kt)("wrapper",(0,r.Z)({},u,n,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("p",null,"In come cases one might want to implement some kind of function wrapper or transformation at runtime. E.g. automatically wrapping and registering async Task / Rx functions."),(0,i.kt)("p",null,"The latest check-ins for Excel-DNA (check-in 79681 and eventually version 0.32) implement the support required to do this."),(0,i.kt)("p",null,"The key method that has been added is ",(0,i.kt)("inlineCode",{parentName:"p"},"ExcelIntegration.RegisterDelegates(...)"),", which allows you to pass in a list of delegates, together with lists of ",(0,i.kt)("inlineCode",{parentName:"p"},"ExcelFunction")," / ",(0,i.kt)("inlineCode",{parentName:"p"},"ExcelArgument")," attributes. Because this takes ",(0,i.kt)("inlineCode",{parentName:"p"},"Delegates")," and not just ",(0,i.kt)("inlineCode",{parentName:"p"},"MethodInfos"),", you can easily wrap an existing method with to include your processing code for the optional / default values."),(0,i.kt)("p",null,"A useful helper that complements this is ",(0,i.kt)("inlineCode",{parentName:"p"},"ExcelIntegration.GetExportedAssemblies()")," which returns the ",(0,i.kt)("inlineCode",{parentName:"p"},"Assemblies")," that were considered for registration by Excel-DNA - either from ",(0,i.kt)("inlineCode",{parentName:"p"},"ExternalLibrary")," tags in the .dna file or from runtime-compiled source projects inside the .dna file."),(0,i.kt)("p",null,"The basic idea would be:"),(0,i.kt)("p",null,"In your ",(0,i.kt)("inlineCode",{parentName:"p"},"AutoOpen"),", call some kind of ",(0,i.kt)("inlineCode",{parentName:"p"},"UpdateRegistrations()")," which works like this:"),(0,i.kt)("ol",null,(0,i.kt)("li",{parentName:"ol"},"Get all the methods you\u2019re interested in via Reflection (from the assemblies returned by ",(0,i.kt)("inlineCode",{parentName:"li"},"ExcelIntegration.GetExportedAssemblies()"),")."),(0,i.kt)("li",{parentName:"ol"},"Build delegates using lambda expressions that add the optional handling (or using the Expression Tree API for even more control)."),(0,i.kt)("li",{parentName:"ol"},"Register the delegates with the right attributes via ",(0,i.kt)("inlineCode",{parentName:"li"},"ExcelIntegration.RegisterDelegates"),".")),(0,i.kt)("p",null,"This code should be a start:"),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-csharp"},'<DnaLibrary Name="Dynamic Function Tests" Language="C#" RuntimeVersion="v4.0">\n<Reference Name="System.Windows.Forms" />\n<![CDATA[\nusing System;\nusing System.Collections.Generic;\nusing System.Linq;\nusing System.Linq.Expressions;\nusing System.Reflection;\nusing System.Windows.Forms;\nusing ExcelDna.Integration;\n\npublic class TestAddIn : IExcelAddIn\n{\n    public void AutoOpen()\n    {\n        try\n        {\n            MessageBox.Show("In AutoOpen");\n\n            var helloDel = MakeDelegate("Hello ");\n            var byeDel = MakeDelegate("Goodbye ");\n\n            var helloAtt = new ExcelFunctionAttribute\n            {\n                Name = "delHello",\n            };\n            var helloArgAtt = new ExcelArgumentAttribute\n            {\n                Name = "theName",\n                Description = "is the name of the person to say \'Hello\' to."\n            };\n\n            var byeAtt = new ExcelFunctionAttribute\n            {\n                Name = "delGoodbye",\n            };\n            var byeArgAtt = new ExcelArgumentAttribute\n            {\n                Name = "theName",\n                Description = "is the name of the person to say \'Goodbye\' to."\n            };\n\n            var add3Del = MakeAddNumber(3);\n            var add3Att = new ExcelFunctionAttribute\n            {\n                Name = "delAdd3",\n                Description = "Adds 3 to a number",\n                IsThreadSafe = true,\n                IsExceptionSafe = true\n            };\n            var add3ArgAtt = new ExcelArgumentAttribute\n            {\n                Name = "theNumber",\n                Description = "is the number to which the adding is done."\n            };\n\n            ExcelIntegration.RegisterDelegates(\n              new List<Delegate> { helloDel, byeDel, add3Del },\n              new List<object>   { helloAtt, byeAtt, add3Att },\n              new List<List<object>> { new List<object> {helloArgAtt},\n                                       new List<object> {byeArgAtt},\n                                       new List<object> {add3ArgAtt},\n                                     } );\n        }\n        catch (Exception ex)\n        {\n              MessageBox.Show(ex.ToString());\n        }\n    }\n\n    public void AutoClose() {}\n\n    static Func<string, string> MakeDelegate(string sayWhat)\n    {\n        Func<string, string> saySomethingToName = name => sayWhat + name;\n        return saySomethingToName;\n    }\n\n    static Func<double, object> MakeAddNumber(double numberToAdd)\n    {\n      return x =>\n      {\n        try\n        {\n          return x + numberToAdd;\n        }\n        catch (Exception ex)\n        {\n          return double.NaN;\n        }\n      };\n    }\n}\n\n]]>\n</DnaLibrary>\n')))}d.isMDXComponent=!0}}]);