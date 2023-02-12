"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[5357],{3905:(e,t,n)=>{n.d(t,{Zo:()=>s,kt:()=>H});var r=n(7294);function a(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function i(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);t&&(r=r.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,r)}return n}function o(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?i(Object(n),!0).forEach((function(t){a(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):i(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function l(e,t){if(null==e)return{};var n,r,a=function(e,t){if(null==e)return{};var n,r,a={},i=Object.keys(e);for(r=0;r<i.length;r++)n=i[r],t.indexOf(n)>=0||(a[n]=e[n]);return a}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(r=0;r<i.length;r++)n=i[r],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(a[n]=e[n])}return a}var h=r.createContext({}),f=function(e){var t=r.useContext(h),n=t;return e&&(n="function"==typeof e?e(t):o(o({},t),e)),n},s=function(e){var t=f(e.components);return r.createElement(h.Provider,{value:t},e.children)},c={inlineCode:"code",wrapper:function(e){var t=e.children;return r.createElement(r.Fragment,{},t)}},D=r.forwardRef((function(e,t){var n=e.components,a=e.mdxType,i=e.originalType,h=e.parentName,s=l(e,["components","mdxType","originalType","parentName"]),D=f(n),H=a,P=D["".concat(h,".").concat(H)]||D[H]||c[H]||i;return n?r.createElement(P,o(o({ref:t},s),{},{components:n})):r.createElement(P,o({ref:t},s))}));function H(e,t){var n=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var i=n.length,o=new Array(i);o[0]=D;var l={};for(var h in t)hasOwnProperty.call(t,h)&&(l[h]=t[h]);l.originalType=e,l.mdxType="string"==typeof e?e:a,o[1]=l;for(var f=2;f<i;f++)o[f]=n[f];return r.createElement.apply(null,o)}return r.createElement.apply(null,n)}D.displayName="MDXCreateElement"},1129:(e,t,n)=>{n.r(t),n.d(t,{assets:()=>h,contentTitle:()=>o,default:()=>c,frontMatter:()=>i,metadata:()=>l,toc:()=>f});var r=n(7462),a=(n(7294),n(3905));const i={title:"Excel UDF IntelliSense for Excel-DNA and VBA",date:"2016-11-24 11:44:00 -0000",authors:"govert"},o=void 0,l={permalink:"/blog/2016/11/24/excel-udf-intellisense-for-excel-dna-and-vba",source:"@site/blog/2016-11-24-excel-udf-intellisense-for-excel-dna-and-vba/index.md",title:"Excel UDF IntelliSense for Excel-DNA and VBA",description:"I'm happy to announce the first official release of the IntelliSense extension!",date:"2016-11-24T11:44:00.000Z",formattedDate:"November 24, 2016",tags:[],readingTime:1.71,hasTruncateMarker:!1,authors:[{name:"Govert van Drimmelen",url:"https://github.com/Excel-DNA",imageURL:"https://avatars.githubusercontent.com/u/414659",key:"govert"}],frontMatter:{title:"Excel UDF IntelliSense for Excel-DNA and VBA",date:"2016-11-24 11:44:00 -0000",authors:"govert"},prevItem:{title:"Excel-DNA version 0.34",permalink:"/blog/2017/05/31/excel-dna-0-34-final-testing"},nextItem:{title:"Add-in spotlight: ACQ for interpolation",permalink:"/blog/2016/06/10/add-in-spotlight-acq-for-interpolation"}},h={authorsImageUrls:[void 0]},f=[],s={toc:f};function c(e){let{components:t,...i}=e;return(0,a.kt)("wrapper",(0,r.Z)({},s,i,{components:t,mdxType:"MDXLayout"}),(0,a.kt)("p",null,"I'm happy to announce the first official release of the IntelliSense extension!"),(0,a.kt)("p",null,(0,a.kt)("strong",{parentName:"p"},"Excel-DNA IntelliSense")," provides on-sheet help for UDF functions as they are entered into a cell formula, similar to the help available for built-in Excel functions."),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"Intellisense Release v1 01",src:n(1556).Z,width:"557",height:"287"}),"\n",(0,a.kt)("img",{alt:"Intellisense Release v1 02",src:n(2281).Z,width:"412",height:"130"})),(0,a.kt)("p",null,(0,a.kt)("strong",{parentName:"p"},"For Excel-DNA add-ins")," (v0.32 and later) that already provide descriptions in the ",(0,a.kt)("inlineCode",{parentName:"p"},"[ExcelFunction]")," and ",(0,a.kt)("inlineCode",{parentName:"p"},"[ExcelArgument]")," attributes, no extra work is needed. Just download and open (or install) the latest ",(0,a.kt)("inlineCode",{parentName:"p"},"ExcelDna.IntelliSense.xll")," add-in from the GitHub (",(0,a.kt)("a",{parentName:"p",href:"https://github.com/Excel-DNA/IntelliSense/releases"},"https://github.com/Excel-DNA/IntelliSense/releases"),"), and the IntelliSense will light up. (There is also a NuGet package for embedding the service into your add-in, making distribution a bit easier.)"),(0,a.kt)("p",null,(0,a.kt)("strong",{parentName:"p"},"For VBA functions"),", you can add an extra sheet with the IntelliSense descriptions, or add an external .xml file with the information, or embed as a the ",(0,a.kt)("inlineCode",{parentName:"p"},"CustomXML")," part in the Workbook or ",(0,a.kt)("inlineCode",{parentName:"p"},".xlam")," add-in.\nThen open (or install) the ",(0,a.kt)("inlineCode",{parentName:"p"},"ExcelDna.IntelliSense.xll")," add-in to provide the display service. Charles Williams, of ",(0,a.kt)("a",{parentName:"p",href:"http://www.decisionmodels.com/fastexcelD.htm"},"FastExcel")," fame, has a detailed write-up on adding IntelliSense for your VBA function - see ",(0,a.kt)("a",{parentName:"p",href:"https://fastexcel.wordpress.com/2016/10/07/writing-efficient-vba-udfs-part-15-adding-intellisense-to-your-udfs/"},"https://fastexcel.wordpress.com/2016/10/07/writing-efficient-vba-udfs-part-15-adding-intellisense-to-your-udfs/"),"."),(0,a.kt)("p",null,(0,a.kt)("strong",{parentName:"p"},"For PyXLL users"),", the latest PyXLL 3.1 release offer built-in support for IntelliSense with the ExcelDna.IntelliSense.xll add-in installed. See ",(0,a.kt)("a",{parentName:"p",href:"https://enthought.pyxll.com/whatsnew.html#intellisense"},"https://enthought.pyxll.com/whatsnew.html#intellisense"),"."),(0,a.kt)("p",null,"Other native .xll add-ins can also provide IntelliSense through an external .xml file."),(0,a.kt)("p",null,"Details and downloads are on GitHub:"),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},"Home: ",(0,a.kt)("a",{parentName:"li",href:"https://github.com/Excel-DNA/IntelliSense"},"https://github.com/Excel-DNA/IntelliSense")),(0,a.kt)("li",{parentName:"ul"},"Releases: ",(0,a.kt)("a",{parentName:"li",href:"https://github.com/Excel-DNA/IntelliSense/releases"},"https://github.com/Excel-DNA/IntelliSense/releases")),(0,a.kt)("li",{parentName:"ul"},"Getting Started: ",(0,a.kt)("a",{parentName:"li",href:"https://github.com/Excel-DNA/IntelliSense/wiki/Getting-Started"},"https://github.com/Excel-DNA/IntelliSense/wiki/Getting-Started")),(0,a.kt)("li",{parentName:"ul"},"Detailed Usage Instructions: ",(0,a.kt)("a",{parentName:"li",href:"https://github.com/Excel-DNA/IntelliSense/wiki/Usage-Instructions"},"https://github.com/Excel-DNA/IntelliSense/wiki/Usage-Instructions")," including details for incorporating the library into your own add-in for easier distribution.")),(0,a.kt)("p",null,(0,a.kt)("strong",{parentName:"p"},"Public support and bug reports"),":\nThe Excel-DNA Google group (",(0,a.kt)("a",{parentName:"p",href:"https://groups.google.com/forum/#!forum/exceldna"},"https://groups.google.com/forum/#!forum/exceldna"),") is the best place for general questions, comments etc. Detailed bug reports and feature requests can be added to the GitHub issues list: ",(0,a.kt)("a",{parentName:"p",href:"https://github.com/Excel-DNA/IntelliSense/issues"},"https://github.com/Excel-DNA/IntelliSense/issues")),(0,a.kt)("p",null,(0,a.kt)("strong",{parentName:"p"},"Corporate support and private donations"),":\nIf you find Excel-DNA and extensions like the IntelliSense service useful, please support the project by arranging a corporate support agreement, or making a donation via PayPal. See ",(0,a.kt)("a",{parentName:"p",href:"/support/"},"https://excel-dna.net/support/")," for details and contact information."))}c.isMDXComponent=!0},1556:(e,t,n)=>{n.d(t,{Z:()=>r});const r="data:image/png;base64,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"},2281:(e,t,n)=>{n.d(t,{Z:()=>r});const r="data:image/png;base64,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"}}]);