"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[2267],{3905:(e,t,r)=>{r.d(t,{Zo:()=>p,kt:()=>m});var a=r(7294);function n(e,t,r){return t in e?Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[t]=r,e}function l(e,t){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),r.push.apply(r,a)}return r}function o(e){for(var t=1;t<arguments.length;t++){var r=null!=arguments[t]?arguments[t]:{};t%2?l(Object(r),!0).forEach((function(t){n(e,t,r[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):l(Object(r)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(r,t))}))}return e}function i(e,t){if(null==e)return{};var r,a,n=function(e,t){if(null==e)return{};var r,a,n={},l=Object.keys(e);for(a=0;a<l.length;a++)r=l[a],t.indexOf(r)>=0||(n[r]=e[r]);return n}(e,t);if(Object.getOwnPropertySymbols){var l=Object.getOwnPropertySymbols(e);for(a=0;a<l.length;a++)r=l[a],t.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(n[r]=e[r])}return n}var c=a.createContext({}),s=function(e){var t=a.useContext(c),r=t;return e&&(r="function"==typeof e?e(t):o(o({},t),e)),r},p=function(e){var t=s(e.components);return a.createElement(c.Provider,{value:t},e.children)},u={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},d=a.forwardRef((function(e,t){var r=e.components,n=e.mdxType,l=e.originalType,c=e.parentName,p=i(e,["components","mdxType","originalType","parentName"]),d=s(r),m=n,g=d["".concat(c,".").concat(m)]||d[m]||u[m]||l;return r?a.createElement(g,o(o({ref:t},p),{},{components:r})):a.createElement(g,o({ref:t},p))}));function m(e,t){var r=arguments,n=t&&t.mdxType;if("string"==typeof e||n){var l=r.length,o=new Array(l);o[0]=d;var i={};for(var c in t)hasOwnProperty.call(t,c)&&(i[c]=t[c]);i.originalType=e,i.mdxType="string"==typeof e?e:n,o[1]=i;for(var s=2;s<l;s++)o[s]=r[s];return a.createElement.apply(null,o)}return a.createElement.apply(null,r)}d.displayName="MDXCreateElement"},2954:(e,t,r)=>{r.r(t),r.d(t,{assets:()=>c,contentTitle:()=>o,default:()=>u,frontMatter:()=>l,metadata:()=>i,toc:()=>s});var a=r(7462),n=(r(7294),r(3905));const l={title:"Tutorial: COM server support for VBA integration",date:"2014-03-21 23:16:00 -0000",authors:"govert",tags:[".net","com","excel","excel-vba","exceldna","vba","xll"]},o=void 0,i={permalink:"/blog/2014/03/21/tutorial-com-server-support-for-vba-integration",source:"@site/blog/2014-03-21-tutorial-com-server-support-for-vba-integration.md",title:"Tutorial: COM server support for VBA integration",description:'Functions and macros created in an Excel-DNA add-in can be called directly from Excel VBA by using Application.Run(...). However, .NET also supports creating rich object models that are exported as COM libraries, which can be Tools->Referenced in VBA. Excel-DNA has some advanced support to host COM-exported objects from Excel-DNA add-ins, giving some advantages over the regular .NET "Register for COM interop" hosting approach:',date:"2014-03-21T23:16:00.000Z",formattedDate:"March 21, 2014",tags:[{label:".net",permalink:"/blog/tags/net"},{label:"com",permalink:"/blog/tags/com"},{label:"excel",permalink:"/blog/tags/excel"},{label:"excel-vba",permalink:"/blog/tags/excel-vba"},{label:"exceldna",permalink:"/blog/tags/exceldna"},{label:"vba",permalink:"/blog/tags/vba"},{label:"xll",permalink:"/blog/tags/xll"}],readingTime:1.165,hasTruncateMarker:!1,authors:[{name:"Govert van Drimmelen",url:"https://github.com/Excel-DNA",imageURL:"https://avatars.githubusercontent.com/u/414659",key:"govert"}],frontMatter:{title:"Tutorial: COM server support for VBA integration",date:"2014-03-21 23:16:00 -0000",authors:"govert",tags:[".net","com","excel","excel-vba","exceldna","vba","xll"]},prevItem:{title:"Excel-DNA 0.32 - Breaking changes to integer and boolean parameter handling",permalink:"/blog/2014/05/03/excel-dna-0-32-breaking-changes-to-integer-and-boolean-parameter-handling"},nextItem:{title:"Excel-DNA 0.32 Release Candidate",permalink:"/blog/2014/03/03/excel-dna-0-32-release-candidate"}},c={authorsImageUrls:[void 0]},s=[],p={toc:s};function u(e){let{components:t,...r}=e;return(0,n.kt)("wrapper",(0,a.Z)({},p,r,{components:t,mdxType:"MDXLayout"}),(0,n.kt)("p",null,"Functions and macros created in an Excel-DNA add-in can be called directly from Excel VBA by using ",(0,n.kt)("inlineCode",{parentName:"p"},"Application.Run(...)"),'. However, .NET also supports creating rich object models that are exported as COM libraries, which can be Tools->Referenced in VBA. Excel-DNA has some advanced support to host COM-exported objects from Excel-DNA add-ins, giving some advantages over the regular .NET "Register for COM interop" hosting approach:'),(0,n.kt)("ul",null,(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"COM objects that are created via the Excel-DNA COM server support will be active in the same AppDomain as the rest of the add-in, allowing direct shared access to static variables, internal caches etc.")),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"COM registration for classes hosted by Excel-DNA does not require administrative access (even when registered via ",(0,n.kt)("inlineCode",{parentName:"p"},"RegSvr32.exe"),").")),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"Everything needed for the COM server can be packed in a single-file .xll add-in, including the type library used for IntelliSense support in VBA."))),(0,n.kt)("p",null,(0,n.kt)("a",{parentName:"p",href:"http://mikejuniperhill.blogspot.com/"},"Mikael Katajam\xe4ki")," has written some detailed tutorial posts on his ",(0,n.kt)("a",{parentName:"p",href:"http://mikejuniperhill.blogspot.com/"},"Excel in Finance")," blog that explore this Excel-DNA feature, with detailed explanation, step-by-step instructions, screen shots and further links. See:"),(0,n.kt)("ul",null,(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("a",{parentName:"li",href:"http://mikejuniperhill.blogspot.com/2014/03/interfacing-c-and-vba-with-exceldna-no.html"},"Interfacing C# and VBA with Excel-DNA (no intellisense support)")),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("a",{parentName:"li",href:"http://mikejuniperhill.blogspot.com/2014/03/interfacing-c-and-vba-with-exceldna_16.html"},"Interfacing C# and VBA with Excel-DNA (with intellisense support)"))),(0,n.kt)("p",null,"Note that these techniques would work equally well with code written in VB.NET, allowing you to port VB/VBA libraries to VB.NET with Excel-DNA and then use these from VBA."),(0,n.kt)("p",null,"Thank you Mikael for the great write-up!"))}u.isMDXComponent=!0}}]);