"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[6135],{3905:(e,t,r)=>{r.d(t,{Zo:()=>p,kt:()=>h});var n=r(7294);function a(e,t,r){return t in e?Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[t]=r,e}function i(e,t){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),r.push.apply(r,n)}return r}function o(e){for(var t=1;t<arguments.length;t++){var r=null!=arguments[t]?arguments[t]:{};t%2?i(Object(r),!0).forEach((function(t){a(e,t,r[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):i(Object(r)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(r,t))}))}return e}function c(e,t){if(null==e)return{};var r,n,a=function(e,t){if(null==e)return{};var r,n,a={},i=Object.keys(e);for(n=0;n<i.length;n++)r=i[n],t.indexOf(r)>=0||(a[r]=e[r]);return a}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(n=0;n<i.length;n++)r=i[n],t.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(a[r]=e[r])}return a}var s=n.createContext({}),l=function(e){var t=n.useContext(s),r=t;return e&&(r="function"==typeof e?e(t):o(o({},t),e)),r},p=function(e){var t=l(e.components);return n.createElement(s.Provider,{value:t},e.children)},u={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},d=n.forwardRef((function(e,t){var r=e.components,a=e.mdxType,i=e.originalType,s=e.parentName,p=c(e,["components","mdxType","originalType","parentName"]),d=l(r),h=a,f=d["".concat(s,".").concat(h)]||d[h]||u[h]||i;return r?n.createElement(f,o(o({ref:t},p),{},{components:r})):n.createElement(f,o({ref:t},p))}));function h(e,t){var r=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var i=r.length,o=new Array(i);o[0]=d;var c={};for(var s in t)hasOwnProperty.call(t,s)&&(c[s]=t[s]);c.originalType=e,c.mdxType="string"==typeof e?e:a,o[1]=c;for(var l=2;l<i;l++)o[l]=r[l];return n.createElement.apply(null,o)}return n.createElement.apply(null,r)}d.displayName="MDXCreateElement"},6203:(e,t,r)=>{r.r(t),r.d(t,{assets:()=>s,contentTitle:()=>o,default:()=>u,frontMatter:()=>i,metadata:()=>c,toc:()=>l});var n=r(7462),a=(r(7294),r(3905));const i={},o=void 0,c={unversionedId:"archive/wiki/COM-exports-for-VBA-access",id:"archive/wiki/COM-exports-for-VBA-access",title:"COM-exports-for-VBA-access",description:"Excel-DNA supports registering the .xll as a regular COM library, which can then be accessed from VBA (either late-bound via CreateObject or referenced in a project via Tools->References. This allows COM-visible classes in the add-in to be instantiated and accessed from VBA, with two advantages over regular COM libraries:",source:"@site/docs/archive/wiki/COM-exports-for-VBA-access.md",sourceDirName:"archive/wiki",slug:"/archive/wiki/COM-exports-for-VBA-access",permalink:"/docs/archive/wiki/COM-exports-for-VBA-access",draft:!1,tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"Build-Output-Customization",permalink:"/docs/archive/wiki/Build-Output-Customization"},next:{title:"COM-object-model-notes",permalink:"/docs/archive/wiki/COM-object-model-notes"}},s={},l=[],p={toc:l};function u(e){let{components:t,...r}=e;return(0,a.kt)("wrapper",(0,n.Z)({},p,r,{components:t,mdxType:"MDXLayout"}),(0,a.kt)("p",null,"Excel-DNA supports registering the .xll as a regular COM library, which can then be accessed from VBA (either late-bound via ",(0,a.kt)("inlineCode",{parentName:"p"},"CreateObject")," or referenced in a project via Tools->References. This allows COM-visible classes in the add-in to be instantiated and accessed from VBA, with two advantages over regular COM libraries:"),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"The COM types are registered in the registry under the user hive if there is no access to the machine hive - this means that users with limited permissions in the registry can still use the COM objects.")),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"The COM objects are created in the same AppDomain as the rest of the Excel-DNA add-in. Among other things, this means that static references are shared between the UDF functions and the COM objects. So a cache or settings that are used by the UDF functions can be shared by objects instantiated and called from VBA."))),(0,a.kt)("p",null,"More details in the particular .dna settings to add and step-by-step instructions for trying this out, can be found here:"),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("a",{parentName:"li",href:"http://mikejuniperhill.blogspot.co.za/2014/03/interfacing-c-and-vba-with-exceldna-no.html"},"http://mikejuniperhill.blogspot.co.za/2014/03/interfacing-c-and-vba-with-exceldna-no.html")),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("a",{parentName:"li",href:"http://mikejuniperhill.blogspot.co.za/2014/03/interfacing-c-and-vba-with-exceldna_16.html"},"http://mikejuniperhill.blogspot.co.za/2014/03/interfacing-c-and-vba-with-exceldna_16.html"))))}u.isMDXComponent=!0}}]);