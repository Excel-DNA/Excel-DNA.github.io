"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[1859],{3905:(e,t,n)=>{n.d(t,{Zo:()=>u,kt:()=>m});var r=n(7294);function i(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function a(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);t&&(r=r.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,r)}return n}function o(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?a(Object(n),!0).forEach((function(t){i(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):a(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,r,i=function(e,t){if(null==e)return{};var n,r,i={},a=Object.keys(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||(i[n]=e[n]);return i}(e,t);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(i[n]=e[n])}return i}var c=r.createContext({}),l=function(e){var t=r.useContext(c),n=t;return e&&(n="function"==typeof e?e(t):o(o({},t),e)),n},u=function(e){var t=l(e.components);return r.createElement(c.Provider,{value:t},e.children)},p="mdxType",d={inlineCode:"code",wrapper:function(e){var t=e.children;return r.createElement(r.Fragment,{},t)}},f=r.forwardRef((function(e,t){var n=e.components,i=e.mdxType,a=e.originalType,c=e.parentName,u=s(e,["components","mdxType","originalType","parentName"]),p=l(n),f=i,m=p["".concat(c,".").concat(f)]||p[f]||d[f]||a;return n?r.createElement(m,o(o({ref:t},u),{},{components:n})):r.createElement(m,o({ref:t},u))}));function m(e,t){var n=arguments,i=t&&t.mdxType;if("string"==typeof e||i){var a=n.length,o=new Array(a);o[0]=f;var s={};for(var c in t)hasOwnProperty.call(t,c)&&(s[c]=t[c]);s.originalType=e,s[p]="string"==typeof e?e:i,o[1]=s;for(var l=2;l<a;l++)o[l]=n[l];return r.createElement.apply(null,o)}return r.createElement.apply(null,n)}f.displayName="MDXCreateElement"},5674:(e,t,n)=>{n.r(t),n.d(t,{assets:()=>c,contentTitle:()=>o,default:()=>d,frontMatter:()=>a,metadata:()=>s,toc:()=>l});var r=n(7462),i=(n(7294),n(3905));const a={title:"Function Registration"},o=void 0,s={unversionedId:"guides-basic/Function Registration",id:"guides-basic/Function Registration",title:"Function Registration",description:"Note: This document reflects changes made in Excel-DNA v1.9.",source:"@site/docs/guides-basic/Function Registration.md",sourceDirName:"guides-basic",slug:"/guides-basic/Function Registration",permalink:"/docs/guides-basic/Function Registration",draft:!1,tags:[],version:"current",frontMatter:{title:"Function Registration"},sidebar:"tutorialSidebar",previous:{title:"Excel C API",permalink:"/docs/guides-basic/excel-programming-interfaces/excel-c-api"},next:{title:"IntelliSense",permalink:"/docs/guides-basic/Intellisense"}},c={},l=[{value:"Changes from earlier versions",id:"changes-from-earlier-versions",level:3}],u={toc:l},p="wrapper";function d(e){let{components:t,...n}=e;return(0,i.kt)(p,(0,r.Z)({},u,n,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("p",null,(0,i.kt)("strong",{parentName:"p"},"Note:")," This document reflects changes made in Excel-DNA v1.9."),(0,i.kt)("p",null,"Excel-DNA support the creation of user-defined functions for Excel in .NET. This document describes how functions are selected for registration, supported method signatures, method conversions and extension points."),(0,i.kt)("h3",{id:"changes-from-earlier-versions"},"Changes from earlier versions"),(0,i.kt)("p",null,"In v1.9 we incorporate the functionality previously exposes in the separate ",(0,i.kt)("inlineCode",{parentName:"p"},"ExcelDna.Registration")," library (and package) into the main Excel-DNA library.\nTo use the extended registration features under older versions, explict registration was required.\nUnder v1.9 we have:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"expanded the supported parameter and return types for functions markes with ",(0,i.kt)("inlineCode",{parentName:"li"},"[ExcelFunction]"),"."),(0,i.kt)("li",{parentName:"ul"},"added support for async and streaming functions and object handles (with the ",(0,i.kt)("inlineCode",{parentName:"li"},"[return:ExcelHandle]")," and ",(0,i.kt)("inlineCode",{parentName:"li"},"[ExcelHandle]")," attributes) in the main library"),(0,i.kt)("li",{parentName:"ul"},"migrated registration extension points like ",(0,i.kt)("inlineCode",{parentName:"li"},"FunctionExecutionHandler")," from the ExcelDna.Registration package to the main ExcelDna.Integration library.")))}d.isMDXComponent=!0}}]);