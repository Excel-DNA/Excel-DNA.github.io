"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[2754],{3905:(e,t,n)=>{n.d(t,{Zo:()=>u,kt:()=>f});var r=n(7294);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function a(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);t&&(r=r.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,r)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?a(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):a(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function c(e,t){if(null==e)return{};var n,r,o=function(e,t){if(null==e)return{};var n,r,o={},a=Object.keys(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||(o[n]=e[n]);return o}(e,t);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var l=r.createContext({}),s=function(e){var t=r.useContext(l),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},u=function(e){var t=s(e.components);return r.createElement(l.Provider,{value:t},e.children)},p={inlineCode:"code",wrapper:function(e){var t=e.children;return r.createElement(r.Fragment,{},t)}},m=r.forwardRef((function(e,t){var n=e.components,o=e.mdxType,a=e.originalType,l=e.parentName,u=c(e,["components","mdxType","originalType","parentName"]),m=s(n),f=o,d=m["".concat(l,".").concat(f)]||m[f]||p[f]||a;return n?r.createElement(d,i(i({ref:t},u),{},{components:n})):r.createElement(d,i({ref:t},u))}));function f(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var a=n.length,i=new Array(a);i[0]=m;var c={};for(var l in t)hasOwnProperty.call(t,l)&&(c[l]=t[l]);c.originalType=e,c.mdxType="string"==typeof e?e:o,i[1]=c;for(var s=2;s<a;s++)i[s]=n[s];return r.createElement.apply(null,i)}return r.createElement.apply(null,n)}m.displayName="MDXCreateElement"},1502:(e,t,n)=>{n.r(t),n.d(t,{assets:()=>l,contentTitle:()=>i,default:()=>p,frontMatter:()=>a,metadata:()=>c,toc:()=>s});var r=n(7462),o=(n(7294),n(3905));const a={title:"Enumerating Excel COM Automation Collections"},i=void 0,c={unversionedId:"tips-n-tricks/enumerating-excel-com-automation-collections",id:"tips-n-tricks/enumerating-excel-com-automation-collections",title:"Enumerating Excel COM Automation Collections",description:"When referencing COM Automation collections late-bound, the enumeration via For Each does not automatically work. An explicitly cast or set to a variable of type IEnumerable will work, though:",source:"@site/docs/tips-n-tricks/enumerating-excel-com-automation-collections.md",sourceDirName:"tips-n-tricks",slug:"/tips-n-tricks/enumerating-excel-com-automation-collections",permalink:"/docs/tips-n-tricks/enumerating-excel-com-automation-collections",draft:!1,tags:[],version:"current",frontMatter:{title:"Enumerating Excel COM Automation Collections"},sidebar:"tutorialSidebar",previous:{title:"Creating a Threaded Modal Dialog",permalink:"/docs/tips-n-tricks/creating-a-threaded-modal-dialog"},next:{title:"Archive",permalink:"/docs/category/archive"}},l={},s=[],u={toc:s};function p(e){let{components:t,...n}=e;return(0,o.kt)("wrapper",(0,r.Z)({},u,n,{components:t,mdxType:"MDXLayout"}),(0,o.kt)("p",null,"When referencing COM Automation collections late-bound, the enumeration via ",(0,o.kt)("inlineCode",{parentName:"p"},"For Each")," does not automatically work. An explicitly cast or set to a variable of type ",(0,o.kt)("inlineCode",{parentName:"p"},"IEnumerable")," will work, though:"),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb"},"Dim app As Object = ExcelDnaUtil.Application\n\nDim sh As Object\nDim flg As Boolean\n\nFor Each sh In CType(app.Worksheets, IEnumerable)\n    ' Do stuff with sh here\nNext\n")),(0,o.kt)("p",null,"This should not be needed if an interop library is referenced."))}p.isMDXComponent=!0}}]);