"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[8554],{3905:(e,n,t)=>{t.d(n,{Zo:()=>i,kt:()=>f});var r=t(7294);function l(e,n,t){return n in e?Object.defineProperty(e,n,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[n]=t,e}function a(e,n){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);n&&(r=r.filter((function(n){return Object.getOwnPropertyDescriptor(e,n).enumerable}))),t.push.apply(t,r)}return t}function A(e){for(var n=1;n<arguments.length;n++){var t=null!=arguments[n]?arguments[n]:{};n%2?a(Object(t),!0).forEach((function(n){l(e,n,t[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):a(Object(t)).forEach((function(n){Object.defineProperty(e,n,Object.getOwnPropertyDescriptor(t,n))}))}return e}function s(e,n){if(null==e)return{};var t,r,l=function(e,n){if(null==e)return{};var t,r,l={},a=Object.keys(e);for(r=0;r<a.length;r++)t=a[r],n.indexOf(t)>=0||(l[t]=e[t]);return l}(e,n);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(r=0;r<a.length;r++)t=a[r],n.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(l[t]=e[t])}return l}var o=r.createContext({}),u=function(e){var n=r.useContext(o),t=n;return e&&(t="function"==typeof e?e(n):A(A({},n),e)),t},i=function(e){var n=u(e.components);return r.createElement(o.Provider,{value:n},e.children)},c={inlineCode:"code",wrapper:function(e){var n=e.children;return r.createElement(r.Fragment,{},n)}},d=r.forwardRef((function(e,n){var t=e.components,l=e.mdxType,a=e.originalType,o=e.parentName,i=s(e,["components","mdxType","originalType","parentName"]),d=u(t),f=l,p=d["".concat(o,".").concat(f)]||d[f]||c[f]||a;return t?r.createElement(p,A(A({ref:n},i),{},{components:t})):r.createElement(p,A({ref:n},i))}));function f(e,n){var t=arguments,l=n&&n.mdxType;if("string"==typeof e||l){var a=t.length,A=new Array(a);A[0]=d;var s={};for(var o in n)hasOwnProperty.call(n,o)&&(s[o]=n[o]);s.originalType=e,s.mdxType="string"==typeof e?e:l,A[1]=s;for(var u=2;u<a;u++)A[u]=t[u];return r.createElement.apply(null,A)}return r.createElement.apply(null,t)}d.displayName="MDXCreateElement"},4558:(e,n,t)=>{t.r(n),t.d(n,{assets:()=>D,contentTitle:()=>g,default:()=>B,frontMatter:()=>m,metadata:()=>h,toc:()=>y});var r=t(7462),l=t(7294),a=t(3905),A=t(6010),s=t(2389),o=t(7392),u=t(7094),i=t(2466);const c="tabList__CuJ",d="tabItem_LNqP";function f(e){var n;const{lazy:t,block:a,defaultValue:s,values:f,groupId:p,className:b}=e,v=l.Children.map(e.children,(e=>{if((0,l.isValidElement)(e)&&"value"in e.props)return e;throw new Error(`Docusaurus error: Bad <Tabs> child <${"string"==typeof e.type?e.type:e.type.name}>: all children of the <Tabs> component should be <TabItem>, and every <TabItem> should have a unique "value" prop.`)})),m=f??v.map((e=>{let{props:{value:n,label:t,attributes:r}}=e;return{value:n,label:t,attributes:r}})),g=(0,o.l)(m,((e,n)=>e.value===n.value));if(g.length>0)throw new Error(`Docusaurus error: Duplicate values "${g.map((e=>e.value)).join(", ")}" found in <Tabs>. Every value needs to be unique.`);const h=null===s?s:s??(null==(n=v.find((e=>e.props.default)))?void 0:n.props.value)??v[0].props.value;if(null!==h&&!m.some((e=>e.value===h)))throw new Error(`Docusaurus error: The <Tabs> has a defaultValue "${h}" but none of its children has the corresponding value. Available values are: ${m.map((e=>e.value)).join(", ")}. If you intend to show no default tab, use defaultValue={null} instead.`);const{tabGroupChoices:D,setTabGroupChoices:y}=(0,u.U)(),[x,B]=(0,l.useState)(h),w=[],{blockElementScrollPositionUntilNextRender:k}=(0,i.o5)();if(null!=p){const e=D[p];null!=e&&e!==x&&m.some((n=>n.value===e))&&B(e)}const I=e=>{const n=e.currentTarget,t=w.indexOf(n),r=m[t].value;r!==x&&(k(n),B(r),null!=p&&y(p,String(r)))},F=e=>{var n;let t=null;switch(e.key){case"Enter":I(e);break;case"ArrowRight":{const n=w.indexOf(e.currentTarget)+1;t=w[n]??w[0];break}case"ArrowLeft":{const n=w.indexOf(e.currentTarget)-1;t=w[n]??w[w.length-1];break}}null==(n=t)||n.focus()};return l.createElement("div",{className:(0,A.Z)("tabs-container",c)},l.createElement("ul",{role:"tablist","aria-orientation":"horizontal",className:(0,A.Z)("tabs",{"tabs--block":a},b)},m.map((e=>{let{value:n,label:t,attributes:a}=e;return l.createElement("li",(0,r.Z)({role:"tab",tabIndex:x===n?0:-1,"aria-selected":x===n,key:n,ref:e=>w.push(e),onKeyDown:F,onClick:I},a,{className:(0,A.Z)("tabs__item",d,null==a?void 0:a.className,{"tabs__item--active":x===n})}),t??n)}))),t?(0,l.cloneElement)(v.filter((e=>e.props.value===x))[0],{className:"margin-top--md"}):l.createElement("div",{className:"margin-top--md"},v.map(((e,n)=>(0,l.cloneElement)(e,{key:n,hidden:e.props.value!==x})))))}function p(e){const n=(0,s.Z)();return l.createElement(f,(0,r.Z)({key:String(n)},e))}const b="tabItem_Ymn6";function v(e){let{children:n,hidden:t,className:r}=e;return l.createElement("div",{role:"tabpanel",className:(0,A.Z)(b,r),hidden:t},n)}const m={title:"IntelliSense"},g=void 0,h={unversionedId:"guides-basic/Intellisense",id:"guides-basic/Intellisense",title:"IntelliSense",description:"",source:"@site/docs/guides-basic/Intellisense.md",sourceDirName:"guides-basic",slug:"/guides-basic/Intellisense",permalink:"/TestDocs/docs/guides-basic/Intellisense",draft:!1,tags:[],version:"current",frontMatter:{title:"IntelliSense"},sidebar:"tutorialSidebar",previous:{title:"Guides - Basic",permalink:"/TestDocs/docs/category/guides---basic"},next:{title:"Asynchronous Functions",permalink:"/TestDocs/docs/guides-basic/asynchronous-functions"}},D={},y=[{value:"Usage",id:"usage",level:2},{value:"Additional Remarks",id:"additional-remarks",level:2}],x={toc:y};function B(e){let{components:n,...l}=e;return(0,a.kt)("wrapper",(0,r.Z)({},x,l,{components:n,mdxType:"MDXLayout"}),(0,a.kt)("p",null,"Excel includes a feature that enables users to utilise functions more effectively by displaying pop-up description regarding the function and its parameters. Examples of IntelliSense in action can be seen in the figure below:"),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"IntelliSense Examples",src:t(564).Z,width:"691",height:"111"})),(0,a.kt)("p",null,"By default, Excel does not provide native IntelliSense functionality to user-defined functions (UDFs). As a result, a custom ",(0,a.kt)("a",{parentName:"p",href:"https://github.com/Excel-DNA/IntelliSense"},"Excel-DNA IntelliSense")," library was created as part of the Excel-DNA project."),(0,a.kt)("h2",{id:"usage"},"Usage"),(0,a.kt)("p",null,"Using Excel-DNA's IntelliSense capabilities is relatively simple and requires just 3 steps."),(0,a.kt)("ol",null,(0,a.kt)("li",{parentName:"ol"},"Depending on the language of choice, in the .csproj, .vbproj, or .fsproj file, add the following under ",(0,a.kt)("em",{parentName:"li"},"</PropertyGroup",">"),":")),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-xml"},'<ItemGroup>\n    <PackageReference Include="ExcelDna.IntelliSense" Version="*-*" />\n</ItemGroup>\n')),(0,a.kt)("ol",{start:2},(0,a.kt)("li",{parentName:"ol"},"Add a new class file to the solution by pressing Ctrl+Shift+A and include the following code:")),(0,a.kt)(p,{mdxType:"Tabs"},(0,a.kt)(v,{value:"csharp",label:"C#",mdxType:"TabItem"},(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-csharp"},"using ExcelDna.Integration;\nusing ExcelDna.IntelliSense;\n\npublic class IntelliSenseAddIn : IExcelAddIn\n{\n    public void AutoOpen()\n    {\n        IntelliSenseServer.Install();\n    }\n    public void AutoClose()\n    {\n        IntelliSenseServer.Uninstall();\n    }   \n}\n"))),(0,a.kt)(v,{value:"vbnet",label:"VB.Net",mdxType:"TabItem"},(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vbnet"},"Imports ExcelDna.Integration\nImports ExcelDna.IntelliSense\n\nPublic Class IntelliSenseAddIn\n    Implements IExcelAddIn\n\n    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen\n        IntelliSenseServer.Install()\n    End Sub\n\n    Public Sub AutoClose() Implements IExcelAddIn.AutoClose\n        IntelliSenseServer.Uninstall()\n    End Sub\nEnd Class\n"))),(0,a.kt)(v,{value:"fsharp",label:"F#",mdxType:"TabItem"},(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-fsharp"},"namespace TestFsIntelliSense\n\nopen ExcelDna.Integration\nopen ExcelDna.IntelliSense\n\ntype IntelliSenseAddIn() =\n    interface IExcelAddIn with\n        member _.AutoOpen() =\n            IntelliSenseServer.Install()\n        member _.AutoClose() =\n            IntelliSenseServer.Uninstall()\n")))),(0,a.kt)("ol",{start:3},(0,a.kt)("li",{parentName:"ol"},"Decorate the UDF with the ",(0,a.kt)("inlineCode",{parentName:"li"},"ExcelFunction")," and ",(0,a.kt)("inlineCode",{parentName:"li"},"ExcelArgument")," as per the example code snippet below:")),(0,a.kt)(p,{mdxType:"Tabs"},(0,a.kt)(v,{value:"csharp",label:"C#",mdxType:"TabItem"},(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-csharp"},'[ExcelFunction(Description = "A useful test function that adds two numbers, and returns the sum.")]\npublic static double AddThem(\n    [ExcelArgument(Name = "Augend", Description = "is the first number, to which will be added")] \n    double v1,\n    [ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]     \n    double v2)\n{\n    return v1 + v2;\n}\n'))),(0,a.kt)(v,{value:"vbnet",label:"VB.Net",mdxType:"TabItem"},(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vbnet"},'<ExcelFunction(Description:="A useful test function that adds two numbers, and returns the sum.")>\nPublic Shared Function AddThem(\n    <ExcelArgument(Name:="Augend", Description:="is the first number, to which will be added")>\n    v1 As Double,\n    <ExcelArgument(Name:="Addend", Description:="is the second number that will be added")>\n    v2 As Double) As Double\n\n    Return v1 + v2\nEnd Function\n'))),(0,a.kt)(v,{value:"fsharp",label:"F#",mdxType:"TabItem"},(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-fsharp"},'[<ExcelFunction(Description="A useful test function that adds two numbers, and returns the sum.")>]\nlet AddThem (\n    [<ExcelArgument(Name="Augend", Description="is the first number, to which will be added")>]\n    v1:double) (\n    [<ExcelArgument(Name="Addend", Description="is the second number that will be added")>]\n    v2:double) = v1 + v2\n')))),(0,a.kt)("p",null,"The example above would generate the following IntelliSense output:"),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"IntelliSense UDF example",src:t(1032).Z,width:"937",height:"130"})),(0,a.kt)("h2",{id:"additional-remarks"},"Additional Remarks"),(0,a.kt)("p",null,"If the UDF is decorated as per step 3 above but the first 2 steps are skipped, it would still be possible to get limited information regarding the UDF. The information can be seen by writing the name of the function in Excel's Forumla Bar and clicking on the ",(0,a.kt)("inlineCode",{parentName:"p"},"fx")," button (highlighted in red). The result can be seen in the figure below:"),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"IntelliSense Default",src:t(3751).Z,width:"519",height:"318"})))}B.isMDXComponent=!0},3751:(e,n,t)=>{t.d(n,{Z:()=>r});const r=t.p+"assets/images/intellisense_default-9d08004049848de333b49db260de1103.png"},564:(e,n,t)=>{t.d(n,{Z:()=>r});const r="data:image/png;base64,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"},1032:(e,n,t)=>{t.d(n,{Z:()=>r});const r="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAA6kAAACCCAIAAAArPhmMAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAACAzSURBVHhe7Z3PjxRHlsf9r+x/seLqZTRHc7c5rZDp64416wOSxbFloT4hjTkbDTI+GDUcGJk57KFlzmaslgEPoJnDaIRES7RGaBrvi4xfLyMiKzOrKjuyqj4flVwvX7x48c2sqqgv5eruD34FAAAAANgN8L4AAAAAsCvgfQEAAABgV8D7AgAAAMCugPcFAAAAgF0B7wsAAAAAuwLeFwAAAAB2BbwvAAAAAOwKeF8AAAAA2BXwvgAAAACwK+B9AQAAAGBXwPsCAAAAwK6A9wUAAACAXQHvCwAAAAC7At4XAAAAAHYFvC8AAAAA7Ap4XwAAAADYFfC+AAAAALAr4H0BAAAAYFfA+wIAAADAroD3BQAAwwcAABuF27xGgvcFAABD8kby17/+1UWbCfrrgv667IJ+vC8AAKwE3ndWoL8u6K/LEP14XwAAWAm876xAf13QX5ch+if3vrIAAABsBG7jHkkykff+uqC/LuivyxD9y+917r6PrgV4ctQF/XVBf13QXwTva0F/XdBfl13Qj/ddEvTXBf11QX9dJtK/Fu/7H//zm+24ufPZQHh61wX9dRmiH++7JOivC/rrgv66TKR/Fe8rkiyJg9zcmzsfANg66njfZIvZ3Js7nw1EHnsXbSborwv66zKR/lW8r4v89u4OoAY8veuC/roM0X9O3lekaIJ33PSbOx8AgA0nbNdLgPcdwtKXdyz20dxc0F+XXdB/Tt43YQs2x00/BV6cdUF/XdBfBO87KXjfgaC/LrugH++7JJt+Crw464L+uqC/CN53UvC+A0F/XXZBP953STb9FHhx1gX9dUF/EbzvpOB9B4L+uuyCfrzvkmz6KfDirAv664L+IrPzvg+ufHDh4Gd3oOjKO34+uPDBlQfuYI1I30XL9oH3HQj667IL+vG+S7Lpp8CLsy7orwv6i5yz9xUHKywyqT3e1zZIkH4Ted9UjVnmg04TbkctVowEzcjk8PSuC/rrMkT/0i/GodOKC6zHOPZsi9OynlOoBy/OuqC/Lugvsvz7wTLe1+zUV64sdKkDN3ljM9PjBV2XpLWoHMjRg2ThiEjwA40JNnLwvgNBf112QT/ed0nWcwr14MVZF/TXBf1FztX72o06t6mNVWy4cHCgNvOuvGCG0uMrD4w9dfXtIU9YV5JSE+olH8riXDPcVpovXMZXSTuXmRie3nVBf12G6F/6xTh0WnGB9RjHLo/blV8r6zmFevDirAv664L+Isu/H3R53+AnA35z9vt0Yh/1YTPZHXTlG1ILao5Dhan2nlXHapKttwM2VgeqRq/RUExm+FWlrctMDE/vuqC/LkP0L/1iHDqtuMAg4+i2IIPfrHRy8EcC0zDoFGYML866oL8u6C8iG6iLRqInDt7eS6bSHMT93rhGO9SVt7RaCO3qVhNdVu7eOvC2VUeKtGOJWLP05R0LT++6oL8uQ/Qvv9e5+z6KC7Q2R7OltDH7xIMrfkdRu4veaJppPfkJGbS/zxhenHVBf13QX0T2TheNRE8csje2PGO6fyuHKYd2pCtvabUTzHGsDsUmnRAHQn3rIK6bKLCkC+eYab5C1mvuJ4end13QX5ch+pd+MQ6dVlxgnHEMW1FrT0p2tFJ+SsadwvzgxVkX9NcF/UWWfz/o8r7G+rUxm7PZshPcnr3cJm9G0+NY3dUk0hpoHRj99iBGinThBDNHDctpumhieHrXBf11GaJ/6Rfj0GnFBYYYR7OpRJo9J9l8wo7WlZ+SIacwZ3hx1gX9dUF/EdlqXTQSPbF/b0wsodvt7R6uxux7gDvoyjekFtQcx7eE+I6Q1nla9a0Ds6w9KM5Nk7G8GUomLH15x8LTuy7or8sQ/cvvde6+j+ICrc3R7BZtZL9QW0jcilp7ktrRuvJT0r+/zxtenHVBf13QX0R2XxeNRE/s3Rv17u5RucY0Nlx5IHHYzLvyghlKj+MKrXeE2MXg8q361oHS1ZZtjjS2U6xJx01e/mOGpoend13QX5ch+pd+MQ6dVlyg3zgmO47aTtRmFQ668hPSfwrzhhdnXdBfF/QXWf79YIz33VTkfWbxO0vLNRdY+vKOhad3XdBflyH6l9/r3H0fxQWGbI6N5TU0v7jBbyjW2hoGfyQwDZu+v/PirAv664L+IrKBumgkeuKm743d9Jjf3neepS/vWHh61wX9dRmif/m9zt33UVxgCzbHTT8FXpx1QX9d0F8E77uYFT9XwfsOBP112QX9eN8dhRdnXdBfF/QXwftOCt53IOivyy7ox/suiT2FzT0LXpx1QX9d0F9kFe8rkixhb9y4m5xIknGnBADbBd53ScLm6I43DXnsXbSZoL8u6K/LRPpX8b4uyuzjBt1y8faMNg6e3nVBf12G6D8n7ytSNMn+skE3OZck404JAGCTCdv1EiQTbbdNJNneXXbT2Nzrb0F/XXZB/zl534Rkf9mgWy7entHGwYuzLuivC/qL4H0t6K8L+uuyC/rreF9hcy8u3ncOoL8u6K/LRPrxvhb01wX9ddkF/TP1vv8NHndF1g0vzrqgvy7oL4L3taC/Luivyy7on6/3dVHG/fuHe3tXL1z4z3CTQ0m64e0C79sF+uuC/rpMpL+W9/0/AIDVcLvJsP1nk7zv3//+98uXPxaze/36F48eff/48Q9y+/bbu59//ntJypAUuNJtAe/bBfrrgv66TKS/ovd98eKFOwAAGIPsHlvrfcXX/uY3//XJJx//5S9P3nnEBN+48aUYX0nKkBRsmf3F+3aB/rqgvy4T6a/ofV0EADCerfW+ly9/LO729evXzva+e3fv3ndXr34qwePHP8h/ZUgKpMxN2Arwvl2gvy7or8tE+mfofY/2f/vb/SN3kCKDe3deuQOPS5qJbfbu3CnVL8mrO3ta2EKda6d44gojzlBUZAeToWJyOGF6nzKASdhO73v//qH9cLcxvQZxuhcvfij21x03SIGUbdN3f/G+XaC/Luivy0T6Z+d9xVLt7e/vdbmpotFKkvpwfcbMCFOdenSunZ4TsU5UKBYFn6opJoejpptrcV7XAcCxnd53b+/q9etfOIfbeNxPPjFf/LWf+2qkTIrdNMvPBxc+uHDwszv69dcHV0RT4MoDlx6W9+lSMub8cnlmJHjfLtBfF/TXZSL9slW5aCTJxLHyuryvtVHdZqqa900k9elcO4tPRISI7RUvXja/zfB03rdRt3QjgKWQPeTk5OT09FR84PPnz9+/f+8GOlh+r3P3fXQtsHhzTDyf2NxHj7637tZy+/bXYn8lePnypf0ihP1UWMqk2E1rEOt75coV5T3FkIYj44uVde3LGy9rs7q4IXXYpcx48L5doL8u6K/LRPpn5n29mfT3Duuzku8wFJOGzPsehUpf5Od689Yqk1QYV4vpJfxhK52tG6a6VgXxLQEy3hBbFufmNGUy7O91OmJXKiZbWdtAFJl4z+V9187p3eoApkD2kOPj41evXr158+bp06fb433tl3oDn332uxs3vpSkOOCvvvrD9etfXLr0keQl0/a+xvo+EKsaXWhiW8PhkLzEC7yvM8uOPDMevG8X6K8L+usykf55ed9oobSZMi5M+bZozbKkRflO6+DSGargaL+JkjI/3JribJ6hU2cexxZt8b5SCwhLZVF2ji3iaNPG1TWxnR/DYtKGaiUT2+GmV8h1TLdHPgQ4F2QPOTw8PDo6Evv7008/nZ2duYEONsb76i/7Chcvfmg/CbYf9IYv/qbe11pfZVoz2xosan/epFwoyYBrbEbDQUOeGQvetwv01wX9dZlIv+xXLhpJMnGsvKL3FY8V/J3xW9ZN6azxWN6O5UmHPizF1soFnKdbOMUE0duVdRYndol3azd0CSjOLZC00zY4XkIbF5Muq2mX2krpWp7ujrrkAUyC7CG3bt0S+3t8fPzkyZPt8b76Ow/W4NrY/nzby5cv7WHynYdoeVvRWO/riTY2KQ40xa2RPDMCvG8X6K8L+usykX7Zq1w0kmTiWHkl71u0YImtkprmoJh06MNSbFyb92yOvikmCFMkTrBDpYld4ocIWHSOCtMtfNgbD2RCUGazEheTrQYBVRoKitMbZMSHAOeC7CEHBwd3794V4/vjjz/++9//dgMdLL/Xufs+uhZYvDkmni/5WTf7JQcbS/DJJx/fu/edNcftn3VTttVgrWvbtgaL25+PqbS4hS6z5Jmh4H27QH9d0F+XifTLLumikSQTx8oreF+xTy3/ZcxV46eM41J+yxYVkxbdqBibuXrCgCmyiE/qcUPQGQK9gomK4ls9Wk1DXJyb0gzFsXCo8jY0rYpJHRqO7phxs7jLhVnl6baiyQKcG7KH3Lhx45tvvhHjuz3eN/yOM/vNBzG7L1++vH3760uXPhLXe+PGlxcvfmhHpSz+jjMxqMpy+iNtWyUOBwPyXUa5Bd73PEB/XdBfl4n0z8f7avdnMfbKeivrw4zrUj/yVUwadKeO2Dm3BrNE/5Tg7vSwJdf52/39WOWTHT/rtlBAPlcXGmwjlbJTfK0m0WixyZYePbkZ1muUp0vWRQDnxHZ6X8H+bQvxuFevfmo/4n39+nX4qoP9bQ9SoP+2hfjTluN0x8bXBlTBgny0ucbGmqNWsUmYAY+dnGfGg/ftAv11QX9dJtIve5WLRpJMHCuv8LnvzBED2DKd4zEWclWHOEuTufqlARjN1nrf4t80DkhShvibxsPhvb8u6K8L+ovgfYezonc1n5iuahHnaH3NefGhL5w7W+t9BfG1ly+bP2lx/foXjx59//jxD2J5JZBDScrQlhlfAe/bBfrrgv66TKR/Fe8rkpZmE73vUtjvC1hwiABrQ/aQa9eu3bx58+HDh+J9nz175jaXDjbJ+1ru3z/c27sqZjfc5HCb/o6xBu/bBfrrgv66TKR/Fe/rooax8nbG+wLAJMz0c1/ZCkcxnefbLOQ6uCvSII8uAMAquN0kQzYcvC8AbCKyh2ztdx4sQz73/eB/f9qOW3Id5NF98eKFOwAAGIPsHotdJt4XADaRbfa++fd95fbtt3c///z3ktTf900c5Obecu/rIgCA8eB9AWD72FrvK742/z0PYoJv3PhSjG/yex4SB7m5t+Hed/0/XWs6LvhJZBme4BfZ9Cw6HvOzJWtpN835Goqdp1tuImYieJ7XbUaqNs77yrVjZyvAznZOzETwPK/bjFRtrfe93Px+39evXzvb++7dvXvfXb36qQSPH/8g/9W/3zdxkJt7G+p9zW9U3N9f529VlOf04nec/Ek/6mVQLO5ddCCjlCxA9xnYc4mli1N6+yyx0KTMRM9YGdPJ1p1n9GBtmPdlZ2sxSskCdJ+BPZdYujilt88SC03KTPSMlTGdbN15Rg/Wdnrf+/7vujWm1yBO9+LFD8X+uuMGKZAyKU4c5ObeBnpf8wbR/IXJ9T0Ne5/TecGol0GxeF0vpCn6DOy5xNLFKb19llhoUmaiZ6yM6WTrzjN6sGQPOTk5OT09lQ3z7Ozs/fv3bqBhbt6Xna3NFH0G9lxi6eKU3j5LLDQpM9EzVsZ0snXnGT1Y2+l99/auXr/+hTW4gv2Gg9hc+7mvRsqkuOUgD/7586//OjhQmXj724PCkE3KfzW2rFg/4W2Y9/VvDeW3iOIzVUrN/4IT3KcRrYSUOZrqYgedFOIU1zFdQR9nxYaY7Fn0jitUyy9orvrEqrBsV0Oh1OfIddDKHK7GE8/L0C5r8GKKf9pUJfWpNTWG1kJy4IckVK1stnDWgaY+PX3VJMZN4E9f+oSurVnZ9cnkm7L9fcnpHrkwh1Q7XEWzSv54+U6ti+nQKyZ6YvtmsGnu5oZYT2+S6erFsyh17r84gWQgiNFxq6eUhTmxfwnZQ46Pj1+9evXmzRvZMOftfeWcmrPx922KlyW7pq2ElDma6mIHnRTiFNcxXUEfZ8WGmOxZNH9iL2qu+sSqsGxXQ6HUh52tqZQ+oWtrFjubo9R5NjvbFnpfsbn27xgHbt/+WuyvBPavGUtgPxWWMinW9vHCo389+PHk50fPddLfFnjfPM4PJ78N8r7y1HBPihgpSk8vuW89J1XN0b6vyGa1Yp20dEyxDSXRtWKko0OMJfCvJhPayLw02r065zoJcUaxYSDp4198iyr1lBxdptQvSEouWcqhFgo1cqmFJuufCrFn4To1o2GuikJZiJvKLGzP8tnW5XWFEjaRKYtznbJeQp9mlYJgF6mLGTCjUURBT6gvxsn0bE0JbCYldLCx7xIm6gInxpP2TFrFRnnYxGVBFtlDDg8Pj46OxP6+ffv27OzMDTTMy/v6J7GOFKXLIvddl27EIx6Slo4ptqEkOh+sQEeHGEuQPbUKr9jOuU5C66WXNwwkffInZyBU6ik5ukypX5CUXLKUQy0UauRSs7OFFR1mNIoo6An1xTiZnq0pgc2khA429l3CRF3gxHjSnkmr2CgPm7gsyLK13td+qTfw2We/u3HjS0mKA/7qqz9cv/7FpUsfSV4ybe/7/OAfJ1fEs/7jnxdiUixsw48nbZtbTGbe99GJLVR+Osz9Wyh78OO/bM2VH/XQuNsQ7+v3AhdnT4/S0yt5JTWHEdOhNKsV66RFZfKGyYqF6UJxIR2XkoVz7i+TbHNQqox0japYQofN5E0akjIRE6v8lK5k67oF1EJ+otln5J/LJvYpMz0/64BqEuPhSR3rpDkwCzXiIyalyjpPTSHljnwVH/tzbSUj6YqRRE857ivoPItBEyP6gUl7Flv19e9A9pBbt26J/T0+Pj45OZmz99UPrLkk7edu+ayTS1e4yL3XTSctKpM3TFYsTBeKC+m4lCycc3+ZZJuDUmWka1TFEjpsJm/SkJSJmFjlp3QlW9ctoBbyE9nZ2gWWdMVIoqcc9xV0nsWgiRH9wKQ9i636+newtd5Xf9lXuHjxQ/tJsP2gN3zxN/W+B//8uTGdYkAf/NElJXa29Y/iYp21LSZL3te7WFMmrtrkfXPx2bbYlJlu5usWzZAJXPGo2wDvK0+IhPYm0P3ssU9R/0oeOCvErVYNKlNoaIgrFqYLxYV0XEoOFN8uk2wmQ8eWrlEfx54S2dG8SakslBtUtzzZYBqkj6susDPtP7El3j8KrczM/KwDukmIhyd1rJN+1fbqDbrMYEqyp6wjTpcoX8XHYVAnIyrTo6cY9xYYSmfRN7EgpoXqWWzV178D2UMODg7u3r375MmTN2/eJO8Qc/K+ciIJAx87delWfcQtKtPxqHU8WIHiQjouJQeKb5dJNpOhY0vXqI9jT4nsaN6kVBbKDapbnmwwDRa9cOxMdjaVjKhMj55i3FtgKJ1F38SCmBaqZ7FVX/8Ottb76u88WINrY/vzbS9fvrSHyXceouUVq6o+lM2sbTGZxF1DEkSa5fRQEoy79Xvf9PmQP+9URoqz6uY4G2j1LXZIF25n8oYOtWI+rpO9i4bYVLZ7FctMF9cw9i5WBrpGfRxyUULepKvMKTC5RUmHZNp9WwuZUfX/BPf3w/9pij2bpi706CYhVnVmtk0WK3UcS1UDlXTouQ4jPklZQq1paCM9PcRmla7rlk1JhvXocifuyM6id6IE7RYZvudywnQyot8hZu19U/nqIjiKl8XhL1020Orbe2EtOpM3dKgV83Gd7F00xKay3atYZrq4hrF3sTLQNerjkIsS8iZdZU6ByS1KOiTT7ttayIyys9nCtH8yJRnWo8uduCM7i96JErRbZPieywnTych2et/kZ93slxxsLMEnn3x879531hy3f9at5Ur9x7Tahha9aVfcNZTUFIfymkG3Xu+bPwvUs8kjRRbZOWx5yIRSM81jcu3GeYfCyr7KtkwaZiu2ih3tnmFKedEkdth+7tgcqLIoqatJiB3lPjH2Hc3PDLhRNSVQKPN10kj9EEOe9Jmko9BayKzgm+hYKJx1QHqEpIrDooOufCxofnKimRe0xtVtVs2V0NGU666O/LrpokKrrp8IiZlUj59qFfs+I048TFEnbXEjyVnrOBPjyHuGzMhHxCcjm+J9c/XmaiXXWIqSyxIyoXTBM1DIOxRW9lW2ZdIwW7FV7Gj3DFMGPZoO288dmwNVFiV1NQmxo9wnxr4jO1sTsLNp3Ehy1jrOxDjyniEz8hHxych2et/wO87sNx/E7L58+fL27a8vXfpIXO+NG19evPihHZWy+DvO4me95qa+mZB/vaGYlFviWfVhiNXctCwPxt36P/cF2ApkP0u3WJiATfG+ANsBO9v5sJ3eV7jc/G0L8bhXr35qP+J9/fp1+KqD/W0PUiBlUmyNY/zCg70FK9x8Ddegf6ytmEw9qz5UcZibfrScB+NueF/YDXiDOCfwvgDnCDvbObG13rf4N40DkpShXf6bxgAAveB9AWD72FrvK4ivvXzZ/EmL69e/ePTo+8ePfxDLK4EcSlKGrPEVEge5uTe8LwCsEbwvAGwf2+x9LffvH+7tXRWzG25yKEk33JA4yM295d73xYsX7gAAYAyye+B9AWD72H7vO4TEQW7uLfe+AACrgPcFgC1D72yb7X3B4q5IiZOTk+Pj48PDw1u3bh0cHMgDP5xr1665aDNBf13QX5dV9E/0d91kS18avC8ArILsIbIr3rx58+HDh+J9nz175jaXDmbqfefPHPSfnp6+evXq6OhI3snu3r0r/+IZjjxFXLSZoL8u6K/LKvplu5BNQ7aOt2/fzvlvGnuO7K8ALfxKT6EZTIeKyYEsmGt+a+nwrvP7Af8JFa2r9WQS42MXlsjXWvfqxX4DF1m3lq1npp/7GpsN6+bp06fyGMvb2J///Oc//elP8s8dAIDFyHYhm4ZsHbKB/PLLL243adzqDL1v/DX5JStgR5ORYnIgq8xtGRZptKT7nooJFa3UermLtqw3DEvka637AhX7DVxk3Vq2n5l635yxm+PcmIP+s7Ozd+/evX379uTk5M1I5Mnhos0E/XVBf11W0S/bxenpqWwdsoG8f//e7SYNs/O+YgGMFb1j7gpewA23PVAxOZBV5rYNy/w+tzsKfw947axysstdND1rFGGJfK11P2TFCz7wUZjwwdpS8L7nxBz0y/uWIG9ggjzSo3j27JmLNhP01wX9dVlFv90x7O7hthLP3LyvuJHGiSaW1B56nEkoJpsGFpPxfXza+4t8brKiQVK2PlS3huNSbr7cuZxfRhclLivrmVeGjG2XHHatGBoLyZrlKXIY6kIslftHrpUp9F3tqJQ1f3TXEFeOCmOZu/R6AYfrmuspnEJ7lqORaMNmoSayHeN9rFLlDpk15CwU0sMNqcHhybxzsQwGgfc9J9BfF/TXBf11mUj/vLxv4wSsC2lsQnAv3hnEsJhsJjU5P927DUn2zFVJh2QaAf4+R0b8BDvdzpdFbToErdCQ9swr5U7PUIcSNnPN/Z4V4FONDN+48Elix5RQJ8m4iLv+ZlxoSty4Pdmm0oR2SpxbKFNI1qdsga0I003SC4qnoGZ5JOVVpQLjWB4EJDPkLAKmyLcIwoYn887FMhgK3vecQH9d0F8X9NdlIv2z8r5iCdoYO2AMQtsFS7wgGfDuwrsv291/mJnMNQUpwayYmaUamR3SUqPjZqJd0WI7Rdo980rfw9E6tAdmdTkblWmXJR0MpSkSBB3xhGTUZ8018o1crRoNfeQ+4Mp1WWDARfOrqVjPClg1MmQe1H0buqowMw8CkimtHsjLQybEo5KBBWUwmI3xvgAAMCtm5H2Nx4oGoDkyh41niI7Gxr1Jh0otblggGimDqUzdieR8ShdL3KTbDQqEnnllkmkd2gOd8rHOKXGe0pTkJFzc0ciFpT465yikhPZ6uk+T1jkp9bGaFWm8rzLAsSh08W1jENAr+VjnEvRQEJYkw1KLk5ZiGQxnjt4XAAA2ArdxjySZuLL3FSdgralH7EBjS+OAjVRYTDo/cXRH7mNZ7Fec25rqUJ+HNmT+RCaFCjUY0qZpu0WGm5ZXmoxaTR36UMvxUmOZOdts7f4pflxVqlBqzSR3ZzBz7KjpkyyolwuoyYXWqR5fqmdFTNZ9g0OqBdXZhXngUS0XnoWnKGxcst25VBZz0MvsvC8AAGw36/W+5j0/Oq8GYwhsxkZiD/ZUUTHp2hiMl7CH0VV0z9WVnqZq785R6GknR5RoqQ2jKlZ6WpNjvrsyZLT+1nGQq1b0p2Z+UUb7dPqnHCmjqIWpSTb0MwRfJkSFNquaKFyVGdEFKvbd9SmoWQqT9RVmVhxuepihPAj4dQTVNz0LhZ/QurbDk3nnvMzUpMtCB3hfAAA4V9b9uS8AwAjwvgAAcK7gfQGgInhfAAA4V/C+AFARvC8AAJwreF8AqAjeFwAAzpXVve+LFy/cAQDAGGT3wPsCAMC5srr3BQBYBbwvAACcHyt634STk5Pj4+PDw8Nbt24dHBzIW9o5c+3aNRdtJuivC/qrINuFbBqydTx58uTs7MztJh3gfQEAYCXW631PT09fvXp1dHQk72R379795ty5efOmizYT9NcF/VWQ7UI2Ddk6fvrpJ7wvAABMy3q977t37968eSPvYfYjHPs/Mc+Thw8fumgzQX9d0F8F2S5k05Ct4+nTp+/fv3e7SQd4XwAAWIn1et+zszOxv2/fvj05OZF3svNH3kddtJmgvy7or4JsF6enp7J1PH/+HO8LAADTIm8k4nfXxS8N8gYmPAMAGIDdMezu4baSbvC+AACwEuv93Lc66K8L+uuyC/rxvgAAsBJ431mB/rqgvy5D9ON9AQBgJfC+swL9dUF/XYbox/sCAMBK4H1nBfrrgv66DNGP9wUAgJXA+84K9NcF/XUZoh/vCwAAK4H3nRXorwv66zJEP94XAABWAu87K9BfF/TXZYh+vC8AAKyEvJEAAGwQbvMaCd4XAAAAAHYFvC8AAAAA7Ap4XwAAAADYFfC+AAAAALAr4H0BAAAAYFfA+wIAAADAroD3BQAAAIBdAe8LAAAAALsC3hcAAAAAdgW8LwAAAADsCnhfAAAAANgV8L4AAAAAsCvgfQEAAABgV8D7AgAAAMBu8Ouv/w9+XTh0JS55FwAAAABJRU5ErkJggg=="}}]);