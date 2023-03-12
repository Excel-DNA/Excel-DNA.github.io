"use strict";(self.webpackChunkexcel_dna=self.webpackChunkexcel_dna||[]).push([[4769],{3905:(e,n,t)=>{t.d(n,{Zo:()=>c,kt:()=>g});var o=t(7294);function r(e,n,t){return n in e?Object.defineProperty(e,n,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[n]=t,e}function i(e,n){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);n&&(o=o.filter((function(n){return Object.getOwnPropertyDescriptor(e,n).enumerable}))),t.push.apply(t,o)}return t}function a(e){for(var n=1;n<arguments.length;n++){var t=null!=arguments[n]?arguments[n]:{};n%2?i(Object(t),!0).forEach((function(n){r(e,n,t[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):i(Object(t)).forEach((function(n){Object.defineProperty(e,n,Object.getOwnPropertyDescriptor(t,n))}))}return e}function s(e,n){if(null==e)return{};var t,o,r=function(e,n){if(null==e)return{};var t,o,r={},i=Object.keys(e);for(o=0;o<i.length;o++)t=i[o],n.indexOf(t)>=0||(r[t]=e[t]);return r}(e,n);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(o=0;o<i.length;o++)t=i[o],n.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(r[t]=e[t])}return r}var l=o.createContext({}),d=function(e){var n=o.useContext(l),t=n;return e&&(t="function"==typeof e?e(n):a(a({},n),e)),t},c=function(e){var n=d(e.components);return o.createElement(l.Provider,{value:n},e.children)},p={inlineCode:"code",wrapper:function(e){var n=e.children;return o.createElement(o.Fragment,{},n)}},u=o.forwardRef((function(e,n){var t=e.components,r=e.mdxType,i=e.originalType,l=e.parentName,c=s(e,["components","mdxType","originalType","parentName"]),u=d(t),g=r,f=u["".concat(l,".").concat(g)]||u[g]||p[g]||i;return t?o.createElement(f,a(a({ref:n},c),{},{components:t})):o.createElement(f,a({ref:n},c))}));function g(e,n){var t=arguments,r=n&&n.mdxType;if("string"==typeof e||r){var i=t.length,a=new Array(i);a[0]=u;var s={};for(var l in n)hasOwnProperty.call(n,l)&&(s[l]=n[l]);s.originalType=e,s.mdxType="string"==typeof e?e:r,a[1]=s;for(var d=2;d<i;d++)a[d]=t[d];return o.createElement.apply(null,a)}return o.createElement.apply(null,t)}u.displayName="MDXCreateElement"},5478:(e,n,t)=>{t.r(n),t.d(n,{assets:()=>l,contentTitle:()=>a,default:()=>p,frontMatter:()=>i,metadata:()=>s,toc:()=>d});var o=t(7462),r=(t(7294),t(3905));const i={title:"Checking and Downloading Updates in .NET"},a=void 0,s={unversionedId:"guides-advanced/checking-and-downloading-updates-in-dotnet",id:"guides-advanced/checking-and-downloading-updates-in-dotnet",title:"Checking and Downloading Updates in .NET",description:"Following is a simple method to check for available updates of Add-ins/Programs:",source:"@site/docs/guides-advanced/checking-and-downloading-updates-in-dotnet.md",sourceDirName:"guides-advanced",slug:"/guides-advanced/checking-and-downloading-updates-in-dotnet",permalink:"/docs/guides-advanced/checking-and-downloading-updates-in-dotnet",draft:!1,tags:[],version:"current",frontMatter:{title:"Checking and Downloading Updates in .NET"},sidebar:"tutorialSidebar",previous:{title:"Building Excel-DNA From Source",permalink:"/docs/guides-advanced/building-excedna-from-source"},next:{title:"Configuring NLog Logging",permalink:"/docs/guides-advanced/configuring-nlog-logging"}},l={},d=[],c={toc:d};function p(e){let{components:n,...t}=e;return(0,r.kt)("wrapper",(0,o.Z)({},c,t,{components:n,mdxType:"MDXLayout"}),(0,r.kt)("p",null,"Following is a simple method to check for available updates of Add-ins/Programs:"),(0,r.kt)("p",null,"The procedure can be called on start up of Excel or on displaying an About Dialog-box (as I decided in my case being the less intrusive variant). The parameter ",(0,r.kt)("inlineCode",{parentName:"p"},"doUpdate")," decides whether only the check for the new version is performed or the new version is actually downloaded.\nThe update check/download requires a continuously increasing version number being available via a URL (here on githubs tag/release archive) of the form ",(0,r.kt)("inlineCode",{parentName:"p"},"https://domain.name/path/1.0.0.<release>"),".\nA local update folder can also be provided to allow for a central update by an administrator."),(0,r.kt)("p",null,"The UserMsg and QuestionMsg are just wrappers of ",(0,r.kt)("inlineCode",{parentName:"p"},"MsgBox")," providing a headless mode (without pop-ups) and logging so you can replace them with your choice of ",(0,r.kt)("inlineCode",{parentName:"p"},"MsgBox(theMessage, msgboxIcon + questionType, questionTitle)"),"."),(0,r.kt)("p",null,"First, the necessary settings are fetched, you can also hard-code the default values (being the second argument of ",(0,r.kt)("inlineCode",{parentName:"p"},"fetchSetting"),", described in more detail in ",(0,r.kt)("a",{parentName:"p",href:"/docs/guides-advanced/user-settings-and-the-xllconfig-file"},"User settings and the .xll.config file"),")"),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vbnet"},'    Public Sub checkForUpdate(doUpdate As Boolean)\n        Const updateFilenameZip = "downloadedVersion.zip"\n        Dim localUpdateFolder As String = fetchSetting("localUpdateFolder", "")\n        Dim localUpdateMessage As String = fetchSetting("localUpdateMessage", "A new version is available in the local update folder, after quitting Excel (is done next) start deployAddin.cmd to install it.")\n        Dim updatesMajorVersion As String = fetchSetting("updatesMajorVersion", "1.0.0.")\n        Dim updatesDownloadFolder As String = fetchSetting("updatesDownloadFolder", "C:\\temp\\")\n\n        \' put your UrlBase here, where the release zip files can be found\n        Dim updatesUrlBase As String = fetchSetting("updatesUrlBase", "https://github.com/rkapl123/DBAddin/archive/refs/tags/")\n        Dim response As Net.HttpWebResponse = Nothing\n        Dim urlFile As String = ""\n')),(0,r.kt)("p",null,"Then, the procedure checks check for the zip file of the next higher revisions until no higher version can be found. ",(0,r.kt)("inlineCode",{parentName:"p"},"Net.SecurityProtocolType.Tls12")," is only available starting with .NET 4.5"),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vbnet"},'        Dim curRevision As Integer = My.Application.Info.Version.Revision\n        \' try with highest possible Security protocol\n        Try\n            Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12 Or Net.SecurityProtocolType.SystemDefault\n        Catch ex As Exception\n            UserMsg("Error setting the SecurityProtocol: " + ex.Message())\n            Exit Sub\n        End Try\n        \' always accept url certificate as valid\n        Net.ServicePointManager.ServerCertificateValidationCallback = AddressOf ValidationCallbackHandler\n\n        Do\n            urlFile = updatesUrlBase + updatesMajorVersion + (curRevision + 1).ToString() + ".zip"\n            Dim request As Net.HttpWebRequest\n            Try\n                request = Net.WebRequest.Create(urlFile)\n                response = Nothing\n                request.Method = "HEAD"\n                response = request.GetResponse()\n            Catch ex As Exception\n            End Try\n            If response IsNot Nothing Then\n                curRevision += 1\n                response.Close()\n            End If\n        Loop Until response Is Nothing\n')),(0,r.kt)("p",null,"get out, if no newer version could be found. In my case, I set a TextBox (",(0,r.kt)("inlineCode",{parentName:"p"},"TextBoxDescription"),") and a button (",(0,r.kt)("inlineCode",{parentName:"p"},"CheckForUpdates"),") to notify the user. for the notification step (",(0,r.kt)("inlineCode",{parentName:"p"},"doUpdate = False"),") stop here and let the user decide by clicking on the button ",(0,r.kt)("inlineCode",{parentName:"p"},"CheckForUpdates"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vbnet"},'        If curRevision = My.Application.Info.Version.Revision Then\n            Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "You have the latest version (" + updatesMajorVersion + curRevision.ToString() + ")."\n            Me.TextBoxDescription.BackColor = Drawing.Color.FromKnownColor(Drawing.KnownColor.Control)\n            Me.CheckForUpdates.Text = "no Update ..."\n            Me.CheckForUpdates.Enabled = False\n            Me.Refresh()\n            Exit Sub\n        Else\n            Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "A new version (" + updatesMajorVersion + curRevision.ToString() + ") is available " + IIf(localUpdateFolder <> "", "in " + localUpdateFolder, "on github")\n            Me.TextBoxDescription.BackColor = Drawing.Color.DarkOrange\n            Me.CheckForUpdates.Text = "get Update ..."\n            Me.CheckForUpdates.Enabled = True\n            Me.Refresh()\n            If Not doUpdate Then Exit Sub\n        End If\n')),(0,r.kt)("p",null,"If there is a maintained local update folder, open it and let user update from there."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vbnet"},'        If localUpdateFolder <> "" Then\n            Try\n                If QuestionMsg(localUpdateMessage, MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then\n                    System.Diagnostics.Process.Start("explorer.exe", localUpdateFolder)\n                    Me.quitExcelAfterwards = True\n                    Me.Close()\n                End If\n            Catch ex As Exception\n                UserMsg("Error when opening local update folder: " + ex.Message())\n            End Try\n            Exit Sub\n        End If\n')),(0,r.kt)("p",null,"Otherwise continue and download newest version. Progress information is put into the TextBoxDescription"),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vbnet"},'        urlFile = updatesUrlBase + updatesMajorVersion + curRevision.ToString() + ".zip"\n\n        \' create the download folder\n        Try\n            IO.Directory.CreateDirectory(updatesDownloadFolder)\n        Catch ex As Exception\n            UserMsg("Couldn\'t create file download folder (" + updatesDownloadFolder + "): " + ex.Message())\n            Exit Sub\n        End Try\n\n        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "Downloading new version from " + urlFile\n        Me.Refresh()\n        \' get the new version zip-file\n        Dim requestGet As Net.HttpWebRequest = Net.WebRequest.Create(urlFile)\n        requestGet.Method = "GET"\n        Try\n            response = requestGet.GetResponse()\n        Catch ex As Exception\n            UserMsg("Error when downloading new version: " + ex.Message())\n            Exit Sub\n        End Try\n        \' save the version as zip file\n        If response IsNot Nothing Then\n            Dim receiveStream As Stream = response.GetResponseStream()\n            Using downloadFile As IO.FileStream = File.Create(updatesDownloadFolder + updateFilenameZip)\n                receiveStream.CopyTo(downloadFile)\n            End Using\n        End If\n        response.Close()\n        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "Extracting " + urlFile + " to " + updatesDownloadFolder\n        Me.Refresh()\n        \' now extract the downloaded file and open the Distribution folder, first remove any existing folder...\n        Try\n            Directory.Delete(updatesDownloadFolder + "DBAddin-" + updatesMajorVersion + curRevision.ToString(), True)\n        Catch ex As Exception : End Try\n        Try\n            IO.Compression.ZipFile.ExtractToDirectory(updatesDownloadFolder + updateFilenameZip, updatesDownloadFolder)\n        Catch ex As Exception\n            UserMsg("Error when extracting new version: " + ex.Message())\n        End Try\n        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "New version in " + updatesDownloadFolder + "DBAddin-" + updatesMajorVersion + curRevision.ToString() + "\\Distribution, start deployAddin.cmd to install the new Version."\n        Me.Refresh()\n')),(0,r.kt)("p",null,"Finally, open the windows explorer to let the user start the update process. After leaving the procedure, a hint to close Excel (or do this automatically) is probably a good idea (the Addin won't be deployed as long Excel is open)..."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vbnet"},'        Try\n            System.Diagnostics.Process.Start("explorer.exe", updatesDownloadFolder + "DBAddin-" + updatesMajorVersion + curRevision.ToString() + "\\Distribution")\n        Catch ex As Exception\n            UserMsg("Error when opening Distribution folder of new version: " + ex.Message())\n        End Try\n    End Sub\n')),(0,r.kt)("p",null,"The ValidationCallbackHandler just returns true, if some more checks are needed put them here."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vbnet"},"    Private Function ValidationCallbackHandler() As Boolean\n        Return True\n    End Function\n")),(0,r.kt)("p",null,"For information, the fetchSetting function is shown here, for more details see ",(0,r.kt)("a",{parentName:"p",href:"/docs/guides-advanced/user-settings-and-the-xllconfig-file"},"User settings and the .xll.config file")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vbnet"},'    Public Function fetchSetting(Key As String, defaultValue As String) As String\n        Dim UserSettings As Collections.Specialized.NameValueCollection = Nothing\n        Dim AddinAppSettings As Collections.Specialized.NameValueCollection = Nothing\n        Try : UserSettings = ConfigurationManager.GetSection("UserSettings") : Catch ex As Exception : End Try\n        Try : AddinAppSettings = ConfigurationManager.AppSettings : Catch ex As Exception : End Try\n        \' user specific settings are in UserSettings section in separate file\n        If UserSettings(Key) Is Nothing Then\n            If AddinAppSettings IsNot Nothing Then\n                fetchSetting = AddinAppSettings(Key)\n            Else\n                fetchSetting = Nothing\n            End If\n        Else\n            fetchSetting = UserSettings(Key)\n        End If\n        If fetchSetting Is Nothing Then fetchSetting = defaultValue\n    End Function\n')))}p.isMDXComponent=!0}}]);