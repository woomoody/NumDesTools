# NumDesTools简介
游戏数值计算过程中，需要用到的一些小工具
# 工程引用缺失
https://blog.csdn.net/weixin_43650637/article/details/114318408<br> <备注>工程升级到dotnet6之后应该不存在这个问题了<br>
步骤1：将.csproj文件中的类似下列代码删除
~~~ Html
  <Target Name="EnsureNuGetPackageBuildImports" BeforeTargets="PrepareForBuild">
    <PropertyGroup>
      <ErrorText>这台计算机上缺少此项目引用的 NuGet 程序包。使用“NuGet 程序包还原”可下载这些程序包。有关更多信息，请参见 http://go.microsoft.com/fwlink/?LinkID=322105。缺少的文件是 {0}。</ErrorText>
    </PropertyGroup>
    <Error Condition="!Exists('..\packages\ExcelDna.AddIn.1.6.0\build\ExcelDna.AddIn.props')" Text="$([System.String]::Format('$(ErrorText)', '..\packages\ExcelDna.AddIn.1.6.0\build\ExcelDna.AddIn.props'))" />
    <Error Condition="!Exists('..\packages\ExcelDna.AddIn.1.6.0\build\ExcelDna.AddIn.targets')" Text="$([System.String]::Format('$(ErrorText)', '..\packages\ExcelDna.AddIn.1.6.0\build\ExcelDna.AddIn.targets'))" />
    <Error Condition="!Exists('..\packages\ExcelDna.Interop.15.0.1\build\ExcelDna.Interop.targets')" Text="$([System.String]::Format('$(ErrorText)', '..\packages\ExcelDna.Interop.15.0.1\build\ExcelDna.Interop.targets'))" />
    <Error Condition="!Exists('..\packages\SixLabors.ImageSharp.3.0.1\build\SixLabors.ImageSharp.props')" Text="$([System.String]::Format('$(ErrorText)', '..\packages\SixLabors.ImageSharp.3.0.1\build\SixLabors.ImageSharp.props'))" />
  </Target>
~~~
步骤2：工具\Nuget包管理器\程序包管理器控制台<br>
`update-package -reinstall`
# 打包与使用
  packFromBin文件夹中.xll引用到Excel中即可
## 依赖项
本项目使用了以下外部库：
| 库名称 | 版本 | 许可证 |
| ------ | ---- | ------ |
| BouncyCastle.Cryptography | 1.8.9 | MIT |
| Enums.NET | 5.0.0 | MIT |
| EPPlus | 7.0.4 | Polyform Noncommercial |
| EPPlus.Interfaces | 6.1.1 | Polyform Noncommercial |
| ExcelDna.AddIn | 1.8.0 | MIT |
| ExcelDna.Integration | 1.8.0 | MIT |
| ExcelDna.IntelliSense | 1.8.0 | MIT |
| ExcelDna.Interop | 15.0.1 | MIT |
| GraphX | 3.0.0 | MIT |
| KeraLua | 1.4.1 | MIT |
| MathNet.Numerics.Signed | 4.15.0 | MIT |
| Microsoft.CSharp | 4.7.0 | MIT |
| Microsoft.IO.RecyclableMemoryStream | 2.1.3 | MIT |
| NLua | 1.7.2 | MIT |
| NPOI | 2.6.2 | Apache 2.0 |
| SharpZipLib | 1.4.2 | GPL |
| SixLabors.Fonts | 2.0.3 | Apache 2.0 |
| SixLabors.ImageSharp | 3.1.0 | Apache 2.0 |
| stdole | 17.9.37000 | MIT |
| System.Configuration.ConfigurationManager | 4.7.0 | MIT |
| System.Data.OleDb | 6.0.0 | MIT |
| System.Runtime.CompilerServices.Unsafe | 6.0.0 | MIT |
| System.Runtime.Handles | 4.3.0 | MIT |
## 许可证
本项目采用 [Creative Commons Attribution-NonCommercial 4.0 International License](https://creativecommons.org/licenses/by-nc/4.0/deed.zh) 进行许可。
