# NumDesTools
游戏数值计算过程中，需要用到的一些小工具
# clone代码后引用缺失问题
https://blog.csdn.net/weixin_43650637/article/details/114318408
步骤1：将.csproj文件中的类似下列代码删除
  ~~~
  <Target Name="EnsureNuGetPackageBuildImports" BeforeTargets="PrepareForBuild">
    <PropertyGroup>
      <ErrorText>这台计算机上缺少此项目引用的 NuGet 程序包。使用“NuGet 程序包还原”可下载这些程序包。有关更多信息，请参见 http://go.microsoft.com/fwlink/?LinkID=322105。缺少的文件是 {0}。</ErrorText>
    </PropertyGroup>
    <Error Condition="!Exists('..\packages\ExcelDna.AddIn.1.6.0\build\ExcelDna.AddIn.props')" Text="$([System.String]::Format('$(ErrorText)', '..\packages\ExcelDna.AddIn.1.6.0\build\ExcelDna.AddIn.props'))" />
    <Error Condition="!Exists('..\packages\ExcelDna.AddIn.1.6.0\build\ExcelDna.AddIn.targets')" Text="$([System.String]::Format('$(ErrorText)', '..\packages\ExcelDna.AddIn.1.6.0\build\ExcelDna.AddIn.targets'))" />
    <Error Condition="!Exists('..\packages\ExcelDna.Interop.15.0.1\build\ExcelDna.Interop.targets')" Text="$([System.String]::Format('$(ErrorText)', '..\packages\ExcelDna.Interop.15.0.1\build\ExcelDna.Interop.targets'))" />
    <Error Condition="!Exists('..\packages\SixLabors.ImageSharp.3.0.1\build\SixLabors.ImageSharp.props')" Text="$([System.String]::Format('$(ErrorText)', '..\packages\SixLabors.ImageSharp.3.0.1\build\SixLabors.ImageSharp.props'))" />
  </Target>
  ~~~<br>
步骤2：工具\Nuget包管理器\程序包管理器控制台
    update-package -reinstall
# 打包与使用
  pack文件夹中点击packtool会生成两个.XLL的Pack，与XllConfig文件一起放入到任意路径中，Excel中引用适合自己电脑位数的.XLL即可