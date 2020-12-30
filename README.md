# CSharp_in_Depth

## 第1章 C#开发的进化史

> Microsoft.Office.Interop.Excel with .Net Core:
> 1. First, Remove 'Microsoft.Office.Interop.Excel' nuget package, and then add COM reference to 'Microsoft Excel 16.0 Object Library' (right click the project -> 'Add' -> 'Reference' -> search and add 'Microsoft Excel 16.0 Object Library' ).
> 2. the reference to 'Interop.Microsoft.Office.Interop.Excel' appears in the project 'Dependencies' under 'COM'
> 3. Click 'Interop.Microsoft.Office.Interop.Excel' and then set both "Copy Local" and "Embed Interop Types" to "Yes" in 'Properties' window.
> Refer to: https://social.msdn.microsoft.com/Forums/en-US/690930f1-7856-4f5f-b073-6cf2c40baa19/microsoftofficeinteropexcel-with-net-core

> Note: .NET Core version of MSBuild not support COM reference. You should use the .NET Framework version of MSBuild.