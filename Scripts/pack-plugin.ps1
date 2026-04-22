$version = git describe --always --tags --exclude latest
$pluginPath = "..\Build\Package\QuickLook.Plugin\QuickLook.Plugin.ImageViewer"
$outputPath = "..\Build\QuickLook-ImageViewerPlugin-$version.zip"

if (!(Test-Path $pluginPath)) {
    Write-Output "Plugin directory not found: $pluginPath"
    exit 1
}

# Compress-Archive -Path $pluginPath\* -DestinationPath $outputPath -Force
# 使用 7-Zip 进行打包（假设 7z 已在 PATH 中）
7z.exe a $outputPath $pluginPath\* -t7z -mx=9 -ms=on -m0=lzma2 -mf=BCJ2 -r -y -mmt

Write-Output "Plugin package created: $outputPath"