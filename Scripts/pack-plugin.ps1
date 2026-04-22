$version = git describe --always --tags --exclude latest
$pluginPath = "..\Build\Package\QuickLook.Plugin\QuickLook.Plugin.ImageViewer"
$outputPath = "..\Build\QuickLook-ImageViewerPlugin-$version.zip"

if (!(Test-Path $pluginPath)) {
    Write-Output "Plugin directory not found: $pluginPath"
    exit 1
}

Compress-Archive -Path $pluginPath\* -DestinationPath $outputPath -Force
Write-Output "Plugin package created: $outputPath"