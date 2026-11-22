param(
    [Parameter(Mandatory=$true)][string]$InputPptx,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [int]$PreviewWidth=1200
)
$ErrorActionPreference='Stop'
$taskSource=(Resolve-Path -LiteralPath $InputPptx).Path
$taskOutput=[System.IO.Path]::GetFullPath($OutputDirectory)
[void][System.IO.Directory]::CreateDirectory($taskOutput)
$taskCopy=Join-Path $taskOutput 'editability-test-copy.pptx'
$taskPreview=Join-Path $taskOutput 'preview.png'
if((Test-Path -LiteralPath $taskCopy) -or (Test-Path -LiteralPath $taskPreview)){throw 'Use a fresh output directory; existing QA files will not be overwritten.'}
$app=$null; $pres=$null
try {
    $app=New-Object -ComObject PowerPoint.Application
    $pres=$app.Presentations.Open($taskSource,0,0,0)
    $slide=$pres.Slides.Item(1)
    $height=[int][Math]::Round($PreviewWidth*$pres.PageSetup.SlideHeight/$pres.PageSetup.SlideWidth)
    $slide.Export($taskPreview,'PNG',$PreviewWidth,$height)
    $text=$null; $line=$null; $textCount=0
    foreach($shape in $slide.Shapes){
        if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1){
            $textCount++
            if($null -eq $text){$text=$shape}
        }
        # msoFreeform = 5. Do not confuse an image with native vector geometry.
        if($null -eq $line -and $shape.Type -eq 5){$line=$shape}
    }
    if($null -eq $text -or $null -eq $line){throw 'The test needs native text and a native freeform.'}
    $textName=$text.Name; $lineName=$line.Name
    $text.TextFrame.TextRange.Font.Size=18
    $line.Line.Weight=1.75
    $pres.SaveAs($taskCopy,24)
    $pres.Saved=-1; $pres.Close(); $pres=$null
    $pres=$app.Presentations.Open($taskCopy,0,0,0)
    $savedText=$pres.Slides.Item(1).Shapes.Item($textName)
    $savedLine=$pres.Slides.Item(1).Shapes.Item($lineName)
    if([Math]::Abs($savedText.TextFrame.TextRange.Font.Size-18) -gt .01){throw 'Font change was not preserved'}
    if([Math]::Abs($savedLine.Line.Weight-1.75) -gt .01){throw 'Line change was not preserved'}
    $report=[pscustomobject]@{PowerPointVersion=$app.Version;Textboxes=$textCount;FontEditSaved=$true;LineEditSaved=$true;Preview=$taskPreview;VisualInspection='Pending human or agent image review'}
    $report | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $taskOutput 'powerpoint-check.json') -Encoding UTF8
    $report | ConvertTo-Json
} finally {
    if($null -ne $pres){$pres.Saved=-1;$pres.Close()}
    # Never quit an application that may contain the user's unrelated presentations.
    if($null -ne $app){[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
}
