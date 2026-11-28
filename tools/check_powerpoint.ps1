param(
    [Parameter(Mandatory=$true)][string]$InputPptx,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [int]$PreviewWidth=0,
    [int]$Dpi=600
)
$ErrorActionPreference='Stop'
function Get-Leaves($items) {
    foreach($shape in $items){
        if($shape.Type -eq 6){Get-Leaves $shape.GroupItems}
        else{Write-Output -NoEnumerate $shape}
    }
}
function Get-NamedLeaf($slide,$name){
    foreach($s in (Get-Leaves $slide.Shapes)){if($s.Name -eq $name){return $s}}
    throw "Missing object: $name"
}
$taskSource=(Resolve-Path -LiteralPath $InputPptx).Path
$taskOutput=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $taskOutput){throw 'Use a fresh output directory'}
[void][IO.Directory]::CreateDirectory($taskOutput)
$taskCopy=Join-Path $taskOutput 'editability-test-copy.pptx'
$app=$null;$pres=$null
try {
    $app=New-Object -ComObject PowerPoint.Application
    $pres=$app.Presentations.Open($taskSource,0,0,0)
    $slide=$pres.Slides.Item(1)
    if($PreviewWidth -le 0){$PreviewWidth=[int][Math]::Round($pres.PageSetup.SlideWidth/72*$Dpi)}
    $height=[int][Math]::Round($PreviewWidth*$pres.PageSetup.SlideHeight/$pres.PageSetup.SlideWidth)
    $slide.Export((Join-Path $taskOutput 'preview.png'),'PNG',$PreviewWidth,$height)
    $leaves=@(Get-Leaves $slide.Shapes);$leafCount=$leaves.Count
    $text=$null;$line=$null;$textCount=0
    foreach($s in $leaves){
        if($s.HasTextFrame -eq -1 -and $s.TextFrame.HasText -eq -1){$textCount++;if($null -eq $text){$text=$s}}
        if($null -eq $line -and $s.Type -eq 5 -and $s.Line.Visible -ne 0){$line=$s}
    }
    if($null -eq $text -or $null -eq $line){throw 'Test needs native text and freeform geometry'}
    $textName=$text.Name;$lineName=$line.Name
    if($slide.Shapes.Count -gt 1){$g=$slide.Shapes.Range().Group()}
    elseif($slide.Shapes.Item(1).Type -eq 6){$g=$slide.Shapes.Item(1)}
    else{throw 'Test needs at least two groupable objects'}
    $ungrouped=$g.Ungroup()
    if($ungrouped.Count -lt 2){throw 'Group needs at least two direct children'}
    $g=$ungrouped.Group()
    # Also test selecting all individual objects, not just grouping panel groups.
    do {
        $nested=$null
        foreach($s in $slide.Shapes){if($s.Type -eq 6){$nested=$s;break}}
        if($null -ne $nested){$null=$nested.Ungroup()}
    } while($null -ne $nested)
    if($slide.Shapes.Count -ne $leafCount){throw 'Ungrouping lost leaf objects'}
    $g=$slide.Shapes.Range().Group();$g.Name='QA-regrouped-figure'
    if($g.GroupItems.Count -ne $leafCount){throw 'Not all leaf objects were grouped'}
    if(@(Get-Leaves $slide.Shapes).Count -ne $leafCount){throw 'Grouping changed leaf count'}
    $slide.Export((Join-Path $taskOutput 'regrouped.png'),'PNG',$PreviewWidth,$height)
    $beforeWidth=$g.Width;$beforeHeight=$g.Height
    $g.ScaleWidth(.85,0,0);$g.ScaleHeight(.85,0,0)
    $g.Left=$g.Left+12;$g.Top=$g.Top+8
    if([Math]::Abs($g.Width-$beforeWidth*.85) -gt .1 -or [Math]::Abs($g.Height-$beforeHeight*.85) -gt .1){throw 'Group scaling failed'}
    $expected=@($g.Left,$g.Top,$g.Width,$g.Height)
    (Get-NamedLeaf $slide $textName).TextFrame.TextRange.Font.Size=18
    (Get-NamedLeaf $slide $lineName).Line.Weight=1.75
    $pres.SaveAs($taskCopy,24);$pres.Saved=-1;$pres.Close();$pres=$null
    $pres=$app.Presentations.Open($taskCopy,0,0,0);$slide=$pres.Slides.Item(1)
    if($slide.Shapes.Count -ne 1 -or $slide.Shapes.Item(1).Type -ne 6){throw 'Group was not preserved'}
    $g=$slide.Shapes.Item(1);$actual=@($g.Left,$g.Top,$g.Width,$g.Height)
    for($j=0;$j -lt 4;$j++){if([Math]::Abs($actual[$j]-$expected[$j]) -gt .1){throw 'Group transform not preserved'}}
    if([Math]::Abs((Get-NamedLeaf $slide $textName).TextFrame.TextRange.Font.Size-18) -gt .01){throw 'Font edit not preserved'}
    if([Math]::Abs((Get-NamedLeaf $slide $lineName).Line.Weight-1.75) -gt .01){throw 'Line edit not preserved'}
    if(@(Get-Leaves $slide.Shapes).Count -ne $leafCount){throw 'Objects lost on save'}
    $slide.Export((Join-Path $taskOutput 'edited.png'),'PNG',$PreviewWidth,$height)
    # Confirm ungroup/regroup still works after reopening the saved file.
    $r=$g.Ungroup();$null=$r.Group()
    $build=Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name VersionToReport -ErrorAction SilentlyContinue
    $report=[pscustomobject]@{PowerPointVersion=$app.Version;OfficeBuild=$build.VersionToReport;LeafObjects=$leafCount;Textboxes=$textCount;GroupUngroupRegroup=$true;AllLeafObjectsGrouped=$true;GroupMoveResizeSaved=$true;FontEditSaved=$true;LineEditSaved=$true;ReopenedRegroup=$true;PixelWidth=$PreviewWidth;PixelHeight=$height;EffectiveExportDpi=$PreviewWidth/$pres.PageSetup.SlideWidth*72;VisualInspection='Pending image review'}
    $report | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $taskOutput 'powerpoint-check.json') -Encoding UTF8
    $report | ConvertTo-Json
} finally {
    if($null -ne $pres){$pres.Saved=-1;$pres.Close()}
    if($null -ne $app){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
}
