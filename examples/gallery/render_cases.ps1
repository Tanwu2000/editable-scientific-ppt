param([Parameter(Mandatory=$true)][string]$Artifacts,[switch]$NativeOnly)
$ErrorActionPreference='Stop'
$folder=(Resolve-Path -LiteralPath $Artifacts).Path
$cases=Get-Content -Raw -LiteralPath (Join-Path $folder 'cases.json') | ConvertFrom-Json
$app=$null;$pres=$null;$results=@()
try {
    $app=New-Object -ComObject PowerPoint.Application
    foreach($case in $cases){
        $dir=Join-Path $folder $case.case
        if(-not $NativeOnly -and (Test-Path -LiteralPath (Join-Path $dir 'svg-converted.pptx'))){throw 'Use a fresh output folder'}
        $pw=1620;$ph=[int][Math]::Round($pw*$case.height/$case.width)
        if(-not $NativeOnly){
        # A window is required for the Office SVGEdit command to target our selection.
        $pres=$app.Presentations.Add(-1)
        $pres.PageSetup.SlideWidth=$case.width;$pres.PageSetup.SlideHeight=$case.height
        $slide=$pres.Slides.Add(1,12)
        $picture=$slide.Shapes.AddPicture((Join-Path $dir $case.svg),0,-1,0,0,$case.width,$case.height)
        $slide.Export((Join-Path $dir 'svg-inserted.png'),'PNG',$pw,$ph)
        $pres.Windows.Item(1).Activate();$pres.Windows.Item(1).View.GotoSlide(1);$picture.Select()
        $app.CommandBars.ExecuteMso('SVGEdit')
        $slide.Export((Join-Path $dir 'svg-converted.png'),'PNG',$pw,$ph)
        $pres.SaveAs((Join-Path $dir 'svg-converted.pptx'),24)
        $pres.Saved=-1;$pres.Close();$pres=$null
        }
        $pres=$app.Presentations.Open((Join-Path $dir $case.native),0,0,0)
        $slide=$pres.Slides.Item(1)
        $slide.Export((Join-Path $dir 'native.png'),'PNG',$pw,$ph)
        $title=$null;$curve=$null;$texts=0
        foreach($s in $slide.Shapes){
            if($s.HasTextFrame -eq -1 -and $s.TextFrame.HasText -eq -1){
                $texts++
                if($null -eq $title -and $s.TextFrame.TextRange.Font.Size -ge 11.9){$title=$s}
            }
            if($null -eq $curve -and $s.Type -eq 5 -and $s.Line.Weight -gt .6 -and $s.Width -gt 40 -and $s.Line.ForeColor.RGB -ne 0){$curve=$s}
        }
        if($null -eq $title -or $null -eq $curve){throw 'Native text or curve not found'}
        $titleName=$title.Name;$curveName=$curve.Name
        $title.TextFrame.TextRange.Font.Size=14;$curve.Line.Weight=2.5
        $edit=Join-Path $dir 'native-edited.pptx';$pres.SaveAs($edit,24)
        $pres.Saved=-1;$pres.Close();$pres=$null
        $pres=$app.Presentations.Open($edit,0,0,0);$slide=$pres.Slides.Item(1)
        if([Math]::Abs($slide.Shapes.Item($titleName).TextFrame.TextRange.Font.Size-14) -gt .01){throw 'Text edit failed'}
        if([Math]::Abs($slide.Shapes.Item($curveName).Line.Weight-2.5) -gt .01){throw 'Line edit failed'}
        $slide.Export((Join-Path $dir 'native-edited.png'),'PNG',$pw,$ph)
        $results += [pscustomobject]@{case=$case.case;nativeTextboxes=$texts;fontEditSaved=$true;lineEditSaved=$true;editedText=$titleName;editedLine=$curveName;svgConversion='SVGEdit'}
        $pres.Saved=-1;$pres.Close();$pres=$null
    }
    $build=Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name VersionToReport -ErrorAction SilentlyContinue
    [pscustomobject]@{powerPointVersion=$app.Version;officeBuild=$build.VersionToReport;cases=$results} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $folder 'powerpoint-report.json') -Encoding UTF8
} finally {
    if($null -ne $pres){$pres.Saved=-1;$pres.Close()}
    if($null -ne $app){[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
}
