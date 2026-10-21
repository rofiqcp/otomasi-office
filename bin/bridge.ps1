# OTOMASI Office COM bridge (Windows PowerShell 5.1, interactive session).
# No dynamic user scripts. All actions are explicit allowlisted operations.
param([Parameter(Mandatory=$true)][string]$Payload)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
function Fail($message) { throw [System.Exception]::new([string]$message) }
function ColorValue($hex) {
  if([string]$hex -cnotmatch '^#[0-9A-Fa-f]{6}$'){Fail 'Color must be #RRGGBB'}
  $red=[Convert]::ToInt32($hex.Substring(1,2),16)
  $green=[Convert]::ToInt32($hex.Substring(3,2),16)
  $blue=[Convert]::ToInt32($hex.Substring(5,2),16)
  return [int]($red+256*$green+65536*$blue)
}
function IdFor($kind,$doc) {
  $key=[string]$kind+'|'+[string]$doc.FullName
  $bytes=[Text.Encoding]::UTF8.GetBytes($key)
  $sha=[Security.Cryptography.SHA256]::Create()
  try { $hash=$sha.ComputeHash($bytes) } finally {$sha.Dispose()}
  return ([BitConverter]::ToString($hash).Replace('-','').ToLowerInvariant().Substring(0,24))
}
function Applications {
  $out=@{}
  foreach($pair in @(@('word','Word.Application'),@('excel','Excel.Application'),@('powerpoint','PowerPoint.Application'))) {
    try {$out[$pair[0]]=[Runtime.InteropServices.Marshal]::GetActiveObject($pair[1])} catch {}
  }
  return $out
}
function Items($apps) {
  $out=New-Object Collections.Generic.List[object]
  foreach($kind in @('word','excel','powerpoint')) {
    if (-not $apps.ContainsKey($kind)) {continue}
    $app=$apps[$kind]
    if ($kind -eq 'word') {$collection=$app.Documents}
    elseif ($kind -eq 'excel') {$collection=$app.Workbooks}
    else {$collection=$app.Presentations}
    for($i=1;$i -le $collection.Count;$i++) {
      $doc=$collection.Item($i)
      $readonly=$false
      try {$readonly=[bool]$doc.ReadOnly} catch {}
      $full=''
      try {$full=[string]$doc.FullName} catch {}
      $out.Add(@{document_id=(IdFor $kind $doc);type=$kind;name=[string]$doc.Name;path=$full;read_only=$readonly;item=$doc})
    }
  }
  return $out
}
function GetTarget($docs,$doc_id) {
  if (-not $doc_id -or $doc_id.Length -gt 80) {Fail 'document_id required'}
  $found=@($docs | Where-Object {$_.document_id -ceq $doc_id})
  if ($found.Count -ne 1) {Fail 'Document not open or ambiguous; refresh office_list_documents'}
  return $found[0]
}
function RequireWrite($d) {
  if ($d.read_only) {Fail 'Document is read-only. Open an editable copy first.'}
}
function ReadWord($doc,$limit) {
  $out=@()
  $count=[int]$doc.Paragraphs.Count
  for($i=1;$i -le [Math]::Min($count,$limit);$i++) {
    $p=$doc.Paragraphs.Item($i)
    $style=try {[string]$p.Style.NameLocal} catch {'Normal'}
    $isList=$false
    try {$isList=([int]$p.Range.ListFormat.ListType -ne 0)} catch {}
    $out+=@{index=$i;style=$style;list=$isList;text=[string]$p.Range.Text}
  }
  return @{paragraph_count=$count;table_count=[int]$doc.Tables.Count;paragraphs=$out}
}
function ReadExcel($wb,$limit) {
  $sheets=@()
  for($s=1;$s -le $wb.Worksheets.Count;$s++){
    $ws=$wb.Worksheets.Item($s)
    $used=$ws.UsedRange
    $data=@()
    $rc=[Math]::Min([int]$used.Rows.Count,[Math]::Min($limit,50))
    $cc=[Math]::Min([int]$used.Columns.Count,14)
    $r0=[int]$used.Row;$c0=[int]$used.Column
    for($r=0;$r -lt $rc;$r++){
      $row=@()
      for($c=0;$c -lt $cc;$c++){
        $cell=$ws.Cells.Item($r0+$r,$c0+$c)
        $value=$cell.Value2
        $row+= $(if ($null -eq $value) {''} else {[string]$value})
      }
      $data+= ,$row
    }
    $sheets+=@{name=[string]$ws.Name;used_address=[string]$used.Address();rows=[int]$used.Rows.Count;columns=[int]$used.Columns.Count;shapes=[int]$ws.Shapes.Count;preview=$data}
  }
  return @{sheet_count=[int]$wb.Worksheets.Count;sheets=$sheets}
}
function ReadPpt($ppt,$limit) {
  $all=@()
  for($i=1;$i -le [Math]::Min([int]$ppt.Slides.Count,$limit);$i++){
    $sl=$ppt.Slides.Item($i);$shapes=@()
    for($j=1;$j -le [Math]::Min([int]$sl.Shapes.Count,80);$j++){
      $shape=$sl.Shapes.Item($j)
      $text=''
      try {if($shape.HasTextFrame -and $shape.TextFrame.HasText){$text=[string]$shape.TextFrame.TextRange.Text}}catch {}
      $shapes+=@{shape_index=$j;name=[string]$shape.Name;text=$text;type=[int]$shape.Type}
    }
    $all+=@{slide_number=$i;shapes=$shapes}
  }
  return @{slide_count=[int]$ppt.Slides.Count;slides=$all}
}
function Process($r) {
  $name=[string]$r.name
  $args=$r.arguments
  $allow=@('list','read','verify','save','word_clear','word_insert','word_replace','word_format','word_table','excel_set_range','excel_add_chart','ppt_add_slide','ppt_edit_text','word_comment','word_track_changes','excel_add_sheet','excel_delete_sheet','excel_format_range','ppt_add_textbox','ppt_delete_slide','excel_list_shapes','excel_clear_sheet','excel_freeze_rows','excel_delete_shape','ppt_add_shape','ppt_set_background','ppt_insert_image','ppt_delete_shape','excel_insert_image','word_read_paragraphs','excel_read_range','ppt_read_slide','ppt_preview_slide')
  if ($name -notin $allow) {Fail 'Unsupported operation'}
  $apps=Applications
  $docs=@(Items $apps)
  if ($name -eq 'list') {
    return @{ok=$true;count=$docs.Count;documents=@($docs | ForEach-Object {
      @{document_id=$_.document_id;type=$_.type;name=$_.name;path=$_.path;read_only=$_.read_only}
    })}
  }
  $d=GetTarget $docs ([string]$args.document_id)
  $doc=$d.item
  if($name -eq 'ppt_preview_slide'){
    if($d.type -ne 'powerpoint'){Fail 'Select an open PowerPoint presentation'}
    $index=[int]$args.slide_number
    if($index -lt 1 -or $index -gt [int]$doc.Slides.Count){Fail 'Slide number out of range'}
    $namePart=[Guid]::NewGuid().ToString('N')
    $tempFile=Join-Path ([IO.Path]::GetTempPath()) ('otomasi_mcp_preview_'+$namePart+'.jpg')
    try{
      $doc.Slides.Item($index).Export($tempFile,'JPG',960,540)
      if(-not [IO.File]::Exists($tempFile)){Fail 'PowerPoint preview render missing'}
      $bytes=[IO.File]::ReadAllBytes($tempFile)
      if($bytes.Length -gt 800000){Fail 'Preview is too large'}
      return @{ok=$true;document_id=$d.document_id;slide_number=$index;mime_type='image/jpeg';base64=[Convert]::ToBase64String($bytes)}
    }finally{
      if([IO.File]::Exists($tempFile)){[IO.File]::Delete($tempFile)}
    }
  }
  if($name -eq 'word_read_paragraphs'){
    if($d.type -ne 'word'){Fail 'Select an open Word document'}
    $start=[int]$args.start;$count=[int]$args.count
    if($start -lt 1 -or $count -lt 1 -or $count -gt 75){Fail 'Invalid paragraph window'}
    $total=[int]$doc.Paragraphs.Count
    $out=@()
    for($i=$start;$i -le [Math]::Min($total,$start+$count-1);$i++){
      $p=$doc.Paragraphs.Item($i)
      $style=try{[string]$p.Style.NameLocal}catch{'Normal'}
      $list=$false
      try{$list=([int]$p.Range.ListFormat.ListType -ne 0)}catch{}
      $out+=@{index=$i;style=$style;list=$list;text=[string]$p.Range.Text}
    }
    return @{ok=$true;type='word';document_id=$d.document_id;paragraph_count=$total;start=$start;returned=$out.Count;paragraphs=$out}
  }
  if($name -eq 'excel_read_range'){
    if($d.type -ne 'excel'){Fail 'Select an open Excel workbook'}
    $sheet=[string]$args.sheet;$address=[string]$args.address
    if($address -cnotmatch '^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?$'){
      Fail 'Use a bounded Excel A1 address such as B6:F25'
    }
    $ws=$doc.Worksheets.Item($sheet)
    $range=$ws.Range($address)
    $rows=[int]$range.Rows.Count;$cols=[int]$range.Columns.Count
    if($rows -gt 30 -or $cols -gt 20 -or $rows*$cols -gt 400){Fail 'Range exceeds 30 rows, 20 columns or 400 cells'}
    $values=@();$formulas=@()
    for($i=1;$i -le $rows;$i++){
      $vrow=@();$frow=@()
      for($j=1;$j -le $cols;$j++){
        $cell=$range.Cells.Item($i,$j)
        $raw=$cell.Value2
        $vrow+= $(if($null -eq $raw){''}else{[string]$raw})
        if($args.include_formulas){$frow+=[string]$cell.Formula}
      }
      $values+= ,$vrow
      if($args.include_formulas){$formulas+= ,$frow}
    }
    return @{ok=$true;type='excel';document_id=$d.document_id;sheet=$sheet;address=[string]$range.Address();row_count=$rows;column_count=$cols;values=$values;formulas=$formulas}
  }
  if($name -eq 'ppt_read_slide'){
    if($d.type -ne 'powerpoint'){Fail 'Select an open PowerPoint presentation'}
    $index=[int]$args.slide_number
    if($index -lt 1 -or $index -gt $doc.Slides.Count){Fail 'Slide number out of range'}
    $slide=$doc.Slides.Item($index);$items=@()
    for($j=1;$j -le [Math]::Min(100,[int]$slide.Shapes.Count);$j++){
      $shape=$slide.Shapes.Item($j);$text=''
      try{if($shape.HasTextFrame -and $shape.TextFrame.HasText){$text=[string]$shape.TextFrame.TextRange.Text}}catch{}
      if($text.Length -gt 5000){$text=$text.Substring(0,5000)}
      $items+=@{shape_index=$j;name=[string]$shape.Name;type=[int]$shape.Type;text=$text;left=[double]$shape.Left;top=[double]$shape.Top;width=[double]$shape.Width;height=[double]$shape.Height}
    }
    return @{ok=$true;type='powerpoint';document_id=$d.document_id;slide_number=$index;slide_count=[int]$doc.Slides.Count;page_width_points=[double]$doc.PageSetup.SlideWidth;page_height_points=[double]$doc.PageSetup.SlideHeight;shape_count=[int]$slide.Shapes.Count;truncated=([int]$slide.Shapes.Count -gt 100);shapes=$items}
  }
  if($name -eq 'verify' -and $d.type -eq 'powerpoint'){
    $summary=@()
    for($i=1;$i -le $doc.Slides.Count;$i++){
      $sl=$doc.Slides.Item($i)
      $summary+=@{slide_number=$i;shape_count=[int]$sl.Shapes.Count}
    }
    return @{ok=$true;document_id=$d.document_id;type='powerpoint';slide_count=[int]$doc.Slides.Count;slides=$summary}
  }
  if($name -eq 'verify' -and $d.type -eq 'excel'){
    return @{ok=$true;document_id=$d.document_id;type='excel';sheet_count=[int]$doc.Worksheets.Count}
  }
  if ($name -eq 'read' -or $name -eq 'verify') {
    $max=if ($name -eq 'verify') {300} else {[Math]::Max(1,[Math]::Min(300,[int]$args.limit))}
    $data=switch($d.type) {
      'word' {ReadWord $doc $max}
      'excel' {ReadExcel $doc $max}
      'powerpoint' {ReadPpt $doc $max}
    }
    if ($name -eq 'verify' -and $d.type -eq 'word') {
      $styles=@{};$listCount=0
      foreach($p in $data.paragraphs) {
        $styles[$p.style]=1+([int]$styles[$p.style])
        if($p.list){$listCount++}
      }
      return @{ok=$true;document_id=$d.document_id;paragraph_count=$data.paragraph_count;table_count=$data.table_count;styles=$styles;list_paragraphs=$listCount;inspection_limited_to=$max}
    }
    return @{ok=$true;document_id=$d.document_id;type=$d.type;data=$data}
  }
  RequireWrite $d
  if ($name -eq 'save') {
    if ([string]$doc.Path -eq '') {Fail 'Unsaved new document: use Office Save As first'}
    $doc.Save()
    return @{ok=$true;document_id=$d.document_id;saved=$true}
  }
  if ($name.StartsWith('word_') -and $d.type -ne 'word') {Fail 'Selected document is not Word'}
  if ($name.StartsWith('excel_') -and $d.type -ne 'excel') {Fail 'Selected document is not Excel'}
  if ($name.StartsWith('ppt_') -and $d.type -ne 'powerpoint') {Fail 'Selected document is not PowerPoint'}
  if ($name -eq 'word_clear') {
    [void]$doc.Content.Delete()
    return @{ok=$true;count=1;paragraphs_after=[int]$doc.Paragraphs.Count}
  }
  if ($name -eq 'word_insert') {
    $items=@($args.items);$position=[string]$args.position
    if ($position -notin @('start','end')) {Fail 'Position must be start or end'}
    if($items.Count -gt 150) {Fail 'Too many paragraphs'}
    $ordered=$items
    # For 'start' preserve original order through index reverse (implemented below).
    if($position -eq 'start'){$ordered=@();for($i=$items.Count-1;$i -ge 0;$i--){$ordered+= $items[$i]}}
    $n=0
    foreach($item in $ordered){
      $text=[string]$item.text
      $style=[string]$item.style
      if (-not $style) {$style='Normal'}
      if($style -notin @('Normal','Title','Subtitle','Heading 1','Heading 2','Heading 3','List Bullet','List Number')){Fail ('Unsupported style: '+$style)}
      $where=if($position -eq 'start') {0} else {[Math]::Max(0,$doc.Content.End-1)}
      $range=$doc.Range($where,$where)
      $range.InsertBefore($text+[char]13)
      $applied=$doc.Range($where,($where+$text.Length))
      $applied.Style=$style
      $n++
    }
    return @{ok=$true;count=$n;total_paragraphs=[int]$doc.Paragraphs.Count}
  }
  if($name -eq 'word_replace') {
    $range=$doc.Content
    $find=$range.Find
    $find.ClearFormatting()
    $find.Replacement.ClearFormatting()
    $find.Text=[string]$args.find
    $find.Replacement.Text=[string]$args.replace
    $find.Forward=$true
    $find.Wrap=0
    $mode=if($args.replace_all){2}else{1}
    $did=$find.Execute([string]$args.find,$false,$false,$false,$false,$false,$true,0,$false,[string]$args.replace,$mode)
    return @{ok=$true;matched=[bool]$did;all=[bool]$args.replace_all}
  }
  if($name -eq 'word_format') {
    $range=$doc.Content
    $range.Find.ClearFormatting();$range.Find.Text=[string]$args.find;$range.Find.Wrap=0
    $matched=$range.Find.Execute()
    if(-not $matched){Fail 'Text not found'}
    $style=[string]$args.style
    if($style -notin @('Normal','Title','Subtitle','Heading 1','Heading 2','Heading 3','List Bullet','List Number')){Fail 'Invalid style'}
    $range.Paragraphs.Item(1).Style=$style
    $range.Bold=if($args.bold){1}else{0}
    $range.Italic=if($args.italic){1}else{0}
    $size=[int]$args.font_size
    if($size -lt 8 -or $size -gt 72){Fail 'Font size must be 8-72'}
    $range.Font.Size=$size
    return @{ok=$true;matched=$true;style=$style}
  }
  if($name -eq 'word_table') {
    $rows=@($args.rows)
    $nr=$rows.Count
    $nc=@($rows[0]).Count
    if($nr -lt 1 -or $nr -gt 100 -or $nc -lt 1 -or $nc -gt 25){Fail 'Invalid table size'}
    $pos=[Math]::Max(0,$doc.Content.End-1)
    $range=$doc.Range($pos,$pos)
    $table=$doc.Tables.Add($range,$nr,$nc)
    for($i=1;$i -le $nr;$i++){for($j=1;$j -le $nc;$j++){$table.Cell($i,$j).Range.Text=[string]$rows[$i-1][$j-1]}}
    $table.Borders.Enable=1
    return @{ok=$true;rows=$nr;columns=$nc;table_count=[int]$doc.Tables.Count}
  }
  if ($name -eq 'excel_set_range') {
    $ws=$doc.Worksheets.Item([string]$args.sheet)
    $target=$ws.Range([string]$args.address)
    $rows=@($args.values)
    $nr=$rows.Count;$nc=@($rows[0]).Count
    if($nr -ne [int]$target.Rows.Count -or $nc -ne [int]$target.Columns.Count){Fail 'Values matrix dimensions must match address'}
    for($i=0;$i -lt $nr;$i++){
      if(@($rows[$i]).Count -ne $nc){Fail 'Ragged value matrix'}
      for($j=0;$j -lt $nc;$j++){
        $v=$rows[$i][$j]
        $cell=$target.Cells.Item($i+1,$j+1)
        if($v -is [string]) {
          if($v.StartsWith('=')) {$cell.Formula=[string]$v}
          else {$cell.Value2=[string]$v}
        } elseif ($null -eq $v) {$cell.ClearContents()}
        elseif ($v -is [bool]) {$cell.Value2=[bool]$v}
        else {$cell.Value2=[double]$v}
      }
    }
    return @{ok=$true;written=$nr*$nc;range=$target.Address()}
  }
  if ($name -eq 'excel_add_chart') {
    $ws=$doc.Worksheets.Item([string]$args.sheet)
    $source=$ws.Range([string]$args.data_range)
    $chartObj=$ws.ChartObjects().Add(100,60,430,265)
    $chart=$chartObj.Chart;$chart.SetSourceData($source)
    $chart.ChartType=switch ([string]$args.chart_type) {
      'line' {4} 'pie' {5} 'bar' {57} default {51}
    }
    return @{ok=$true;chart_name=[string]$chartObj.Name}
  }
  if($name -eq 'word_comment') {
    $range=$doc.Content
    $range.Find.ClearFormatting()
    $range.Find.Text=[string]$args.find
    if(-not $range.Find.Execute()) {Fail 'Text not found'}
    $comment=$doc.Comments.Add($range,[string]$args.comment)
    return @{ok=$true;comment_count=[int]$doc.Comments.Count}
  }
  if($name -eq 'word_track_changes') {
    $doc.TrackRevisions=[bool]$args.enabled
    return @{ok=$true;track_revisions=[bool]$doc.TrackRevisions}
  }
  if($name -eq 'excel_add_sheet') {
    $sheetName=[string]$args.name
    if($sheetName.Length -lt 1 -or $sheetName.Length -gt 31 -or $sheetName -match '[\\/\?\*\[\]:]') {Fail 'Invalid sheet name'}
    $ws=$doc.Worksheets.Add()
    $ws.Name=$sheetName
    return @{ok=$true;name=[string]$ws.Name;sheet_count=[int]$doc.Worksheets.Count}
  }
  if($name -eq 'excel_delete_sheet') {
    if($doc.Worksheets.Count -lt 2) {Fail 'Cannot delete the final worksheet'}
    $ws=$doc.Worksheets.Item([string]$args.name)
    $app=$doc.Application
    $old=$app.DisplayAlerts
    try {$app.DisplayAlerts=$false;[void]$ws.Delete()}
    finally {$app.DisplayAlerts=$old}
    return @{ok=$true;sheet_count=[int]$doc.Worksheets.Count}
  }
  if($name -eq 'excel_format_range') {
    $range=$doc.Worksheets.Item([string]$args.sheet).Range([string]$args.address)
    $size=[int]$args.font_size
    if($size -lt 8 -or $size -gt 72) {Fail 'Font must be 8-72'}
    $range.Font.Bold=[bool]$args.bold
    $range.Font.Size=$size
    $range.NumberFormat=[string]$args.number_format
    return @{ok=$true;range=[string]$range.Address();font_size=$size}
  }
  if($name -eq 'ppt_add_textbox') {
    $sn=[int]$args.slide_number
    if($sn -lt 1 -or $sn -gt $doc.Slides.Count) {Fail 'Slide index out of range'}
    $vals=@([double]$args.left,[double]$args.top,[double]$args.width,[double]$args.height)
    if($vals[2] -le 0 -or $vals[3] -le 0 -or $vals[2] -gt 2000 -or $vals[3] -gt 2000) {Fail 'Invalid shape size'}
    $shape=$doc.Slides.Item($sn).Shapes.AddTextbox(1,$vals[0],$vals[1],$vals[2],$vals[3])
    $shape.TextFrame.TextRange.Text=[string]$args.text
    return @{ok=$true;slide_number=$sn;shape_index=[int]$shape.ZOrderPosition}
  }
  if($name -eq 'ppt_delete_slide') {
    $sn=[int]$args.slide_number
    if($sn -lt 1 -or $sn -gt $doc.Slides.Count) {Fail 'Slide index out of range'}
    [void]$doc.Slides.Item($sn).Delete()
    return @{ok=$true;slide_count=[int]$doc.Slides.Count}
  }
  if($name -eq 'excel_insert_image') {
    $sheet=$doc.Worksheets.Item([string]$args.sheet)
    $source=[string]$args.local_image_path
    if($source -match '^\\\\' -or $source -match '^[a-zA-Z]+://') {Fail 'Only local image paths accepted'}
    $ext=[IO.Path]::GetExtension($source).ToLowerInvariant()
    if($ext -notin @('.png','.jpg','.jpeg','.bmp')){Fail 'Only PNG/JPEG/BMP supported'}
    if(-not [IO.File]::Exists($source)){Fail 'Image not found on Windows device'}
    if((Get-Item -LiteralPath $source).Length -gt 25000000){Fail 'Image too large'}
    $l=[double]$args.left;$t=[double]$args.top;$w=[double]$args.width;$h=[double]$args.height
    if($w -le 0 -or $h -le 0 -or $w -gt 4000 -or $h -gt 4000){Fail 'Invalid dimensions'}
    $shape=$sheet.Shapes.AddPicture($source,0,-1,$l,$t,$w,$h)
    return @{ok=$true;sheet=[string]$sheet.Name;shape_name=[string]$shape.Name;shapes=[int]$sheet.Shapes.Count}
  }
  if($name -eq 'excel_list_shapes') {
    $sheet=$doc.Worksheets.Item([string]$args.sheet)
    $shapes=@()
    for($i=1;$i -le $sheet.Shapes.Count;$i++) {
      $shape=$sheet.Shapes.Item($i)
      $shapes+=@{name=[string]$shape.Name;type=[int]$shape.Type}
    }
    return @{ok=$true;sheet=[string]$sheet.Name;shape_count=$shapes.Count;shapes=$shapes}
  }
  if($name -eq 'excel_clear_sheet') {
    $sheet=$doc.Worksheets.Item([string]$args.sheet)
    $prior=[int]$sheet.Shapes.Count
    [void]$sheet.Cells.Clear()
    for($i=$sheet.Shapes.Count;$i -ge 1;$i--) {[void]$sheet.Shapes.Item($i).Delete()}
    $app=$doc.Application
    $previous=$app.ActiveSheet
    try {
      $sheet.Activate()
      $window=$app.ActiveWindow
      $window.FreezePanes=$false
      $window.SplitColumn=0
      $window.SplitRow=0
    } finally {
      try {$previous.Activate()}catch {}
    }
    return @{ok=$true;sheet=[string]$sheet.Name;shapes_deleted=$prior;remaining_shapes=[int]$sheet.Shapes.Count;freeze_rows=0}
  }
  if($name -eq 'excel_delete_shape') {
    $sheet=$doc.Worksheets.Item([string]$args.sheet)
    $shapeName=[string]$args.shape_name
    if(-not $shapeName -or $shapeName.Length -gt 200) {Fail 'Shape name required'}
    $shape=$sheet.Shapes.Item($shapeName)
    [void]$shape.Delete()
    return @{ok=$true;sheet=[string]$sheet.Name;deleted=$shapeName;remaining=[int]$sheet.Shapes.Count}
  }
  if($name -eq 'excel_freeze_rows') {
    $sheet=$doc.Worksheets.Item([string]$args.sheet)
    $n=[int]$args.rows
    if($n -lt 0 -or $n -gt 50){Fail 'Freeze rows limit 0..50'}
    $app=$doc.Application
    $previous=$app.ActiveSheet
    try {
      $sheet.Activate()
      $window=$app.ActiveWindow
      $window.FreezePanes=$false
      $window.SplitColumn=0
      $window.SplitRow=$n
      if($n -gt 0){$window.FreezePanes=$true}
      return @{ok=$true;sheet=[string]$sheet.Name;rows=$n;frozen=[bool]$window.FreezePanes}
    } finally {
      try {$previous.Activate()}catch {}
    }
  }
  if($name -eq 'ppt_set_background') {
    $sn=[int]$args.slide_number
    if($sn -lt 1 -or $sn -gt $doc.Slides.Count){Fail 'Slide index out of range'}
    $slide=$doc.Slides.Item($sn)
    $slide.FollowMasterBackground=$false
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB=ColorValue ([string]$args.color)
    return @{ok=$true;slide_number=$sn;color=[string]$args.color}
  }
  if($name -eq 'ppt_add_shape') {
    $sn=[int]$args.slide_number
    if($sn -lt 1 -or $sn -gt $doc.Slides.Count){Fail 'Slide index out of range'}
    $l=[double]$args.left;$t=[double]$args.top;$w=[double]$args.width;$h=[double]$args.height
    if($l -lt -50 -or $t -lt -50 -or $l -gt 4000 -or $t -gt 4000 -or $w -le 0 -or $h -le 0 -or $w -gt 4000 -or $h -gt 4000){Fail 'Invalid shape dimensions'}
    $kind=[string]$args.shape_type
    $shapeType=switch($kind){'rectangle'{1}'rounded_rectangle'{5}'ellipse'{9}'line'{0} default{Fail 'Unsupported shape type'}}
    if($kind -eq 'line') {
      $shape=$doc.Slides.Item($sn).Shapes.AddLine($l,$t,($l+$w),($t+$h))
      $shape.Line.ForeColor.RGB=ColorValue ([string]$args.fill)
    }else{
      $shape=$doc.Slides.Item($sn).Shapes.AddShape($shapeType,$l,$t,$w,$h)
      $shape.Fill.Solid()
      $shape.Fill.ForeColor.RGB=ColorValue ([string]$args.fill)
      $shape.Line.Visible=0
      if([string]$args.text){
        $shape.TextFrame.TextRange.Text=[string]$args.text
        $shape.TextFrame.TextRange.Font.Color.RGB=ColorValue ([string]$args.text_color)
        $size=[int]$args.font_size
        if($size -lt 8 -or $size -gt 96){Fail 'Font size 8..96'}
        $shape.TextFrame.TextRange.Font.Size=$size
      }
    }
    return @{ok=$true;slide_number=$sn;shape_name=[string]$shape.Name;shape_index=[int]$shape.ZOrderPosition}
  }
  if($name -eq 'ppt_insert_image') {
    $sn=[int]$args.slide_number
    if($sn -lt 1 -or $sn -gt $doc.Slides.Count){Fail 'Slide index out of range'}
    $source=[string]$args.local_image_path
    if($source -match '^\\\\' -or $source -match '^[a-zA-Z]+://') {Fail 'Only local image paths are accepted'}
    $ext=[IO.Path]::GetExtension($source).ToLowerInvariant()
    if($ext -notin @('.png','.jpg','.jpeg','.bmp')){Fail 'Only PNG/JPEG/BMP supported'}
    if(-not [IO.File]::Exists($source)){Fail 'Image not found on Windows device'}
    if((Get-Item -LiteralPath $source).Length -gt 25000000){Fail 'Image too large'}
    $l=[double]$args.left;$t=[double]$args.top;$w=[double]$args.width;$h=[double]$args.height
    if($w -le 0 -or $h -le 0 -or $w -gt 4000 -or $h -gt 4000){Fail 'Invalid image dimensions'}
    $shape=$doc.Slides.Item($sn).Shapes.AddPicture($source,0,-1,$l,$t,$w,$h)
    return @{ok=$true;slide_number=$sn;shape_name=[string]$shape.Name;shape_index=[int]$shape.ZOrderPosition}
  }
  if($name -eq 'ppt_delete_shape') {
    $sn=[int]$args.slide_number;$si=[int]$args.shape_index
    if($sn -lt 1 -or $sn -gt $doc.Slides.Count){Fail 'Slide index out of range'}
    $slide=$doc.Slides.Item($sn)
    if($si -lt 1 -or $si -gt $slide.Shapes.Count){Fail 'Shape index out of range'}
    $shapeName=[string]$slide.Shapes.Item($si).Name
    [void]$slide.Shapes.Item($si).Delete()
    return @{ok=$true;deleted=$shapeName;remaining_shapes=[int]$slide.Shapes.Count}
  }
  if($name -eq 'ppt_add_slide') {
    $layout=[int]$args.layout
    if($layout -notin @(1,2,11,12)){Fail 'Unsupported slide layout'}
    $slide=$doc.Slides.Add($doc.Slides.Count+1,$layout)
    try{$slide.Shapes.Title.TextFrame.TextRange.Text=[string]$args.title}catch{}
    try{if($slide.Shapes.Placeholders.Count -ge 2){$slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text=[string]$args.body}}catch{}
    return @{ok=$true;slide_count=[int]$doc.Slides.Count;slide_number=[int]$slide.SlideIndex}
  }
  if($name -eq 'ppt_edit_text') {
    $sn=[int]$args.slide_number;$si=[int]$args.shape_index
    if($sn -lt 1 -or $sn -gt $doc.Slides.Count) {Fail 'Slide index out of range'}
    $slide=$doc.Slides.Item($sn)
    if($si -lt 1 -or $si -gt $slide.Shapes.Count) {Fail 'Shape index out of range'}
    $shape=$slide.Shapes.Item($si)
    if(-not $shape.HasTextFrame){Fail 'Shape has no text frame'}
    $shape.TextFrame.TextRange.Text=[string]$args.text
    return @{ok=$true;slide_number=$sn;shape_index=$si}
  }
  Fail 'Operation not implemented'
}
try{
  if($Payload.Length -gt 350000) {Fail 'Payload too large'}
  $json=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Payload))
  $inputObject=$json|ConvertFrom-Json
  $answer=Process $inputObject
  $out=ConvertTo-Json -InputObject $answer -Depth 24 -Compress
  [Console]::Out.WriteLine($out)
}catch{
  [Console]::Out.WriteLine((ConvertTo-Json -Compress -InputObject @{ok=$false;error=$_.Exception.Message;detail=($_.InvocationInfo.ScriptLineNumber)}))
  exit 1
}
