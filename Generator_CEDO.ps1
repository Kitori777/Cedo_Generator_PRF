# CEDO Packaging Request Generator - AUTO DETECT DATABASE SCHEMAS, WHITE/BROWN, STACKABLE DEFAULT NO / TWO DATABASE SCHEMAS - OFFLINE, WITHOUT EXCEL COM
# Requires Windows PowerShell 5.1 and .NET 4.5+; does not launch Excel.
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()
$base=Split-Path -Parent $MyInvocation.MyCommand.Path
$ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:Log=New-Object System.Collections.Generic.List[string]
$script:Count=0
$script:Fail=0

function Ask-File($title){$d=New-Object System.Windows.Forms.OpenFileDialog;$d.Title=$title;$d.Filter='Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm';if($d.ShowDialog() -eq 'OK'){return $d.FileName};return $null}
function Ask-Folder($title){$d=New-Object System.Windows.Forms.FolderBrowserDialog;$d.Description=$title;if($d.ShowDialog() -eq 'OK'){return $d.SelectedPath};return $null}
function Unzip($file,$to){[System.IO.Compression.ZipFile]::ExtractToDirectory($file,$to)}
function Zip($from,$file){
 # Explicit '/' ZIP paths; Windows/.NET versions may otherwise store backslashes.
 $root=[System.IO.Path]::GetFullPath($from).TrimEnd([char]'\',[char]'/')
 $archive=[System.IO.Compression.ZipFile]::Open($file,[System.IO.Compression.ZipArchiveMode]::Create)
 try{
  $items=@(Get-ChildItem -LiteralPath $root -Recurse -File -Force)
  foreach($item in $items){
   $relative=$item.FullName.Substring($root.Length).TrimStart([char]'\',[char]'/').Replace('\','/')
   if([string]::IsNullOrWhiteSpace($relative)){throw 'Niepoprawna sciezka w archiwum'}
   [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive,$item.FullName,$relative,[System.IO.Compression.CompressionLevel]::Optimal)
  }
 }finally{$archive.Dispose()}
}
function Read-Xml($path){$x=New-Object System.Xml.XmlDocument;$x.PreserveWhitespace=$true;$x.Load($path);return $x}
function Save-Xml($x,$path){$settings=New-Object System.Xml.XmlWriterSettings;$settings.Encoding=New-Object System.Text.UTF8Encoding($false);$settings.Indent=$false;$writer=[System.Xml.XmlWriter]::Create($path,$settings);try{$x.Save($writer)}finally{$writer.Close()}}
function Child($node,$local){if($null -eq $node){return $null};foreach($a in $node.ChildNodes){if($a.LocalName -eq $local){return $a}};return $null}
function ColNum($letter){$n=0;foreach($ch in $letter.ToUpperInvariant().ToCharArray()){$n=$n*26 + ([int][char]$ch-64)};return $n}
function Letter($n){
 $number=[int]$n; $result=''
 while($number -gt 0){
  $number--; $letter=[char]([int](65 + ($number % 26)))
  $result=([string]$letter)+$result
  $number=[int][math]::Floor($number / 26.0)
 }
 return $result
}
function Cell-Col($a){$s=([string]$a) -replace '\d','';return (ColNum $s)}
function SText($x){$s='';foreach($n in $x.SelectNodes(".//*[local-name()='t']")){$s+=$n.InnerText};return $s}
function Sheet-Path($dir,$sheetName){
 $book=Read-Xml (Join-Path $dir 'xl/workbook.xml');$rels=Read-Xml (Join-Path $dir 'xl/_rels/workbook.xml.rels')
 $id=$null;foreach($sh in $book.SelectNodes("//*[local-name()='sheet']")){if($sh.GetAttribute('name') -eq $sheetName){$id=$sh.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships');break}}
 if(-not $id){throw "Nie znaleziono arkusza '$sheetName'."}
 foreach($rel in $rels.SelectNodes("//*[local-name()='Relationship']")){if($rel.GetAttribute('Id') -eq $id){$target=$rel.GetAttribute('Target').TrimStart('/');if($target -match '^xl/') {return (Join-Path $dir ($target -replace '/','\'))};return (Join-Path $dir ('xl\'+($target -replace '/','\')))}}
 throw "Nie znaleziono powiazania arkusza '$sheetName'."
}
function Read-WorkbookData($dir){
 # Workbooks can have arbitrary worksheet names and header-row positions.
 $shared=@();$sfile=Join-Path $dir 'xl/sharedStrings.xml'
 if(Test-Path $sfile){$sx=Read-Xml $sfile;foreach($si in $sx.SelectNodes("//*[local-name()='si']")){$shared+= (SText $si)}}
 $wb=Read-Xml (Join-Path $dir 'xl/workbook.xml')
 $sheetNames=@($wb.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']") | ForEach-Object {$_.GetAttribute('name')})
 $bestScore=-1;$bestRows=$null;$bestSheet='';$bestHead=0
 foreach($sheetName in $sheetNames){
  try{$path=Sheet-Path $dir $sheetName;$doc=Read-Xml $path}catch{continue}
  $candidate=New-Object System.Collections.ArrayList
  foreach($r in $doc.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")){
   $vals=@{}
   foreach($c in $r.SelectNodes("./*[local-name()='c']")){
    $col=Cell-Col $c.GetAttribute('r');$v=Child $c 'v';$inline=Child $c 'is';$value=''
    if($null -ne $v){$value=$v.InnerText;if($c.GetAttribute('t') -eq 's'){$idx=0;if([int]::TryParse($value,[ref]$idx) -and $idx -ge 0 -and $idx -lt $shared.Count){$value=$shared[$idx]}}}
    elseif($null -ne $inline){$value=SText $inline}
    $vals[$col]=$value
   }
   [void]$candidate.Add([pscustomobject]@{No=[int]$r.GetAttribute('r');V=$vals})
  }
  $limit=[math]::Min(30,$candidate.Count)
  for($h=0;$h -lt $limit;$h++){
   $head=$candidate[$h].V;$score=0
   foreach($v in $head.Values){
    $n=Normalize-Head $v
    if($n -in @('newocprintcode','newoutercartonprintcode')){$score+=12}
    if($n -in @('octype','packagingtype','cartontype')){$score+=8}
    if($n -in @('componentsname','componentname','partdesc','description')){$score+=4}
    if($n -in @('partno','partnumber','sku','productcode')){$score+=4}
    if($n -in @('internallengthmm','internallength','internalwidthmm','internalheightmm')){$score+=2}
    if($n -match '^fc[0-9]' -or $n -in @('annualvolume','requiredannualvolume')){$score++}
   }
   # Prefer header rows with recognized fields and at least one data record below.
   if($candidate.Count -gt ($h+1) -and $score -gt $bestScore){$bestScore=$score;$bestRows=$candidate;$bestSheet=$sheetName;$bestHead=$h}
  }
 }
 if($null -eq $bestRows -or $bestScore -lt 8){throw 'Nie rozpoznano tabeli z danymi. Sprawdz, czy baza ma jeden wiersz naglowkow i kolumny identyfikacyjne; mozna przygotowac profil mapowania dla nowego ukladu.'}
 $script:DetectedSheet=$bestSheet;$script:DetectedHeaderRow=$bestRows[$bestHead].No
 $result=New-Object System.Collections.ArrayList
 for($i=$bestHead;$i -lt $bestRows.Count;$i++){[void]$result.Add($bestRows[$i])}
 return ,$result
}
function Field($row,$key){
 if(-not $script:Mapping.ContainsKey($key)){return ''};$m=$script:Mapping[$key]
 switch($m.Mode){'Pomin'{return ''};'Recznie'{return ''};'Stala'{return [string]$m.Stala};default {if($m.Col -gt 0 -and $row.V.ContainsKey([int]$m.Col)){return [string]$row.V[[int]$m.Col]};return ''}}
}
function Has($v){return (-not [string]::IsNullOrWhiteSpace([string]$v) -and [string]$v -ne '0')}
# CEDO business rule: only explicit WHITE means WHITE; all other values (including GREY,
# N/A, empty, BROWN) mean BROWN. Manual and skip modes are respected.
function Board-Colour($row,$key){
 if(-not $script:Mapping.ContainsKey($key)){return 'Brown'}
 $mode=[string]$script:Mapping[$key].Mode
 if($mode -in @('Pomin','Recznie')){return ''}
 $raw=([string](Field $row $key)).Trim()
 if($raw -match '^(?i:white)$'){return 'White'}
 return 'Brown'
}
# Stacking is ALLOWED only when the source explicitly says YES/TRUE/TAK/1.
# Blank, N/A, 0, NO and ambiguous text -> NO.
function Stackable-Value($row){
 if(-not $script:Mapping.ContainsKey('Stackable')){return 'no'}
 $mode=[string]$script:Mapping['Stackable'].Mode
 if($mode -in @('Pomin','Recznie')){return ''}
 $raw=([string](Field $row 'Stackable')).Trim()
 if($raw -match '^(?i:yes|y|true|tak|t|1|stackable|can be stacked|mozna|może być piętrowana|mozna pietrowac)$'){return 'yes'}
 return 'no'
}

function Set-Cell($d,$address,$value,$percent=$false){
 if($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)){return}
 $cell=$null;foreach($x in $d.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']/*[local-name()='c']")){if($x.GetAttribute('r') -eq $address){$cell=$x;break}}
 if($null -eq $cell){throw "W szablonie nie znaleziono komorki $address."}
 if($null -ne (Child $cell 'f')){throw "Ochrona formuly: nie mozna nadpisac komorki $address."}
 while($cell.HasChildNodes){$null=$cell.RemoveChild($cell.FirstChild)}
 $numeric=0.0;$str=[string]$value;$isNum=([double]::TryParse($str,[System.Globalization.NumberStyles]::Float,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$numeric))
 if($percent){$numeric=if($str -match '^(?i:yes|printed)$'){0.8}elseif($str -match '^(?i:no|perforated)$'){0}elseif($isNum){if($numeric -gt 1){$numeric/100}else{$numeric}}else{0};$isNum=$true}
 if($isNum){$cell.RemoveAttribute('t');$v=$d.CreateElement('v',$ns);$v.InnerText=$numeric.ToString('G17',[System.Globalization.CultureInfo]::InvariantCulture);[void]$cell.AppendChild($v)}
 else{$cell.SetAttribute('t','inlineStr');$is=$d.CreateElement('is',$ns);$t=$d.CreateElement('t',$ns);$t.InnerText=$str;[void]$is.AppendChild($t);[void]$cell.AppendChild($is)}
}
function Set-FormulaCache($d,$address,$value){
 $cell=$null;foreach($x in $d.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']/*[local-name()='c']")){if($x.GetAttribute('r') -eq $address){$cell=$x;break}}
 if($null -eq $cell){throw "Brak komorki $address"}
 if($null -eq (Child $cell 'f')){throw "Brak wymaganej formuly w $address"}
 $old=Child $cell 'v';if($null -ne $old){$null=$cell.RemoveChild($old)}
 $v=$d.CreateElement('v',$ns);$v.InnerText=([double]$value).ToString('G17',[System.Globalization.CultureInfo]::InvariantCulture);[void]$cell.AppendChild($v)
}
function Var-Choice($value,$label){
 $s=([string]$value).Trim().ToLowerInvariant();$l=([string]$label).Trim().ToLowerInvariant().Replace('_',' ')
 if(-not $s -or $s -eq '0'){return $false}
 if($l -eq 'flexo'){return $s.Contains('flexo')}
 if($l -eq 'litho'){return ($s.Contains('offset') -or $s.Contains('litho'))}
 if($l -eq 'digital'){return $s.Contains('digital')}
 if($l -eq 'no varnish'){return $s -in @('n/a','na','none','no varnish','no')}
 if($l -eq 'std machine varnish'){return ($s.Contains('std') -and -not $s.Contains('matt'))}
 if($l -eq 'matt varnish'){return $s.Contains('matt')}
 if($l -eq 'printed'){return $s -in @('yes','printed')}
 if($l -eq 'perforated'){return $s.Contains('perforat')}
 if($l -eq 'white'){return $s -eq 'white'}
 if($l -eq 'brown'){return $s -eq 'brown'}
 if($l -in @('coated','un coated')){return (($l -eq 'coated' -and $s -eq 'coated') -or ($l -eq 'un coated' -and $s -in @('uncoated','un_coated','un coated')))}
 if($l -eq 'automatic'){return $s -eq 'automatic'}
 if($l -eq 'manual'){return $s -eq 'manual'}
 if($l -eq 'to match tray'){return $s -match 'match tray'}
 return $false
}
function Set-Options($temp,$row,$kind,[ref]$warnings){
 $vmlPath=Join-Path $temp 'xl/drawings/vmlDrawing1.vml';if(-not (Test-Path $vmlPath)){return}
 $d=Read-Xml $vmlPath
 $left=@{12=(Field $row 'Print');14=(Field $row 'Varnish');15=(Field $row 'Printed');17=(Board-Colour $row 'External');18=(Field $row 'Coating');19=(Board-Colour $row 'Internal');21=(Field $row 'PackProcess')}
 $right=@{12=(Field $row 'LidPrint');14=(Field $row 'LidVarnish');15=(Field $row 'LidPrinted');17=(Board-Colour $row 'LidExternal');18=(Field $row 'LidCoating');19=(Board-Colour $row 'LidInternal')}
 $choices=@{}; $vmlNs='urn:schemas-microsoft-com:office:excel'
 foreach($sh in $d.SelectNodes("//*[local-name()='shape']")){
  $cd=$null;foreach($n in $sh.ChildNodes){if($n.LocalName -eq 'ClientData'){$cd=$n;break}}
  if($null -eq $cd -or $cd.GetAttribute('ObjectType') -ne 'Radio'){continue}
  $anchor=Child $cd 'Anchor';if($null -eq $anchor){continue};$a=@($anchor.InnerText -split ',' | ForEach-Object {$_.Trim()})
  if($a.Count -lt 4){continue};$col=[int]$a[0];$rr=[int]$a[2]+1
  $side=if($col -le 2){'left'}else{'right'}
  if($side -eq 'right' -and ($kind -ne 'Tray' -or -not (Has (Field $row 'LidID')))){continue}
  $map=if($side -eq 'left'){$left}else{$right};if(-not $map.ContainsKey($rr)){continue}
  $raw=$map[$rr];if(-not (Has $raw)){continue}
  $txt=Child $sh 'textbox';$label='';if($null -ne $txt){$label=($txt.InnerText -replace '\s+',' ').Trim()}
  $key="$side|$rr";if(-not $choices.ContainsKey($key)){$choices[$key]=New-Object System.Collections.ArrayList}
  [void]$choices[$key].Add([pscustomobject]@{Shape=$sh;Client=$cd;Label=$label;Selected=(Var-Choice $raw $label);Raw=$raw})
 }
 $shapeMap=@{};$rels=Read-Xml (Join-Path $temp 'xl/worksheets/_rels/sheet1.xml.rels')
 $rpaths=@{};foreach($re in $rels.SelectNodes("//*[local-name()='Relationship']")){$rpaths[$re.GetAttribute('Id')]=$re.GetAttribute('Target')}
 $sheetDoc=Read-Xml (Join-Path $temp 'xl/worksheets/sheet1.xml')
 foreach($ctrl in $sheetDoc.SelectNodes("//*[local-name()='control']")){$sid=$ctrl.GetAttribute('shapeId');$rid=$ctrl.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships');if($rpaths.ContainsKey($rid)){$shapeMap[$sid]=$rpaths[$rid]}}
 foreach($key in $choices.Keys){
  $group=@($choices[$key]);$selected=@($group | Where-Object {$_.Selected})
  if($selected.Count -ne 1){$warnings.Value+="Nieprzypisana opcja $key = '$($group[0].Raw)'; ";continue}
  foreach($option in $group){
   $check=Child $option.Client 'Checked'
   if($option.Selected){if($null -eq $check){$check=$d.CreateElement('x','Checked',$vmlNs);[void]$option.Client.AppendChild($check)}}elseif($null -ne $check){$null=$option.Client.RemoveChild($check)}
   $sid= $option.Shape.GetAttribute('id') -replace '^_x0000_s',''
   if($shapeMap.ContainsKey($sid)){
    $target=$shapeMap[$sid];if($target -match 'ctrlProp\d+\.xml'){$p=Join-Path $temp ('xl/ctrlProps/'+[System.IO.Path]::GetFileName($target));if(Test-Path $p){$cp=Read-Xml $p;if($option.Selected){$cp.DocumentElement.SetAttribute('checked','Checked')}else{$cp.DocumentElement.RemoveAttribute('checked')};Save-Xml $cp $p}}
   }
  }
 }
 Save-Xml $d $vmlPath
}
# Default mappings are discovered by header name, not hard-coded column numbers.
$spec=@(
 @('PartNo','Part No (identyfikator wiersza / nazwa pliku)',1),@('PrintCode','New OC print code',1),@('Name','Components name',1),@('Type','OC type',1),@('Cutter','Cutter',1),
 @('Length','Internal lenght [mm]',1),@('Width','Internal width [mm]',1),@('Height','Internal height [mm]',1),
 @('Print','Printing technique',1),@('Colours','Number of colours',1),@('Varnish','Varnish',1),
 @('Printed','is SKU number printed/perforated',1),@('External','External board ,colour',1),
 @('Coating','Coating (only for exteral white layer)',1),@('Internal','Internal board ,colour',1),
 @('Fluting','Fluting',1),@('LidID','LID',1),@('LidName','Components name',2),
 @('LidType','OC type',2),@('LidCutter','Cutter',2),
 @('LidLength','Internal lenght [mm]',2),@('LidWidth','Internal width [mm]',2),
 @('LidHeight','Internal height [mm]',2),@('LidPrint','Printing technique',2),
 @('LidColours','Number of colours',2),@('LidVarnish','Varnish',2),
 @('LidPrinted','is SKU number printed/perforated',2),@('LidExternal','External board ,colour',2),
 @('LidCoating','Coating (only for exteral white layer)',2),@('LidInternal','Internal board ,colour',2),
 @('LidFluting','Fluting',2),@('Pieces','number pieces in carton [pcs]',1),
 @('Gross','approx. Gross carton weight [kg]',1),@('Carries','whether the product carries the load',1),
 @('LayerCartons','number cartons on layer [pcs]',1),@('Layers','number of layers',1),
 @('Pallet','pallet type',1),@('PalletHeight','pallet height',1),
 @('Stackable','is pallet stackable',1),
 @('Area','(brak - ręcznie)',0),@('PackProcess','(brak - ręcznie)',0),
 @('Annual','Required annual volume [pcs of cartons] (FC)',1)
)
function Normalize-Head($s){
 $v=([string]$s).Trim().ToLowerInvariant()
 $v=$v.Replace('lenght','length').Replace('exteral','external').Replace('colour','color')
 $v=[regex]::Replace($v,'[^a-z0-9]+','')
 return $v
}
function Auto-Col($header,$key,$needle,$occurrence){
 if($occurrence -lt 1){return 0}
 $aliases=@{
  'PrintCode'=@('newocprintcode','newoutercartonprintcode','ocprintcode','printcode')
  'PartNo'=@('partno','partnumber','productnumber','sku','fgnumber')
  'Name'=@('componentsname','componentname','componentdescription','productdescription','partdesc')
  'Type'=@('octype','outercartontype','packagingtype','cartontype')
  'Length'=@('internallengthmm','internallength')
  'Width'=@('internalwidthmm','internalwidth')
  'Height'=@('internalheightmm','internalheight')
 }
 $target=Normalize-Head $needle
 # Exact source field wins over aliases (e.g. Components name must NOT become Part Desc).
 $found=0
 foreach($i in @($header.Keys | Sort-Object {[int]$_})){
  if((Normalize-Head $header[$i]) -eq $target){$found++;if($found -eq $occurrence){return [int]$i}}
 }
 if($occurrence -ne 1 -or $key -like 'Lid*' -or -not $aliases.ContainsKey($key)){return 0}
 foreach($alias in $aliases[$key]){
  foreach($i in @($header.Keys | Sort-Object {[int]$_})){
   if((Normalize-Head $header[$i]) -eq $alias){return [int]$i}
  }
 }
 return 0
}
# Forecast columns have dated headers: FC 04.08.26 (OCTOPUS 2), FC 13.10.26 (R05).
# Locate the dated FC column by its header, not by its physical location.
# The user may override the selected column in the mapping form.
function Auto-AnnualCol($header){
 foreach($i in @($header.Keys | Sort-Object {[int]$_})){
  $h=([string]$header[$i]).Trim()
  if($h -match '(?i)^FC\s+\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}$'){return [int]$i}
 }
 foreach($i in @($header.Keys | Sort-Object {[int]$_})){
  $h=Normalize-Head $header[$i]
  if($h -in @('requiredannualvolume','annualvolume','yearlyvolume')){return [int]$i}
 }
 return 0
}
function Configure($header,$config){
 $previous=@{};if(Test-Path $config){foreach($r in @(Import-Csv $config -Delimiter ';' -Encoding UTF8)){$previous[$r.Klucz]=$r}}
 $f=New-Object System.Windows.Forms.Form;$f.Text='CEDO - Generator PRF | Mapowanie kolumn - WSZYSTKIE TYPY';$f.Width=1050;$f.Height=740;$f.StartPosition='CenterScreen'
 $lbl=New-Object System.Windows.Forms.Label;$lbl.Dock='Top';$lbl.Height=37;$lbl.Text='Wykryto arkusz: '+$script:DetectedSheet+' (naglowki w wierszu '+$script:DetectedHeaderRow+'). Sprawdz przypisania; Z bazy / Stala / Recznie / Pomin.'
 $g=New-Object System.Windows.Forms.DataGridView;$g.Dock='Fill';$g.AllowUserToAddRows=$false;$g.RowHeadersVisible=$false;$g.AutoSizeColumnsMode='Fill';$g.EditMode='EditOnEnter'
 $c=New-Object System.Windows.Forms.DataGridViewTextBoxColumn;$c.Name='Key';$c.HeaderText='Klucz';$c.ReadOnly=$true;$c.FillWeight=17;[void]$g.Columns.Add($c)
 $c=New-Object System.Windows.Forms.DataGridViewTextBoxColumn;$c.Name='Label';$c.HeaderText='Pole w PRF';$c.ReadOnly=$true;$c.FillWeight=22;[void]$g.Columns.Add($c)
 $c=New-Object System.Windows.Forms.DataGridViewComboBoxColumn;$c.Name='Mode';$c.HeaderText='Tryb';$c.FillWeight=12;foreach($v in @('Z bazy','Stala','Recznie','Pomin')){[void]$c.Items.Add($v)};[void]$g.Columns.Add($c)
 $c=New-Object System.Windows.Forms.DataGridViewComboBoxColumn;$c.Name='Source';$c.HeaderText='Kolumna w bazie';$c.FillWeight=38;[void]$c.Items.Add('(brak)')
 $options=@{};for($i=1;$i -le 80;$i++){$name='';if($header.ContainsKey($i)){$name=[string]$header[$i]};$opt=('{0:000} | {1} | {2}' -f $i,(Letter $i),$name);$options[$i]=$opt;[void]$c.Items.Add($opt)};[void]$g.Columns.Add($c)
 $c=New-Object System.Windows.Forms.DataGridViewTextBoxColumn;$c.Name='Value';$c.HeaderText='Stala';$c.FillWeight=15;[void]$g.Columns.Add($c)
 foreach($s in $spec){
  $key=$s[0]
  $num=if($key -eq 'Annual'){Auto-AnnualCol $header}else{Auto-Col $header $key $s[1] ([int]$s[2])}
  $mode=if($num -gt 0){'Z bazy'}else{'Recznie'};$value=''
  # New headers may need manual mapping; missing print code is validated per generated record.
  if($previous.ContainsKey($key) -and $key -ne 'PrintCode'){
   $p=$previous[$key]
   # Earlier mappings may have Annual=Recznie because automatic FC detection was unavailable.
   # Migrate that default once, but keep genuinely configured values and sources.
   if($key -eq 'Annual' -and $p.Mode -eq 'Recznie' -and [int]$p.NumerKolumny -eq 0 -and [string]::IsNullOrWhiteSpace($p.Stala)){
    # Keep new automatic detection
   } else {$mode=$p.Mode;$num=[int]$p.NumerKolumny;$value=$p.Stala
    if($mode -eq 'Z bazy' -and $num -gt 0 -and $p.PSObject.Properties.Name -contains 'NaglowekZrodla' -and -not [string]::IsNullOrWhiteSpace([string]$p.NaglowekZrodla)){
     $expected=Normalize-Head $p.NaglowekZrodla
     if(-not $header.ContainsKey($num) -or (Normalize-Head $header[$num]) -ne $expected){
      $matches=@($header.Keys | Where-Object {(Normalize-Head $header[$_]) -eq $expected} | Sort-Object {[int]$_})
      if($matches.Count -eq 1){$num=[int]$matches[0]}else{$num=0;$mode='Recznie'}
     }
    }}
  }
  $source=if($options.ContainsKey($num)){$options[$num]}else{'(brak)'}
  $visibleLabel=if($key -eq 'PrintCode'){'Component Master Part Number (B3) = New OC print code'}else{$s[1]}
  $index=$g.Rows.Add($key,$visibleLabel,$mode,$source,$value)
  if($key -eq 'PrintCode'){$g.Rows[$index].DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(230,245,230)}
 }
 $panel=New-Object System.Windows.Forms.Panel;$panel.Dock='Bottom';$panel.Height=52
 $ok=New-Object System.Windows.Forms.Button;$ok.Text='Zapisz i generuj';$ok.Width=170;$ok.Left=10;$ok.Top=8;$ok.Height=33;$ok.Add_Click({try{$g.EndEdit();$f.DialogResult='OK';$f.Close()}catch{}})
 $cancel=New-Object System.Windows.Forms.Button;$cancel.Text='Anuluj';$cancel.Width=100;$cancel.Left=195;$cancel.Top=8;$cancel.Height=33;$cancel.Add_Click({$f.DialogResult='Cancel';$f.Close()})
 $panel.Controls.Add($ok);$panel.Controls.Add($cancel);$f.Controls.Add($g);$f.Controls.Add($panel);$f.Controls.Add($lbl)
 if($f.ShowDialog() -ne 'OK'){return $null};$mapped=@{};$lines=@()
 foreach($gr in $g.Rows){$key=[string]$gr.Cells['Key'].Value;$mode=[string]$gr.Cells['Mode'].Value;$src=[string]$gr.Cells['Source'].Value;$value=[string]$gr.Cells['Value'].Value;$num=0;if($src -match '^(\d+) \|'){$num=[int]$Matches[1]};if($mode -eq 'Z bazy' -and $num -eq 0){$mode='Recznie'};$sourceHeader=if($header.ContainsKey($num)){[string]$header[$num]}else{''};$item=[pscustomobject]@{Klucz=$key;Mode=$mode;NumerKolumny=$num;NaglowekZrodla=$sourceHeader;Stala=$value};$mapped[$key]=[pscustomobject]@{Mode=$mode;Col=$num;Stala=$value};$lines+=$item}
 $lines|Export-Csv $config -Delimiter ';' -NoTypeInformation -Encoding UTF8
 return $mapped
}
try{
 $db=Ask-File 'Wskaz plik bazy Octopus';if(-not $db){exit};$output=Ask-Folder 'Wybierz folder docelowy';if(-not $output){exit}
 # ZAWSZE przetwarzaj wszystkie rozpoznane typy: Tray, American Case i Diecut.
 $dbTmp=Join-Path ([System.IO.Path]::GetTempPath()) ('cedo_db_'+[guid]::NewGuid().ToString('N'))
 [void](New-Item $dbTmp -ItemType Directory);try{Unzip $db $dbTmp;$rows=Read-WorkbookData $dbTmp}finally{Remove-Item $dbTmp -Recurse -Force}
 if($rows.Count -lt 2){throw 'Baza nie zawiera danych pod naglowkami.'}
 # Different source schemas MUST NOT share a mapping file.
 $headerSignature=@($rows[0].V.Keys | Sort-Object | ForEach-Object { ([string]$_)+'='+([string]$rows[0].V[$_]) }) -join '|'
 $sha=[System.Security.Cryptography.SHA256]::Create()
 try{$hashBytes=$sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($headerSignature));$schemaId=([System.BitConverter]::ToString($hashBytes)).Replace('-','').Substring(0,12)}finally{$sha.Dispose()}
 $mapPath=Join-Path $base ('MAPOWANIE_'+$schemaId+'.csv')
 $script:Mapping=Configure $rows[0].V $mapPath;if($null -eq $script:Mapping){exit}
 foreach($required in @('PartNo','Name','Type')){if($script:Mapping[$required].Mode -eq 'Z bazy' -and $script:Mapping[$required].Col -le 0){throw "Brak zrodla dla wymaganego pola $required"}}

 $run=Join-Path $output ('CEDO_WYNIKI_'+(Get-Date -Format 'yyyyMMdd_HHmmss'));[void](New-Item $run -ItemType Directory)
 $script:Log.Add('Wiersz;PartNo;Typ;Status;NewOCPrintCode;Uwagi')
 $script:Log.Add(('0;;BAZA;INFO;;Arkusz zrodlowy: '+$script:DetectedSheet+' | wiersz naglowkow: '+$script:DetectedHeaderRow))
 foreach($row in $rows){if($row.No -le $script:DetectedHeaderRow){continue}
  $part=Field $row 'PartNo';if(-not (Has $part)){continue};$printCode=Field $row 'PrintCode';$type=Field $row 'Type';$kind=''
  if($type -match '(?i)tray'){$kind='Tray'}elseif($type -match '(?i)american|case'){$kind='American_Case'}elseif($type -match '(?i)die.?cut'){$kind='Diecut'}
  if(-not $kind){$script:Log.Add(($row.No.ToString()+';'+$part+';NIEZNANY;POMINIETY;'+$printCode+';Nierozpoznany typ: '+$type));continue}
  $warnings=''
  foreach($colorKey in @('External','Internal')){
   $original=([string](Field $row $colorKey)).Trim()
   if((Board-Colour $row $colorKey) -eq 'Brown' -and $original -notmatch '^(?i:brown)$'){$warnings+=($colorKey+" '"+$original+"' => Brown; ")}
  }
  if($kind -eq 'Tray' -and (Has (Field $row 'LidID'))){
   foreach($colorKey in @('LidExternal','LidInternal')){
    $original=([string](Field $row $colorKey)).Trim()
    if((Board-Colour $row $colorKey) -eq 'Brown' -and $original -notmatch '^(?i:brown)$'){$warnings+=($colorKey+" '"+$original+"' => Brown; ")}
   }
  }
  $stackOriginal=([string](Field $row 'Stackable')).Trim()
  if((Stackable-Value $row) -eq 'no' -and $stackOriginal -notmatch '^(?i:no|n|false|nie|0)$'){$warnings+=("Stackable '"+$stackOriginal+"' => no; ")}
  foreach($needed in @('Name','Length','Width','Height','Print','Cutter')){if(-not (Has (Field $row $needed))){$warnings+="Brak pola $needed; "}}
  if($kind -eq 'Tray' -and (Has (Field $row 'LidID'))){foreach($needed in @('LidName','LidLength','LidWidth','LidHeight')){if(-not (Has (Field $row $needed))){$warnings+="Brak pola $needed; "}}}
  $temp=Join-Path ([System.IO.Path]::GetTempPath()) ('cedo_one_'+[guid]::NewGuid().ToString('N'))
  try{
   $tpl=Join-Path $base ('templates\'+$kind+'.xlsm');if(-not(Test-Path $tpl)){throw "Brak szablonu $tpl"}
   [void](New-Item $temp -ItemType Directory);Unzip $tpl $temp
   $path=Sheet-Path $temp 'outer carton';$d=Read-Xml $path
   # Component Master Part Number ALWAYS receives New OC print code, not Part No.
   if(Has $printCode){Set-Cell $d 'B3' $printCode}else{throw ('Wiersz '+$row.No+': puste New OC print code. Nie generuje PRF z pustym Component Master Part Number.')}
   Set-Cell $d 'B5' (Field $row 'Name');Set-Cell $d 'B6' $type
   $cutter=Field $row 'Cutter';if(Has $cutter){Set-Cell $d 'B7' 'Existing';Set-Cell $d 'B8' $cutter}else{Set-Cell $d 'B7' 'New'}
   Set-Cell $d 'B9' (Field $row 'Length');Set-Cell $d 'B10' (Field $row 'Width');Set-Cell $d 'B11' (Field $row 'Height')
   Set-Cell $d 'B13' (Field $row 'Colours')
   $printed=Field $row 'Printed';if($printed -match '(?i)yes|printed'){Set-Cell $d 'B16' 0.8}elseif($printed -match '(?i)^no$|perforat'){Set-Cell $d 'B16' 0}
   Set-Cell $d 'B20' (Field $row 'Fluting')
   if($kind -eq 'Tray' -and (Has (Field $row 'LidID'))){
    $lc=Field $row 'LidCutter';if(Has $lc){Set-Cell $d 'D7' 'Existing';Set-Cell $d 'D8' $lc}
    Set-Cell $d 'D9' (Field $row 'LidLength');Set-Cell $d 'D10' (Field $row 'LidWidth');Set-Cell $d 'D11' (Field $row 'LidHeight')
    Set-Cell $d 'D13' (Field $row 'LidColours')
    $lp=Field $row 'LidPrinted';if($lp -match '(?i)printed|yes'){Set-Cell $d 'D16' 0.8}elseif($lp -match '(?i)perforat|^no$'){Set-Cell $d 'D16' 0}
    Set-Cell $d 'D20' (Field $row 'LidFluting')
   }
   # Logistics: target addresses follow the original template. Do not overwrite formulas.
   Set-Cell $d 'C29' (Field $row 'Pieces');Set-Cell $d 'C30' (Field $row 'Gross');Set-Cell $d 'C31' (Field $row 'Carries')
   Set-Cell $d 'C32' (Field $row 'LayerCartons');Set-Cell $d 'C33' (Field $row 'Layers')
   if((Has (Field $row 'LayerCartons')) -and (Has (Field $row 'Layers'))){$product=[int][double](Field $row 'LayerCartons') * [int][double](Field $row 'Layers');Set-FormulaCache $d 'C34' $product}
   Set-Cell $d 'C35' (Field $row 'Pallet');Set-Cell $d 'C36' (Field $row 'PalletHeight');Set-Cell $d 'C37' (Stackable-Value $row);Set-Cell $d 'C38' (Field $row 'Area')
   # Required annual volume: ALWAYS round UP to complete cartons.
   # Source can be fractional and Excel provides invariant numeric values.
   $annual=([string](Field $row 'Annual')).Trim()
   if(-not [string]::IsNullOrWhiteSpace($annual)){
    $annualNumber=[decimal]0
    $styles=[System.Globalization.NumberStyles]::Float
    $ok=[decimal]::TryParse($annual,$styles,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$annualNumber)
    if(-not $ok){$ok=[decimal]::TryParse($annual,$styles,[System.Globalization.CultureInfo]::GetCultureInfo('pl-PL'),[ref]$annualNumber)}
    if($ok -and $annualNumber -ge 0){
     $fullCartons=[decimal]::Ceiling($annualNumber)
     Set-Cell $d 'B27' ($fullCartons.ToString('0',[System.Globalization.CultureInfo]::InvariantCulture))
     if($fullCartons -ne $annualNumber){$warnings+=("Annual volume: " + $annual + " -> " + $fullCartons.ToString('0') + " (zaokraglono w gore); ")}
    }else{$warnings+="Niepoprawny Required annual volume: '$annual'; "}
   }else{$warnings+='Brak Required annual volume (B27); '}
   Save-Xml $d $path
   Set-Options $temp $row $kind ([ref]$warnings)
   $folder=Join-Path $run $kind;[void](New-Item $folder -ItemType Directory -Force)
   $safe=([string]$part -replace '[^a-zA-Z0-9_-]','_');$dest=Join-Path $folder ('CEDO_'+$kind+'_'+$safe+'_wiersz'+$row.No+'.xlsm')
   Zip $temp $dest
   $testZip=[System.IO.Compression.ZipFile]::OpenRead($dest)
   try{
    $entries=@($testZip.Entries | ForEach-Object {$_.FullName.Replace('\','/')})
    foreach($mandatory in @('[Content_Types].xml','xl/workbook.xml','xl/vbaProject.bin','xl/drawings/vmlDrawing1.vml')){
     if($mandatory -notin $entries){throw ('Niekompletny wynik: brak '+$mandatory)}
    }
    if($entries.Count -lt 20){throw 'Niekompletny wynik: za malo elementow XLSM'}
   }finally{$testZip.Dispose()}
   $script:Count++
   $script:Log.Add(($row.No.ToString()+';'+$safe+';'+$kind+';OK;'+$printCode+';'+($warnings -replace ';',' | ')))
  }catch{$script:Fail++;$script:Log.Add(($row.No.ToString()+';'+$part+';'+$kind+';BLAD;'+$printCode+';'+($_.Exception.Message -replace ';',' | ')))}
  finally{if(Test-Path $temp){Remove-Item $temp -Recurse -Force}}
 }
 if($script:Count -eq 0){$script:Log.Add('0;;BRAK;UWAGA;;Nie wygenerowano zadnego formularza - sprawdz mapowanie typu opakowania')}
 $report=Join-Path $run 'RAPORT_CEDO.csv';$script:Log|Set-Content $report -Encoding UTF8
 [System.Windows.Forms.MessageBox]::Show("Utworzono: $script:Count; bledy: $script:Fail.`nFolder: $run`nRaport: RAPORT_CEDO.csv`n`nWeryfikuj kontrolki w Excelu przed uzyciem produkcyjnym.",'CEDO - generowanie zakonczone')|Out-Null
}catch{[System.Windows.Forms.MessageBox]::Show("Blad: $($_.Exception.Message)`n`nGenerator nie korzysta z Excel COM.",'CEDO - blad')|Out-Null;throw}
