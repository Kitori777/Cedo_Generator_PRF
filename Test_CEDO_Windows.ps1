# CEDO - locally executed diagnostics. Read-only inspection.
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$dlg=New-Object System.Windows.Forms.FolderBrowserDialog
$dlg.Description='Wybierz folder CEDO_WYNIKI zawierajacy wygenerowane PRF-y'
if($dlg.ShowDialog() -ne 'OK'){exit}
$folder=$dlg.SelectedPath
$lines=New-Object System.Collections.Generic.List[string]
$lines.Add('CEDO - RAPORT KONTROLI WINDOWS')
$lines.Add(('Data: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
$lines.Add(('System: '+[Environment]::OSVersion.VersionString))
$lines.Add(('Folder: '+$folder))
$files=@(Get-ChildItem -LiteralPath $folder -Recurse -Filter '*.xlsm' -File)
$lines.Add(('Liczba formularzy: '+$files.Count))
$bad=0;$ok=0;$options=0
foreach($file in $files){
 try{
  $z=[System.IO.Compression.ZipFile]::OpenRead($file.FullName)
  try{
   $rawNames=@($z.Entries | ForEach-Object {$_.FullName})
   $names=@($rawNames | ForEach-Object {$_.Replace('\','/')})
   if(@($rawNames | Where-Object {$_ -match '\\'}).Count -gt 0){$lines.Add(('OSTRZEZENIE: '+$file.Name+' | niestandardowe separatory w ZIP; sprawdzono po normalizacji'))}
   foreach($required in @('[Content_Types].xml','xl/workbook.xml','xl/vbaProject.bin','xl/drawings/vmlDrawing1.vml')){
    if($required -notin $names){throw ('Brak '+$required)}
   }
   $ctl=@($names | Where-Object {$_ -match '^xl/ctrlProps/ctrlProp\d+\.xml$'})
   if($ctl.Count -eq 0){throw 'Brak definicji kontrolek wyboru'}
   $options+=$ctl.Count
   # Inspect each ZIP entry via decompression to detect broken streams.
   foreach($entry in $z.Entries){if($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')){continue};$stream=$entry.Open();try{$buffer=New-Object byte[] 65536;while($stream.Read($buffer,0,$buffer.Length) -gt 0){}}finally{$stream.Dispose()}}
   $ok++
  }finally{$z.Dispose()}
 }catch{$bad++;$lines.Add(('BLAD STRUKTURY: '+$file.Name+' | '+$_.Exception.Message))}
}
$lines.Add(('Struktura XLSM poprawna: '+$ok+', bledy: '+$bad+', definicje kontrolek: '+$options))
$lines.Add('UWAGA: Strukturalny test ZIP nie weryfikuje wygladu ani dzialania formularza w programie Excel.')
$lines.Add('UWAGA: Obecnosc kontrolek w ZIP nie dowodzi poprawnego dzialania klikniec w Excelu.')
$excel=$null;$opened=0;$desktopError=''
try{
 $excel=New-Object -ComObject Excel.Application -ErrorAction Stop
 $excel.Visible=$false;$excel.DisplayAlerts=$false
 $excel.AutomationSecurity=3 # force disable macros
 foreach($file in $files){
  $w=$null
  try{
   $w=$excel.Workbooks.Open($file.FullName,0,$true)
   $sheet=$w.Worksheets.Item('outer carton')
   $code=[string]$sheet.Range('B3').Text
   $name=[string]$sheet.Range('B5').Text
   $shapes=[int]$sheet.Shapes.Count
   if([string]::IsNullOrWhiteSpace($code)){throw 'Pusty Component Master Part Number (B3)'}
   $lines.Add(('Excel READONLY OK: '+$file.Name+' | B3='+$code+' | B5='+$name+' | obiekty='+$shapes))
   $opened++
  }catch{$desktopError+=$file.Name+': '+$_.Exception.Message+'; '}
  finally{if($null -ne $w){$w.Close($false);[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($w)}}
 }
}catch{$desktopError=$_.Exception.Message}
finally{if($null -ne $excel){$excel.Quit();[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel);[GC]::Collect();[GC]::WaitForPendingFinalizers()}}
if($opened -eq 0){$lines.Add('TEST EXCELA: NIE WYKONANO. Brak dzialajacego interfejsu Excel COM / polityka IT. '+$desktopError);$lines.Add('BRAK TESTU EXCELA NIE OZNACZA BLEDU WSZYSTKICH FORMULARZY. Otworz probke recznie w aplikacji Excel.')}
else{$lines.Add(('TEST EXCELA: Otworzono tylko do odczytu '+$opened+' formularzy. '+$desktopError))}
$lines.Add('Wymagane potwierdzenie uzytkownika: klikniecia opcji, niezaleznosc grup i wyglad 1:1 w Excelu.')
$dest=Join-Path $folder 'RAPORT_TEST_WINDOWS.txt'
$lines|Set-Content -LiteralPath $dest -Encoding UTF8
[System.Windows.Forms.MessageBox]::Show(('Weryfikacja zakonczona. Poprawne pliki: '+$ok+'; bledy: '+$bad+'; otwarte w Excelu: '+$opened+'.' + [Environment]::NewLine + 'Raport: '+$dest),'CEDO - test Windows')|Out-Null
