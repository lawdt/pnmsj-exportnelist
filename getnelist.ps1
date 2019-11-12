# скрипт анализирующий базу элементов и выгружающий инвентори по элементам прям как в ручной выгрузке NeList.csv
# создал Лаврухин Дмитрий 09.06.2016
# 20.06.16 успешно протестировано
# 12.08.16 модифицировано
# 19.10.16 скрипт модифицирован, добавлены ePasolink, NEO Extention

# определяем входящие параметры
$pnms_path = "C:\PNMSj" # путь к PNMSj
$exit_file = "C:\PNMSj\NeList.csv" # выходной файл
#определяем массив типов элементов - 
$eqlist =@{
"0" = "PASOLINK V3"
"3" = "PASOLINK V4"
"20" = "PASOLINK+ STM-1"
"60" = "PASOLINK Mx"
"100" = "PASOLINK NEO"
"110" = "5000S Terminal"
"150" = "PASOLINK NEO HP"
"170" = "PASOLINK NEO HP/A"
"200" = "iPASOLINK 200"
"210" = "iPASOLINK 200"
"131272" = "iPASOLINK 200"
"131282" = "iPASOLINK 200"
"400" = "iPASOLINK 400"
"520" = "iPASOLINK EX"
"65646" = "5000S Terminal(2x)"
"16777216" = "NEO Extension"
"33554432" = "ePASOLINK"
}

# вычисляем пути к файлам конфигурации
$Scr_file = $pnms_path + "\config\PnmsScrData.pdt"

# получаем элементы со всеми параметрами
$header_elements = "id","IP Address","Network Element Name","group","d","e","f","g","eqtype_id","oppositeneip","j","PMC Type","l","m","n","o","p","q","r","s","t","u","subnet","w","x","y","z","AA","AB","ac","ad","ae","af","ag","ah","ai","aj","ak","al","am","an","ao","ap","aq","ar","as","at","au","av","aw","ax","ay","az"
$elements = Get-Content $Scr_file|Select-String -Pattern 'Object\d*=NE,'|%{$_ -replace 'Object','' -replace '=NE',''} |ConvertFrom-Csv -Header $header_elements
# получаем подсети
$header_subnets = "id","num","name","network","e","f","g","h","i","j","k","l"
$subnets = Get-Content $Scr_file|Select-String -Pattern 'Object\d*=SUBNET,'|%{$_ -replace 'Object','' -replace '=SUBNET',''}|ConvertFrom-Csv -Header $header_subnets
#Write-Output $subnets
# получаем сети
$header_networks = "id","num","name","region_id","e","f","g"
$networks = Get-Content $Scr_file|Select-String -Pattern 'Object\d*=NETWORK,'|%{$_ -replace 'Object','' -replace '=NETWORK',''}|ConvertFrom-Csv -Header $header_networks
# получаем регионы
$header_regions = "id","num","name","d","e","f","g"
$regions = Get-Content $Scr_file|Select-String -Pattern 'Object\d*=REGION,'|%{$_ -replace 'Object','' -replace '=REGION',''}|ConvertFrom-Csv -Header $header_regions

function getregion_name
{
$test_ne = $args[0]
$test_subnet = $elements|Where-Object {$_."id" -eq $test_ne}|Select-Object subnet
#Write-Host subnet is $test_subnet.subnet
$test_network = $subnets|Where-Object {$_."num" -eq $test_subnet.subnet}|Select-Object network
#Write-Host network is $test_network.network
$test_region = $networks|Where-Object {$_."num" -eq $test_network.network}|Select-Object region_id
#Write-Output region_id is $test_region.region_id
$test_regionname = $regions|Where-Object {$_."num" -eq $test_region.region_id}|Select-Object name
Write-Output $test_regionname.name
}

function getnetwork_name
{
$test_ne = $args[0]
$test_subnet = $elements|Where-Object {$_."id" -eq $test_ne}|Select-Object subnet
#Write-Host subnet is $test_subnet.subnet
$test_network = $subnets|Where-Object {$_."num" -eq $test_subnet.subnet}|Select-Object network
#Write-Host network is $test_network.network
$test_networkname = $networks|Where-Object {$_."num" -eq $test_network.network}|Select-Object name
Write-Output $test_networkname.name
#$test_region = $networks|Where-Object {$_."num" -eq $test_network.network}|Select-Object region_id
#Write-Output region_id is $test_region.region_id
#$test_regionname = $regions|Where-Object {$_."num" -eq $test_region.region_id}|Select-Object name
#Write-Output $test_regionname.name
}

function getequipmenttype
{
$test_ne = $args[0]
$test_eqtype_id_orig = $elements|Where-Object {$_."id" -eq $test_ne}|Select-Object eqtype_id
$test_eqtype_id = $elements|Where-Object {$_."id" -eq $test_ne}|Select-Object "IP Address"
$inventoryfile_path = $pnms_path + "\inventory\" + $test_eqtype_id."IP Address"
$p1 = $inventoryfile_path + "\*.csv"
$marker = 0
#Write-Host $test_eqtype_id_orig.eqtype_id
switch  ($test_eqtype_id_orig.eqtype_id)
    {
        0 { #PASOLINK V3
            if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[0] -Header "a1","a2","a3","a4","a5","a6"
                               if ($t2.a4 -eq "PASOLINK V3")  {
                                                                 $t3 = $t1[17].Split(",")
                                                                 $eq_print = ($t2.a4 + " " + $t3[1])
                                                                 $marker = 1
                                                                 }
                               }
                }
        3 { #PASOLINK V4
            if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[0] -Header "a1","a2","a3","a4","a5","a6"
                               if ($t2.a4 -eq "PASOLINK V4")  {
                                                                 $t3 = $t1[17].Split(",")
                                                                 $eq_print = ($t2.a4 + " " + $t3[29])
                                                                 $marker = 1
                                                                 }
                               }
                }
        {($_ -eq 20) -or ($_ -eq 60)} { # PASOLINK+ STM-1 (20)  or PASOLINK Mx (60)
           # Write-Host "PASOLINK+ STM-1 (20)  or PASOLINK Mx (60)"
            if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[2] -Header "a1","a2","a3","a4","a5","a6"
                               if (($t2.a2 -eq "PASOLINK+ STM-1") -or ($t2.a2 -eq "PASOLINK Mx")) {
                                                                 $eq_print = ($t2.a2 + " " + $t2.a3)|%{$_ -replace '\(',' ' -replace '\)',''}
                                                                 $marker = 1
                                                                 }
                               }
                }
        {($_ -eq 100) -or ($_ -eq 170) } { #PASOLINK NEO or PASOLINK NEO HP/A
        #Write-Host "PASOLINK NEO or PASOLINK NEO HP/A"
                if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[2] -Header "a1","a2","a3","a4","a5","a6"
                               if (($t2.a1 -eq "PASOLINK NEO") -or ($t2.a1 -eq "PASOLINK NEO HP/A")) {
                                                                                                     $eq_print = ($t2.a1 + " " + $t2.a2)|%{$_ -replace '\(',' ' -replace '\)',''}
                                                                                                     $marker = 1
                                                                                                       }
                                    }
                    }
        150 { #PASOLINK NEO HP
                if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[2] -Header "a1","a2","a3","a4","a5","a6"
                               if ($t2.a1 -eq "PASOLINK NEO HP") {
                                                                                                     $eq_print = ($t2.a1 + " " + $t2.a2)
                                                                                                     $marker = 1
                                                                                                       }
                                    }
                    }
        {($_ -eq 200) -or ($_ -eq 210) -or ($_ -eq 131282) -or ($_ -eq 131272)} { # iPASOLINK 200
               # Write-Host "iPASOLINK 200"
                    $eq_print = "iPASOLINK 200"
                    }
        default { #если вдруг встретится другое значение , заполняем из таблички.
        #Write-Host "default"
                $eq_print =  $eqlist[$test_eqtype_id_orig.eqtype_id]
                }
    }

# если все же где-то не нашлось директории или файлов в директории, заполняем из таблички.
if ($marker -eq 0) { $eq_print =  $eqlist[$test_eqtype_id_orig.eqtype_id] }
# выводим результат
Write-Output $eq_print
}


function getoppositenename
{
$test_ne = $args[0]
$test_oppositeneip = $elements|Where-Object {$_."id" -eq $test_ne}|Select-Object oppositeneip
$test_oppositenename = $elements|Where-Object {$_."IP Address" -eq $test_oppositeneip.oppositeneip}|Select-Object "Network Element Name"
Write-Output  $test_oppositenename."Network Element Name"
}

function getconnect
{
$test_ne = $args[0]
$test_eqtype_id = $elements|Where-Object {$_."id" -eq $test_ne}|Select-Object "IP Address"
$inventoryfile_path = $pnms_path + "\inventory\" + $test_eqtype_id."IP Address"
$p1 = $inventoryfile_path + "\*.csv"
if (Test-Path $p1)
    {
    Write-Output Connect
    }
else {
# Если не находим файлов в inventory То также пишем что все хорошо.
     Write-Output Connect

     }
}

function getpmcsw
{
$test_ne = $args[0]
$test_eqtype_id_orig = $elements|Where-Object {$_."id" -eq $test_ne}|Select-Object eqtype_id
$test_eqtype_id = $elements|Where-Object {$_."id" -eq $test_ne}|Select-Object "IP Address"
$inventoryfile_path = $pnms_path + "\inventory\" + $test_eqtype_id."IP Address"
$p1 = $inventoryfile_path + "\*.csv"
$marker = 0
#Write-Host $test_eqtype_id_orig.eqtype_id
switch  ($test_eqtype_id_orig.eqtype_id)
    {
        0 { # PASOLINK V3
            if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[11] -Header "a1","a2","a3","a4","a5","a6","a7","a8","a9","a10","a11","a12"
                               $eq_print = $t2.a2
                               $marker = 1               
                               }
                }
        3 { # PASOLINK V4
            if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[11] -Header "a1","a2","a3","a4","a5","a6","a7","a8","a9","a10","a11","a12"
                               $eq_print = $t2.a2
                               $marker = 1               
                               }
                }
        20 { # PASOLINK+ STM-1 (20)
            if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[11] -Header "a1","a2","a3","a4","a5","a6","a7","a8","a9","a10","a11","a12"
                               $eq_print = $t2.a12
                               $marker = 1               
                               }
                }
        60 { #PASOLINK Mx (60)
            if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[8] -Header "a1","a2","a3","a4","a5","a6","a7","a8","a9","a10","a11","a12"
                               $eq_print = $t2.a4
                               $marker = 1
                                                                 
                               }
                }
        100 { #PASOLINK NEO
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[8].Split(",")
                                    $eq_print = $t2[54]
                                    $marker = 1
                                    }
                    }
        150 { #PASOLINK NEO HP
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[8].Split(",")
                                    $eq_print = $t2[54]
                                    $marker = 1
                                    }
                    }
        170 { #PASOLINK NEO HP/A
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[8].Split(",")
                                    $eq_print = $t2[47]
                                    $marker = 1
                                    }
                    }
        {($_ -eq 200) -or ($_ -eq 210) -or ($_ -eq 131282) -or ($_ -eq 131272)} { # iPASOLINK 200
                        if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[8].Split(",")
                                    $eq_print = $t2[5]
                                    $marker = 1
                                    }
                    }
        400 { #iPASOLINK 400
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[8].Split(",")
                                    $eq_print = $t2[4]
                                    $marker = 1
                                    }
                    }
        520 { #iPASOLINK EX
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[8].Split(",")
                                    $eq_print = $t2[0]
                                    $marker = 1
                                    }
                    }
        65646 { # 5000S Terminal(2x)
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[125].Split(",")
                                    $eq_print = $t2[8]
                                    $marker = 1
                                    }
                    }
        110 { # 5000S Terminal
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[125].Split(",")
                                    $eq_print = $t2[8]
                                    $marker = 1
                                    }
                    }
        default { #если вдруг встретится другое значение 
        #Write-Host "default"
                $eq_print =  $eqlist[$test_eqtype_id_orig.eqtype_id]
                }
    }

# выводим результат
Write-Output $eq_print
}

function getpmcserial
{
$test_ne = $args[0]
$test_eqtype_id_orig = $elements|Where-Object {$_."id" -eq $test_ne}|Select-Object eqtype_id
$test_eqtype_id = $elements|Where-Object {$_."id" -eq $test_ne}|Select-Object "IP Address"
$inventoryfile_path = $pnms_path + "\inventory\" + $test_eqtype_id."IP Address"
$p1 = $inventoryfile_path + "\*.csv"
$marker = 0
#Write-Host $test_eqtype_id_orig.eqtype_id
switch  ($test_eqtype_id_orig.eqtype_id)
    {
        0 { # PASOLINK V3
            if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[11] -Header "a1","a2","a3","a4","a5","a6","a7","a8","a9","a10","a11","a12"
                               $eq_print = $t2.a4
                               $marker = 1               
                               }
                }
        3 { # PASOLINK V4
            if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[11] -Header "a1","a2","a3","a4","a5","a6","a7","a8","a9","a10","a11","a12"
                               $eq_print = $t2.a4
                               $marker = 1               
                               }
                }
        20 { # PASOLINK+ STM-1 (20)
            if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[11] -Header "a1","a2","a3","a4","a5","a6","a7","a8","a9","a10","a11","a12"
                               $eq_print = $t2.a6
                               $marker = 1               
                               }
                }
        60 { #PASOLINK Mx (60)
            if (Test-Path $p1) {
                               $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                               $inventoryfiletarget = $inventoryfile[0].FullName
                               $t1 = Get-Content $inventoryfiletarget
                               $t2 = ConvertFrom-Csv $t1[8] -Header "a1","a2","a3","a4","a5","a6","a7","a8","a9","a10","a11","a12"
                               $eq_print = $t2.a2
                               $marker = 1
                                                                 
                               }
                }
        100 { #PASOLINK NEO 
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[8].Split(",")
                                   $eq_print = $t2[12]
                                   $marker = 1
                                   }
                    }
        150 { #PASOLINK NEO HP
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[8].Split(",")
                                   $eq_print = $t2[12]
                                   $marker = 1
                                   }
                    }
        170 { #PASOLINK NEO HP/A
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[8].Split(",")
                                   $eq_print = $t2[10]
                                   $marker = 1
                                   }
                    }
        400 { #iPASOLINK 400
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[8].Split(",")
                                   $eq_print = $t2[2]
                                   $marker = 1
                                   }
                    }
        520 { #iPASOLINK EX
                if (Test-Path $p1) {
                                   $inventoryfile_path = $pnms_path + "\Software\" + $test_eqtype_id."IP Address"
                                   $p1 = $inventoryfile_path + "\*.csv"
                                   if (Test-Path $p1) {
                                                       $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                                       #Write-Host $inventoryfile
                                                       $inventoryfiletarget = $inventoryfile[0].FullName
                                                       #Write-Host $inventoryfiletarget
                                                       $t1 = Get-Content $inventoryfiletarget
                                                       ##Write-Host $t1[0]
                                                       #Write-Host $t1[1]
                                                       $t2 = $t1.Split(",")
                                                       $eq_print = $t2[13]
                                                       $marker = 1
                                                       }
                                   }
                    }
        65646 { # 5000S Terminal(2x)
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[125].Split(",")
                                   $eq_print = $t2[2]
                                   $marker = 1
                                   }
                    }
        110 { # 5000S Terminal
                if (Test-Path $p1) {
                                   $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                   $inventoryfiletarget = $inventoryfile[0].FullName
                                   $t1 = Get-Content $inventoryfiletarget
                                   $t2 = $t1[125].Split(",")
                                   $eq_print = $t2[2]
                                   $marker = 1
                                   }
                    }
        {($_ -eq 200) -or ($_ -eq 210) -or ($_ -eq 131282) -or ($_ -eq 131272)} { # iPASOLINK 200
                if (Test-Path $p1) {
                                       $inventoryfile = Get-ChildItem $inventoryfile_path |Sort-Object -Descending LastWriteTime
                                       $inventoryfiletarget = $inventoryfile[0].FullName
                                       $t1 = Get-Content $inventoryfiletarget
                                       $t2 = ConvertFrom-Csv $t1[8] -Header "a1","a2","a3","a4","a5","a6","a7","a8","a9","a10","a11","a12"
                                       $eq_print = $t2.a3
                                       $marker = 1
                                       }
                                        }
        default { #если вдруг встретится другое значение 
        #Write-Host "default"
                $eq_print =  $eqlist[$test_eqtype_id_orig.eqtype_id]
                }
    }

# выводим результат
Write-Output $eq_print
}

#добавляем необходимый свойства в массив элементов
$elements |Add-Member NoteProperty "Region Name" ([PSObject]$null)
$elements |Add-Member NoteProperty "Network Name" ([PSObject]$null)
$elements |Add-Member NoteProperty "Equipment Type" ([PSObject]$null)
$elements |Add-Member NoteProperty "Opposite Network Element" ([PSObject]$null)
$elements |Add-Member NoteProperty Connect ([PSObject]$null)
$elements |Add-Member NoteProperty Severity ([PSObject]$null)
$elements |Add-Member NoteProperty Maintenance ([PSObject]$null)
$elements |Add-Member NoteProperty "Link Performance Monitor" ([PSObject]$null)
$elements |Add-Member NoteProperty "AUX. I/O" ([PSObject]$null)
$elements |Add-Member NoteProperty "Commissioning Data" ([PSObject]$null)
$elements |Add-Member NoteProperty "Commissioning Data(RX Level)" ([PSObject]$null)
$elements |Add-Member NoteProperty "Link Performance Data Collect" ([PSObject]$null)
$elements |Add-Member NoteProperty "PMC S/W Version" ([PSObject]$null)
$elements |Add-Member NoteProperty "Backward PMC S/W Version" ([PSObject]$null)
$elements |Add-Member NoteProperty "PMC Serial No." ([PSObject]$null)
$elements |Add-Member NoteProperty Comment ([PSObject]$null)
$elements |Add-Member NoteProperty "PNMTj Connect Status" ([PSObject]$null)

# вычисляем и добавляем свойство имени региона для каждого объекта элемента.
function add_info
{
$elements|ForEach-Object {
                        $_."Region Name" = (getregion_name $_.id)
                        $_."Network Name" = (getnetwork_name $_.id)
                        $_."Equipment Type" = (getequipmenttype $_.id)
                        $_."Opposite Network Element" = (getoppositenename $_.id)
                        $_.Connect = (getconnect $_.id)    
                        $_.Severity = "Normal"
                        $_.Maintenance = "Off"
                        $_."Link Performance Monitor" = "Normal"
                        $_."AUX. I/O" = "Normal"
                        $_."Commissioning Data" = ""
                        $_."Commissioning Data(RX Level)" = ""
                        $_."Link Performance Data Collect" = "Collect"
                        $_."PMC S/W Version" = (getpmcsw $_.id)
                        $_."Backward PMC S/W Version" = $_."PMC S/W Version"
                        $_."PMC Serial No." = (getpmcserial $_.id)
                        $_.Comment = ""
                        $_."PNMTj Connect Status" = ""
                        }
$elements|ForEach-Object { if ($_.Connect -eq "Disconnect") {
                                                            $_.Severity = "---"
                                                            $_.Maintenance = "---"
                                                            $_."Link Performance Monitor" = "---"
                                                            $_."AUX. I/O" = "---"
                                                            $_."PMC S/W Version" = "---"
                                                            $_."Backward PMC S/W Version" = "---"
                                                            $_."PMC Serial No." = "---"
                                                            }
                          }                     
}

add_info

#выводим список элементов и регионов
#$elements|Select-Object "Network Element Name","Region Name","Network Name","Equipment Type","PMC Type","Opposite Network Element","IP Address",Connect,Severity,Maintenance,"Link Performance Monitor","AUX. I/O","Commissioning Data","Commissioning Data(RX Level)","Link Performance Data Collect","PMC S/W Version","Backward PMC S/W Version","PMC Serial No.",Comment,"PNMTj Connect Status" | ConvertTo-Csv –NoTypeInformation|% { $_ -replace '"',''} |Out-File -Encoding utf8 $exit_file
$elements|Select-Object "Network Element Name","Region Name","Network Name","Equipment Type","PMC Type","Opposite Network Element","IP Address",Connect,Severity,Maintenance,"Link Performance Monitor","AUX. I/O","Commissioning Data","Commissioning Data(RX Level)","Link Performance Data Collect","PMC S/W Version","Backward PMC S/W Version","PMC Serial No.",Comment,"PNMTj Connect Status" | ConvertTo-Csv –NoTypeInformation|% { $_ -replace '"',''}|Out-File -Encoding utf8 $exit_file