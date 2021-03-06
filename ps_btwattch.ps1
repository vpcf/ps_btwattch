﻿<#
.SYNOPSIS
    ps_btwattch : REX-BTWATTCH1測定クライアント

.DESCRIPTION
    REX-BTWATTCH1の毎秒測定値を取得してテーブル表示します
    フレンドリ名"WATT CHECKER"のデバイスから対応するBluetooth仮想COMポートを探して接続します
    あらかじめWindows上でREX-BTWATTCH1とペアリングを行い、公式ツールで安定して接続できることを確認してください
    接続に成功すると測定値が出力ウインドウに表示され、実行ディレクトリ下のCSVファイルに出力されます

.EXAMPLE
    .\ps_btwattch.ps1

.NOTES
    LICENSE : MIT
    AUTHOR  : @vpcf90
    VERSION : 20190916
#>

Set-StrictMode -Version Latest

function get_crc8([byte[]]$bytes, [byte]$crc_init=0x00){
    function get_crc1([byte]$crc_init, $times=0){
        if($times -ge 8){
            $crc_init
        }else{
            if($crc_init -band 0x80){
                get_crc1 (($crc_init -shl 1) -bxor 0x85) ($times + 1)
            }else{
                get_crc1 ($crc_init -shl 1) ($times + 1)
            }
        }
    }
    $head, $tail = $bytes
    [byte]$crc_curr = get_crc1 ($crc_init -bxor $head)
    if($null -eq $tail){
        $crc_curr
    }else{
        get_crc8 $tail $crc_curr
    }
}

function get_port_name{
    # フレンドリ名"WATT CHECKER"の機器をリストし、BDアドレスを取得
    $bt_dev_address =
        Get-WmiObject -Class Win32_PnPEntity |
        Where-Object{$_.Name -match "WATT CHECKER"} |
        Select-Object Description, Name, HardWareID |
        Out-GridView -Title "接続する機器を選択してください" -OutputMode Single |
        ForEach-Object{$_.HardWareID -replace ".+(\w{12}).*",'$1'}

    if($null -eq $bt_dev_address){
        Read-Host "No device selected"
        break
    }

    # BDアドレスに対応するCOMポート名を取得
    $com_port_name =
        Get-WmiObject -Class Win32_PnPEntity |
        Where-Object{($_.Name -match "COM") -and ($_.DeviceID -match $bt_dev_address)} |
        ForEach-Object{$_.Name -replace ".+?\((COM\d+).+?",'$1'}

    if($null -eq $com_port_name){
        Read-Host "No device found"
        break
    }

    Write-Output $com_port_name
}

function set_serialport([string]$port_name){
    Write-Host -NoNewline "Configuring $port_name... "
    $port = New-Object System.IO.Ports.SerialPort $port_name,115200
    $port.ReadTimeout = 1000  #ms
    $port.WriteTimeout = 1000 #ms
    Write-Host "done"
    Write-Output $port
}

function communicate($port, [byte[]]$payload, [int]$receive_length){
    [byte]$cmd_header = 0xAA
    [byte]$cmd_lobyte = $payload.length
    [byte]$cmd_hibyte = ($payload.length -shr 8) -band 0xFF
    [byte]$cmd_crc8 = get_crc8 $payload
    [byte[]]$cmd_array = $cmd_header, $cmd_lobyte, $cmd_hibyte, $payload, $cmd_crc8 | ForEach-Object{$_}

    [void]$port.ReadExisting()
    $port.Write($cmd_array, 0, $cmd_array.length)
    $read_length = $port.Read(($buf = New-Object byte[] 256), 0, $receive_length)

    $read_packet = [int[]]$buf[0..($read_length - 1)]
    Write-Output $read_packet
}

function init_wattch1($port){
    Write-Host -NoNewline "Initializing... "
    $timer_payload =
        0x01,
        ($now = Get-Date).Second,
        $now.Minute,
        $now.Hour,
        $now.Day,
        $now.Month,
        ($now.Year%100),
        $now.DayOfWeek

    $init_received = communicate $port $timer_payload 6
    if($init_received[4] -eq 0x00){
        Write-Host "done"
    }else{
        Write-Host "failed"
    }
}

function start_measure($port){
    Write-Host -NoNewline "Starting... "
    $start_received = communicate $port 0x02,0x1e 6
    if($start_received[4] -eq 0x00){
        Write-Host "done"
    }else{
        Write-Host "failed"
    }
}

function stop_measure($port){
    Write-Host -NoNewline "Stopping... "
    $stop_received = communicate $port 0x03 6
    if($stop_received[4] -eq 0x00){
        Write-Host "done"
    }else{
        Write-Host "failed"
    }
}

function format_value([int[]]$data){
    if(($data[0] -eq 0xAA) -and ($data[4] -eq 0x00) -and ($data[1] -ne 0x02)){
        $wattage = if($data[13] -le 5){(($data[13] -shl 16) + ($data[12] -shl 8) + $data[11]) * 5/1000}else{0}
        $voltage = (($data[10] -shl 16) + ($data[9] -shl 8) + $data[8]) * 1/1000
        $current = (($data[7] -shl 16) + ($data[6] -shl 8) + $data[5]) * 1/128
        $value = [PSCustomObject]@{
            'Datetime' = [DateTime]::New((2000 + $data[19]), $data[18], $data[17], $data[16], $data[15], $data[14]);
            'Wattage(W)' = [Double]$wattage;
            'Voltage(V)' = [Double]$voltage;
            'Current(mA)' = [Double]$current
        }
        Write-Output $value
    }
}

function request_measure($port){
    $measured_data = communicate $port 0x08 21
    $value = format_value $measured_data
    Write-Output $value
}

function make_thread($cmd, $arg){
    $functions = Get-ChildItem function: | ForEach-Object{$_.ScriptBlock.Ast.ToString()}

    $ps = [PowerShell]::Create()
    $ps.AddScript($cmd).AddArgument($arg).AddArgument($functions) | Out-Null

    $result = $ps.BeginInvoke()
    try{
        $ps.EndInvoke($result)
    }catch{
        Write-Output "Measurement halted"
    }
    $ps.Dispose()
}

function resume_measure($port){
    Write-Host "Resuming connection..."
    $port.close()
    $port.open()
    init_wattch1 $port
    start_measure $port
}

function measure_value($port){
    $outname = Get-Date -Format "'.\\'yyyyMMdd_HHmmss'.csv'"

    # 1秒おきに測定値の取得
    $port | ForEach-Object{
        while($true){
            Start-Sleep -Milliseconds (1e3 - (Get-Date).Millisecond)

            try{
                $current_value = request_measure $_
            }catch [TimeoutException]{
                Write-Host "Connection timed out"
                resume_measure $_
                continue
            }catch{
                Write-Host "Invalid Value Received"
                continue
            }
            
            if($current_value){
                Write-Output $current_value
                $current_value | Export-Csv -Path $outname -NoTypeInformation -Append -Encoding "UTF8"
            }                
        }
    } | Out-GridView -Title "REX-BTWATTCH1"
}

$call_measure = {
    param($port, $functions)

    $func = Get-ChildItem function: | ForEach-Object{$_.ScriptBlock.Ast.ToString()}
    compare-object $func $functions | ForEach-Object inputobject | Invoke-Expression

    measure_value $port
}

$COMport_name = get_port_name
$wattch1 = set_serialport $COMport_name

try{
    $wattch1.open()
    Write-Host "Successfully opened device"
}catch{
    Read-Host "Failed to open device"
    break
}

init_wattch1 $wattch1

make_thread $call_measure $wattch1

$wattch1.close()