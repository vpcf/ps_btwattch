<#
.SYNOPSIS
    ps_btwattch : REX-BTWATTCH1測定値取得ツール

.DESCRIPTION
    REX-BTWATTCH1のリアルタイム値計測値を取得してテーブル表示します
    フレンドリ名"WATT CHECKER"のデバイスから対応するBluetooth仮想COMポートを探して接続します
    あらかじめWindows上でREX-BTWATTCH1とペアリングを行い、公式ツールで安定して接続できることを確認してください
    接続に成功すると測定値が出力ウインドウに表示されるので、適宜Excelやテキストエディタに貼り付けて利用してください

.EXAMPLE
    .\ps_btwattch.ps1

.NOTES
    LICENSE : MIT
    AUTHOR  : @vpcf90
    VERSION : 20190209
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

    # BDアドレスに対応するCOMポート名を取得
    $com_port_name =
        Get-WmiObject -Class Win32_PnPEntity |
        Where-Object{($_.Name -match "COM") -and ($_.DeviceID -match $bt_dev_address)} |
        ForEach-Object{$_.Name -replace ".+?\((COM\d+).+?",'$1'}

    Write-Output $com_port_name
}

function open_serialport([string]$port_name){
    $bt_device = New-Object System.IO.Ports.SerialPort $port_name,115200
    $bt_device.ReadTimeout = 1000  #ms
    $bt_device.WriteTimeout = 1000 #ms
    try{
        $bt_device.Open()
        Write-Output $bt_device
    }catch{
        Write-Host "failed to open device"
        break
    }
}

function communicate($bt_device, [byte[]]$payload, [int]$receive_length){
    [byte]$cmd_header = 0xAA
    [byte]$cmd_lobyte = $payload.length
    [byte]$cmd_hibyte = ($payload.length -shr 8) -band 0xFF
    [byte]$cmd_crc8 = get_crc8 $payload
    [byte[]]$cmd_array = $cmd_header, $cmd_lobyte, $cmd_hibyte, $payload, $cmd_crc8 | ForEach-Object{$_}

    try{
        $bt_device.Write($cmd_array, 0, $cmd_array.length)
    }catch [TimeoutException]{
        $bt_device.close()
        Write-Host "connection timed out. terminated."
        break
    }
    $count = $bt_device.Read(($buf = New-Object byte[] 256), 0, $receive_length)

    Write-Output ([int[]]$buf[0..($count - 1)])
}

function init_wattch1($bt_device){
    Write-Host -NoNewline "initializing... "
    $timer_payload =
        0x01,
        ($now = Get-Date).Second,
        $now.Minute,
        $now.Hour,
        $now.Day,
        $now.Month,
        ($now.Year%100),
        $now.DayOfWeek

    $init_received = communicate $bt_device $timer_payload 6
    if($init_received[4] -eq 0x00){
        Write-Host "done"
    }else{
        Write-Host "failed"
    }
}

function start_measure($bt_device){
    Write-Host -NoNewline "starting... "
    $start_received = communicate $bt_device 0x02,0x1e 6
    if($start_received[4] -eq 0x00){
        Write-Host "done"
    }else{
        Write-Host "failed"
    }
}

function stop_measure($bt_device){
    Write-Host -NoNewline "stopping... "
    $stop_received = communicate $bt_device 0x03 6
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
        try{
            $value = [PSCustomObject]@{
                'Datetime' = [DateTime]::New((2000 + $data[19]), $data[18], $data[17], $data[16], $data[15], $data[14]);
                'Wattage(W)' = [Double]$wattage;
                'Voltage(V)' = [Double]$voltage;
                'Current(mA)' = [Double]$current
            }
        }catch{
            exit
        }
        Write-Output $value
    }
}

function request_measure($bt_device){
    $measured_data = communicate $bt_device 0x08 21
    $value = format_value $measured_data
    Write-Output $value
}

function make_thread($bt_device, $cmd, $function_list){
    $ps = [PowerShell]::Create()
    $ps.AddScript($cmd).AddArgument($bt_device).AddArgument($function_list) | Out-Null

    $result = $ps.BeginInvoke()
    try{
        $ps.EndInvoke($result)
    }catch{
    }
    $ps.Dispose()
}

$measure_value = {
    param($bt_device, $function_list)

    $function_list | Invoke-Expression

    $outname = Get-Date -Format "'.\\'yyyyMMddHHmmss'.csv'"
    $pastsec = (Get-Date).Second

    # 1秒おきに測定値の取得
    &{
        do{
            $nowsec = (Get-Date).Second
            if($nowsec -eq $pastsec){
                Start-Sleep -Milliseconds 10
            }else{
                $pastsec = $nowsec
                ($current_value = request_measure $bt_device)
                $current_value | Export-Csv -Path $outname -Append -NoTypeInformation -Encoding "UTF8"
            }
        }while($true)
    } | Out-GridView -Title "REX-BTWATTCH1"
}

$COMport_name = get_port_name
$wattch1 = open_serialport $COMport_name

init_wattch1 $wattch1
start_measure $wattch1

# 計測用スレッドで使用する関数
$function_list = (
    $function:get_crc8.Ast.ToString(),
    $function:communicate.Ast.ToString(),
    $function:format_value.Ast.ToString(),
    $function:request_measure.Ast.ToString()
)

# 別スレッドで測定値の取得
make_thread $wattch1 $measure_value $function_list

stop_measure $wattch1
$wattch1.close()
