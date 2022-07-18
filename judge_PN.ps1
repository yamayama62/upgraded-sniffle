using namespace Microsoft.VisualBasic
Set-PSDebug -Strict
Add-Type -Assembly Microsoft.VisualBasic

function Is-Prime_number([System.Int64]$n){
    $sqr = [System.Math]::Sqrt($n)
    if ($n -lt 1){
        return $false
    }
    if ($n -eq 2){
        return $true
    }elseif(($n%2) -eq 0){
        return $false
    }

    for ($i=3; $i -le $sqr; $i += 2){
        if (($n % $i) -eq 0){
        return $false
        break
        }

    }
    return $true
}

$n = [Interaction]::InputBox("整数を入力してください:")
[bool]$pn = Is-Prime_number $n
if ($pn -eq $true){
    [Interaction]::MsgBox("素数です。") 
 }else{
    [Interaction]::MsgBox( "素数ではありません。")
 } 