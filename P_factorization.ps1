using namespace Microsoft.VisualBasic
Set-PSDebug -Strict
Add-Type -Assembly Microsoft.VisualBasic

function Factorization_Prime_Number([long]$n){
    if($n -ge 2){
        
        [System.String]$str = ""
        $sqr = [System.Math]::Sqrt($n)
        while(($n%2) -eq 0 -and 2 -le $sqr){
            $str += "2 × "
            [long]$n = $n / 2
            $sqr = [System.Math]::Sqrt($n)
        }

        for($i=3; $i -le $sqr; $i += 2){
            while (($n % $i) -eq 0 -and $i -le $sqr){
                $str += [string]$i + " × "
                [long]$n = $n / $i
                $sqr = [System.Math]::Sqrt($n)
            }
        }
        if ($str-ne ""){
            $str += [string]$n
        }
        return $str
    }else{
    return "2以上の整数を入力してください。"
    }
}

[int16]$j = 1
while($j -le 10){
    $n = [Interaction]::InputBox("整数を入力してください:")
    if ($n -eq ""){
        break
    }
    $str = Factorization_Prime_Number $n

    if ($str -eq ""){
        [Interaction]::MsgBox("素数です。")
     }elseif($str -eq "2以上の整数を入力してください。"){
        [Interaction]::MsgBox($str)
     }else{
        [Interaction]::MsgBox( "素数ではありません。`r`n" + $str)
     }
    $j += 1
}