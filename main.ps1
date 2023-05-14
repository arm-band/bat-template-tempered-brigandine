# Shift-JIS でコード記述

######################
# 環境設定            #
######################

# ファイル出力時の文字コード設定
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

######################
# グローバル変数      #
######################
[String]$configSamplePath = Join-Path ( Convert-Path . ) 'config.json.sample'
[String]$configPath = Join-Path ( Convert-Path . ) 'config.json'

######################
# 関数               #
######################

##
# Find-ErrorMessage: エラーメッセージ組み立て
#
# @param {Int} $code: エラーコード
# @param {String} $someStr: 一部エラーメッセージで使用する文字列
#
# @return {String} $msg: 出力メッセージ
#
function Find-ErrorMessage([Int]$code, [String]$someStr) {
    $msg = ''
    $errMsgObj = [Hashtable]@{
        # 1x: 設定ファイル系
        11 = '設定ファイル (config.json) が存在しません。'
        12 = 'パラメータ文字列の長さが空です。'
        13 = 'パラメータ includeSunday が Boolean 型ではありません。'
        14 = 'パラメータ lastSunday が Int 型ではありません。'
        15 = 'パラメータ category が Array 型ではありません。'
        16 = 'パラメータ category が 既定のキーを含んでいません。'
        # 2x: 読み込み系
        21 = 'テンプレートディレクトリ ########## にアクセスできませんでした。'
        22 = 'テンプレートファイル ########## にアクセスできませんでした。'
        # 3x: 出力系
        31 = '出力先ディレクトリ ########## にアクセスできませんでした。'
        # 4x: 入力系
        41 = 'パラメータが数値ではありません。'
        42 = 'パラメータは 1-52 の間で入力してください。'
        45 = 'パラメータは y(Y) または n(N) で入力してください。'
        # 9x: その他、処理中エラー
        99 = '##########'
    }
    $msg = $errMsgObj[$code]
    if ($someStr.Length -gt 0) {
        $msg = $msg.Replace('##########', $someStr)
    }

    return $msg
}
##
# Show-ErrorMessage: エラーメッセージ出力
#
# @param {Int} $code: エラーコード
# @param {Boolean} $exitFlag: exit するかどうかのフラグ
# @param {String} $someStr: 一部エラーメッセージで使用する文字列
#
function Show-ErrorMessage([Int]$code, [Boolean]$exitFlag, [String]$someStr) {
    $msg = Find-ErrorMessage $code $someStr
    Write-Host('ERROR ' + $code + ': ' + $msg) -BackgroundColor DarkRed
    Write-Host `r`n

    if ($exitFlag) {
        exit
    }
}

##
# Assert-ParamStrGTZero: パラメータ文字列長さチェック
#
# @param {String} $paramStr: ファイルパス
#
# @return {Boolean} : 文字長が0より大きければ True, そうでなければ False
#
function Assert-ParamStrGTZero([String]$paramStr) {
    return ($paramStr.Length -gt 0)
}
##
# Assert-ParamArrayHasKeys: パラメータ配列のキーチェック
#
# @param {Object} $paramObj: ファイルパス
#
# @return {Boolean} : 配列が name, slug の2つのキーを持っていれば True, そうでなければ False
#
function Assert-ParamArrayHasKeys([Object]$paramObj) {
    if (-not ($paramObj -is [Object])) {
        Show-ErrorMessage 15 $True ''
    }
    for ($i = 0; $i -lt $paramObj.Length; $i++) {
        if (-not ($paramObj[$i] -is [Object])) {
            Show-ErrorMessage 15 $True ''
        }
        if($paramObj[$i].name -eq $null -Or $paramObj[$i].slug -eq $null) {
            return $False
        }
    }
    return $True
}
##
# Assert-ExistFile: ファイル存在チェック
#
# @param {String} $filePath: ファイルパス
#
# @return {Boolean} : ファイルが存在すれば True, そうれなければ False
#
function Assert-ExistFile([String]$filePath) {
    return (Test-Path $filePath)
}

######################
# main process       #
######################

if (Assert-ExistFile $configPath) {
    Write-Host '処理を開始します ...'
    Write-Host `r`n

    # 設定ファイル読み込み
    $configData = Get-Content -Path $configPath -Raw -Encoding UTF8 | ConvertFrom-JSON
    ######################
    # check              #
    ######################
    # default
    if (-not ($configData.default.includeSunday -is [Boolean])) {
        Show-ErrorMessage 13 $True ''
    }
    if (-not ($configData.default.lastSunday -is [Int])) {
        Show-ErrorMessage 14 $True ''
    }
    if (-not (Assert-ParamArrayHasKeys $configData.category)) {
        Show-ErrorMessage 16 $True ''
    }
    # テンプレート読み込み元ディレクトリ
    if (-not (Assert-ParamStrGTZero $configData.path.templateDir)) {
        Show-ErrorMessage 21 $True ''
    }
    elseif (-not (Assert-ExistFile $configData.path.templateDir)) {
        Show-ErrorMessage 21 $True $configData.path.templateDir
    }
    # テンプレート読み込み元ファイル名
    if (-not (Assert-ParamStrGTZero $configData.path.templateFile)) {
        Show-ErrorMessage 12 $True ''
    }
    elseif (-not (Assert-ExistFile (Join-Path $configData.path.templateDir $configData.path.templateFile))) {
        Show-ErrorMessage 22 $True (Join-Path $configData.path.templateDir $configData.path.templateFile)
    }
    # 出力先ディレクトリ
    if (-not (Assert-ParamStrGTZero $configData.path.distDir)) {
        Show-ErrorMessage 12 $True ''
    }
    elseif (-not (Assert-ExistFile $configData.path.distDir)) {
        Show-ErrorMessage 31 $True $configData.path.distDir
    }

    ######################
    # process            #
    ######################
    $period = 0
    # 直前の日曜日は当日を除くので7,巻き戻るので負の整数とする
    $lastSunday = -7
#    $date =  '2023/05/10'
    $todaysDayOfWeek = (Get-Date).DayOfWeek.value__
#    $todaysDayOfWeek = (Get-Date $date).DayOfWeek.value__
    if (-not ($todaysDayOfWeek -eq 0)) {
        # 日曜日(0)以外ならば曜日の日付を負の整数にする
        $lastSunday = -1 * $todaysDayOfWeek
    }
    # $lastSunday の日数分遡った日付とする
    Write-Host('直前の日曜日は' + (Get-Date).AddDays($lastSunday).ToString("MM/dd") + '、')
    Write-Host([String]$configData.default.lastSunday + '週間前の日曜日は' + (Get-Date).AddDays($lastSunday * $configData.default.lastSunday).ToString("MM/dd") + 'です。')
    Write-Host `r`n
    # 何週間前かはデフォルト値をセット
    $beforeSunday = $configData.default.lastSunday
    $beforeSundayStr = Read-Host('何週間分生成しますか？ [1-52], default:' + [String]$configData.default.lastSunday + ' ')
    if (Assert-ParamStrGTZero $beforeSundayStr) {
        # 入力文字列がある場合は値を書き換える
        if (-not ($beforeSundayStr -match '\d')) {
            # 数値でない場合
            Show-ErrorMessage 41 $True ''
        }
        $beforeSundayTemp = [Int]$beforeSundayStr
        if($beforeSundayTemp -lt 1 -Or $beforeSundayTemp -gt 53) {
            # 範囲外の場合
            Show-ErrorMessage 42 $True ''
        }
        # エラーにならなかった場合は値を書き換える
        $beforeSunday = $beforeSundayTemp
    }
    Write-Host([String]$beforeSunday + '週間前の日曜日は' + (Get-Date).AddDays($lastSunday * $beforeSunday).ToString("MM/dd") + 'です。')
    Write-Host `r`n
    # 日曜日を含むかどうかはデフォルト値をセット
    $includeSunday = $configData.default.includeSunday
    $includeSundayStr = Read-Host((Get-Date).AddDays($lastSunday * $beforeSunday).ToString("MM/dd") + 'を生成する日付に含めますか？ [yes(y)/no(n)], default:' + $configData.default.includeSunday + ' ')
    if (Assert-ParamStrGTZero $includeSundayStr) {
        # 入力文字列があった場合に書き換え
        if (-not ($includeSundayStr -eq 'y') -And -not ($includeSundayStr -eq 'Y') -And -not ($includeSundayStr -eq 'n') -And -not ($includeSundayStr -eq 'N')) {
            # y, n 以外の文字列
            Show-ErrorMessage 45 $True ''
        }
        if ($includeSundayStr -eq 'y' -Or $includeSundayStr -eq 'Y') {
            # y, Y ならば $True
            $includeSunday = $True
        }
        else {
            # それ以外は $False
            $includeSunday = $False
        }
    }
    $days = $lastSunday * $beforeSunday
    if (-not $includeSunday) {
        # 日曜日を含まない場合は +1日 する
        $days = $days + 1
    }
    Write-Host((Get-Date).AddDays($days).ToString("MM/dd") + 'から' + (Get-Date).ToString("MM/dd") + 'までの日付を生成します。')
    Write-Host `r`n

    ######################
    # template           #
    ######################
    # make dir
    $distDateDir = Join-Path $configData.path.distDir (Get-Date).ToString("yyyyMMdd")
    if (-not (Assert-ExistFile $distDateDir)) {
        New-Item $distDateDir -ItemType Directory
    }
    for ($i = 0; $i -lt $configData.category.Length; $i++) {
        $template = Get-Content -Path (Join-Path $configData.path.templateDir $configData.path.templateFile) -Raw -Encoding UTF8
        # title
        $categoryNameStr = $configData.category[$i].name
        if (Assert-ParamStrGTZero $categoryNameStr) {
            $categoryNameStr = ' (' + $categoryNameStr + ')'
        }
        $dateStr = ' (' + (Get-Date).AddDays($days).ToString("MM/dd") + '～' + (Get-Date).ToString("MM/dd") + ')'
        $template = $template.Replace('<!-- TITLE -->', '# 雑記' + $categoryNameStr + $dateStr)
        # slug
        $categorySlugStr = $configData.category[$i].slug
        if (Assert-ParamStrGTZero $categorySlugStr) {
            $categorySlugStr = '-' + $categorySlugStr
        }
        $dateStr = '-' + (Get-Date).AddDays($days).ToString("MM-dd") + '-' + (Get-Date).ToString("MM-dd")
        $slugName = 'zakki' + $categorySlugStr + $dateStr
        $template = $template.Replace('<!-- SLUG -->', $slugName)
        # subtitle (date)
        $contentStr = ''
        for ($j = $days; $j -le 0; $j++) {
            $contentStr = $contentStr + '## ' + (Get-Date).AddDays($j).ToString("MM/dd")
            $contentStr = $contentStr + "`r`n`r`n`r`n`r`n"
        }
        $template = $template.Replace('<!-- DATE -->', $contentStr)
        $filename = $slugName + '.md'
        $filepath = Join-Path $distDateDir $filename
        # BOMなし UTF-8 で出力
        $template `
            | % { [Text.Encoding]::UTF8.GetBytes($_) } `
            | Set-Content -Path $filepath -Encoding Byte
    }
}
else {
    Show-ErrorMessage 11 $True $configPath
}
