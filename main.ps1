# Shift-JIS �ŃR�[�h�L�q

######################
# ���ݒ�            #
######################

# �t�@�C���o�͎��̕����R�[�h�ݒ�
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

######################
# �O���[�o���ϐ�      #
######################
[String]$configSamplePath = Join-Path ( Convert-Path . ) 'config.json.sample'
[String]$configPath = Join-Path ( Convert-Path . ) 'config.json'

######################
# �֐�               #
######################

##
# Find-ErrorMessage: �G���[���b�Z�[�W�g�ݗ���
#
# @param {Int} $code: �G���[�R�[�h
# @param {String} $someStr: �ꕔ�G���[���b�Z�[�W�Ŏg�p���镶����
#
# @return {String} $msg: �o�̓��b�Z�[�W
#
function Find-ErrorMessage([Int]$code, [String]$someStr) {
    $msg = ''
    $errMsgObj = [Hashtable]@{
        # 1x: �ݒ�t�@�C���n
        11 = '�ݒ�t�@�C�� (config.json) �����݂��܂���B'
        12 = '�p�����[�^������̒�������ł��B'
        13 = '�p�����[�^ includeSunday �� Boolean �^�ł͂���܂���B'
        14 = '�p�����[�^ lastSunday �� Int �^�ł͂���܂���B'
        15 = '�p�����[�^ category �� Array �^�ł͂���܂���B'
        16 = '�p�����[�^ category �� ����̃L�[���܂�ł��܂���B'
        # 2x: �ǂݍ��݌n
        21 = '�e���v���[�g�f�B���N�g�� ########## �ɃA�N�Z�X�ł��܂���ł����B'
        22 = '�e���v���[�g�t�@�C�� ########## �ɃA�N�Z�X�ł��܂���ł����B'
        # 3x: �o�͌n
        31 = '�o�͐�f�B���N�g�� ########## �ɃA�N�Z�X�ł��܂���ł����B'
        # 4x: ���͌n
        41 = '�p�����[�^�����l�ł͂���܂���B'
        42 = '�p�����[�^�� 1-52 �̊Ԃœ��͂��Ă��������B'
        45 = '�p�����[�^�� y(Y) �܂��� n(N) �œ��͂��Ă��������B'
        # 9x: ���̑��A�������G���[
        99 = '##########'
    }
    $msg = $errMsgObj[$code]
    if ($someStr.Length -gt 0) {
        $msg = $msg.Replace('##########', $someStr)
    }

    return $msg
}
##
# Show-ErrorMessage: �G���[���b�Z�[�W�o��
#
# @param {Int} $code: �G���[�R�[�h
# @param {Boolean} $exitFlag: exit ���邩�ǂ����̃t���O
# @param {String} $someStr: �ꕔ�G���[���b�Z�[�W�Ŏg�p���镶����
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
# Assert-ParamStrGTZero: �p�����[�^�����񒷂��`�F�b�N
#
# @param {String} $paramStr: �t�@�C���p�X
#
# @return {Boolean} : ��������0���傫����� True, �����łȂ���� False
#
function Assert-ParamStrGTZero([String]$paramStr) {
    return ($paramStr.Length -gt 0)
}
##
# Assert-ParamArrayHasKeys: �p�����[�^�z��̃L�[�`�F�b�N
#
# @param {Object} $paramObj: �t�@�C���p�X
#
# @return {Boolean} : �z�� name, slug ��2�̃L�[�������Ă���� True, �����łȂ���� False
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
# Assert-ExistFile: �t�@�C�����݃`�F�b�N
#
# @param {String} $filePath: �t�@�C���p�X
#
# @return {Boolean} : �t�@�C�������݂���� True, ������Ȃ���� False
#
function Assert-ExistFile([String]$filePath) {
    return (Test-Path $filePath)
}

######################
# main process       #
######################

if (Assert-ExistFile $configPath) {
    Write-Host '�������J�n���܂� ...'
    Write-Host `r`n

    # �ݒ�t�@�C���ǂݍ���
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
    # �e���v���[�g�ǂݍ��݌��f�B���N�g��
    if (-not (Assert-ParamStrGTZero $configData.path.templateDir)) {
        Show-ErrorMessage 21 $True ''
    }
    elseif (-not (Assert-ExistFile $configData.path.templateDir)) {
        Show-ErrorMessage 21 $True $configData.path.templateDir
    }
    # �e���v���[�g�ǂݍ��݌��t�@�C����
    if (-not (Assert-ParamStrGTZero $configData.path.templateFile)) {
        Show-ErrorMessage 12 $True ''
    }
    elseif (-not (Assert-ExistFile (Join-Path $configData.path.templateDir $configData.path.templateFile))) {
        Show-ErrorMessage 22 $True (Join-Path $configData.path.templateDir $configData.path.templateFile)
    }
    # �o�͐�f�B���N�g��
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
    # ���O�̓��j���͓����������̂�7,�����߂�̂ŕ��̐����Ƃ���
    $lastSunday = -7
#    $date =  '2023/05/10'
    $todaysDayOfWeek = (Get-Date).DayOfWeek.value__
#    $todaysDayOfWeek = (Get-Date $date).DayOfWeek.value__
    if (-not ($todaysDayOfWeek -eq 0)) {
        # ���j��(0)�ȊO�Ȃ�Ηj���̓��t�𕉂̐����ɂ���
        $lastSunday = -1 * $todaysDayOfWeek
    }
    # $lastSunday �̓������k�������t�Ƃ���
    Write-Host('���O�̓��j����' + (Get-Date).AddDays($lastSunday).ToString("MM/dd") + '�A')
    Write-Host([String]$configData.default.lastSunday + '�T�ԑO�̓��j����' + (Get-Date).AddDays($lastSunday * $configData.default.lastSunday).ToString("MM/dd") + '�ł��B')
    Write-Host `r`n
    # ���T�ԑO���̓f�t�H���g�l���Z�b�g
    $beforeSunday = $configData.default.lastSunday
    $beforeSundayStr = Read-Host('���T�ԕ��������܂����H [1-52], default:' + [String]$configData.default.lastSunday + ' ')
    if (Assert-ParamStrGTZero $beforeSundayStr) {
        # ���͕����񂪂���ꍇ�͒l������������
        if (-not ($beforeSundayStr -match '\d')) {
            # ���l�łȂ��ꍇ
            Show-ErrorMessage 41 $True ''
        }
        $beforeSundayTemp = [Int]$beforeSundayStr
        if($beforeSundayTemp -lt 1 -Or $beforeSundayTemp -gt 53) {
            # �͈͊O�̏ꍇ
            Show-ErrorMessage 42 $True ''
        }
        # �G���[�ɂȂ�Ȃ������ꍇ�͒l������������
        $beforeSunday = $beforeSundayTemp
    }
    Write-Host([String]$beforeSunday + '�T�ԑO�̓��j����' + (Get-Date).AddDays($lastSunday * $beforeSunday).ToString("MM/dd") + '�ł��B')
    Write-Host `r`n
    # ���j�����܂ނ��ǂ����̓f�t�H���g�l���Z�b�g
    $includeSunday = $configData.default.includeSunday
    $includeSundayStr = Read-Host((Get-Date).AddDays($lastSunday * $beforeSunday).ToString("MM/dd") + '�𐶐�������t�Ɋ܂߂܂����H [yes(y)/no(n)], default:' + $configData.default.includeSunday + ' ')
    if (Assert-ParamStrGTZero $includeSundayStr) {
        # ���͕����񂪂������ꍇ�ɏ�������
        if (-not ($includeSundayStr -eq 'y') -And -not ($includeSundayStr -eq 'Y') -And -not ($includeSundayStr -eq 'n') -And -not ($includeSundayStr -eq 'N')) {
            # y, n �ȊO�̕�����
            Show-ErrorMessage 45 $True ''
        }
        if ($includeSundayStr -eq 'y' -Or $includeSundayStr -eq 'Y') {
            # y, Y �Ȃ�� $True
            $includeSunday = $True
        }
        else {
            # ����ȊO�� $False
            $includeSunday = $False
        }
    }
    $days = $lastSunday * $beforeSunday
    if (-not $includeSunday) {
        # ���j�����܂܂Ȃ��ꍇ�� +1�� ����
        $days = $days + 1
    }
    Write-Host((Get-Date).AddDays($days).ToString("MM/dd") + '����' + (Get-Date).ToString("MM/dd") + '�܂ł̓��t�𐶐����܂��B')
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
        $dateStr = ' (' + (Get-Date).AddDays($days).ToString("MM/dd") + '�`' + (Get-Date).ToString("MM/dd") + ')'
        $template = $template.Replace('<!-- TITLE -->', '# �G�L' + $categoryNameStr + $dateStr)
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
        # BOM�Ȃ� UTF-8 �ŏo��
        $template `
            | % { [Text.Encoding]::UTF8.GetBytes($_) } `
            | Set-Content -Path $filepath -Encoding Byte
    }
}
else {
    Show-ErrorMessage 11 $True $configPath
}
