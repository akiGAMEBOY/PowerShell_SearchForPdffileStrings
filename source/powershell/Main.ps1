#################################################################################
# 処理名　｜SearchForPdffileStrings（メイン処理）
# 機能　　｜PDFファイル内の文字列を検索し存否を判定するツール
#--------------------------------------------------------------------------------
# 戻り値　｜下記の通り。
# 　　　　｜   0: 正常終了
# 　　　　｜-001: エラー 設定ファイル読み込み
# 　　　　｜-211: エラー 参照できないファイル
# 　　　　｜-401: エラー テキストデータの書き出し失敗
# 　　　　｜-411: エラー PDFファイル内の処理結果が異常終了
# 　　　　｜-901: エラー メイン - 処理中断
# 引数　　｜-
# 注意事項｜iTextSharp 5.5.13 を使用    : https://www.nuget.org/packages/iTextSharp/5.5.13
# 　　　　｜iTextSharp 5 は「AGPLv3」   : https://opensource.org/license/agpl-v3/
# 　　　　｜ライセンスの詳細は          : https://github.com/itext/itextsharp/blob/develop/LICENSE.md
#################################################################################
# 設定
# 定義されていない変数があった場合にエラーとする
Set-StrictMode -Version Latest
# アセンブリ読み込み（フォーム用）
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# try-catchの際、例外時にcatchの処理を実行する
$ErrorActionPreference = "Stop"
# itextsharp_5.5.13 を使用
Add-Type -Path '.\source\lib\itextsharp.dll'
# 定数
[System.String]$c_config_file = "setup.ini"
# Function
#################################################################################
# 処理名　｜ExpandString
# 機能　　｜文字列を展開（先頭桁と最終桁にあるダブルクォーテーションを削除）
#--------------------------------------------------------------------------------
# 戻り値　｜String（展開後の文字列）
# 引数　　｜target_str: 対象文字列
#################################################################################
Function ExpandString {
    param ([System.String]$target_str)
    [System.String]$expand_str = $target_str
    
    If ($target_str.Length -ge 2) {
        if (($target_str.Substring(0, 1) -eq "`"") -and
                ($target_str.Substring($target_str.Length - 1, 1) -eq "`"")) {
            # ダブルクォーテーション削除
            $expand_str = $target_str.Substring(1, $target_str.Length - 2)
        }
    }

    return $expand_str
}

#################################################################################
# 処理名　｜ConfirmYesno_winform
# 機能　　｜YesNo入力（Windowsフォーム）
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True: 正常終了, False: 処理中断）
# 引数　　｜prompt_message: 入力応答待ち時のメッセージ内容
#################################################################################
Function ConfirmYesno_winform {
    param (
        [System.String]$prompt_message
    )
    [System.Boolean]$return = $false

    # フォームの作成
    [System.Windows.Forms.Form]$form = New-Object System.Windows.Forms.Form
    $form.Text = "実行前の確認"
    $form.Size = New-Object System.Drawing.Size(460,210)
    $form.StartPosition = "CenterScreen"
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("${root_dir}\source\icon\shell32-296.ico")
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    # ピクチャボックス作成
    [System.Windows.Forms.PictureBox]$pic = New-Object System.Windows.Forms.PictureBox
    $pic.Size = New-Object System.Drawing.Size(32, 32)
    $pic.Image = [System.Drawing.Image]::FromFile("${root_dir}\source\icon\shell32-296.ico")
    $pic.Location = New-Object System.Drawing.Point(30,30)
    $pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    # ラベル作成
    [System.Windows.Forms.Label]$label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(85,30)
    $label.Size = New-Object System.Drawing.Size(350,80)
    $label.Text = $prompt_message
    $font = New-Object System.Drawing.Font("ＭＳ ゴシック",12)
    $label.Font = $font
    # OKボタンの作成
    [System.Windows.Forms.Button]$btnOkay = New-Object System.Windows.Forms.Button
    $btnOkay.Location = New-Object System.Drawing.Point(255,120)
    $btnOkay.Size = New-Object System.Drawing.Size(75,30)
    $btnOkay.Text = "OK"
    $btnOkay.DialogResult = [System.Windows.Forms.DialogResult]::OK
    # Cancelボタンの作成
    [System.Windows.Forms.Button]$btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(345,120)
    $btnCancel.Size = New-Object System.Drawing.Size(75,30)
    $btnCancel.Text = "キャンセル"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    # ボタンの紐づけ
    $form.AcceptButton = $btnOkay
    $form.CancelButton = $btnCancel
    # フォームに紐づけ
    $form.Controls.Add($pic)
    $form.Controls.Add($label)
    $form.Controls.Add($btnOkay)
    $form.Controls.Add($btnCancel)
    # フォーム表示
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $return = $true
    } else {
        $return = $false
    }
    $pic.Image.Dispose()
    $pic.Image = $null
    $form = $null

    return $return
}

#################################################################################
# 処理名　｜RetrieveMessage
# 機能　　｜メッセージ内容を取得
#--------------------------------------------------------------------------------
# 戻り値　｜String（メッセージ内容）
# 引数　　｜target_code; 対象メッセージコード, append_message: 追加メッセージ（任意）
#################################################################################
Function RetrieveMessage {
    param (
        [System.String]$target_code,
        [System.String]$append_message=''
    )
    [System.String]$return = ''

    [System.String[][]]$messages = @(
        ("0", "正常終了"),
        ("-1", "設定ファイルの読み込みでエラー。"),
        ("-111", "必須項目が未入力。"),
        ("-112", "数値項目で数値以外が入力。"),
        ("-211", "参照できないファイルがあり。"),
        ("-212", "参照できないフォルダがあり。"),
        ("-311", "取り込んだデータが0件。"),
        ("-401", "テキストデータの書き出し失敗。"),
        ("-411", "PDFファイル内の処理結果が異常終了。"),
        ("-901", "処理をキャンセル。"),
        ("-999", "例外が発生。")
    )

    for ([System.Int32]$i = 0; $i -lt $messages.Length; $i++) {
        if ($messages[$i][0] -eq $target_code) {
            $sbtemp=New-Object System.Text.StringBuilder
            @("$($messages[$i][1])`r`n",`
              "${append_message}`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $return = $sbtemp.ToString()
            break
        }
    }
    
    return $return
}

#################################################################################
# 処理名　｜SearchPdffile
# 機能　　｜PDFファイルの検索処理
#--------------------------------------------------------------------------------
# 戻り値　｜Int
# 　　　　｜   0: 正常終了
# 　　　　｜-401: エラー テキストデータの書き出し失敗
# 　　　　｜-411: エラー PDFファイル内の処理結果が異常終了
# 引数　　｜target_path; 対象XMLファイル
#################################################################################
Function SearchPdffile {
    param (
        [System.String]$target_path,
        [System.String]$target_text
    )
    [System.Int32]$result = 0

    # PDFファイルのテキストデータ読み込み
    try {
        [System.String]$tmp_textfile = ".\source\tmp\pdf_textdata.txt"
        $reader = New-Object iTextSharp.text.pdf.PdfReader($target_path)
        [System.Int32]$totalpages = $reader.NumberOfPages
        New-Item $tmp_textfile -ItemType file -Force 2>&1>$null

        ## テキストデータを一時ファイルに書き出し
        for ([System.Int32]$i = 1; $i -le $totalpages; $i++) {
            $line = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $i)
            Write-Output $line | Add-Content $tmp_textfile -Encoding UTF8
        }
        $reader.Close()
    } catch {
        $result = -401
    }

    # 指定文字列の検索
    if ($result -eq 0) {
        [System.String]$textdata = (Get-Content -Raw $tmp_textfile)
        $pdfresult = [Regex]::Matches($textdata, $target_text) | ForEach-Object {$_.Value}

        if ($null -eq $pdfresult) {
            $result = -411
        }
    }

    return $result
}

#################################################################################
# 処理名　｜メイン処理
# 機能　　｜同上
#--------------------------------------------------------------------------------
# 　　　　｜-
#################################################################################
[System.Int32]$result = 0
[System.String]$prompt_message = ''
[System.String]$result_message = ''
[System.String]$append_message = ''
[System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

# 初期設定
## ディレクトリの取得
[System.String]$current_dir=Split-Path ( & { $myInvocation.ScriptName } ) -parent
Set-Location $current_dir"\..\.."
[System.String]$root_dir = (Convert-Path .)
## 設定ファイル読み込み
$sbtemp=New-Object System.Text.StringBuilder
@("$current_dir",`
"\",`
"$c_config_file")|
ForEach-Object{[void]$sbtemp.Append($_)}
[System.String]$config_fullpath = $sbtemp.ToString()
try {
    [System.Collections.Hashtable]$param = Get-Content $config_fullpath -Raw -Encoding UTF8 | ConvertFrom-StringData
    # 対象ファイル
    [System.String]$Targetfile=ExpandString($param.Targetfile)
    # 対象検索文字
    [System.String]$Targettext=ExpandString($param.Targettext)

    $sbtemp=New-Object System.Text.StringBuilder
    @("通知　　　: 設定ファイル読み込み`r`n",`
    "　　　　　　設定ファイルの読み込みが正常終了しました。`r`n",`
    "　　　　　　対象: [${config_fullpath}]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $prompt_message = $sbtemp.ToString()
    Write-Host $prompt_message
}
catch {
    $result = -1
    $append_message = "　　　　　　エラー内容: [${config_fullpath}$($_.Exception.Message)]`r`n"
    $result_message = RetrieveMessage $result $append_message
}
## 対象ファイルの存在チェック
if ($result -eq 0) {
    [System.String]$target_path = "${root_dir}\${Targetfile}"
    if (-Not(Test-Path $target_path)) {
        $result = -211
        $append_message = "　　　　　　対象: [${target_path}]`r`n"
        $result_message = RetrieveMessage $result $append_message
    }
}

# 実行前のポップアップ
if ($result -eq 0) {
    $sbtemp=New-Object System.Text.StringBuilder
    # 実行有無の確認
    @("PDFファイル内の検索を実行します。`r`n",`
      "処理を続行しますか？`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $prompt_message = $sbtemp.ToString()
    If (ConfirmYesno_winform $prompt_message) {
        $result = SearchPdffile $target_path $Targettext
    } else {
        $result = -901
    }
    if ($result -ne 0) {
        $result_message = RetrieveMessage $result
    }
}

# 処理結果の表示
$sbtemp=New-Object System.Text.StringBuilder
if ($result -eq 0) {
    @("処理結果　: 正常終了`r`n",`
      "　　　　　　メッセージコード: [${result}]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()
    Write-Host $result_message
}
else {
    @("処理結果　: 異常終了`r`n",`
      "　　　　　　メッセージコード: [${result}]`r`n",`
      "　　　　　　",`
      $result_message)|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()
    Write-Host $result_message -ForegroundColor DarkRed
}

# 終了
exit $result
