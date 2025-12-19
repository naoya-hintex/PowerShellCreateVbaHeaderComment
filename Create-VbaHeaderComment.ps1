# ＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
# ファイル名    Create-VbaHeaderComment
# 概要         VBAのプロシージャ用のコメントを作成します。
#              作成したコメントはクリップボードに設定します。
# 作成者        naoya-hintex
# 引数          なし
# 戻り値        なし
# ＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

# ユーザにコンソールから「プロシージャ名」を入力してもらう
$funcName = Read-Host "プロシージャ名を入力してください"
if ([string]::IsNullOrWhiteSpace($funcName)) {
    $funcName = "[プロシージャ名を記載]"
}

# ユーザにコンソールから「概要」を入力してもらう
$description = Read-Host "処理概要を入力してください"
if ([string]::IsNullOrWhiteSpace($description)) {
    $description = "[処理概要を記載]"
}

# コメントを作成
$separator = "' ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝"
$content = @(
    $separator
    "' プロシージャ名   $funcName                                                             "
    "' 処理概要         $description                                                          "
    "' 作成者           [名前を記載]                                                           "
    "' 引数             [引数を記載]                                                           "
    "' 戻り値           [戻り値を記載]                                                         "
    $separator
)

($content -join "`r`n") | Set-clipboard