#############################################################
# WinSCPを使用してN日を経過したファイルを転送し
# 結果をメール通知する。
############################################################

# 転送元Windowsフォルダを指定
$SRC_DIR="D:\xxx"

# SCP転送先情報
$SCP_IPADDR="xx.xx.x.x"
$SCP_USER="xxx"
$SCP_PASS="xxx"
$SCP_DIR="/work/xxx"

# ファイルの作成日時がDAY_OLD以内の
# ファイルを送信対象とする
$DAY_OLD=1  

# WinSCPのインストールパスを指定(.exeを含む)
$SCP="C:\Program Files\WinSCP\WinSCP.exe"

# メール設定(例:下記はgmail)
$From="xxxx@gmail.com"
$To="xxxx@gmail.com"
$SMTP_Server="smtp.gmail.com"
$SMTP_User="xxxx@gmail.com"
$SMTP_Pass="password"
$SMTPClient=New-Object Net.Mail.SmtpClient($Smtp_Server,587)
$SMTPClient.EnableSsl = $true 
$SMTPClient.Credentials=New-Object System.Net.NetworkCredential($SMTP_User,$SMTP_Pass)


# scp スクリプトファイル、およびログ
$BATCH_SCRIPT="$SRC_DIR\batch.script"
$BATCH_LOG="$SRC_DIR\batch.log"
$BATCH_ERR="$SRC_DIR\batch.err"

##
## ここから処理開始
##
$CUR_PWD=pwd
cd $SRC_DIR

# ログ削除
if(Test-Path $BATCH_LOG){
	rm -force $BATCH_LOG
}
# スクリプト開始処理作成
echo "option batch on" > $BATCH_SCRIPT
#echo "option batch abort" >> $BATCH_SCRIPT
echo "option confirm on" >> $BATCH_SCRIPT
echo "open ${SCP_USER}:${SCP_PASS}@${SCP_IPADDR}" >> $BATCH_SCRIPT
echo "cd $SCP_DIR" >> $BATCH_SCRIPT
echo "option transfer binary" >> $BATCH_SCRIPT

# 対象ファイル数分繰り返す
$files=Get-ChildItem -R $SRC_DIR -exclude batch.*  | where { $_.fullname -notlike "*done*" }
foreach($file in $files)
{
	$d=((Get-Date) - $file.LastWriteTime).Days
	if($d -ge $DAY_OLD -and $file.PsISContainer -ne $True){
		$rfile=$file | Resolve-Path -Relative
		echo "target [$rfile]"

		$parents=Split-Path $rfile -parent
		$pdir=""
		foreach($p in $parents)
		{
			if ( $p -eq "."){ continue}
		        $tmp=$p.replace(".`\","")
			echo "mkdir $tmp" >> $BATCH_SCRIPT
			echo "cd $tmp" >> $BATCH_SCRIPT
		}

		echo "put `"$rfile`"" >> $BATCH_SCRIPT
		echo "cd $SCP_DIR" >> $BATCH_SCRIPT
	}
}

# スクリプト終了処理作成
echo "close" >> $BATCH_SCRIPT
echo "exit" >> $BATCH_SCRIPT

$arg="`/console `/script=$BATCH_SCRIPT `/log=$BATCH_LOG"
Start-Process -FilePath $SCP  -ArgumentList $arg -PassThru -Wait

$ret=Get-Content $BATCH_LOG -Encoding UTF8 | select-string  "^\*"
#echo $ret.line
#echo $ret.length

$Subject="Result SCP"
if($ret.length -ne 0){
	$msg=""
	foreach($m in $ret)
	{
		$msg=$msg + $m.line + "`r`n"
	}
	$SMTPClient.Send($From, $To, $Subject, "$msg")
}else{
	$SMTPClient.Send($From, $To, $Subject, "Success")
}

# doneフォルダへ移動
if(!(test-path $SRC_DIR\done)){
	mkdir $SRC_DIR\done
}

foreach($file in $files)
{
	$d=((Get-Date) - $file.LastWriteTime).Days
	if($d -ge $DAY_OLD -and $file.PsISContainer -ne $True){
		$rfile=$file | Resolve-Path -Relative

		$parent=Split-Path $rfile -parent
		if(!(test-path $SRC_DIR\done\$parent)){
			mkdir $SRC_DIR\done\$parent
		}
		
		mv $rfile $SRC_DIR\done\$parent
	}
}

cd $CUR_PWD
exit 0

