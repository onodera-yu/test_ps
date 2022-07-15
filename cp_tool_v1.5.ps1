$ErrorActionPreference = "stop"
Set-PSDebug -Strict
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

[xml]$xaml = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        
        Title="copy_tool" Height="199" Width="354">
    <Grid Margin="0,0,14,8">
        <ComboBox x:Name="comboBoxYears" HorizontalAlignment="Left" Height="20" Margin="30,34,0,0" VerticalAlignment="Top" Width="80"/>
        <Label x:Name="label" Content="年" HorizontalAlignment="Left" Height="25" Margin="125,34,0,0" VerticalAlignment="Top" Width="40" Background="{x:Null}" RenderTransformOrigin="0.5,-0.249" FontSize="9" IsEnabled="False"/>
        <TextBox x:Name="textBox" HorizontalAlignment="Left" Height="20" Margin="30,84,0,0"  Text="フォルダを選択してください" VerticalAlignment="Top" Width="210" Background="#FFECECEC" FontSize="10"/>
        <Button x:Name="button_OK" Content="OK" HorizontalAlignment="Left" Height="20" Margin="120,124,0,0" VerticalAlignment="Top" Width="80" RenderTransformOrigin="0.5,-0.498"/>
        <ComboBox x:Name="comboBoxMonth" HorizontalAlignment="Left" Height="20" Margin="180,34,0,0" VerticalAlignment="Top" Width="80" RenderTransformOrigin="0.5,1.502"/>
        <Label x:Name="label_Copy" Content="月" HorizontalAlignment="Left" Height="25" Margin="270,34,0,0" VerticalAlignment="Top" Width="40" RenderTransformOrigin="0.5,-0.249" FontSize="9" Background="{x:Null}"/>
        <Label x:Name="label_Copy1" Content="コピー先フォルダの選択" HorizontalAlignment="Left" Height="25" Margin="20,59,0,0" VerticalAlignment="Top" Width="160" Background="{x:Null}" RenderTransformOrigin="0.5,-0.249" FontSize="10"/>
        <Label x:Name="label_Copy2" Content="日付選択" HorizontalAlignment="Left" Height="25" Margin="20,4,0,0" VerticalAlignment="Top" Width="160" Background="{x:Null}" RenderTransformOrigin="0.5,-0.249" FontSize="10"/>
        <Button x:Name="button_reference" Content="参照" HorizontalAlignment="Left" Height="20" Margin="255,84,0,0" VerticalAlignment="Top" Width="55" RenderTransformOrigin="0.5,-0.498"/>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$frm = [System.Windows.Markup.XamlReader]::Load($reader)

#日付の取得
$todays = Get-Date -format "yyyyMM"


###########################################################
#ここから下、年数の設定
###########################################################
#年数の取得（今年から前後50年分くらい（適当））
$years = @();
$thisYear = get-date -Format "yyyy"
for($i = 0; $i -lt 50; $i++) {
    $years += ([int]$thisYear + $i)
    $years += ([int]$thisYear - $i)
}
#年数のソートと重複する年数の削除
$years = $years | Sort-Object | Get-Unique
#コンボボックスに年数値を設定
$comboBoxYears = $frm.FindName("comboBoxYears")
foreach ($year in $years) {
    [void]$comboBoxYears.Items.Add($year)
}
#初期値の設定(SelectedIndex)
for ($i = 0; $i -lt $years.Count; $i++) {
    if($years[$i] -eq $thisYear) {
        $comboBoxYears.SelectedIndex = $i
    }
}

###########################################################
#ここから下、月の設定
###########################################################
#月の取得
$months = @("01","02","03","04","05","06","07","08","09","10","11","12");
$thisMonth = get-date -Format "MM"
#コンボボックスに年数値を設定
$comboBoxMonth = $frm.FindName("comboBoxMonth")
foreach ($month in $months) {
    [void]$comboBoxMonth.Items.Add($month)
}
#初期値の設定(SelectedIndex)
for ($i = 0; $i -lt $months.Count; $i++) {
    if($months[$i] -eq $thisMonth - 1) {
        $comboBoxMonth.SelectedIndex = $i
    }
}

###########################################################
#ここから下、フォルダ選択機能の設定
###########################################################
#参照ボタンとフォルダ選択メソッドの紐づけ
$button_reference = $frm.FindName("button_reference")
$button_reference.Add_Click({Get-FolderPathG})
$textBox = $frm.FindName("textBox")
<#
.SYNOPSIS
    フォルダ選択ダイアログ表示

.DESCRIPTION
    フォルダ選択ダイアログを表示し、選択したフォルダパスを返す。

.PARAMETER Description
    ダイアログに表示する説明文（省略可）

.PARAMETER CurrentDefault
    カレントディレクトリをダイアログの初期フォルダとするか否か（省略可）

.OUTPUTS
    選択したフォルダパス。キャンセル時はnull
#>
function Get-FolderPathG {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [string]$Description = "フォルダを選択してください",
        [boolean]$CurrentDefault = $false
    )
    # メインウィンドウ取得
    $process = [Diagnostics.Process]::GetCurrentProcess()
    $window = New-Object Windows.Forms.NativeWindow
    $window.AssignHandle($process.MainWindowHandle)

    $fd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fd.Description = $Description

    if($CurrentDefault -eq $true){
        # カレントディレクトリを初期フォルダとする
        $fd.SelectedPath = (Get-Item $PWD).FullName
    }

    # フォルダ選択ダイアログ表示
    $ret = $fd.ShowDialog($window)
    if($ret -eq [System.Windows.Forms.DialogResult]::OK){
        $textBox.Text = $fd.SelectedPath
    }
    <#else{
        $textBox.Text = $textBox
    }#>
}

###########################################################
#ここから下、OKボタンの設定
###########################################################
#OKボタンの処理（copy機能を実行） 
$button_OK = $frm.FindName("button_OK")
$button_OK.Add_Click({copyRun})
#結果判定ラベル
$label_result = $frm.FindName("label_result")
$count = 0;
function copyRun {
    $dayStr = $comboBoxYears.Text + $comboBoxMonth.text
    [string]$msg = "";
    #ディレクトリの取得
    $targetPath = "\\GSFS.iijglobal.co.jp\OE\NETOPS\Public\RegKisso個人ファイル保管"
    $targetFolder = Get-ChildItem $targetPath -Name -Include G*

    #コピー先のフォルダを確認し、あれば開く
    if(Test-Path $textBox.Text ) {
        $copyFolder = $textBox.Text
        Invoke-Item　-Path $copyFolder
        #Write-Host($copyFolder);
    }else {
        #Write-Host("コピー先のフォルダがありません。");
        return
    }

    #ループ処理
    #対象ディレクトリ配下の特定文字列を含むファイルに対して、リネーム＆コピーを実行
    $targetFolder | ForEach-Object {
        $target =  $targetPath + "\" + $_ | Get-ChildItem -Name | Where-Object { $_ -match "worktime_$dayStr" }
        if($target -match $dayStr) {
            Copy-Item -Path $targetPath\$_\$target -Destination "$copyFolder\$_`_$target"
            $count = $count + 1;
        }
    }
    
}

#ダイアログの表示
$frm.ShowDialog()

