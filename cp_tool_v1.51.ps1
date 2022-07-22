$ErrorActionPreference = "stop"
Set-PSDebug -Strict
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

###########################################################
#ここから下、GUI部分
###########################################################
[xml]$xaml = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        
        Title="copy_tool" Height="300" Width="354">
    <Grid Margin="0,0,14,8">

        <Label    x:Name="label_origin" Content="コピー元フォルダの選択" HorizontalAlignment="Left" Height="25" Margin="25,10,0,0" VerticalAlignment="Top" Width="160" Background="{x:Null}" RenderTransformOrigin="0.5,-0.249" FontSize="10"/>
        <TextBox  x:Name="textBox_origin" HorizontalAlignment="Left" Height="20" Margin="30,35,0,0"  Text="" VerticalAlignment="Top" Width="210" Background="#FFECECEC" FontSize="10"/>
        <Button   x:Name="button_ref_origin" Content="参照" HorizontalAlignment="Left" Height="20" Margin="255,35,0,0" VerticalAlignment="Top" Width="55" RenderTransformOrigin="0.5,-0.498"/>


        <Label    x:Name="label_selectDay" Content="日付選択" HorizontalAlignment="Left" Height="25" Margin="25,65,0,0" VerticalAlignment="Top" Width="160" Background="{x:Null}" RenderTransformOrigin="0.5,-0.249" FontSize="10"/>

        <ComboBox x:Name="comboBox_years" HorizontalAlignment="Left" Height="20" Margin="30,90,0,0" VerticalAlignment="Top" Width="80"/>
        <Label    x:Name="label" Content="年" HorizontalAlignment="Left" Height="25" Margin="120,90,0,0" VerticalAlignment="Top" Width="40" Background="{x:Null}" RenderTransformOrigin="0.5,-0.249" FontSize="9" IsEnabled="False"/>

        <ComboBox x:Name="comboBox_month" HorizontalAlignment="Left" Height="20" Margin="170,90,0,0" VerticalAlignment="Top" Width="80" RenderTransformOrigin="0.5,1.502"/>
        <Label    x:Name="label_Copy" Content="月" HorizontalAlignment="Left" Height="25" Margin="260,90,0,0" VerticalAlignment="Top" Width="40" RenderTransformOrigin="0.5,-0.249" FontSize="9" Background="{x:Null}"/>


        <Label    x:Name="label_toCopy" Content="コピー先フォルダの選択" HorizontalAlignment="Left" Height="25" Margin="25,120,0,0" VerticalAlignment="Top" Width="160" Background="{x:Null}" RenderTransformOrigin="0.5,-0.249" FontSize="10"/>
        <TextBox  x:Name="textBox_toCopy" HorizontalAlignment="Left" Height="20" Margin="30,145,0,0"  Text="" VerticalAlignment="Top" Width="210" Background="#FFECECEC" FontSize="10"/>
        <Button   x:Name="button_ref_toCopy" Content="参照" HorizontalAlignment="Left" Height="20" Margin="255,145,0,0" VerticalAlignment="Top" Width="55" RenderTransformOrigin="0.5,-0.498"/>


        <Button   x:Name="button_OK" Content="OK" HorizontalAlignment="Left" Height="20" Margin="120,200,0,0" VerticalAlignment="Top" Width="80" RenderTransformOrigin="0.5,-0.498"/>

    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$frm = [System.Windows.Markup.XamlReader]::Load($reader)

###########################################################
#ここから下、設定ファイルの作成
###########################################################
#json設定ファイルが無ければ設定ファイルを作成する
$configpath = "$PSScriptRoot\config.json"

if(!(Test-Path $configpath -PathType leaf)){
    
    $json = @{originPath="フォルダを選択してください"; toCopyPath="フォルダを選択してください"}
    ConvertTo-Json $json | Out-File $configpath -Encoding utf8
}

#jsonファイルを読み込んで変数へ格納
$configJson = ConvertFrom-Json -InputObject (Get-Content $configpath -Raw)

###########################################################
#ここから下、年数の設定
###########################################################
#日付の取得
$todays = Get-Date -format "yyyyMM"
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
$comboBox_years = $frm.FindName("comboBox_years")
foreach ($year in $years) {
    [void]$comboBox_years.Items.Add($year)
}
#初期値の設定(SelectedIndex)
for ($i = 0; $i -lt $years.Count; $i++) {
    if($years[$i] -eq $thisYear) {
        $comboBox_years.SelectedIndex = $i
    }
}

###########################################################
#ここから下、月の設定
###########################################################
#月の取得
$months = @("01","02","03","04","05","06","07","08","09","10","11","12");
$thisMonth = get-date -Format "MM"
#コンボボックスに年数値を設定
$comboBox_month = $frm.FindName("comboBox_month")
foreach ($month in $months) {
    [void]$comboBox_month.Items.Add($month)
}
#初期値の設定(SelectedIndex)
for ($i = 0; $i -lt $months.Count; $i++) {
    if($months[$i] -eq $thisMonth - 1) {
        $comboBox_month.SelectedIndex = $i
    }
}

###########################################################
#ここから下、フォルダ選択機能の設定
###########################################################
#参照ボタンとフォルダ選択メソッドの紐づけ
$button_ref = $frm.FindName("button_ref_origin")
$button_ref.Add_Click({selectFolder($textBox_origin)})
$textBox_origin = $frm.FindName("textBox_origin")

$button_ref_toCopy = $frm.FindName("button_ref_toCopy")
$button_ref_toCopy.Add_Click({selectFolder($textBox_toCopy)})
$textBox_toCopy = $frm.FindName("textBox_toCopy")

#設定ファイルから前回指定されたフォルダパスがあればその値を設定
if(Test-Path $configpath -PathType leaf){
    $textBox_origin.text = $configJson.originPath
    $textBox_toCopy.text = $configJson.toCopyPath
}

function selectFolder($btn_ref_name) {

    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
        RootFolder = "Desktop"
        Description = "フォルダを選択してください"
    }
    Write-Host $btn_ref_name.object
    # フォルダ選択の有無を判定
    if($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $btn_ref_name.Text = $FolderBrowser.SelectedPath
    }
    Write-Host $btn_ref_name
    Write-Host $textBox_origin.Text

#    if($btn_ref_name -eq ) { 
#
#    } elseif($btn_ref_name -eq ) {
#
#    }

}

$configJson.toCopyPath = $FolderBrowser.SelectedPath
$configJson.toCopyPath = $FolderBrowser.SelectedPath
ConvertTo-Json $configJson | Out-File $configpath -Encoding utf8

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

    $dayStr = $comboBox_years.Text + $comboBox_month.text
    [string]$msg = "";
    #ディレクトリの取得
    $targetPath = $frm.FindName("textBox_origin")
    $targetFolder = Get-ChildItem $targetPath

    #対象ファイルに対してアカウントの判定処理呼び出し
    $targetFolder = 

    #コピー先のフォルダを確認し、あれば開く
    if(Test-Path $textBox_toCopy.Text ) {
        $copyFolder = $textBox_toCopy.Text
        Invoke-Item -Path $copyFolder
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
