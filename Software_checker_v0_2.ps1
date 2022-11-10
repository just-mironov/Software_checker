$SoftwareList = "Java", "Adobe Acrobat Reader", "Google Chrome", "Mozilla Firefox"

$logfile = "c:\scripts\SoftVersion\logfile.log"
$user = ((whoami) -split "\\")[1]
Add-Content $logfile ("`n")
Add-Content $logfile ((Get-Date).ToString() + ": Start by " + $user)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

$MainForm                        = New-Object system.Windows.Forms.Form
$MainForm.ClientSize             = New-Object System.Drawing.Size('380, 380')
$MainForm.text                   = "Software checker"
$MainForm.StartPosition          = "CenterScreen"
$MainForm.ForeColor              = [System.Drawing.Color]::FromArgb(255,0,0,0)

#не используется
$LoadFromFileButton              = New-Object system.Windows.Forms.Button
$LoadFromFileButton.Visible      = $false
$LoadFromFileButton.text         = "Загрузить из файла"
$LoadFromFileButton.width        = 150
$LoadFromFileButton.height       = 30
$LoadFromFileButton.location     = New-Object System.Drawing.Point(180,60)
$LoadFromFileButton.Font         = 'Microsoft Sans Serif,8'

$ComputerList                    = New-Object System.Windows.Forms.TextBox
$ComputerList.multiline          = $true
$ComputerList.width              = 170
$ComputerList.height             = 290
$ComputerList.location           = New-Object System.Drawing.Point(10,10)
$ComputerList.Font               = 'Microsoft Sans Serif,11'
$ComputerList.ScrollBars         = [System.Windows.Forms.ScrollBars]::Vertical

$ComputersCount                  = New-Object system.Windows.Forms.Label
$ComputersCount.AutoSize         = $true
$ComputersCount.width            = 30
$ComputersCount.height           = 20
$ComputersCount.location         = New-Object System.Drawing.Point(10,310)
$ComputersCount.Font             = 'Microsoft Sans Serif,10'
$ComputersCount.text             = "Кол-во компьютеров: "

$CheckButton                     = New-Object system.Windows.Forms.Button
$CheckButton.text                = "Проверить"
$CheckButton.width               = 180
$CheckButton.height              = 30
$CheckButton.location            = New-Object System.Drawing.Point(190,10)
$CheckButton.Font                = 'Microsoft Sans Serif,10'

$ProgressBar                     = New-Object system.Windows.Forms.ProgressBar
$ProgressBar.width               = 360
$ProgressBar.height              = 25
$ProgressBar.Maximum             = 100
$ProgressBar.Minimum             = 0
$ProgressBar.Visible             = $false
$ProgressBar.location            = New-Object System.Drawing.Point(10,340)

$CurrentComputer                 = New-Object system.Windows.Forms.Label
$CurrentComputer.AutoSize        = $true
$CurrentComputer.Visible         = $false
$CurrentComputer.width           = 30
$CurrentComputer.height          = 20
$CurrentComputer.location        = New-Object System.Drawing.Point(190,310)
$CurrentComputer.Font            = 'Microsoft Sans Serif,10'

$OutFileCheckBox                 = New-Object System.Windows.Forms.CheckBox
$OutFileCheckBox.AutoSize        = $true
$OutFileCheckBox.Visible         = $true
$OutFileCheckBox.location        = New-Object System.Drawing.Point(190,50)
$OutFileCheckBox.Font            = 'Microsoft Sans Serif,10'
$OutFileCheckBox.Text            = 'Сохранить в файл'

$OutDeskCheckBox                 = New-Object System.Windows.Forms.CheckBox
$OutDeskCheckBox.AutoSize        = $true
$OutDeskCheckBox.Visible         = $true
$OutDeskCheckBox.Checked         = $true
$OutDeskCheckBox.location        = New-Object System.Drawing.Point(190,80)
$OutDeskCheckBox.Font            = 'Microsoft Sans Serif,10'
$OutDeskCheckBox.Text            = 'Показать в таблице'

# Создание TreeView необычный, так как есть глюк при двойном тыке
# Оригинал https://www.reddit.com/r/PowerShell/comments/a8087s/formsgui_treeview_checking_or_unchecking_parent/ec7flz4/
################################################################################################################################################################
$ref = [reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')

$MyTreeView = @'
using System;
using System.Windows.Forms;

public class MyTreeView : TreeView {
    protected override void WndProc(ref Message m) {
        // Suppress WM_LBUTTONDBLCLK
        if (m.Msg == 0x203) { m.Result = IntPtr.Zero; }
        else base.WndProc(ref m);
    }
}
'@

Add-Type -TypeDefinition $MyTreeView -ReferencedAssemblies $ref -WarningAction Ignore
################################################################################################################################################################

$TreeView                        = New-Object 'MyTreeView'
$TreeView.CheckBoxes             = $True
$TreeView.Size                   = New-Object System.Drawing.Size(180,190)
$TreeView.location               = New-Object System.Drawing.Point(190,110)

$MainNode                        = New-Object System.Windows.Forms.TreeNode
$MainNode.Text                   = "Software"
$MainNode.Checked                = $True
$MainNode.ExpandAll()


function NodeCreation ([String[]]$SoftwareList) {
foreach ($Soft in $SoftwareList) {
    $ChildNode                       = New-Object System.Windows.Forms.TreeNode
    $ChildNode.Text                  = $Soft
    $ChildNode.Checked               = $True
    $MainNode.Nodes.Add($ChildNode) | Out-Null
    }
}

$TreeView.Nodes.Add($MainNode) | Out-Null

NodeCreation $SoftwareList

# Собираем выбранный софт
function Get-CheckedNode($nodes) {
    foreach ($n in $Nodes) {
        if ($n.nodes.count -gt 0) {
            Get-CheckedNode $n.nodes
        }
        if ($n.checked -and $n.parent) {
            $n.Text
        }
    }
}

# выставляем детям тру или фолс в зависимости от родителя
$treeView.Add_AfterCheck({
    $node = $_.node
    if($node.nodes) {
        foreach($subnode in $node.nodes) {
            $subnode.checked = $node.checked
            }
    }
})

$MainForm.controls.AddRange(@($LoadFromFileButton, $ComputerList, $ComputersCount, `
$CheckButton,$ProgressBar,$TreeView,$CurrentComputer,$OutFileCheckBox,$OutDeskCheckBox))

$OutFileCheckBox.Add_Click({
    if (!($OutFileCheckBox.checked) -and !($OutDeskCheckBox.Checked))  {
        $CheckButton.Enabled = $false
    } else {
        $CheckButton.Enabled = $true
    }
})

$OutDeskCheckBox.Add_Click({
    if (!($OutFileCheckBox.checked) -and !($OutDeskCheckBox.Checked))  {
        $CheckButton.Enabled = $false
    } else {
        $CheckButton.Enabled = $true
    }
})

#Основная логика здесь

    # код просмотра реестра
    function Get-SoftWare([string[]]$ComputerNames) {
        $obj = @()

        foreach ($ComputerName in $ComputerNames) {
            $CurrentComputer.text = "Проверяю... " + $ComputerName
            [int]$Percent = ++$i/$ComputerNames.Count*100
            $ProgressBar.Value = $Percent
            [void]$MainForm.Update()

            if (Test-Connection $ComputerName -Count 1 -Quiet) {
                try {
                    $ErrorActionPreference = 'Stop'
                    $WindowsVersion = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName | select Caption, Version, OSArchitecture
                    } catch { $ExcMsg = $_.Exception.Message }
                } else { $ExcMsg = "Ping failed" }

                function Get-Registry($ComputerName) {
                    $obj = @()
                    if ($WindowsVersion.OSArchitecture -eq "32-bit") {
                        $Registry_Location = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
                        } else {
                        $Registry_Location = 'Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
                    }

                    $regHive = [Microsoft.Win32.RegistryHive]::LocalMachine
                    $Remote_Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regHive, $ComputerName)

                    foreach ($Sub_Key in $Remote_Registry.OpenSubKey($Registry_Location).GetSubKeyNames()) {
                        $properties = New-Object PSObject -Property ([Ordered]@{
                            CopmuterName = $ComputerName
                            Caption = $WindowsVersion.Caption
                            Version = $WindowsVersion.Version
                            DisplayName =  $Remote_Registry.OpenSubKey("$Registry_Location\$Sub_Key\").GetValue('DisplayName')
                            DisplayVersion = $Remote_Registry.OpenSubKey("$Registry_Location\$Sub_Key\").GetValue('DisplayVersion')
                            Error = $ExcMsg
                        })
                        $obj += $properties
                    }
                    return $obj
                }

                if ($ExcMsg) {
                    $obj += New-Object -TypeName psobject -Property ([Ordered]@{
                        CopmuterName = $ComputerName
                        Caption = ""
                        Version = ""
                        DisplayName = ""
                        DisplayVersion = ""
                        Error = $ExcMsg
                    })
                    $ExcMsg = ""
                    }
                else {
                    $obj += (Get-Registry $ComputerName)
                    }
            }
            return $obj #| select -Unique
    }

    # Загрузка файла из .txt, кнопка не активна
    $LoadFromFileButton.Add_Click({
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "Txt files (*.txt)| *.txt"
        $OpenFileDialog.ShowDialog() | Out-Null
        if($OpenFileDialog.FileName) {
            $ComputerList.Text = Get-Content $OpenFileDialog.FileName
            }
     })

    #Собираем и считаем введеные компьютеры
    $ComputerList.Add_TextChanged({
        $result = $ComputerList.Text | Select-String "W[S,M,N]-.{4}" -AllMatches
        $ComputersCount.text = "Кол-во компьютеров: " + [string]$result.Matches.Count
        $Global:ComputerNameList = [String[]]$result.Matches | select -Unique
    })

    #Запуск проверки
    $CheckButton.Add_Click({
        [int]$PCCount = $ComputersCount.Text -replace "Кол-во компьютеров: ", ""
        if ($PCCount -and $PCCount -gt 0) {

                $CurrentComputer.Visible = $true
                $ProgressBar.Visible     = $true

                $CheckedSoftwareList = Get-CheckedNode($TreeView)

                Add-Content $logfile ((Get-Date).ToString() + ": click check with Arguments[" + ($CheckedSoftwareList -join ", ") + "][" + ($ComputerNameList  -join ", ") + "]" )

                $TimeExecution = Measure-Command { $AllSoft = Get-SoftWare $ComputerNameList } | select @{n = 'String';e = {$_.Minutes,"Minutes",$_.Seconds,"Seconds",$_.Milliseconds,"Milliseconds" -join " "}}
                Add-Content $logfile ((Get-Date).ToString() + ": Execution time " + $TimeExecution.String )

                $Result = $AllSoft | Where-Object {($_.DisplayName -Match ($CheckedSoftwareList -join "|")) -or ($_.Error) }

                if ($OutDeskCheckBox.checked) { $Result | Out-GridView -Title "Найденое ПО" }
                if ($OutFileCheckBox.checked) {
                        $dlg = New-Object System.Windows.Forms.SaveFileDialog
                        $dlg.Filter = "CSV Files (*.csv)|*.csv"
                        $dlg.SupportMultiDottedExtensions = $true;
                        $dlg.InitialDirectory = 'C:\Users\' + $user + '\Desktop'
                        if($dlg.ShowDialog() -eq 'Ok') {
                            $Result | Export-Csv -Path $dlg.FileName -Delimiter ";" -Encoding UTF8 -NoTypeInformation
                        }
                    }

                $ComputerNameList       = ""
                $CurrentComputer.text   = ""
                $ProgressBar.Visible    = $false
            } else {
                [System.Windows.MessageBox]::Show('Не указаны пк для проверки','Ошибка','Ok','Error')
            }
        })

    #Запись в лог закрытия
    $MainForm.Add_Closing({
            Add-Content $logfile ((Get-Date).ToString() + ": closing program `n")
    })
[void]$MainForm.ShowDialog()
[void]$MainForm.Focus()
