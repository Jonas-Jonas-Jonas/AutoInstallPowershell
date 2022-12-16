##############################################################################################################################################
#   Fully automated software installation for Windows 10 Pro x64 and x86
#   Version 1.1 - Created: 30.08.2022 - Updated: 22.09.2022 with bugfixes.
#   The script does the following tasks:
#   ProTouch Installation, SQL Server 2012 SP2 Installation, Teamviewer Installation, Google Chrome Installation
#   Debloating Windows 10 with custom settings, Assigning Teamviewer Client to Amendo Group, Importing Web Browser custom browser settings and writing all console outputs to c:\installlog.txt
#   
#   Made by Younas Sidia
##############################################################################################################################################


# .Net methods for hiding/showing the console in the background
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Show-Console
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()

        # Hide = 0,
        # ShowNormal = 1,
        # ShowMinimized = 2,
        # ShowMaximized = 3,
        # Maximize = 3,
        # ShowNormalNoActivate = 4,
        # Show = 5,
        # Minimize = 6,
        # ShowMinNoActivate = 7,
        # ShowNoActivate = 8,
        # Restore = 9,
        # ShowDefault = 10,
        # ForceMinimized = 11

        [Console.Window]::ShowWindow($consolePtr, 4)
    }

function Hide-Console
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        #0 hide
        [Console.Window]::ShowWindow($consolePtr, 0)

    }

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

#Checking for minimum requirments
function MinHWReq {

    $DiskSpace = $(Get-WmiObject -Class win32_logicaldisk | Where-Object -Property Name -eq C:).FreeSpace / 1GB 

    if ($DiskSpace -gt 29) { 
        Write-Output "Disk Space requirments are satisfied..."
        Write-Output "Starting Download and installation..."
        Download-InstallFiles # Executes function to download and extract files if minimum system requirments are meet.
    }
    else {
        [System.Windows.Forms.MessageBox]::Show('Please free up more disk space. Minimum Requriments are 30 GB')
        Exit
    }
}


function Installation-GUI {

    # This code block is used for the custom user interface.
    # Using WinForms from .Net Framework

    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $Form                            = New-Object system.Windows.Forms.Form
    $Form.FormBorderStyle            = 'Fixed3D'
    $Form.MaximizeBox = $false
    $Form.ClientSize                 = New-Object System.Drawing.Point(692,516)
    $Form.text                       = "ProTouch Software Installation Wizard"
    $Form.TopMost                    = $false
    $Form.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#193750")

    $SWInstall                       = New-Object system.Windows.Forms.Button
    $SWInstall.text                  = "New Customer POS installation"
    $SWInstall.width                 = 464
    $SWInstall.height                = 86
    $SWInstall.visible               = $true
    $SWInstall.location              = New-Object System.Drawing.Point(127,194)
    $SWInstall.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',16)
    $SWInstall.ForeColor             = [System.Drawing.ColorTranslator]::FromHtml("#000000")
    $SWInstall.BackColor             = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
    $SWInstall.Add_Click(
        { 
            Add-Type -AssemblyName PresentationFramework
            Add-Type -AssemblyName System.Drawing
            [System.Windows.Forms.Application]::EnableVisualStyles()

            $jobScript =
            {
                Start-Sleep -Seconds 1
            }


            function Extract() {
                $ProgressBar = New-Object System.Windows.Forms.ProgressBar
                $ProgressBar.Location = New-Object System.Drawing.Point(10, 35)
                $ProgressBar.Size = New-Object System.Drawing.Size(460, 40)
                $ProgressBar.Style = "Marquee"
                $ProgressBar.MarqueeAnimationSpeed = 20

                $main_form.Controls.Add($ProgressBar);

                $Label.Font = $procFont
                $Label.ForeColor = 'red'
                $Label.Text = "Processing ..."
                $ProgressBar.visible

                MinHWReq

                $job = Start-Job -ScriptBlock $jobScript
                do { [System.Windows.Forms.Application]::DoEvents() } until ($job.State -eq "Completed")
                Remove-Job -Job $job -Force


                $Label.Text = "Process Complete"
                $ProgressBar.Hide()
                $StartButton.Hide()
                $EndButton.Visible
            }

            $main_form = New-Object System.Windows.Forms.Form
            $main_form.FormBorderStyle            = 'Fixed3D'
            $main_form.MaximizeBox = $false
            $main_form.Text = 'New Customer POS installation'
            $main_form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#193750")
            $main_form.Width = 520
            $main_form.Height = 180

            $header = New-Object System.Drawing.Font("Verdana", 13, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
            $procFont = New-Object System.Drawing.Font("Verdana", 20, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

            $Label = New-Object System.Windows.Forms.Label
            $Label.Font = $header
            $Label.ForeColor = 'red'
            $Label.Text = "Click Start To begin installation"
            $Label.Location = New-Object System.Drawing.Point(10, 10)
            $Label.Width = 480
            $Label.Height = 50

            $StartButton = New-Object System.Windows.Forms.Button
            $StartButton.Location = New-Object System.Drawing.Size(350, 75)
            $StartButton.Size = New-Object System.Drawing.Size(120, 50)
            $StartButton.Text = "Start"
            $StartButton.height = 40
            $StartButton.BackColor = 'red'
            $StartButton.ForeColor = 'white'
            $StartButton.Add_click( { EXTRACT });

            $EndButton = New-Object System.Windows.Forms.Button
            $EndButton.Location = New-Object System.Drawing.Size(350, 75)
            $EndButton.Size = New-Object System.Drawing.Size(120, 50)
            $EndButton.Text = "OK"
            $EndButton.height = 40
            $EndButton.BackColor = 'green'
            $EndButton.ForeColor = 'white'
            $EndButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

            $main_form.Controls.AddRange(($Label, $StartButton, $EndButton))

            $main_form.StartPosition = "manual"
            $main_form.Location = New-Object System.Drawing.Size(500, 300)
            $result = $main_form.ShowDialog()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $main_form.Close()
            }

        }

    )

    
    $Button2                         = New-Object system.Windows.Forms.Button
    $Button2.text                    = "Transfer Ownership and reinstall ProTouch"
    $Button2.width                   = 464
    $Button2.height                  = 86
    $Button2.location                = New-Object System.Drawing.Point(124,345)
    $Button2.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',16)
    $Button2.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
    $Button2.Add_Click(
        { 
            Add-Type -AssemblyName PresentationFramework
            Add-Type -AssemblyName System.Drawing
            [System.Windows.Forms.Application]::EnableVisualStyles()

            $jobScript2 =
            {
                Start-Sleep -Seconds 1
            }


            function Extract() {
                $ProgressBar = New-Object System.Windows.Forms.ProgressBar
                $ProgressBar.Location = New-Object System.Drawing.Point(10, 35)
                $ProgressBar.Size = New-Object System.Drawing.Size(460, 40)
                $ProgressBar.Style = "Marquee"
                $ProgressBar.MarqueeAnimationSpeed = 20

                $main_form.Controls.Add($ProgressBar);

                $Label.Font = $procFont
                $Label.ForeColor = 'red'
                $Label.Text = "Processing ..."
                $ProgressBar.visible

                Eierskifte

                $job2 = Start-Job -ScriptBlock $jobScript2
                do { [System.Windows.Forms.Application]::DoEvents() } until ($job2.State -eq "Completed")
                Remove-Job -Job $job2 -Force
                

                $Label.Text = "Process Complete"
                $ProgressBar.Hide()
                $StartButton.Hide()
                $EndButton.Visible
            }

            $main_form = New-Object System.Windows.Forms.Form
            $main_form.Text = 'New Customer POS installation'
            $main_form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#193750")
            $main_form.Width = 520
            $main_form.Height = 180

            $header = New-Object System.Drawing.Font("Verdana", 13, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
            $procFont = New-Object System.Drawing.Font("Verdana", 20, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

            $Label = New-Object System.Windows.Forms.Label
            $Label.Font = $header
            $Label.ForeColor = 'red'
            $Label.Text = "Click Start To begin ownership transfer"
            $Label.Location = New-Object System.Drawing.Point(10, 10)
            $Label.Width = 480
            $Label.Height = 50

            $StartButton = New-Object System.Windows.Forms.Button
            $StartButton.Location = New-Object System.Drawing.Size(350, 75)
            $StartButton.Size = New-Object System.Drawing.Size(120, 50)
            $StartButton.Text = "Start"
            $StartButton.height = 40
            $StartButton.BackColor = 'red'
            $StartButton.ForeColor = 'white'
            $StartButton.Add_click( { EXTRACT });

            $EndButton = New-Object System.Windows.Forms.Button
            $EndButton.Location = New-Object System.Drawing.Size(350, 75)
            $EndButton.Size = New-Object System.Drawing.Size(120, 50)
            $EndButton.Text = "OK"
            $EndButton.height = 40
            $EndButton.BackColor = 'green'
            $EndButton.ForeColor = 'white'
            $EndButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

            $main_form.Controls.AddRange(($Label, $StartButton, $EndButton))

            $main_form.StartPosition = "manual"
            $main_form.Location = New-Object System.Drawing.Size(500, 300)
            $result = $main_form.ShowDialog()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $main_form.Close()
            }

        }
    )


    $TextBox1                        = New-Object system.Windows.Forms.TextBox
    $TextBox1.multiline              = $false
    $TextBox1.text                   = "Please Select a installation option"
    $TextBox1.width                  = 550
    $TextBox1.height                 = 20
    $TextBox1.location               = New-Object System.Drawing.Point(101,58)
    $TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',24,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $TextBox1.ForeColor              = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
    $TextBox1.BackColor              = [System.Drawing.ColorTranslator]::FromHtml("#193750")


    $Form.controls.AddRange(@($SWInstall,$Button2,$TextBox1))
    
    #region Logic 
    #endregion
    [void]$Form.ShowDialog()

}



# Downloads Install files from official site.
function Download-InstallFiles {


    # Sets UAC to 0
    powershell Set-Itemproperty -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\policies\system' -Name 'EnableLUA' -value 0
    
    $Url = 'https://autoupdate.tellix.no/ProTouchSetupsSupport/install.zip'
    $ZipFile = 'C:\Install.zip' + $(Split-Path -Path $Url -Leaf) 
    $Destination = 'C:\' 

    $ProgressPreference = "SilentlyContinue"
    Invoke-WebRequest -Uri $Url -OutFile $ZipFile 
    
    #Unzips install folder.
    $ExtractShell = New-Object -ComObject Shell.Application 
    $Files = $ExtractShell.Namespace($ZipFile).Items() 
    $ExtractShell.NameSpace($Destination).CopyHere($Files) 

    # Installation-GUI # Executes function for the main GUI interface to popup after downlad and unzip of installation files.
    BeginInstall
    
}


function BeginInstall {
    Start-Transcript -Append C:\InstallLog.txt


    # Checks if Chrome is installed and if its not, then install it and set Amendo's settings.

    $chromeInstalled = Test-Path -Path "C:\Program Files\Google\Chrome\Application\chrome.exe"

    if (!$chromeInstalled) {

        Write-Host "Chrome is not installed. Installing now..."
        Start-Process -Wait -FilePath 'C:\Install\ChromeSetup.exe' -ArgumentList '/Silent /Install' -PassThru
        Copy-Item C:\Install\initial_preferences "C:\Program Files\Google\Chrome\Application\initial_preferences"
        Copy-Item C:\Install\bookmarks.html "C:\Program Files\Google\Chrome\Application\bookmarks.html"


        Remove-Item "C:\Program Files\Google\Chrome\Application\master_preferences"

        Start-Process -FilePath "C:\Program Files\Google\Chrome\Application\chrome.exe"
        timeout 3
        TASKKILL /F /IM chrome.exe
        Copy-Item C:\Install\Bookmarks "$env:HOMEPATH\AppData\Local\Google\Chrome\User Data\Default\Bookmarks"

    }
    else {
        Write-Host "chrome is already installed"
    }


    #Teamviewer install locations
    $TWx86Test = Test-Path -Path 'C:\Program Files (x86)\TeamViewer\TeamViewer.exe'
    $TWx64Test = Test-Path -Path 'C:\Program Files\TeamViewer\TeamViewer.exe'

    if ($TWx64Test -Or $TWx86Test) {

        Write-Output "Teamviewer already installed. Skipping Teamviewer installation"
        [System.Windows.Forms.MessageBox]::Show('Teamviewer already installed. Skipping Teamviewer installation...')

    }
    else {   
        Write-Output "Installing Teamviewer and assigning API Token"
        
        msiexec.exe /i "C:\Install\TeamViewer_Host.msi" /qn CUSTOMCONFIGID=xxxx
        timeout 3

        if ($TWx64Test) {
            Start-Process -FilePath "C:\Program Files\TeamViewer\TeamViewer.exe" -ArgumentList 'assign --api-token=xxxxx-xxxxxxxxxxxxxxx --grant-easy-access'  
        }
        if ($TWx86Test) {
            Start-Process -FilePath "C:\Program Files (x86)\TeamViewer\TeamViewer.exe" -ArgumentList 'assign --api-token=xxxxxx-xxxxxxxxxxxxx --grant-easy-access'
        }
    }


    # Installing Printer drivers and ProTouch.
    Start-Process -FilePath 'C:\Install\PrintDrivers.exe' -ArgumentList '/Silent /Install' -PassThru
    Start-Process -Wait -FilePath 'C:\Install\ProTouch_Setup_v1.3.3.1_Live\Application Setup - ProTouch Only\setup.exe' -ArgumentList '/s /v/qn' -PassThru

    # Sets the backgroumd image.
    xcopy C:\Install\amendo.jpg C:\Windows\Web\Wallpaper\Theme1 /v /s /e
    reg add "HKEY_CURRENT_USER\Control Panel\Desktop" /v Wallpaper /t REG_SZ /d C:\Windows\Web\Wallpaper\Theme1\amendo.jpg /f 
    RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters

    #   Custom windows settings
    C:/install/bginfo/bginfo.ps1
    Start-Process -Wait -FilePath 'C:\install\bginfo\Bginfo.exe' -ArgumentList '"C:\install\bginfo\customconfig.bgi" /silent /nolicprompt /timer:0' -PassThru

    #   More custom Windows settings and tweaks. Debloats Windows.
    C:/install/Debloat/DeBloater.ps1

    #SQL and Dependencies install
    Start-Process -Wait -FilePath "C:\Install\SQL\setup.exe" -ArgumentList '/CONFIGURATIONFILE=C:\Install\SQL\ConfigurationFile.INI'
    Start-Process -Wait -FilePath "C:\Install\VC_redist.x64.exe" -ArgumentList '/install /quiet /nostart'
    msiexec.exe /quiet /passive /i "C:\Install\SQL\msodbcsql.msi" IACCEPTMSODBCSQLLICENSETERMS=YES ADDLOCAL=ALL

    #  hack to get sqlcmd to work in the same powershell session.
    $Env:PATH = "C:\Program Files\Microsoft SQL Server\110\DTS\Binn\;$Env:PATH"
    $Env:PATH = "C:\Program Files\Microsoft SQL Server\110\Tools\Binn\;$Env:PATH"
    
    # Verifies ProTouch, SQL and DB Scheme is installed and importet correctly.
    $PTinstallLocation = Test-Path -Path "C:\Program Files (x86)\Tellix\ProTouch\PTClient.exe"
    $GetDBVersion = sqlcmd -E -S .\sqlexpress -Q "USE ProTouch SET NOCOUNT ON;"
    
    if ($GetDBVersion.Contains("Changed database context to 'ProTouch'")) { 
        Write-Output "ProTouhch DB Already Exist. Skipping"
    }
    else {
        sqlcmd -E -S .\sqlexpress -i "C:\Install\ProTouch_Setup_v1.3.3.1_Live\Database Script\ProTouchDB-LiveServer.sql"

        Write-Output "SQL ProTouch scheme Imported sucessfully!"
        [System.Windows.Forms.MessageBox]::Show('Software POS Installation Completed!')
    }

    If(!$PTinstallLocation) {

        Write-Warning "ProTouch Installation was not installed successfully. Trying again to install..."
        Start-Process -Wait -FilePath 'C:\Install\ProTouch_Setup_v1.3.3.1_Live\Application Setup - ProTouch Only\setup.exe' -ArgumentList '/s /v/qn' -PassThru
        
    } Else {
   
        Write-Host "ProTouch sucessfully installed."
   
    }

    Stop-Transcript
}


function Eierskifte {

    sqlcmd -E -S .\sqlexpress -Q "BACKUP DATABASE [ProTouch] TO DISK='C:\install\ProTouchBackup.bak'"
    

    if (Test-Path -Path "C:\install\ProTouchBackup.bak") {

        sqlcmd -E -S .\sqlexpress -m 1 -Q "USE ProTouch SET NOCOUNT ON; SELECT PosName, PosNumber, Comments, isActive FROM PosDetails" -o "C:\Install\Eierskiftefra.txt"
        
        Remove-Item -Path HKCU:\SOFTWARE\Protouch -Force -Verbose

        sqlcmd -E -S .\sqlexpress -Q "DROP DATABASE [ProTouch]"

        sqlcmd -E -S .\sqlexpress -i "C:\Install\ProTouch_Setup_v1.3.3.1_Live\Database Script\ProTouchDB-LiveServer.sql"

        $ProTouch = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq "ProTouch"}
        $ProTouch.Uninstall()

        Start-Process -Wait -FilePath 'C:\Install\ProTouch_Setup_v1.3.3.1_Live\Application Setup - ProTouch Only\setup.exe' -ArgumentList '/s /v/qn' -PassThru

        [System.Windows.Forms.MessageBox]::Show('Ownership software changes completed! Please start ProTouch and contact Amendo Support for License key.')
    }
    else {

        Write-Warning "An error occured. Please contact support!"
        [System.Windows.Forms.MessageBox]::Show('An error occured. Please contact support!')
    }

}

Hide-Console
Installation-GUI
