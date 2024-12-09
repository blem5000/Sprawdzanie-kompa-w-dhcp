Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName 'Microsoft.VisualBasic, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'

$global:poprawnyIP = $null
$global:poprawnyMAC = $null
$global:niePoprawnyIP = $null
$global:niePoprawnyMAC = $null
$global:DHCPIP = $null

$FormObject = [System.Windows.Forms.Form]
$LabelObject = [System.Windows.Forms.Label]
$ButtonObject = [System.Windows.Forms.Button]
$TextBoxObject = [System.Windows.Forms.TextBox]

$DHCPListForm = New-Object $FormObject
$DHCPListForm.MinimumSize = '750,250'
#$DHCPListForm.ClientSize = '750,250'
$DHCPListForm.Text = 'Rozwiazywanie hostname w DHCP'
$DHCPListForm.BackColor = "#ffffff"
$DHCPListForm.MaximizeBox = $false
$DHCPListForm.AutoSize = $true
$DHCPListForm.AutoSizeMode = 'GrowAndShrink'


$lblHostname = New-Object $LabelObject
$lblHostname.Text = 'Hostname:'
$lblHostname.AutoSize = $true
$lblHostname.Font = 'Verdana,12,style=Bold'
$lblHostname.Location = New-Object System.Drawing.Point(10, 25)

$txtHostname = New-Object $TextBoxObject
$txtHostname.Autosize = $true
$txtHostname.Font = 'Verdana,12'
$txtHostname.Size = New-Object System.Drawing.Size(210, 20)
$txtHostname.Location = New-Object System.Drawing.Point(170, 30)

$lblIP = New-Object $LabelObject
$lblIP.Text = 'Aktualne IP:'
$lblIP.AutoSize = $true
$lblIP.Font = 'Verdana,12,style=Bold'
$lblIP.Location = New-Object System.Drawing.Point(10, 80)

$lblRozwiazanyIP = New-Object $LabelObject
$lblRozwiazanyIP.Text = ''
$lblRozwiazanyIP.AutoSize = $true
$lblRozwiazanyIP.Font = 'Verdana,12,style=Bold'
$lblRozwiazanyIP.Location = New-Object System.Drawing.Point(220, 80)

$lblMAC = New-Object $LabelObject
$lblMAC.Text = 'Aktualny MAC:'
$lblMAC.AutoSize = $true
$lblMAC.Font = 'Verdana,12,style=Bold'
$lblMAC.Location = New-Object System.Drawing.Point(10, 110)

$lblRozwiazanyMAC = New-Object $LabelObject
$lblRozwiazanyMAC.Text = ''
$lblRozwiazanyMAC.AutoSize = $true
$lblRozwiazanyMAC.Font = 'Verdana,12,style=Bold'
$lblRozwiazanyMAC.Location = New-Object System.Drawing.Point(220, 110)

$lblStareIP = New-Object $LabelObject
$lblStareIP.Text = 'Stare IP:'
$lblStareIP.AutoSize = $true
$lblStareIP.Font = 'Verdana,12,style=Bold'
$lblStareIP.Location = New-Object System.Drawing.Point(10, 140)
$lblStareIP.Visible = $false

$lblStaryIP = New-Object $LabelObject
$lblStaryIP.Text = ''
$lblStaryIP.AutoSize = $true
$lblStaryIP.Font = 'Verdana,12,style=Bold'
$lblStaryIP.Location = New-Object System.Drawing.Point(220, 140)

$lblStaryMAC = New-Object $LabelObject
$lblStaryMAC.Text = 'Stary MAC:'
$lblStaryMAC.AutoSize = $true
$lblStaryMAC.Font = 'Verdana,12,style=Bold'
$lblStaryMAC.Location = New-Object System.Drawing.Point(10, 170)
$lblStaryMAC.Visible = $false

$lblStaryRozwiazanyMAC = New-Object $LabelObject
$lblStaryRozwiazanyMAC.Text = ''
$lblStaryRozwiazanyMAC.AutoSize = $true
$lblStaryRozwiazanyMAC.Font = 'Verdana,12,style=Bold'
$lblStaryRozwiazanyMAC.Location = New-Object System.Drawing.Point(220, 170)

$btnHello = New-Object $ButtonObject
$btnHello.Text = 'Szukaj'
$btnHello.AutoSize = $true
$btnHello.Font = 'Verdana,12'
$btnHello.Location = New-Object System.Drawing.Point(400, 27)

$btnCopyIP = New-Object $ButtonObject
$btnCopyIP.Text = 'Kopiuj IP'
$btnCopyIP.AutoSize = $true
$btnCopyIP.Font = 'Verdana,12'
$btnCopyIP.Location = New-Object System.Drawing.Point(485, 27)
$btnCopyIP.Enabled = $false

$btnCopyMAC = New-Object $ButtonObject
$btnCopyMAC.Text = 'Kopiuj MAC'
$btnCopyMAC.AutoSize = $true
$btnCopyMAC.Font = 'Verdana,12'
$btnCopyMAC.Location = New-Object System.Drawing.Point(585, 27)
$btnCopyMAC.Enabled = $false


$DHCPListForm.Controls.AddRange(@($lblHostname, $btnHello, $txtHostname, $lblIP, $lblRozwiazanyIP, $lblRozwiazanyMAC, $lblMAC, $lblStareIP, $lblStaryIP, $lblStaryMAC, $lblStaryRozwiazanyMAC, $btnCopyIP, $btnCopyMAC))


$DHCPServerListForm = New-Object $FormObject
$DHCPServerListForm.MinimumSize = '250,150'
$DHCPServerListForm.Text = 'Rozwiazywanie hostname w DHCP'
$DHCPServerListForm.BackColor = "#ffffff"
$DHCPServerListForm.MaximizeBox = $false
$DHCPServerListForm.AutoSize = $true
$DHCPServerListForm.AutoSizeMode = 'GrowAndShrink'

$txtServerIP = New-Object $TextBoxObject
$txtServerIP.Autosize = $true
$txtServerIP.Font = 'Verdana,12'
$txtServerIP.Text = '150.150.222.20'
$txtServerIP.Size = New-Object System.Drawing.Size(150, 20)
$txtServerIP.Location = New-Object System.Drawing.Point(440, 32)

$lblServerIP = New-Object $LabelObject
$lblServerIP.Text = 'Aktualny adres serwera DHCP:'
$lblServerIP.AutoSize = $true
$lblServerIP.Font = 'Verdana,12,style=Bold'
$lblServerIP.Location = New-Object System.Drawing.Point(10, 30)

$btnServerIP = New-Object $ButtonObject
$btnServerIP.Text = 'OK'
$btnServerIP.AutoSize = $true
$btnServerIP.Font = 'Verdana,12'
$btnServerIP.Location = New-Object System.Drawing.Point(250, 75)
$btnServerIP.Enabled = $true

$DHCPServerListForm.Controls.AddRange(@($txtServerIP, $lblServerIP, $btnServerIP))


#Logic Section/Functions

function DHCPList {
    param (
        $ComputerName
    )
    $btnHello.Enabled = $false
    $global:poprawnyIP = $null
    $global:poprawnyMAC = $null
    $global:niePoprawnyIP = $null
    $global:niePoprawnyMAC = $null
    $niePoprawnyIP, $niePoprawnyMAC = $null
    $niePoprawnyIP = New-Object System.Collections.Generic.List[System.Object]
    $niePoprawnyMAC = New-Object System.Collections.Generic.List[System.Object]
    $poprawnyIP = New-Object System.Collections.Generic.List[System.Object]
    $poprawnyMAC = New-Object System.Collections.Generic.List[System.Object]

    $lblRozwiazanyIP.Text = ''
    $lblRozwiazanyMAC.Text = ''
    $lblStaryRozwiazanyMAC.Text = ''
    $lblStaryIP.Text = ''

    $lblRozwiazanyIP.Text = 'Szukam...'
    $lblRozwiazanyMAC.Text = 'Szukam...'
    $lblStaryMAC.Visible = $false
    $lblStareIP.Visible = $false
    
    $btnCopyIP.Enabled = $false
    $btnCopyMAC.Enabled = $false

    $wynik = Get-DhcpServerv4Scope -ComputerName $global:DHCPIP.ToString() | Get-DhcpServerv4Lease -ComputerName $global:DHCPIP.ToString() | Where-Object HostName -like ($ComputerName + "*") | Select-Object -Property HostName, IPAddress, ClientId
    foreach ($adres in $wynik) {
        if (Test-Connection $adres.IPAddress -Count 1 -Quiet) {
            $poprawnyIP.Add($adres.IPAddress)
            $poprawnyMAC.Add($adres.ClientId)
        }
        else {
            $niePoprawnyIP.Add($adres.IPAddress)
            $niePoprawnyMAC.Add($adres.ClientId)
        }
    }
 
    $global:niePoprawnyMAC = $niePoprawnyMAC | Select-Object -Unique
    $global:poprawnyIP = $poprawnyIP
    $global:poprawnyMAC = $poprawnyMAC
    $global:niePoprawnyIP = $niePoprawnyIP

    $btnHello.Enabled = $true

    $btnCopyIP.Enabled = $true
    $btnCopyMAC.Enabled = $true
}

function SayHello {
    if ($txtHostname.Text -eq '') {
        $lblRozwiazanyIP.Text = "Brak Hostname!!!"
        [System.Windows.MessageBox]::Show('Brak wpisanej nazwy komputera', 'Toś narobił')
    }
    else {
        #$lblRozwiazanyIP.Text = DHCPList($txtHostname.Text)
        DHCPList($txtHostname.Text)
        if ($global:poprawnyIP -ne $null) {
            $lblRozwiazanyIP.Text = $global:poprawnyIP
        }
        else {
            $lblRozwiazanyIP.Text = 'Brak adresu IP w DHCP'
        }
        if ($global:poprawnyMAC -ne $null) {
            $lblRozwiazanyMAC.Text = $global:poprawnyMAC
        }
        else {
            $lblRozwiazanyMAC.Text = 'Brak adresu MAC w DHCP'
        }
        
        $lblStaryRozwiazanyMAC.Text = $global:niePoprawnyMAC
        $lblStaryIP.Text = $global:niePoprawnyIP
        if ($global:niePoprawnyMAC -ne $null) {
            $lblStaryMAC.Visible = $true
        }
        if ($global:niePoprawnyIP -ne $null) {
            $lblStareIP.Visible = $true
        }

    }
}

function KopiujIP {
    Set-Clipboard -Value $lblRozwiazanyIP.Text
}

function KopiujMAC {
    Set-Clipboard -Value $lblRozwiazanyMAC.Text.Replace("-", "") 
}

function SendServerIP {
    $test = Get-DhcpServerSetting -ComputerName $txtServerIP.Text -ErrorAction SilentlyContinue
    if ($txtServerIP.Text -eq '') {
        $txtServerIP.Text = "150.150.222.20"
        [System.Windows.MessageBox]::Show('Brak wpisanego IP serwera DHCP', 'Toś narobił')
    }
    if ( -not($test)) {
        [System.Windows.MessageBox]::Show('Niepoprawny adres serwera!!!', 'Toś narobił')
        $txtServerIP.Text = "150.150.222.20"
    }
    else {
        $global:DHCPIP = $txtServerIP.Text
        $DHCPServerListForm.Close()
        $DHCPServerListForm.Dispose()
        $DHCPListForm.ShowDialog()
    }

}

# Add the functions to the form
$btnHello.Add_Click({ SayHello })
$btnCopyIP.Add_Click({ KopiujIP })
$btnCopyMAC.Add_Click({ KopiujMAC })
$txtHostname.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            #logic
            SayHello
        }
    })
$btnServerIP.Add_Click({ SendServerIP })


#Display the form
$DHCPServerListForm.ShowDialog()
#$DHCPListForm.ShowDialog()

# Cleans up the form
#$DHCPServerListForm.Dispose()
$DHCPListForm.Dispose()
