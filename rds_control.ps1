chcp 65001 

echo ---------------------------------------------------------------------------
echo ---------------------------SRV003------------------------------------------
echo ---------------------------------------------------------------------------
qwinsta /server:srv003   

echo ---------------------------------------------------------------------------
echo ---------------------------SRV004------------------------------------------
echo ---------------------------------------------------------------------------
qwinsta /server:srv004

echo ---------------------------------------------------------------------------
echo ---------------------------SRV005------------------------------------------
echo ---------------------------------------------------------------------------
qwinsta /server:srv005

echo ---------------------------------------------------------------------------
echo ---------------------------SRV006------------------------------------------
echo ---------------------------------------------------------------------------
qwinsta /server:srv006

echo ---------------------------------------------------------------------------
echo ---------------------------SRV007------------------------------------------
echo ---------------------------------------------------------------------------
qwinsta /server:srv007

echo ---------------------------------------------------------------------------
echo ---------------------------SRV008------------------------------------------
echo ---------------------------------------------------------------------------
qwinsta /server:srv008

echo ---------------------------------------------------------------------------
echo ---------------------------SRV009------------------------------------------
echo ---------------------------------------------------------------------------
qwinsta /server:srv009

echo ---------------------------------------------------------------------------
echo ---------------------------SRV010------------------------------------------
echo ---------------------------------------------------------------------------
qwinsta /server:srv010


$Srv = Read-Host "Please enter ServerName"
$server_name = "SRV0$Srv"


Add-Type -assembly System.Windows.Forms
$Header = "SESSIONNAME", "USERNAME", "ID", "STATUS"
$dlgForm = New-Object System.Windows.Forms.Form
$dlgForm.Text ='Session Connect'
$dlgForm.Width = 400
$dlgForm.AutoSize = $true
$dlgBttn = New-Object System.Windows.Forms.Button
$dlgBttn.Text = 'Connect ask'
$dlgBttn.Width = 100
$dlgBttn.Height = 30
$dlgBttn.Location = New-Object System.Drawing.Point(125,10)
$dlgForm.Controls.Add($dlgBttn)
$dlgBttnsec = New-Object System.Windows.Forms.Button 
$dlgBttnsec.Text = 'Connect confirm' 
$dlgBttnsec.Location = New-Object System.Drawing.Point(15,10) 
$dlgBttnsec.Width = 100
$dlgBttnsec.Height = 30
$dlgForm.Controls.Add($dlgBttn) 
$dlgForm.Controls.Add($dlgBttnsec)
$dlgList = New-Object System.Windows.Forms.ListView
$dlgList.Location = New-Object System.Drawing.Point(0,50)
$dlgList.Width = $dlgForm.ClientRectangle.Width
$dlgList.Height = $dlgForm.ClientRectangle.Height
$dlgList.Anchor = "Top, Left, Right, Bottom"
$dlgList.MultiSelect = $False
$dlgList.View = 'Details'
$dlgList.FullRowSelect = 1;
$dlgList.GridLines = 1
$dlgList.Scrollable = 1
$dlgForm.Controls.add($dlgList)
# Add columns to the ListView
foreach ($column in $Header){
$dlgList.Columns.Add($column) | Out-Null
}

$(qwinsta.exe /server:$server_name | findstr "Active") -replace "^[\s>]" , "" -replace "\s+" , "," | ConvertFrom-Csv -Header $Header | ForEach-Object {
$dlgListItem = New-Object System.Windows.Forms.ListViewItem($_.SESSIONNAME)
$dlgListItem.Subitems.Add($_.USERNAME) | Out-Null
$dlgListItem.Subitems.Add($_.ID) | Out-Null
$dlgListItem.Subitems.Add($_.STATUS) | Out-Null
$dlgList.Items.Add($dlgListItem) | Out-Null
}

$dlgBttnsec.Add_Click( 
{ 
$SelectedItem = $dlgList.SelectedItems[0]
if ($SelectedItem -eq $null){
[System.Windows.Forms.MessageBox]::Show("Select Session")
}else{
$session_id = $SelectedItem.subitems[2].text
$(mstsc /v:$server_name  /shadow:$session_id /control)
#[System.Windows.Forms.MessageBox]::Show($session_id)
}
}
)

$dlgBttn.Add_Click(
{
$SelectedItem = $dlgList.SelectedItems[0]
if ($SelectedItem -eq $null){
[System.Windows.Forms.MessageBox]::Show("Select Session")
}else{
$session_id = $SelectedItem.subitems[2].text
$(mstsc /v:$server_name  /shadow:$session_id /noConsentPrompt /control)
#[System.Windows.Forms.MessageBox]::Show($session_id)
}
}
)
$dlgForm.ShowDialog()
