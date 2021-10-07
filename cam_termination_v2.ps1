Add-Type -AssemblyName PresentationFramework

### Start XAML and Reader to use WPF, as well as declare variables for use
[xml]$xaml = @"
<Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

  Title="CAM Termination" Height="350" Width="525" Topmost="True">

    <Grid Background="#FFC8C8C8">
        <Button Name="UserButton" Content="Select User" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="135" Height="20" TabIndex="0"/>
        <TextBox Name="UserTextBox" HorizontalAlignment="Left" Height="20" Margin="150,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="357" IsReadOnly="True" IsEnabled="False"/>
        <TextBox Name="PasswordTextBox" HorizontalAlignment="Left" Height="20" Margin="150,35,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="357" TabIndex="1"/>
        <Label Content="Enter New Password:" HorizontalAlignment="Left" Margin="29,32,0,0" VerticalAlignment="Top" Width="121" Height="23"/>
        <CheckBox Name="OOOCheckBox" Content="Out Of Office?" HorizontalAlignment="Left" Margin="10,60,0,0" VerticalAlignment="Top" TabIndex="2" IsChecked="True"/>
        <CheckBox Name="SharedCheckBox" Content="Shared Mailbox?" HorizontalAlignment="Left" Margin="196,60,0,0" VerticalAlignment="Top" TabIndex="3" IsChecked="True"/>
        <CheckBox Name="LitigationHoldCheckBox" Content="Litigation Hold?" HorizontalAlignment="Left" Margin="403,60,0,0" TabIndex="4" VerticalAlignment="Top"/>
        <Button Name="ManagerButton" Content="Select Manager" HorizontalAlignment="Left" Margin="10,80,0,0" VerticalAlignment="Top" Width="135" Height="20" TabIndex="5"/>
        <TextBox Name="ManagerTextBox" HorizontalAlignment="Left" Height="20" Margin="150,80,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="357" IsReadOnly="True" IsEnabled="False"/>
        <RichTextBox Name="TerminationRichTextBox" HorizontalAlignment="Left" Height="85" Margin="10,195,0,0" VerticalAlignment="Top" Width="497" Background="Black" Foreground="#FF00C8C8" IsReadOnly="True">
            <FlowDocument/>
        </RichTextBox>
        <Button Name="TerminateGoButton" Content="Terminate User" HorizontalAlignment="Left" Margin="10,285,0,0" VerticalAlignment="Top" Width="497" Height="24" IsEnabled="False" TabIndex="7"/>
        <TextBox Name="OOOTextBox" HorizontalAlignment="Left" Height="85" Margin="10,105,0,0" TextWrapping="Wrap" Text="User is no longer with Crisis Assistance Ministry, and this email is not monitored.&#xD;&#xA;&#xD;&#xA;Please contact Manager and your emails will be delivered to the appropriate department.&#xD;&#xA;&#xD;&#xA;Thank you." VerticalAlignment="Top" Width="497" TabIndex="6"/>
    </Grid>

</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
Try{
    $UserForm = [Windows.Markup.XamlReader]::Load($reader)
}
Catch{
    Write-Host "Unable to load Windows.Markup.XamlReader.  Some possible causes for this problem include: .NET Framework is missing, PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."
    Exit
}

#Create Variables For Use In Script Automatically
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name ($_.Name) -Value $UserForm.FindName($_.Name)}
### End XAML and Variables from XAML

Function Write-RichTextBox {
    Param(
        [System.Windows.Controls.RichTextBox]$TextBox,
        [string]$Text,
        [string]$Color = "Cyan"
    )
    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $TextBox.Document.ContentEnd,$TextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $Text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $Color )
    $TextBox.ScrollToEnd()
}

### Logic for enabling/disabling functionality
$PasswordTextBox.Add_TextChanged({
    if (($PasswordTextBox.Text.Length -ge 8) -and ($UserTextBox.Text.Length -ge 2)){
        $TerminateGoButton.IsEnabled = $true
    }
    else{
        $TerminateGoButton.IsEnabled = $false
    }
})

$OOOCheckBox.Add_Unchecked({
    $ManagerButton.IsEnabled = $false
    $OOOTextBox.IsEnabled = $false
})

$OOOCheckbox.Add_Checked({
    $ManagerButton.IsEnabled = $true
    $OOOTextBox.IsEnabled = $true
})
### End Logic/Functions for enabling/disabling functionality

$UserButton.Add_Click({

})

$ManagerButton.Add_Click({

})

$TerminateGoButton.Add_Click({

})

$null = $UserForm.ShowDialog()