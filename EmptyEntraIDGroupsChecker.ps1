# Load required assemblies for the GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form with increased size to accommodate additional data and make it resizable
$form = New-Object System.Windows.Forms.Form
$form.Text = "Entra ID Security Group Empty Membership Checker"
$form.Size = New-Object System.Drawing.Size(800, 500)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox = $true
$form.MinimumSize = New-Object System.Drawing.Size(800, 500)

# Create a label to show status updates (e.g., group deletion, selected count, or query results)
$label = New-Object System.Windows.Forms.Label
$label.Text = "Loading security groups, please wait..."
$label.Size = New-Object System.Drawing.Size(750, 20)
$label.Location = New-Object System.Drawing.Point(20, 20)
$form.Controls.Add($label)

# Create the 'Rerun Query' button to refresh the list
$rerunButton = New-Object System.Windows.Forms.Button
$rerunButton.Text = "Rerun Query"
$rerunButton.Size = New-Object System.Drawing.Size(100, 30)
$rerunButton.Location = New-Object System.Drawing.Point(20, 50)
$form.Controls.Add($rerunButton)

# Create the 'Delete Selected' button
$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "Delete Selected"
$deleteButton.Size = New-Object System.Drawing.Size(120, 30)
$deleteButton.Location = New-Object System.Drawing.Point(140, 50)
$form.Controls.Add($deleteButton)

# Create the 'Show Selected' button, which will toggle to 'Clear Selected'
$showButton = New-Object System.Windows.Forms.Button
$showButton.Text = "Show Selected"
$showButton.Size = New-Object System.Drawing.Size(100, 30)
$showButton.Location = New-Object System.Drawing.Point(280, 50)
$form.Controls.Add($showButton)

# Create the 'Select All' checkbox
$selectAllCheckBox = New-Object System.Windows.Forms.CheckBox
$selectAllCheckBox.Text = "Select All"
$selectAllCheckBox.Size = New-Object System.Drawing.Size(100, 30)
$selectAllCheckBox.Location = New-Object System.Drawing.Point(400, 55)
$form.Controls.Add($selectAllCheckBox)

# Create the main ListView to display groups with additional columns for Group Type and Membership Type
$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(750, 300)
$listView.Location = New-Object System.Drawing.Point(20, 100)
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.CheckBoxes = $true
$listView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

# Add columns with correct headers and initial widths matching the percentage-based resizing
$listView.Columns.Add("Group Name", 292)
$listView.Columns.Add("Group ID", 182)
$listView.Columns.Add("Group Type", 109)
$listView.Columns.Add("Membership Type", 146)
$form.Controls.Add($listView)

# Global Variable to store the original ListView items
$global:allGroups = @()

# Function to dynamically resize ListView columns based on the form size
function Resize-ListViewColumns {
    $listViewWidth = $listView.ClientSize.Width
    $listView.Columns[0].Width = [math]::Round($listViewWidth * 0.40)
    $listView.Columns[1].Width = [math]::Round($listViewWidth * 0.25)
    $listView.Columns[2].Width = [math]::Round($listViewWidth * 0.15)
    $listView.Columns[3].Width = [math]::Round($listViewWidth * 0.20)
}

# Function to update the selected items count or total count in the label
function Update-Status {
    $selectedCount = $listView.CheckedItems.Count
    if ($selectedCount -eq 0) {
        $label.Text = "$($listView.Items.Count) Cloud-only group(s) loaded."
    } else {
        $label.Text = "$selectedCount item(s) selected."
    }
}

# Function to populate the ListView from stored ListViewItem objects
function Populate-ListView {
    param($items, $clearChecks = $false)  # $clearChecks indicates whether we should reset checkboxes
    $listView.BeginUpdate()
    $listView.Items.Clear()
    foreach ($item in $items) {
        if ($clearChecks) {
            $item.Checked = $false  # Clear checkboxes if $clearChecks is true
        }
        $listView.Items.Add($item)
    }
    $listView.EndUpdate()

    Update-Status
    Resize-ListViewColumns
}

# Function to run the query and populate the ListView (cloud-only groups)
function Run-Query {
    $label.Text = "Running query, please wait..."
    $listView.Items.Clear()

    # Reset the $global:allGroups array ONLY when the query runs
    $global:allGroups = @()

    try {
        # Connect to Microsoft Graph
        Connect-MgGraph -Scopes "Group.ReadWrite.All" -ErrorAction Stop

        # Get all security-enabled groups
        $groups = Get-MgGroup -Filter "securityEnabled eq true" -All

        # Filter to include only cloud-only groups (onPremisesSyncEnabled should be false or null)
        $cloudGroups = $groups | Where-Object { $_.onPremisesSyncEnabled -ne $true }

        # Sort the groups alphabetically by their DisplayName
        $sortedGroups = $cloudGroups | Sort-Object DisplayName

        # For each group, check if it has zero members and populate the ListView with additional information
        $listView.BeginUpdate()
        foreach ($group in $sortedGroups) {
            $members = Get-MgGroupMember -GroupId $group.Id -All
            if ($members.Count -eq 0) {
                $groupType = if ($group.GroupTypes -contains "Unified") { "Microsoft 365" } else { "Security" }
                $membershipType = if ($group.MembershipRule) { "Dynamic" } else { "Assigned" }

                # Create a new ListViewItem for each group
                $item = New-Object System.Windows.Forms.ListViewItem($group.DisplayName)
                $item.SubItems.Add($group.Id)
                $item.SubItems.Add($groupType)
                $item.SubItems.Add($membershipType)
                $item.Checked = $false
                $listView.Items.Add($item)

                # Store the ListViewItem in $global:allGroups
                $global:allGroups += $item
            }
        }
        $listView.EndUpdate()

        if ($listView.Items.Count -eq 0) {
            $label.Text = "No empty cloud-only security groups found."
        } else {
            Update-Status
        }
    }
    catch {
        $label.Text = "Error: $($_.Exception.Message)"
    }

    Resize-ListViewColumns
}

# Event handler for 'Rerun Query' button to refresh the list
$rerunButton.Add_Click({
    Run-Query
})

# Event handler for 'Show Selected' button to toggle between 'Show Selected' and 'Clear Selection'
$showButton.Add_Click({
    if ($showButton.Text -eq "Show Selected") {
        $selectedItems = $listView.CheckedItems
        if ($selectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No groups selected.")
            return
        }

        # Create a new array of selected items
        $selectedGroups = @()
        foreach ($item in $selectedItems) {
            $selectedGroups += $item.Clone()
        }

        Populate-ListView $selectedGroups  # Retain checkboxes for selected items
        $showButton.Text = "Clear Selection"
        $label.Text = "Showing selected groups only."
    }
    else {
        if ($global:allGroups.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No groups stored to restore.", "Error")
        } else {
            Populate-ListView $global:allGroups $true  # Set $clearChecks to true to untick all checkboxes
            
            # Clear the 'Select All' checkbox if it was selected
            $selectAllCheckBox.Checked = $false

            $showButton.Text = "Show Selected"
            $label.Text = "$($global:allGroups.Count) cloud-only group(s) loaded."
        }
    }
})

# Event handler for 'Select All' checkbox
$selectAllCheckBox.Add_CheckedChanged({
    if ($selectAllCheckBox.Checked) {
        $listView.BeginUpdate()
        foreach ($item in $listView.Items) {
            $item.Checked = $true
        }
        $listView.EndUpdate()
    } else {
        $listView.BeginUpdate()
        foreach ($item in $listView.Items) {
            $item.Checked = $false
        }
        $listView.EndUpdate()
    }

    Update-Status
})

# Event handler for 'Delete Selected' button
$deleteButton.Add_Click({
    $selectedItems = $listView.CheckedItems
    if ($selectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No groups selected for deletion.")
        return
    }

    $confirmationMessage = "Are you sure you want to delete the $($selectedItems.Count) selected group(s)?"
    $confirmation = [System.Windows.Forms.MessageBox]::Show($confirmationMessage, "Confirm Deletion", [System.Windows.Forms.MessageBoxButtons]::YesNo)

    if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    $deletedCount = 0
    foreach ($item in $selectedItems) {
        $groupId = $item.SubItems[1].Text
        try {
            Remove-MgGroup -GroupId $groupId -ErrorAction Stop
            $listView.Items.Remove($item)
            $deletedCount++
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error deleting group: $($_.Exception.Message)", "Deletion Failed")
        }
    }

    if ($deletedCount -gt 0) {
        $label.Text = "$deletedCount group(s) deleted."
    } else {
        $label.Text = "No groups were deleted."
    }

    Update-Status
})

# Event handler to track when items are checked or unchecked
$listView.add_ItemChecked({
    Update-Status
})

# Event handler to resize ListView columns when the form is resized
$form.add_SizeChanged({
    Resize-ListViewColumns
})

# Run query when the form is shown
$form.Add_Shown({
    Run-Query
    $form.Activate()
})

# Show the main form
[void]$form.ShowDialog()
