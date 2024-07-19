# Initial message to the user
Write-Host "Enter a letter to start listing files. Keep entering letters to narrow down the list."
Write-Host "Type 'exit' to quit."

# Initialize an array to hold patterns
$patterns = @()

# Infinite loop to keep the script running until the user decides to exit
while ($true) {
    # Read user input
    $input = Read-Host "Enter a letter (or type 'exit' to quit)"
    
    # Exit the loop if the user types 'exit'
    if ($input -eq "exit") {
        break
    }

    # Append the user input to the patterns array
    $patterns += $input

    # Clear the screen to remove the previous list of files
    Clear-Host

    # Display the updated list of files for each pattern
    foreach ($pattern in $patterns) {
        $currentPattern = -join $patterns
        $files = Get-ChildItem -Path . -File -Filter "$currentPattern*"

        if ($files) {
            Write-Host "Files starting with '$currentPattern':"
            $files | ForEach-Object { Write-Host $_.Name }
        } else {
            Write-Host "No files found starting with '$currentPattern'."
        }

        # Remove the last entered letter to narrow down the pattern
        $patterns = $patterns[0..($patterns.Length - 2)]
    }

    # Add a separator between entries
    Write-Host "--------------------"
}

Write-Host "Script terminated. Goodbye!"
