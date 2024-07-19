function New-ChoicePrompt {  [cmdletBinding()]
  param( 
      [parameter(mandatory=$true)]$Choices, 
      $Property, 
      $ReadProperty, 
      $ExprLabel,
      [switch]$AllowManualInput, 
      [Scriptblock]$ReadPropertyExpr, 
      $ManualInputLabel = "Type my own"
  )
  if ( $choices[0] -isnot [string] -and !$property ) {"Please include New-ChoicePrompt -Property unless -Choices is an array of strings."; break}
  if ( $choices[0] -is [string] -and ($property -or $ReadProperty) ) {"When New-ChoicePrompt -Choices is an array of strings, please omit -Property and -ReadProperty."; break}
  #if ( $choices[0] -isnot [string] -and $allowManualInput ) {"When New-ChoicePrompt -Choices is a PSobject, please omit -AllowManualInput"; break}
  $x = 0; $script:options = @()
  $script:propty = $property
  $script:choices = $choices
  $manualInputLabel = "<" + $manualInputLabel + ">"
  foreach ($item in $choices) { $value = $null
    $x += 1
    if ($property) { $value = $item | select -expand $property } `
    else {$value = $item}
    if ($readProperty) {
      $readVal = $item | select -expand $readProperty
      $row = new-object -type psObject -property @{Press = $x; 'to select' = $value; $readproperty = $readVal}
    } ` #close if readProperty
    elseif ($readPropertyExpr) `
    {
      $readVal = & $ReadPropertyExpr
      $row = new-object -type psObject -property @{Press = $x; 'to select' = $value; $ExprLabel = $readVal}
    }` #close if readPropertyExpr
    else { $row = new-object -type psObject -property @{'to select' = $value; Press = $x} }
    $script:options += $row
  } #close foreach
  if ($AllowManualInput) {
    $row = new-object -type psObject -property @{'to select' = $manualInputLabel; Press = ($x + 1) }
    $script:options += $row
  } #close if allowManualInput
  if ($ReadProperty) { $script:options | Select Press, "to select", $readproperty | ft -auto }
  elseif ($ReadPropertyExpr) { $script:options | Select Press, "to select", $ExprLabel | ft -auto }
  else { $script:options | Select Press, "to select" | ft -auto }
} #end function new-choicePrompt
 
 $vmhosts = dir "C:\Users\admin\Desktop\Church"
  if ($vmhosts.count -gt 1) {
    do {
      new-choicePrompt -choices $vmhosts -property name
      $in = read-host -prompt 'Please select a target host'
      $range = $options | select -expand press
    } #close do
    until ($range -contains $in)
    $selection = $options | where {$_.press -eq $in} | select -expand 'To select'
    $choice = $choices | where {$_.@($propty) -eq $selection} 
    $vmHost = $choice
  } else {$vmhost = $vmhosts} #close if multiple hosts
  "Target host: " + $vmhost.name