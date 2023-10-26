
<#
    .SYNOPSIS
        Script to get Personization rule level detail of current item in report format.
    .CreatedBy
        Rinku Jain
    
#>

<#
.SYNOPSIS
    Recurssive function to evaluate rule and insert data in related array based on whether it is a operator or a condition
    
 .Param
  conditionParamNode: Node which needs to evaluate 
  depth: Depth level of element in rule xml, it will help in generating final level report
#>
function Invoke-EvaluateRule {
  param (
      [Parameter(Mandatory=$true)]
      $conditionParamNode,

      [Parameter(Mandatory=$false)]
      [int]$depth = 1
  )

  foreach ($currentNode in $conditionParamNode.ChildNodes) {
      if ($currentNode.LocalName -match $global:Operators) {
          Add-RuleOperator $depth $currentNode.LocalName $currentNode.Uid
          Invoke-EvaluateRule $currentNode ($depth + 1)
      } else {
          Add-ConditionValue $currentNode $depth $currentNode.ParentNode.Uid
      }
  }
}
<#
.SYNOPSIS
    Add all conditions in RulesCondition table for a level and Parent
 
#>
function Add-ConditionValue($conditions, $conditionlevel, $ParentUID ){
  foreach($condition in $conditions) {
    $except=""
    if ($condition.except){    
       $except = "except "                
    }
    
    if(test-path $condition.id){
      $conditionItem = Get-Item "master://" -ID $condition.id 
      $selectedItemsText=$conditionItem.Text
      
      if($condition.selectedItems -ne $null)
      {
          $selectedItemIds= $condition.selectedItems.split("|")
         
          $ct=0
          foreach($id in $selectedItemIds)
          {
              if(test-path $id){
              $selectedItem = Get-Item "master://" -ID $id }
              if ($ct -gt 0){
              $selectedItemsText= $selectedItemsText +', ' +  $selectedItem.Name }
              else
              {  $selectedItemsText=   $selectedItem.Name }
              $ct++
          }
          $selectedItemsText = $conditionItem.Text   -replace  "\[SelectedItems.*\]" , $selectedItemsText
      }
      
      if($condition.rulesid -ne $null)
      {
          $selectedItemIds= $condition.rulesid.split("|")
       
          $ct=0
          foreach($id in $selectedItemIds)
          {
              if(test-path $id){
              $selectedItem = Get-Item "master://" -ID $id }
              if ($ct -gt 0){
              $selectedItemsText= $selectedItemsText +', ' +  $selectedItem.Name }
              else
              {  $selectedItemsText=   $selectedItem.Name }
              $ct++
          }
          $selectedItemsText = $conditionItem.Text   -replace  "\[rulesid.*\]" , $selectedItemsText 
      }
      
      if($condition.operatorid -ne $null)
      {
          $selectedoperatorid= $condition.operatorid
          if(test-path $selectedoperatorid){
              $selectedoperatorItem = Get-Item "master://" -ID $selectedoperatorid 
              $selectedoperatorItemText=   $selectedoperatorItem.Name 
              $selectedvalue= $condition.value
               if(test-path $selectedvalue){
                    $selectedvalueItem = Get-Item "master://" -ID $selectedvalue
                    $opervalue = $selectedoperatorItemText +' ' + $selectedvalueItem.Name
               }
               else{
                   $opervalue =  $selectedoperatorItemText +' ' + $selectedvalue
               }
               
               $selectedItemsText = $conditionItem.Text   -replace  "\[operatorid.*\]" , $opervalue 
              }
        }
      else {
          #check for only Value attribute
           if($condition.Value -ne $null)
              {
                  $selectedItemIds= $condition.Value.split("|")
                 
                  $ct=0
                  foreach($id in $selectedItemIds)
                  {
                      if(test-path $id){
                      $selectedItem = Get-Item "master://" -ID $id }
                      if ($ct -gt 0){
                      $selectedItemsText= $selectedItemsText +', ' +  $selectedItem.Name }
                      else
                      {  $selectedItemsText=   $selectedItem.Name }
                      $ct++
                  }
                  $selectedItemsText = $conditionItem.Text   -replace  "\[value.*\]" , $selectedItemsText
              }
      }    
        
      $selectedCookieName= $condition.CookieName
      if($selectedCookieName -ne $null -and $selectedCookieName -ne '')
      {
       $selectedItemsText = $conditionItem.Text   -replace  "\[CookieName.*\]" , $selectedCookieName 
      }
      
      $condText = $except + $selectedItemsText
      Write-Host "`n            ConditionName: " $conditionItem.Name  
      Write-Host "`n            ConditionText: " $condText
      Add-RuleCondition $conditionlevel $conditionItem.Name  $condText $ParentUID $RenderingNumber $rulecount $false
    }
  }        
}

<#
.SYNOPSIS
    Add a Rule Operator and associated detail in its table
 
#>
function Add-RuleOperator($level, $Operator, $UID){
  $ruleoperator=[RulesOperator]@{
    RenderingNumber=$RenderingNumber
    RuleNo= $rulecount
    Operator = $Operator
    Level= $level
    UID = $UID
  }

  $Script:RulesOperators +=$ruleoperator 
}
<#
.SYNOPSIS
    Add a Rule Condition  and associated detail in its table
 
#>
function Add-RuleCondition($level, $conditionName, $conditionText, $ParentUID, $RenderingNumber, $rulecount, $isNoCondition) {
  $rulesCondition=[RulesCondition]@{
    RenderingNumber = $RenderingNumber
    RuleNo= $rulecount
    ConditionName = $conditionName
    ConditionText = $conditionText
    Level= $level
    ParentUID=$ParentUID
    isNoCondition = $isNoCondition
  } 
    
  $Script:RulesConditions +=$rulesCondition 
}

<#
.SYNOPSIS
    Add a Rule Result  and associated detail in its table
 
#>
function Add-RuleResult($RenderingNumber,$RenderingName, $RuleNo, $ConditionName, $ConditionText, $ActionName, $ActionDataSource) {
  $ruleresult=[RulesResult]@{
      RenderingNumber=$RenderingNumber
      RenderingName=$RenderingName
      RuleNo= $RuleNo
      ConditionName = $ConditionName
      ConditionText = $ConditionText
      ActionName = $ActionName
      ActionDataSource = $ActionDataSource
    } 
  $Script:RulesResults +=$ruleresult 
}

<#
.SYNOPSIS
    Evaluate Rule Conditions table for all condition of a rule, create details based on operator applied and add into Result table 
 
#>
function Invoke-Evaluate-RuleCondition-AddResult ($RenderingNumber, $rullno ){
  #  2. If rule don't have level 1 then merge all condition of it 
  #  3. Else fetch opertor of upper level and in between current condition
  
  $ruleconditiontext =""
  $ruleconditionName =""

  $allrulecondition  = $Script:RulesConditions  | Where-Object {(($_.RenderingNumber -eq $RenderingNumber) -and ($_.RuleNo -eq $rullno) -and ($_.isNoCondition -eq $false)) } 

  $rulelevel= $allrulecondition.level | sort-object -Descending | Get-Unique
  if ($rulelevel  -ne $null){  
      foreach($currentlevel in $rulelevel) {
        if ($currentlevel -eq 1){
          foreach($singlecondition in $allrulecondition) {
          $ruleconditionName = $ruleconditionName +   $singlecondition.ConditionName + "<br>"
          $ruleconditiontext = $ruleconditiontext +   $singlecondition.ConditionText + "<br>"
          }
        }
        else {
            $ruleParentUIDs= ($allrulecondition | Where-Object {$_.level -eq $currentlevel}).ParentUID  | Get-Unique
            #if single parent so we need to club all its condition
            if($ruleParentUIDs.count -eq 1)
            { 
              $prevruleleveloperator = ($script:RulesOperators |  Where-Object {($_.RenderingNumber -eq $RenderingNumber) -and ($_.RuleNo -eq $rullno)  -and ($_.uid -eq ($ruleParentUIDs))}).Operator
              $currentLevelConditions = $allrulecondition  | Where-Object {$_.level -eq $currentlevel}
              foreach($singlecondition in $currentLevelConditions)
              {
                  if ($ruleconditiontext -ne "") {
                    $ruleconditiontext = $ruleconditiontext + " " + $prevruleleveloperator + "<br>"
                  }
                  # Check for dups
                  if($ruleconditionName.Contains($singlecondition.ConditionName) -eq $false){
                    $ruleconditionName = $ruleconditionName +   $singlecondition.ConditionName + "<br>"
                  }
                
                  $ruleconditiontext = $ruleconditiontext + $singlecondition.ConditionText + "<br>"
              }
            } 
            else    # we need to club all its condition based on associated parent and then add parent level operator to it
            { #1. Create condition details for each parent 2. club detail with its level-1 operator
              $parentconditionName=@()
              $parentconditionText=@()
                foreach($currentParentUID in $ruleParentUIDs)
                {
                    $pconditionName=''
                    $pconditiontext=''
                      $parentoperator = ($script:RulesOperators |  Where-Object {($_.RenderingNumber -eq $RenderingNumber) -and ($_.RuleNo -eq $rullno)  -and ($_.uid -eq $currentParentUID)}).Operator
                      $currentParentLevelConditions = $allrulecondition  | Where-Object {($_.level -eq $currentlevel) -and ($_.Parentuid -eq $currentParentUID)}
                      foreach($singlecondition in $currentParentLevelConditions)
                      {
                        if ($pconditiontext -ne "") {
                          $pconditiontext = $pconditiontext + " " + $parentoperator + "<br>"
                        }
                        $pconditionName = $pconditionName +   $singlecondition.ConditionName + "<br>"
                        $pconditiontext = $pconditiontext +   $singlecondition.ConditionText + "<br>"
                      } 
                      $parentconditionName+=$pconditionName
                      $parentconditionText+=$pconditiontext
                      
                } 
                
                $uidtolookup= ($Script:RulesConditions |  Where-Object {($_.RenderingNumber -eq $RenderingNumber) -and ($_.RuleNo -eq $rullno)  -and ($_.level -eq ($currentlevel-1)) -and ($_.isNoCondition -eq $true)}).parentuid
                $parentleveloperator = ($script:RulesOperators |  Where-Object {($_.RenderingNumber -eq $RenderingNumber) -and ($_.RuleNo -eq $rullno)  -and ($_.uid -eq $uidtolookup)}).Operator

                for($parentindex=0; $parentindex -le  $parentconditionName.count-1; $parentindex++){
                  if ($ruleconditiontext.length -ne 0) {
                    $ruleconditiontext = $parentconditionText[$parentindex] + " " + $parentleveloperator + "<br>"
                  }
                  $ruleconditionName = $ruleconditionName +    $parentconditionName[$parentindex] + "<br>"
                  $ruleconditiontext = $ruleconditiontext +   $parentconditionText[$parentindex] + "<br>"
                }
            }
        }
      }
  } 

  $ruleaction = $Script:Rules | Where-Object {($_.RenderingNumber -eq $RenderingNumber) -and ($_.RuleNo -eq $rullno)} 
  $ActionName = $ruleaction.ActionName
  $ActionDataSource =  $ruleaction.ActionDataSource

  Add-RuleResult $RenderingNumber $ruleaction.RenderingName $rullno $ruleconditionName $ruleconditiontext $ActionName $ActionDataSource
} 


class Rule {
  [int] $RenderingNumber
  [string] $RenderingName
  [int]$RuleNo
  [string]$ActionName
  [string]$ActionDataSource
}

class RulesOperator {
  [int] $RenderingNumber
  [int]$RuleNo
  [string]$Operator
  [int]$Level
  [string]$UID
}

class RulesCondition {
  [int] $RenderingNumber 
  [int]$RuleNo
  [int]$Level
  [string]$ConditionName
  [string]$ConditionText
  [string]$ParentUID
  [bool] $isNoCondition 
}  

class RulesResult {
  [int] $RenderingNumber 
  [string] $RenderingName
  [int]$RuleNo
  [string]$ConditionName
  [string]$ConditionText
  [string]$ActionName
  [string]$ActionDataSource
}  

$itemtolookup= Get-Item '/sitecore/content/Wealth Site/Home/Test Pages/Test Researching investments Performance Tab'
#$itemtolookup =  Get-Item . -Language $SitecoreContextItem.Language.Name
$renderings = Get-Rendering -Item $itemtolookup -FinalLayout | Where-Object { ![string]::IsNullOrEmpty($_.Rules)} 
$renderingnumber=0 
$rulecount=0

[Rule[]]$Rules = @()  
[RulesOperator[]] $Script:RulesOperators = @()  
[RulesCondition[]]$Script:RulesConditions = @()  
[RulesResult[]]$Script:RulesResults = @()   
   
Write-Host "Item:  $itemtolookup"

foreach($rendering in $renderings) {
  $renderingnumber++
  if($rendering ) 
  {
      $renderingItem= Get-Item "master://" -ID $rendering.ItemID
      Write-Host "`n##############################  Rendering: "$renderingItem.Name , Rendering number: $renderingnumber " #########################################"

      $xml = New-Object -TypeName System.Xml.XmlDocument
      $xml.LoadXml($rendering.Rules)
      write-host $rendering.Rules
      $operatorslist = @('and','or')
      $operators= [string]::Join('|',$operatorslist)
      $rulecount = 0

      foreach($rule in $xml.ruleset.rule) {
          $rulecount++
          Write-Host "`n  Rule $rulecount.    Rule Name: " $rule.Name
          $conditions=$rule.conditions
          $childnode = $conditions.childnodes
         #if there is no operator then we can directly get all its conditions
          foreach($condition in $conditions) {
            $nodename =  $condition | gm -MemberType property | select Name
            if ($nodename -match $operators ){
                Invoke-EvaluateRule  $condition 
              }
              else{
                  Add-ConditionValue $condition.condition 1
            }
          }
         # To add rule details of current rule in Rule table
          if ($rule.actions.action -ne $null -and $rule.actions.action -ne '')
          {
            $datasourcedetail =''
            $actiondetail=''
            foreach($actionID in  $rule.actions.action ) {   
              $actionItem = Get-Item "master://" -ID $actionID.id
              if(($actionID.Datasource -ne $null) -and ($actionID.Datasource -ne '') ) { 
                $dbname=(get-item  "master://" -ID $actionID.Datasource).name
                $datasourcedetail = if($datasourcedetail -ne ''){ $datasourcedetail + ' , ' + $dbname } else {$dbname }  
              }
              $actiondetail =if($actiondetail -ne ''){ $actiondetail + ' , ' + $actionItem.Name } else {$actionItem.Name }
            }
            
            $ruledetail=[Rule]@{
              RenderingNumber=$RenderingNumber
              RenderingName = $renderingItem.Name
              RuleNo= $rulecount
              ActionName = $actiondetail
              ActionDataSource= $datasourcedetail 
            } 
          }
          else {
            $ruledetail=[Rule]@{
              RenderingNumber=$RenderingNumber
              RenderingName = $renderingItem.Name
              RuleNo= $rulecount
              ActionName = 'Show Rendering'
              ActionDataSource= '' 
            } 
          }

          Write-Host "`n    ActionName: " $actiondetail   ", Datasource: " $datasourcedetail
        
          $Rules +=$ruledetail 
          $dbname=""
      }
  }
}

$renderinglist = $Rules.RenderingNumber |  Sort-Object | Get-Unique

foreach($renderingrule in $renderinglist){
  $activeRenderingNumber =   $renderingrule
  
  #  $Script:RulesConditions | Where-Object {($_.RenderingNumber -eq 11 -and $_.RuleNo -eq 1) }  | format-table | show-listview
  #  $Script:RulesOperators | Where-Object {($_.RenderingNumber -eq 11 -and $_.RuleNo -eq 1) } | format-table 
  
  #  check for any missed entry in RulesCondition due to not having conditions under operator
#  if ($activeRenderingNumber -eq 11) {
  $rop=$script:RulesOperators |  Where-Object {($_.RenderingNumber -eq $activeRenderingNumber) } 
  foreach($rulesOperatoritem in $rop) {
    $operatorcount =  ($Script:RulesConditions | Where-Object {$_.parentuid -eq $rulesOperatoritem.uid}).count
    if ($operatorcount -eq 0) {
        Add-RuleCondition ($rulesOperatoritem.level + 1) ' ' ' ' $rulesOperatoritem.uid $activeRenderingNumber $rulesOperatoritem.ruleno $true
    }
  } 
#  } 

  #To evaluate all rules of active rendering and save into result table
  $allRenderingCondition =  $Script:RulesConditions | Where-Object {($_.RenderingNumber -eq $activeRenderingNumber) } 
  $rclist = $allRenderingCondition.RuleNo |  Sort-Object | Get-Unique
  foreach-object {
    foreach($rno in $rclist){
      Invoke-Evaluate-RuleCondition-AddResult $activeRenderingNumber $rno
    } 
  } 
}

#Finally show result in report
$Script:RulesResults |  Show-ListView -InfoTitle "Personization Report" -InfoDescription (" Here Rendering Number is evaluated w.r.t. to Personization, Item : " + $itemtolookup.fullpath) -Property `
@{Label="Rendering Name"; Expression={$_.RenderingName} }, 
@{Label="Rendering Number"; Expression={$_.RenderingNumber} }, 
@{Label="Rule No"; Expression={$_.RuleNo} }, 
@{Label="Condition"; Expression={$_.ConditionName  } }, 
@{Label="Condition Text"; Expression={$_.ConditionText  } }, 
@{Label="Action"; Expression={$_.ActionName  } }, 
@{Label="Action DataSource"; Expression={$_.ActionDataSource  } }