param
(
	$OutPutFile = (join-path (pwd) ($env:USERDNSDOMAIN+"_attribute_flow.html")),
	[switch]$Debug
	
)

Add-Type -AssemblyName System.Web

function Format-XML-HTML([xml]$xml)
{ 
    $StringWriter = New-Object System.IO.StringWriter 
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
    $xmlWriter.Formatting = "indented" 
    $xmlWriter.Indentation = 2
    $xml.WriteContentTo($XmlWriter) 
    $XmlWriter.Flush() 
    $StringWriter.Flush()
    return [Web.HttpUtility]::HtmlEncode($StringWriter.ToString()).Replace("`r`n","</br>").Replace(" ","&nbsp;")
}

#$PortalParameters = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Forefront Identity Manager\2010\Portal"
$SynchronizationServiceParameters = Get-ItemProperty HKLM:\SYSTEM\CurrentControlSet\services\FIMSynchronizationService\Parameters


#$PortalUrl = $PortalParameters.BaseSiteCollectionURL;
$SQLServerInstans = ("localhost",$SynchronizationServiceParameters.Server)[$SynchronizationServiceParameters.Server.Length -gt 0]+("","\")[$SynchronizationServiceParameters.SQLInstance.Length -gt 0]+("",$SynchronizationServiceParameters.SQLInstance)[$SynchronizationServiceParameters.SQLInstance.Length -gt 0]
$DBName = $SynchronizationServiceParameters.DBName

$ConnectionString = "Data Source=$SQLServerInstans;Initial Catalog=$DBName;Integrated Security=SSPI;"
$CurentDir = (pwd)

$OIDTabel = 
@{
	"1.3.6.1.4.1.1466.115.121.1.12" = "Reference(DN)";
	"1.3.6.1.4.1.1466.115.121.1.15" = "String";
	"1.3.6.1.4.1.1466.115.121.1.5" = "Binary";
	"1.3.6.1.4.1.1466.115.121.1.7" = "Bit";
	"1.3.6.1.4.1.1466.115.121.1.27" = "Integer";
}

#Connection
$Connection = New-Object System.Data.SqlClient.SqlConnection ($ConnectionString)
$Connection.Open()

if($Connection.State -eq "Closed"){
	write-host -ForegroundColor Blue -BackgroundColor Yellow "Error connecting to SQL"
	write-host -ForegroundColor Blue -BackgroundColor Yellow $SQLServerInstans
	write-host -ForegroundColor Blue -BackgroundColor Yellow $DBName
	exit
}


#MV schema and import
$SQLstring = "select mv_schema_xml,import_attribute_flow_xml from mms_server_configuration"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SQLstring,$Connection)
$SqlCmd.CommandTimeout = $Connection.ConnectionTimeout

$reader = $SqlCmd.ExecuteReader()
while ($reader.Read())
{
	[xml]$mv_schema_xml = $reader["mv_schema_xml"]
	if($Debug){ $mv_schema_xml.Save((Join-Path $CurentDir "mv_schema_xml.xml")) }
	
	[xml]$import_attribute_flow_xml = $reader["import_attribute_flow_xml"]
	if($Debug){ $import_attribute_flow_xml.Save((Join-Path $CurentDir "import_attribute_flow_xml.xml")) }
}

$reader.Close()
$SqlCmd.Dispose()


#MV SynchronizationRules
$SynchronizationRules = New-Object hashtable
$SynchronizationRulesvalue = New-Object hashtable 

$SQLstring = "select mv.*"
$SQLstring += ",STUFF((`
			select CAST('|' + RTRIM(mvm.string_value_not_indexable) AS VARCHAR(MAX))`
			from mms_metaverse_multivalue mvm`
			where mv.object_id = mvm.object_id`
			and mvm.attribute_name = 'persistentFlow'`
			FOR XML PATH (''))`
			, 1, 1, ''`
) as persistentFlow "
$SQLstring += "from mms_metaverse mv where object_type = 'synchronizationRule'"

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SQLstring,$Connection)
$SqlCmd.CommandTimeout = $Connection.ConnectionTimeout

$synchronizationRuleTable = New-Object system.Data.DataTable "synchronizationRule"
$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd
$RowCount = $Adapter.Fill($synchronizationRuleTable)
$Adapter.Dispose()
$SqlCmd.Dispose()

if($Debug){
	$synchronizationRuleTable.WriteXml((Join-Path $CurentDir "synchronizationRuleTable.xml"))
}

foreach($row in $synchronizationRuleTable.Rows){
	[void]$SynchronizationRules.Add("{"+$row["object_id"].ToString().ToUpper()+"}",$row["displayname"])
}

# mms_management_agent
$SQLstring = "select * from mms_management_agent"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SQLstring,$Connection)
$SqlCmd.CommandTimeout = $Connection.ConnectionTimeout

$attribute_exportMA = New-Object 'system.collections.generic.dictionary[string,System.Collections.Generic.HashSet[string]]'
$agentList = New-Object 'system.collections.generic.dictionary[string,string]'

$ListMA = New-Object System.Text.StringBuilder
$MaOut = New-Object System.Text.StringBuilder
[void]$MaOut.Append("<h1>Management agent</h1>`r`n")
[void]$ListMA.AppendFormat("<h1>{0}</h1></br>Generated {1}</br><h2>Management agent List</h2>`r`n", ($env:USERDNSDOMAIN),(get-date).ToString("yyy-MM-dd HH:mm:ss"))

$reader = $SqlCmd.ExecuteReader()
while ($reader.Read())
{
	$ma_guid           = $reader["ma_id"].ToString().ToUpper()
	$ma_name           = $reader["ma_name"]
	$ma_type           = $reader["ma_type"]
	$capabilities_mask = $reader["capabilities_mask"]
	$ma_export_type    = $reader["ma_export_type"]
	$ma_description    = $reader["ma_description"]
	
	[void]$agentList.Add("{$ma_guid}",$ma_name)
	[void]$ListMA.AppendFormat("<a href='#{0}'>{0}</a></br>",$ma_name)

	try{[xml]$attribute_inclusion_xml    	= $reader["attribute_inclusion_xml"];if($Debug){ $attribute_inclusion_xml.Save((Join-Path $CurentDir "attribute_inclusion_xml.$ma_name .xml")) }         } catch{}
	try{[xml]$component_mappings_xml    	= $reader["component_mappings_xml"];if($Debug){ $component_mappings_xml.Save((Join-Path $CurentDir "component_mappings_xml.$ma_name .xml")) }          } catch{}
	try{[xml]$controller_configuration_xml 	= $reader["controller_configuration_xml"];if($Debug){ $controller_configuration_xml.Save((Join-Path $CurentDir "controller_configuration_xml.$ma_name .xml")) }    } catch{}
	try{[xml]$dn_construction_xml    		= $reader["dn_construction_xml"];if($Debug){ $dn_construction_xml.Save((Join-Path $CurentDir "dn_construction_xml.$ma_name .xml")) }             } catch{}
	try{[xml]$export_attribute_flow_xml   	= $reader["export_attribute_flow_xml"];if($Debug){ $export_attribute_flow_xml.Save((Join-Path $CurentDir "export_attribute_flow_xml.$ma_name .xml")) }       } catch{}
	try{[xml]$join_rule_xml    				= $reader["join_rule_xml"];if($Debug){ $join_rule_xml.Save((Join-Path $CurentDir "join_rule_xml.$ma_name .xml")) }                   } catch{}
	try{[xml]$ma_extension_xml    			= $reader["ma_extension_xml"];if($Debug){ $ma_extension_xml.Save((Join-Path $CurentDir "ma_extension_xml.$ma_name .xml")) }                } catch{}
	try{[xml]$passwordsync_xml    			= $reader["passwordsync_xml"];if($Debug){ $passwordsync_xml.Save((Join-Path $CurentDir "passwordsync_xml.$ma_name .xml")) }                } catch{}
	try{[xml]$private_configuration_xml    	= $reader["private_configuration_xml"];if($Debug){ $private_configuration_xml.Save((Join-Path $CurentDir "private_configuration_xml.$ma_name .xml")) }       } catch{}
	try{[xml]$projection_rule_xml    		= $reader["projection_rule_xml"];if($Debug){ $projection_rule_xml.Save((Join-Path $CurentDir "projection_rule_xml.$ma_name .xml")) }             } catch{}
	try{[xml]$provisioning_cleanup_xml   	= $reader["provisioning_cleanup_xml"];if($Debug){ $provisioning_cleanup_xml.Save((Join-Path $CurentDir "provisioning_cleanup_xml.$ma_name .xml")) }        } catch{}
	try{[xml]$stay_disconnector_xml    		= $reader["stay_disconnector_xml"];if($Debug){ $stay_disconnector_xml.Save((Join-Path $CurentDir "stay_disconnector_xml.$ma_name .xml")) }           } catch{}
	try{[xml]$ui_settings_xml    			= $reader["ui_settings_xml"];if($Debug){ $ui_settings_xml.Save((Join-Path $CurentDir "ui_settings_xml.$ma_name .xml")) }                 } catch{}

	[void]$MaOut.AppendFormat("<table id='{0}'><tr><td><h1>{0}</h1></td></tr><tr><td>Type:{1}</td><td>Capabilities:{2}</td><td>export_type:{3}</td><td>Description:{4}</td><td>GUID:{5}</td></tr>`r`n", $ma_name,$ma_type,$capabilities_mask,$ma_export_type,$ma_description,$ma_guid)
	$assemblyname = ""
	if($ma_extension_xml.extension.'assembly-name'){
		$assemblyname = $ma_extension_xml.extension.'assembly-name'
	}
	[void]$MaOut.AppendFormat("<tr><td>extension assembly</td><td>{0}</td></tr>`r`n", $assemblyname)
	[void]$MaOut.Append("<tr></tr>`r`n")
	[void]$MaOut.AppendFormat("<tr><th>Type</th><th>CS</th><th>import/export</th><th>MV</th><th>Rule(s)</th></tr>")
	
	#Join object cs - attibute - mv - extension rule name
	$OldjoinName = ""
	if($join_rule_xml)
	{
		foreach($profile in $join_rule_xml.join.'join-profile'){
			$count=1
			foreach($join in $profile.'join-criterion'){
				$mvobject = $join.search.'mv-object-type'
				if($mvobject.Length -lt 1){$mvobject = "Any"}
				$joinName = "{0}{1}" -f $profile.'cd-object-type', $mvobject
				if($joinName -ne $OldjoinName) { [void]$MaOut.AppendFormat("<tr><td>Join Object</td><td>{0}</td><td><-></td><td>{1}</td></tr>",$profile.'cd-object-type',$mvobject) } 
				$OldjoinName = $joinName
				
				$scriptcontext = ""
				if($join.search.'attribute-mapping'.'direct-mapping'){
					if($join.search.'attribute-mapping'.'direct-mapping'.'src-attribute'.'#text'){
						$css = [string]::Join(",",$join.search.'attribute-mapping'.'direct-mapping'.'src-attribute'.'#text')
					}else{
						$css = [string]::Join(",",$join.search.'attribute-mapping'.'direct-mapping'.'src-attribute')
					}
				} else{
					$scriptcontext = $join.search.'attribute-mapping'.'scripted-mapping'.'script-context'
					$css = [string]::Join(",",$join.search.'attribute-mapping'.'scripted-mapping'.'src-attribute')
				}
				[void]$MaOut.AppendFormat("<tr><td>{0}</td><td>{1}</td><td><-></td><td>",$count,$css)
				$joinS = ""
				$join.search.'attribute-mapping'.'mv-attribute' | % { [void]$MaOut.AppendFormat("{1}<a href='#{0}'>{0}</a>",$_,$joinS);$joinS="," }
				[void]$MaOut.AppendFormat("</td><td>{0}</td></tr>`r`n",$scriptcontext)
				$count++
				foreach($resolution in $join.resolution){
					if($resolution.'script-context'){
						#[void]$MaOut.AppendFormat("{0}`r`n",$resolution.'script-context')
						#[void]$MaOut.AppendFormat("{0}`r`n",$resolution.'script-context')
					}
				}
			}
		}
	}
	[void]$MaOut.Append("<tr></tr>`r`n")
	
	#Projektion object cs - attibute - mv - extension rule name
	if($projection_rule_xml){
		if($projection_rule_xml.projection.'class-mapping'.type){
			if($projection_rule_xml.projection.'class-mapping'.type -eq "sync-rule"){
				$SyncRuleID = ($projection_rule_xml.projection.'class-mapping'.'sync-rule-id').ToString().ToUpper()
				$SyncRuleName = $SynchronizationRules[$SyncRuleID]
				$Type = "sync-rule(<a href='#{0}'>{1}</a>)" -f $SyncRuleID,$SyncRuleName
			}else{
				$Type = $projection_rule_xml.projection.'class-mapping'.type
			}
			[void]$MaOut.AppendFormat("<tr><td>Projektion:{0}</td></tr>`r`n", $Type)
		}
		if($projection_rule_xml.projection.'class-mapping'.'cd-object-type'){
			[void]$MaOut.AppendFormat("<tr><td>Projektion:{0}</td></tr>`r`n", $projection_rule_xml.projection.'class-mapping'.'cd-object-type')
		}
		if($projection_rule_xml.projection.'class-mapping'.'mv-object-type'){
			[void]$MaOut.AppendFormat("<tr><td>Projektion:{0}</td></tr>`r`n", $projection_rule_xml.projection.'class-mapping'.'mv-object-type')
		}
		[void]$MaOut.Append("<tr></tr>`r`n")
	}
	
	# DN = dn-construction.attribute
	if($dn_construction_xml){
		#[void]$MaOut.Append()
		#[void]$MaOut.Append("`r`n")
	}
	
	#Flow Rules
	$Objectflow = New-Object 'system.collections.generic.dictionary[string,System.Collections.Generic.HashSet[string]]'
	#Import
	foreach($flow in $import_attribute_flow_xml.SelectNodes("/import-attribute-flow/import-flow-set/import-flows/import-flow[@src-ma='{$ma_guid}']")){
		$rulename = ""
		
		$flowName = "<tr><td>Flow object</td><td>{0}</td><td><-></td><td>{1}</td></tr>`r`n" -f $flow.'cd-object-type',$flow.ParentNode.ParentNode.'mv-object-type'
		$srcattributes=""
		if($flow.'scripted-mapping'.'src-attribute'){
			$list=New-Object System.Collections.Generic.HashSet[string]
			$rulename = $flow.'scripted-mapping'.'script-context'
			
			$flow.'scripted-mapping'.'src-attribute' | %{
				if($_.'#text') { [void]$list.Add("{0}" -f $_.'#text') }
				else { [void]$list.Add("{0}" -f $_) }
			}
			$srcattributes = [string]::Join(",",$list)
		} elseif($flow.'direct-mapping'.'src-attribute'){
			$list=New-Object System.Collections.Generic.HashSet[string]
			$rulename = $flow.'direct-mapping'.'script-context'
			
			$flow.'direct-mapping'.'src-attribute' | %{
				if($_.'#text') { [void]$list.Add("{0}" -f $_.'#text') }
				else { [void]$list.Add("{0}" -f $_) }
			}
			$srcattributes = [string]::Join(",",$list)
		} elseif($flow.'constant-mapping'.'constant-value'){
			$rulename = "constant"
			$srcattributes = "'"+$flow.'constant-mapping'.'constant-value'+"'"
		} elseif($flow.'sync-rule-mapping') {
			$list=New-Object System.Collections.Generic.HashSet[string]
			
			if($flow.'sync-rule-mapping'.'sync-rule-id'){
				$SyncRuleID = ($flow.'sync-rule-mapping'.'sync-rule-id').ToUpper()
				if(-NOT $SynchronizationRulesvalue.ContainsKey($SyncRuleID)){
					$syncrulevalue = $flow.'sync-rule-mapping'.'sync-rule-value'.InnerXml
					[void]$SynchronizationRulesvalue.Add($SyncRuleID,$syncrulevalue)
				}
				$SyncRuleName = $SynchronizationRules[$SyncRuleID]
				if($SyncRuleName.Length -eq 0){
					$SyncRuleName = $SyncRuleID
				}
				$rulename = "(sync-rule '<a href='#{0}'>{1}</a>'){2}" -f $SyncRuleID,$SyncRuleName,($flow.'sync-rule-mapping'.'mapping-type')
				
				#$SyncRuleName = $SynchronizationRules[($flow.'sync-rule-mapping'.'sync-rule-id')]
				#$rulename = "(sync-rule '$SyncRuleName')" + $flow.'sync-rule-mapping'.'mapping-type'
			}

			if($flow.'sync-rule-mapping'.'src-attribute'.Count -gt 0){
				$flow.'sync-rule-mapping'.'src-attribute' | %{
					[void]$list.Add($_) 
				}
			}else{
				if($flow.'sync-rule-mapping'.'src-attribute'.'#text'){
					[void]$list.Add($flow.'sync-rule-mapping'.'src-attribute'.'#text')
				}else{
					[void]$list.Add($flow.'sync-rule-mapping'.'src-attribute')
				}
			}
			
			$srcattributes = [string]::Join(",",$list)
		}

		$flowRule = "<tr><td></td><td>{0}</td><td>-></td><td><a href='#{1}'>{1}</a></td><td>{2}</td></tr>`r`n" -f $srcattributes,$flow.ParentNode.'mv-attribute',$rulename

		if($Objectflow.ContainsKey($flowName)){
			[void]$Objectflow[$flowName].Add($flowRule)
		}
		else{
			$flowRuleList=New-Object System.Collections.Generic.HashSet[string]
			[void]$flowRuleList.Add($flowRule)
			[void]$Objectflow.Add($flowName,$flowRuleList)
		}
	}
	
	#Export
	#Attibute flow cs - mv - extension rule name 	export_attribute_flow_xml
	if($export_attribute_flow_xml){
		if($export_attribute_flow_xml.'export-attribute-flow'){
			foreach($flowset in $export_attribute_flow_xml.'export-attribute-flow'.'export-flow-set'){

				$flowName = "<tr><td>Flow object</td><td>{0}</td><td><-></td><td>{1}</td></tr>`r`n" -f $flowset.'cd-object-type',$flowset.'mv-object-type'
				
				foreach($flow in $flowset.'export-flow'){
					$rulename = ""
					$srcattributes=""
					$attlist=New-Object System.Collections.Generic.HashSet[string]
					
					if($flow.'scripted-mapping'.'src-attribute'){
						$rulename = $flow.'scripted-mapping'.'script-context'
						$list=New-Object System.Collections.Generic.HashSet[string]

						$flow.'scripted-mapping'.'src-attribute' | %{
							if($_.'#text') { 
								[void]$list.Add("<a href='#{0}'>{0}</a>" -f $_.'#text') 
								[void]$attlist.Add($_.'#text') 
								}
							else { 
								[void]$list.Add("<a href='#{0}'>{0}</a>" -f $_) 
								[void]$attlist.Add($_) 
								}
						}
						$srcattributes = [string]::Join(",",$list)
					} elseif($flow.'direct-mapping'.'src-attribute'){
						$rulename = $flow.'direct-mapping'.'script-context'
						$list=New-Object System.Collections.Generic.HashSet[string]

						$flow.'direct-mapping'.'src-attribute' | %{
							if($_.'#text') { 
								[void]$list.Add("<a href='#{0}'>{0}</a>" -f $_.'#text') 
								[void]$attlist.Add($_.'#text') 
							}
							else { 
								[void]$list.Add("<a href='#{0}'>{0}</a>" -f $_) 
								[void]$attlist.Add($_) 
							}
						}
						$srcattributes = [string]::Join(",",$list)
					} elseif($flow.'constant-mapping'.'constant-value'){
						$rulename = "constant"
						$srcattributes = "'"+$flow.'constant-mapping'.'constant-value'+"'"
					}elseif($flow.'sync-rule-mapping') {
						
						if($flow.'sync-rule-mapping'.'sync-rule-id'){
							$SyncRuleID = ($flow.'sync-rule-mapping'.'sync-rule-id').ToUpper()
							if(-NOT $SynchronizationRulesvalue.ContainsKey($SyncRuleID)){
								$syncrulevalue = $flow.'sync-rule-mapping'.'sync-rule-value'.InnerXml
								[void]$SynchronizationRulesvalue.Add($SyncRuleID,$syncrulevalue)
							}
							$SyncRuleName = $SynchronizationRules[$SyncRuleID]
							if($SyncRuleName.Length -eq 0){
								$SyncRuleName = $SyncRuleID
							}
							$rulename = "(sync-rule '<a href='#{0}'>{1}</a>'){2}" -f $SyncRuleID,$SyncRuleName,($flow.'sync-rule-mapping'.'mapping-type')
							
							#$SyncRuleName = $SynchronizationRules[$flow.'sync-rule-mapping'.'sync-rule-id']
							#$rulename = "(sync-rule '$SyncRuleName')" + $flow.'sync-rule-mapping'.'mapping-type'
						}
						
						$list=New-Object System.Collections.Generic.HashSet[string]
						if($flow.'sync-rule-mapping'.'src-attribute'.Count -gt 0){
							$flow.'sync-rule-mapping'.'src-attribute' | %{
								[void]$list.Add("<a href='#{0}'>{0}</a>" -f $_) 
								[void]$attlist.Add($_) 
							}
						}else{
							[void]$list.Add("<a href='#{0}'>{0}</a>" -f $flow.'sync-rule-mapping'.'src-attribute') 
							[void]$attlist.Add($_) 
						}
						$srcattributes = [string]::Join(",",$list)
					}

					if($flow.'suppress-deletions' -eq "false"){ $AllowNull = ",Allow null" } else { $AllowNull="" }
					$CSatt = $flow.'cd-attribute'
					$flowRule = "<tr><td></td><td>{0}</td><td><-</td><td>{1}</td><td>{2}{3}</td></tr>`r`n" -f $CSatt,$srcattributes,$rulename,$AllowNull
					
					foreach($MvAtt in $attlist){
						$MVstring = "<tr><td></td><td>-></td><td><a href='#{0}'>{0}</a>({1})</td><td></td></tr>`r`n" -f $ma_name,$CSatt
						if($attribute_exportMA.ContainsKey($MvAtt)){
							[void]$attribute_exportMA[$MvAtt].Add($MVstring)
						}
						else{
							$objl=New-Object System.Collections.Generic.HashSet[string]
							[void]$objl.Add($MVstring)
							[void]$attribute_exportMA.Add($MvAtt,$objl)
						}
					}
					
					if($Objectflow.ContainsKey($flowName)){
						[void]$Objectflow[$flowName].Add($flowRule)
					}
					else{
						$flowRuleList=New-Object System.Collections.Generic.HashSet[string]
						[void]$flowRuleList.Add($flowRule)
						[void]$Objectflow.Add($flowName,$flowRuleList)
					}
				}
			}
		}
	}
	foreach($key in $Objectflow.Keys){
		[void]$MaOut.AppendFormat("<tr><td>{0}</td></tr>",$key)
		[void]$MaOut.Append([string]::Join("",$Objectflow[$key]))
		[void]$MaOut.Append("<tr></tr>`r`n")
	}
	
	
	#Provisionering object cs ? in 
	
	#Deprovisionering object cs
	if($provisioning_cleanup_xml){
		[void]$MaOut.AppendFormat("<tr><td>Deprovisionering:{0}</td></tr><tr><td>enable-recall:{1}</td></tr>`r`n", $provisioning_cleanup_xml.'provisioning-cleanup'.action, $import_attribute_flow_xml.SelectSingleNode("/import-attribute-flow/per-ma-options/ma-options[@ma-id='{$ma_guid}']").'enable-recall')
	}
	#CS full list
	#CS - object - type?
	#CS - attibutes - type 
	[void]$MaOut.Append("<tr></tr>`r`n")
	if($attribute_inclusion_xml){
		[void]$MaOut.AppendFormat("<tr><th>CS attribute</th></tr>")
		foreach($attribute in $attribute_inclusion_xml.'attribute-inclusion'.'attribute'){
			[void]$MaOut.AppendFormat("<tr><td>{0}</td></tr>",$attribute)
		}
	}
	
	[void]$MaOut.Append("</table></br></br>`r`n")
	
	$attribute_inclusion_xml    	= $null
	$component_mappings_xml    		= $null
	$controller_configuration_xml 	= $null
	$dn_construction_xml    		= $null
	$export_attribute_flow_xml   	= $null
	$join_rule_xml    				= $null
	$ma_extension_xml    			= $null
	$passwordsync_xml    			= $null
	$private_configuration_xml    	= $null
	$projection_rule_xml    		= $null
	$provisioning_cleanup_xml   	= $null
	$stay_disconnector_xml    		= $null
	$ui_settings_xml    			= $null
}
$reader.Close()
$SqlCmd.Dispose()
$Connection.Close()


$schemaOut = New-Object System.Text.StringBuilder
[void]$schemaOut.Append("<h1>MV attribute</h1>`r`n")

foreach($attribute in $mv_schema_xml.dsml.'directory-schema'.'attribute-type'){
	[void]$schemaOut.AppendFormat("<table id='{0}'><tr><td><h1>{0}</h1></td><td></td></tr><tr><td>mulitvalue:{1}</td><td>indexed:{2}</td><td>syntax:{3}</td></tr>`r`n", $attribute.Name,!($attribute.'single-value'), ([Boolean]$attribute.indexed),$OIDTabel[$attribute.syntax])
	
	$ma=New-Object System.Collections.Generic.HashSet[string]
	$aname = $attribute.Name
	$flowss = $import_attribute_flow_xml.SelectNodes("/import-attribute-flow/import-flow-set/import-flows[@mv-attribute='$aname']")
	if($flowss){
		[void]$schemaOut.Append("`r`n")
		foreach($flows in $flowss){
		[void]$schemaOut.AppendFormat("<tr><td>{0}</td></tr>`r`n",$flows.type)
			$count=1
			foreach($flow in $flows.'import-flow'){

				$MAname = $agentList[($flow.'src-ma')]

				$MA_CS_type = $flow.'cd-object-type'
				$srcattributes=""
				if($flow.'scripted-mapping'.'src-attribute'){
					$rulename = $flow.'scripted-mapping'.'script-context'
					$list=New-Object System.Collections.Generic.HashSet[string]

					$flow.'scripted-mapping'.'src-attribute' | %{
					if($_.'#text') { [void]$list.Add($_.'#text') }
					else { [void]$list.Add($_) }
					}
					$srcattributes = [string]::Join(",",$list)
				} elseif($flow.'direct-mapping'.'src-attribute'){
					$rulename = $flow.'direct-mapping'.'script-context'
					$list=New-Object System.Collections.Generic.HashSet[string]

					$flow.'direct-mapping'.'src-attribute' | %{
					if($_.'#text') { [void]$list.Add($_.'#text') }
					else { [void]$list.Add($_) }
					}
					$srcattributes = [string]::Join(",",$list)
				} elseif($flow.'constant-mapping'.'constant-value'){
					$rulename = "constant"
					$srcattributes = "'"+$flow.'constant-mapping'.'constant-value'+"'"
				}
			[void]$schemaOut.AppendFormat("<tr><td>{3}</td><td><-</td><td><a href='#{0}'>{0}</a>({1})</td><td>{2}</td></tr>`r`n",$MAname,$MA_CS_type,$srcattributes,$count)
			$count++
			}
		}
	}

	if($attribute_exportMA[$aname]){
		[void]$schemaOut.Append("<tr><td>Export</td></tr>`r`n")
		foreach($line in $attribute_exportMA[$aname]){
			[void]$schemaOut.Append($line)
		}
	}

	[void]$schemaOut.Append("</table></br>`r`n")
}

$SROut = New-Object System.Text.StringBuilder
[void]$SROut.Append("<h1>Synchronization rules</h1>")
$ColumnNames = $synchronizationRuleTable.Columns.ColumnName
foreach($row in $synchronizationRuleTable.Rows){
	[void]$SROut.AppendFormat("<table id='{{{0}}}'><tr><td><h1>{1}</h1></td></tr>", $row["object_id"].Tostring().ToUpper(), $row["displayname"])
	foreach($name in $ColumnNames){
		if($name -eq "persistentFlow"){
			if($row[$name] -ne [DBNull]::Value){
				foreach($Value in $row[$name].ToString().Split("|")){
					if($Value -is [string] -AND $Value.StartsWith("&lt;")){
						$Value = "</br>" + (Format-XML-HTML ([Web.HttpUtility]::HtmlDecode($Value)))
					}
					[void]$SROut.AppendFormat("<tr><td>{0}: {1}</td></tr>",$name,$Value)
				}
			}
		}elseif($name -ne "displayname"){
			if($row[$name] -ne [DBNull]::Value){
				$Value = $row[$name]
				if($Value -is [string] -AND $Value.StartsWith("<")){
					$Value = "</br>" + (Format-XML-HTML $Value)
				}elseif($name -eq "connectedSystem"){
					$Value  = "<a href='#{0}'>{0}</a>" -f ($agentList[$Value])
				}
				[void]$SROut.AppendFormat("<tr><td>{0}: {1}</td></tr>",$name,$Value)
			}
		}
	}
	$id = "{{{0}}}" -f $row["object_id"].ToString().ToUpper()
	$Value = $SynchronizationRulesvalue[$id]
	$SynchronizationRulesvalue.Remove($id)
	if($Value.Length -gt 0){
		[void]$SROut.AppendFormat("<tr><td>{0}:</br>{1}</td></tr>","sync-rule-value",(Format-XML-HTML $Value))
	}
	[void]$SROut.Append("</table></br>")
}
#Unkown name rules (portal)
foreach($Key in $SynchronizationRulesvalue.Keys){
	[void]$SROut.AppendFormat("<table id='{0}'><tr><td><h1>{0}</h1></td></tr>", $Key.ToUpper())
	[void]$SROut.AppendFormat("<tr><td>{0}: {1}</td></tr></br></br>","object_id",$Key)
	[void]$SROut.AppendFormat("<tr><td>{0}:</br>{1}</td></tr>","sync-rule-value",(Format-XML-HTML ($SynchronizationRulesvalue[$Key])))
	[void]$SROut.Append("</table></br>")
}



"<html><head><style>table{border: 1px solid black;}, th, td {}</style></head><body>" | Out-File -Encoding UTF8 -FilePath $OutPutFile
$ListMA.ToString() | Out-File -Append -Encoding UTF8 -FilePath $OutPutFile
$schemaOut.ToString() | Out-File -Append -Encoding UTF8 -FilePath $OutPutFile
$MaOut.ToString() | Out-File -Append -Encoding UTF8 -FilePath $OutPutFile
$SROut.ToString() | Out-File -Append -Encoding UTF8 -FilePath $OutPutFile
"</body></html>`r`n" | Out-File -Append -Encoding UTF8 -FilePath $OutPutFile
