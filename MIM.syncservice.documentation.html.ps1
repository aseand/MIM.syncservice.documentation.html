function MIM.syncservice.documentation.html {
	param
	(
		$OutPutFile = (join-path (pwd) ($env:USERDNSDOMAIN+"_attribute_flow.html")),
		[switch]$Debug
	)
	write-progress -id 1 -activity "Create html file" -status "Initialize" -percentComplete 0

	Add-Type -AssemblyName System.Web
	Add-Type -Path "C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\UIShell\PropertySheetBase.dll"
	$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)

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

	function SynchronizationRule-fn($fn){

		if($fn.id -eq "+"){
			$returnstring = "("
		}else{
			$returnstring = "{0}(" -f $fn.id
		}
		
		#All arg
		$sep = ""
		foreach($arg in $fn.arg){

			
			if($arg.fn){
				$returnstring += $sep + (SynchronizationRule-fn $arg.fn)
			}else{
				if($importflow){
					$returnstring +=  $sep + $arg
				}else{
					
					if($SynchronizationRuleAttrList.Contains($arg)){
						
						$returnstring +=  $sep + ("<a href='#{0}'>{0}</a>" -f $arg)
					}else{
						$returnstring +=  $sep + $arg
					}
				}
			}
			if($fn.id -eq "+"){
				$sep = "+"
			}else{
				$sep = ","
			}
		}
		return $returnstring +")"
	}

	function SynchronizationRule-flow-HTML([xml]$xml)
	{ 
		$Out = New-Object System.Text.StringBuilder("&nbsp;&nbsp;&nbsp;&nbsp;")
		#[void]$Out.Append((Format-XML-HTML $xml))
		if($xml.'import-flow'){
			$Global:importflow = $true
			$flow = $xml.'import-flow'
			#[void]$Out.Append("import-flow:</br>&nbsp;&nbsp;&nbsp;&nbsp;")
		}else{
			$Global:importflow = $false
			$flow = $xml.'export-flow'
			#$allow = $flow.'allows-null'
			#[void]$Out.Append("export-flow allows-null "+ $allow+" :</br>&nbsp;&nbsp;&nbsp;&nbsp;")
			if($flow.'allows-null' -eq "true"){
				[void]$Out.Append("(allows-null) ")
			}
		}
		
		#Scoping
		#[void]$Out.Append()
		
		#src list
		$Global:SynchronizationRuleAttrList = New-Object System.Collections.Generic.HashSet[string]
		foreach($attr in $flow.src.attr){
			[void]$SynchronizationRuleAttrList.Add($attr)
			#[void]$Out.AppendFormat("<a href='#{0}'>{0}</a>",$flow.src.attr)
		}
		
		#fn
		if($flow.fn){
			[void]$Out.Append((SynchronizationRule-fn $flow.fn))
		}else{
			if($SynchronizationRuleAttrList.Count -eq 0){
				[void]$Out.AppendFormat("`"{0}`"",$flow.src)
			}else{
				if($importflow){
					foreach($attr in $SynchronizationRuleAttrList){
						[void]$Out.Append($attr)
					}
				}else{
					foreach($attr in $SynchronizationRuleAttrList){
						[void]$Out.AppendFormat("<a href='#{0}'>{0}</a>",$attr)
					}
				}
			}
		}
		#Dir
		#[void]$Out.Append(" &lArr; ")
		[void]$Out.Append(" &rArr; ")

		#dest
		if($importflow){
			[void]$Out.AppendFormat("<a href='#{0}'>{0}</a>",($flow.dest))
		}else{
			[void]$Out.Append($flow.dest)
		}
		
		$Out.ToString()
	}

	$CurentDir = (pwd)

	$OIDTabel = 
	@{
		"1.3.6.1.4.1.1466.115.121.1.12" = "Reference(DN)";
		"1.3.6.1.4.1.1466.115.121.1.15" = "String";
		"1.3.6.1.4.1.1466.115.121.1.5" = "Binary";
		"1.3.6.1.4.1.1466.115.121.1.7" = "Bit";
		"1.3.6.1.4.1.1466.115.121.1.27" = "Integer";
	}

	write-progress -id 1 -activity "Create html file" -status "Loading SynchronizationRules" -percentComplete 5

	#MV SynchronizationRules
	$SynchronizationRules = New-Object hashtable
	$SynchronizationRulesvalue = New-Object hashtable 

	[xml]$synchronizationRuleXml = $MMSWebService.SearchMV("<mv-filter collation-order=`"Latin1_General_CI_AS`"><mv-object-type>synchronizationRule</mv-object-type></mv-filter>")

	if($Debug){
		$synchronizationRuleXml.Save((join-path $CurentDir "synchronizationRule.xml"))
	}

	$TotalStep = $synchronizationRuleXml.'mv-objects'.'mv-object'.entry.Length
	$StegCount = 0
	foreach($entry in $synchronizationRuleXml.'mv-objects'.'mv-object'.entry){
		$index = ($entry.attr.name).IndexOf("displayName")
		write-progress -id 2 -activity "SynchronizationRules" -status ($entry.attr[$index].value.'#text') -percentComplete ($StegCount/$TotalStep*100)
		[void]$SynchronizationRules.Add($entry.dn.ToUpper(),$entry.attr[$index].value.'#text')
		$StegCount++
	}

	write-progress -id 1 -activity "Create html file" -status "Loading MV data" -percentComplete 10

	$MVXMLdata = [xml]$MMSWebService.GetMVData([uint32]::MaxValue)
	$MVdata = $MVXMLdata.'mv-data'

	if($Debug){
		$MVXMLdata.Save((join-path $CurentDir "MV.xml"))
	}

	$maGuid = $null
	$maName = $null
	$MMSWebService.GetMAGuidList([ref] $maGuid,[ref] $maName)

	$attribute_exportMA = New-Object 'system.collections.generic.dictionary[string,System.Collections.Generic.HashSet[string]]'
	$agentList = New-Object 'system.collections.generic.dictionary[string,string]'
	$ListMA = New-Object System.Text.StringBuilder
	$MaOut = New-Object System.Text.StringBuilder

	[void]$MaOut.Append("<h1>Management agent</h1>`r`n")
	[void]$ListMA.AppendFormat("<h1>{0}</h1></br>Generated {1}</br>", ($env:USERDNSDOMAIN),(get-date).ToString("yyy-MM-dd HH:mm:ss"))

	$InstallMIMVersions = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |  ? {$_.DisplayName -like "*Identity Manager*Service*"}|%{@{DisplayName = $_.DisplayName;Version = $_.DisplayVersion}}
	#$InstallMIMVersions = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |  ? {$_.DisplayName -like "*Identity Manager*"}|%{@{DisplayName = $_.DisplayName;Version = $_.DisplayVersion}}
	#Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |  ? {$_.displayName -like "*Identity Manager*"}|%{@{DisplayName = $_.DisplayName,Version = $_.DisplayVersion}}
	foreach($Item in $InstallMIMVersions){
		[void]$ListMA.AppendFormat("{0} - {1} </br>", $Item.DisplayName,$Item.Version)
	}
	[void]$ListMA.Append("</br><h2>Management agent List</h2>")


	write-progress -id 1 -activity "Create html file" -status "Loading MA data" -percentComplete 20
	for($i=0;$i -lt $maGuid.Count;$i++){
		write-progress -id 2 -activity "Management Agent" -status ($maName[$i]) -percentComplete ($i/$maGuid.Count*100)

		$MAXmldata = [xml]$MMSWebService.GetMaData($maGuid[$i],[uint32]::MaxValue,[uint32]::MaxValue,[uint32]::MaxValue)
		$MAdata = $MAXmldata.'ma-data'

		$ma_guid           = $maGuid[$i].ToUpper()
		$ma_name           = $maName[$i]
		$ma_type           = $MAdata.category
		$subtype           = $MAdata.subtype
		$capabilities_mask = $MAdata.'capabilities-mask'
		$ma_export_type    = $madata.'export-type'
		$ma_description    = $madata.description
		
		if($Debug){
			$MAXmldata.Save((join-path $CurentDir "$ma_name.xml"))
		}
		
		[void]$agentList.Add("$ma_guid",$ma_name)
		[void]$ListMA.AppendFormat("<a href='#{0}'>{0}</a></br>",$ma_name)

		[void]$MaOut.AppendFormat("<table id='{0}'><tr><td><h1>{0}</h1></td></tr><tr><td>Type:{1}</td><td>Capabilities:{2}</td><td>export_type:{3}</td><td>Description:{4}</td><td>GUID:{5}</td></tr>`r`n", 
		$ma_name,$ma_type,$capabilities_mask,$ma_export_type,$ma_description,$ma_guid)
		
		$assemblyname = $MAdata.'private-configuration'.MAConfig.'extension-config'.filename.'#text'
		if(-not $assemblyname){$assemblyname=""}

		[void]$MaOut.AppendFormat("<tr><td>extension assembly</td><td>{0}</td></tr>`r`n", $assemblyname)
		[void]$MaOut.Append("<tr></tr>`r`n")
		[void]$MaOut.AppendFormat("<tr><th>Type</th><th>CS</th><th>import/export</th><th>MV</th><th>Rule(s)</th></tr>")
		
		#Join object cs - attibute - mv - extension rule name
		$OldjoinName = ""
		if($MAdata.join)
		{
			foreach($profile in $MAdata.join.'join-profile'){
				$count=1
				foreach($join in $profile.'join-criterion'){
					$mvobject = $join.search.'mv-object-type'
					if($mvobject.Length -lt 1){$mvobject = "Any"}
					$joinName = "{0}{1}" -f $profile.'cd-object-type', $mvobject
					if($joinName -ne $OldjoinName) { [void]$MaOut.AppendFormat("<tr><td>Join Object</td><td>{0}</td><td>&hArr;</td><td>{1}</td></tr>",$profile.'cd-object-type',$mvobject) } 
					$OldjoinName = $joinName
					
					$scriptcontext = ""
					if($join.search.'attribute-mapping'.'direct-mapping' -ne $null){
						if($join.search.'attribute-mapping'.'direct-mapping'.'src-attribute'.'#text'){
							$css = [string]::Join(",",$join.search.'attribute-mapping'.'direct-mapping'.'src-attribute'.'#text')
						}else{
							$css = [string]::Join(",",$join.search.'attribute-mapping'.'direct-mapping'.'src-attribute')
						}
					} else{
						$scriptcontext = $join.search.'attribute-mapping'.'scripted-mapping'.'script-context'
						$css = [string]::Join(",",$join.search.'attribute-mapping'.'scripted-mapping'.'src-attribute')
					}
					[void]$MaOut.AppendFormat("<tr><td>{0}</td><td>{1}</td><td>&hArr;</td><td>",$count,$css)
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
		if($MAdata.projection){
			if($MAdata.projection.'class-mapping'.type){
				if($MAdata.projection.'class-mapping'.type -eq "sync-rule"){
					$SyncRuleID = ($MAdata.projection.'class-mapping'.'sync-rule-id').ToString().ToUpper()
					$SyncRuleName = $SynchronizationRules[$SyncRuleID]
					$Type = "sync-rule(<a href='#{0}'>{1}</a>)" -f $SyncRuleID,$SyncRuleName
				}else{
					$Type = $MAdata.projection.'class-mapping'.type
				}
				[void]$MaOut.AppendFormat("<tr><td>Projektion:{0}</td></tr>`r`n", $Type)
			}
			if($MAdata.projection.'class-mapping'.'cd-object-type'){
				[void]$MaOut.AppendFormat("<tr><td>Projektion:{0}</td></tr>`r`n", $MAdata.projection.'class-mapping'.'cd-object-type')
			}
			if($MAdata.projection.'class-mapping'.'mv-object-type'){
				[void]$MaOut.AppendFormat("<tr><td>Projektion:{0}</td></tr>`r`n", $MAdata.projection.'class-mapping'.'mv-object-type')
			}
			[void]$MaOut.Append("<tr></tr>`r`n")
		}
		
		# DN = dn-construction.attribute
		#if($dn_construction_xml){
			#[void]$MaOut.Append()
			#[void]$MaOut.Append("`r`n")
		#}
		
		#Flow Rules
		$Objectflow = New-Object 'system.collections.generic.dictionary[string,System.Collections.Generic.HashSet[string]]'
		#Import
		foreach($flow in $MVdata.SelectNodes("/mv-data/import-attribute-flow/import-flow-set/import-flows/import-flow[@src-ma='$ma_guid']")){
			$rulename = ""
			
			$flowName = "<tr><td>Flow object</td><td>{0}</td><td>&hArr;</td><td>{1}</td></tr>`r`n" -f $flow.'cd-object-type',$flow.ParentNode.ParentNode.'mv-object-type'
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

			$flowRule = "<tr><td></td><td>{0}</td><td>&rArr;</td><td><a href='#{1}'>{1}</a></td><td>{2}</td></tr>`r`n" -f $srcattributes,$flow.ParentNode.'mv-attribute',$rulename

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
		if($MAdata.'export-attribute-flow'){
			foreach($flowset in $MAdata.'export-attribute-flow'.'export-flow-set'){

				$flowName = "<tr><td>Flow object</td><td>{0}</td><td>&hArr;</td><td>{1}</td></tr>`r`n" -f $flowset.'cd-object-type',$flowset.'mv-object-type'
				
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
					$flowRule = "<tr><td></td><td>{0}</td><td>&lArr;</td><td>{1}</td><td>{2}{3}</td></tr>`r`n" -f $CSatt,$srcattributes,$rulename,$AllowNull
					
					foreach($MvAtt in $attlist){
						$MVstring = "<tr><td></td><td>&rArr;</td><td><a href='#{0}'>{0}</a>({1})</td><td></td></tr>`r`n" -f $ma_name,$CSatt
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
		
		foreach($key in $Objectflow.Keys){
			[void]$MaOut.AppendFormat("<tr><td>{0}</td></tr>",$key)
			[void]$MaOut.Append([string]::Join("",$Objectflow[$key]))
			[void]$MaOut.Append("<tr></tr>`r`n")
		}
		
		
		#Provisionering object cs ? in 
		
		#Deprovisionering object cs
		if($MAdata){
			[void]$MaOut.AppendFormat("<tr><td>Deprovisionering:{0}</td></tr><tr><td>enable-recall:{1}</td></tr>`r`n", $MAdata.'provisioning-cleanup'.action, $MVdata.SelectSingleNode("/import-attribute-flow/per-ma-options/ma-options[@ma-id='{$ma_guid}']").'enable-recall')
		}
		#CS full list
		#CS - object - type?
		#CS - attibutes - type 
		[void]$MaOut.Append("<tr></tr>`r`n")
		if($MAdata.'attribute-inclusion'){
			[void]$MaOut.AppendFormat("<tr><th>CS attribute</th></tr>")
			foreach($attribute in $MAdata.'attribute-inclusion'.'attribute'){
				[void]$MaOut.AppendFormat("<tr><td>{0}</td></tr>",$attribute)
			}
		}
		
		[void]$MaOut.Append("</table></br></br>`r`n")
	}

	$schemaOut = New-Object System.Text.StringBuilder
	[void]$schemaOut.Append("<h1>MV attribute</h1>`r`n")

	write-progress -id 1 -activity "Create html file" -status "Processing MV data" -percentComplete 40
	$TotalStep = $MVdata.schema.dsml.'directory-schema'.'attribute-type'.Length
	$StegCount = 0
	foreach($attribute in $MVdata.schema.dsml.'directory-schema'.'attribute-type'){
		write-progress -id 2 -activity "Processing attribute" -status ($attribute.Name) -percentComplete ($StegCount/$TotalStep*100)
		[void]$schemaOut.AppendFormat("<table id='{0}'><tr><td><h1>{0}</h1></td><td></td></tr><tr><td>mulitvalue:{1}</td><td>indexed:{2}</td><td>syntax:{3}</td></tr>`r`n", $attribute.Name,!($attribute.'single-value'), ([Boolean]$attribute.indexed),$OIDTabel[$attribute.syntax])
		
		$ma=New-Object System.Collections.Generic.HashSet[string]
		$aname = $attribute.Name
		$flowss = $MVdata.SelectNodes("/mv-data/import-attribute-flow/import-flow-set/import-flows[@mv-attribute='$aname']")
		
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
						$list = New-Object System.Collections.Generic.HashSet[string]

						$flow.'scripted-mapping'.'src-attribute' | %{
						if($_.'#text') { [void]$list.Add($_.'#text') }
						else { [void]$list.Add($_) }
						}
						$srcattributes = [string]::Join(",",$list)
					} elseif($flow.'direct-mapping'.'src-attribute'){
						$rulename = $flow.'direct-mapping'.'script-context'
						$list = New-Object System.Collections.Generic.HashSet[string]

						$flow.'direct-mapping'.'src-attribute' | %{
						if($_.'#text') { [void]$list.Add($_.'#text') }
						else { [void]$list.Add($_) }
						}
						$srcattributes = [string]::Join(",",$list)
					} elseif($flow.'constant-mapping'.'constant-value'){
						$rulename = "constant"
						$srcattributes = "'"+$flow.'constant-mapping'.'constant-value'+"'"
					}
				[void]$schemaOut.AppendFormat("<tr><td>{3}</td><td>&lArr;</td><td><a href='#{0}'>{0}</a>({1})</td><td>{2}</td></tr>`r`n",$MAname,$MA_CS_type,$srcattributes,$count)
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
		$StegCount++
	}

	$SROut = New-Object System.Text.StringBuilder
	[void]$SROut.Append("<h1>Synchronization rules</h1>")

	write-progress -id 1 -activity "Create html file" -status "Processing synchronizationRule" -percentComplete 60
	$TotalStep = $synchronizationRuleXml.'mv-objects'.'mv-object'.entry.Length
	$StegCount = 0
	foreach($entry in $synchronizationRuleXml.'mv-objects'.'mv-object'.entry){
		write-progress -id 2 -activity "Processing attribute" -status ($attribute.Name) -percentComplete ($StegCount/$TotalStep*100)
		$index = ($entry.attr.name).IndexOf("displayName")
		[void]$SROut.AppendFormat("<table id='{0}'><tr><td><h1>{1} - {0}</h1></td></tr>", $entry.dn.ToUpper(), ($entry.attr[$index].value.'#text'))
		foreach($attr in $entry.attr){
			if($attr.name -ne "displayname"){
				$AttrValue = ""
				switch($attr.name){
					{ @("persistentFlow", "initialFlow") -contains $_ } {
						
						$initalT = ""
						$importT = ""
						$exportT = ""
						foreach($Value in $attr.Value.'#text'){
							[xml]$ValueXml = $Value
							$temptext = (SynchronizationRule-flow-HTML $ValueXml) + "</br>"
							if($attr.name -eq "initialFlow"){
								$initalT += $temptext
							}else{
								if($ValueXml.'import-flow'){
									$importT += $temptext
								}else{
									$exportT += $temptext
								}
							}
						}
						if($attr.name -eq "initialFlow"){
							$AttrValue += $initalT + "</br>"
						}else{
							if($importT.Length -gt 0){
								$AttrValue = "&nbsp;&nbsp;&nbsp;&nbsp;<b>import-flow</b>:</br>" + $importT + "</br>"
							}
							if($exportT.Length -gt 0){
								$AttrValue += "&nbsp;&nbsp;&nbsp;&nbsp;<b>export-flow</b>:</br>" + $exportT + "</br>"
							}
						}
					}
					"relationshipCriteria" {
						foreach($Value in $attr.Value.'#text'){
							[xml]$relationshipCriteria = $Value
							foreach($condition in $relationshipCriteria.conditions.condition){
								$AttrValue  += "&nbsp;&nbsp;&nbsp;&nbsp;{0} &hArr; {1}</br>" -f ("<a href='#{0}'>{0}</a>" -f $condition.ilmAttribute),($condition.csAttribute)
							}
						}
					}
					"connectedSystemScopec" {
						foreach($Value in $attr.Value.'#text'){						
							[xml]$connectedSystemScope = $Value
							foreach($scope in $connectedSystemScope.scoping.scope){
								$AttrValue  += "&nbsp;&nbsp;&nbsp;&nbsp;{0} {1} {2}</br>" -f ($scope.csAttribute),($scope.csOperator),($csValue)
							}
						}
					}
					{ @("connectedSystemScope", "msidmOutboundScopingFilters") -contains $_ }  {
						foreach($Value in $attr.Value.'#text'){
							[xml]$Scoping = $Value
							foreach($scope in $Scoping.scoping.scope){
								$AttrValue  += "&nbsp;&nbsp;&nbsp;&nbsp;{0} {1} '{2}'</br>" -f ($scope.csAttribute),($scope.csOperator),($scope.csValue)
							}
						}
					}
					"connectedSystem" {
						foreach($Value in $attr.Value.'#text'){
							$AttrValue  += "<a href='#{0}'>{0}</a></br>" -f ($agentList[$Value])
						}
					}
					default {
						foreach($Value in $attr.Value.'#text'){
							if($Value -is [string] -AND $Value.StartsWith("<")){
								$AttrValue += (Format-XML-HTML $Value) + "</br>"
							}else{
								$AttrValue += $Value + "</br>"
							}
						}
					}
				}
				[void]$SROut.AppendFormat("<tr><td><b>{0}</b>:</br>{1}</td></tr>",$attr.name,$AttrValue)
			}
		}
		$id = $entry.dn.ToUpper()
		$Value = $SynchronizationRulesvalue[$id]
		$SynchronizationRulesvalue.Remove($id)
		#if($Value.Length -gt 0){
		#	[void]$SROut.AppendFormat("<tr><td>{0}:</br>{1}</td></tr>","sync-rule-value",(Format-XML-HTML $Value))
		#}
		[void]$SROut.Append("</table></br>")
		$StegCount++
	}

	write-progress -id 1 -activity "Create html file" -status "Processing unkown synchronizationRule" -percentComplete 80
	#Unkown name rules (portal?)
	[void]$SROut.AppendFormat("<h1>Unkown synchronizationRule</h1>")
	foreach($Key in $SynchronizationRulesvalue.Keys){
		[void]$SROut.AppendFormat("<table id='{0}'><tr><td><h1>{0}</h1></td></tr>", $Key.ToUpper())
		[void]$SROut.AppendFormat("<tr><td>{0}: {1}</td></tr></br></br>","object_id",$Key)
		[void]$SROut.AppendFormat("<tr><td>{0}:</br>{1}</td></tr>","sync-rule-value",(Format-XML-HTML ($SynchronizationRulesvalue[$Key])))
		[void]$SROut.Append("</table></br>")
	}

	write-progress -id 1 -activity "Create html file" -status "Write content" -percentComplete 90

	"<html><head><style>table{border: 1px solid black;}, th, td {}</style></head><body>" | Out-File -Encoding UTF8 -FilePath $OutPutFile
	$ListMA.ToString() | Out-File -Append -Encoding UTF8 -FilePath $OutPutFile
	$schemaOut.ToString() | Out-File -Append -Encoding UTF8 -FilePath $OutPutFile
	$MaOut.ToString() | Out-File -Append -Encoding UTF8 -FilePath $OutPutFile
	$SROut.ToString() | Out-File -Append -Encoding UTF8 -FilePath $OutPutFile
	"</body></html>`r`n" | Out-File -Append -Encoding UTF8 -FilePath $OutPutFile

	write-progress -id 1 -activity "Create html file" -status "Done" -percentComplete 100
}

MIM.syncservice.documentation.html
