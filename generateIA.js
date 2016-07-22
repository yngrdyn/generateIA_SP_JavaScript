var spLibraries = ["Documents","Form Templates","List Template Gallery","Master Page Gallery","Site Assets","Site Pages","Solution Gallery","Style Library","Theme Gallery","Web Part Gallery","wfpub","fpdatasources","wfsvc","Converted Forms"]
var spFields = ["Approval Status"];
var spRoles = ["Limited Access"];
var spFolders = ["Forms","_w","_t"];
var spInternalNames = {"Series_x0020_Corporate_x0020_IDB":"Series", "Function_x0020_Corporate_x0020_IDB":"Function", "Project_x0020_Number":"Project Number", "System_x0020_Name":"System Name", "Function Corporate IDB":"Function", "Series Corporate IDB":"Series", "Document_x0020_Author":"Document Author","Function_x0020_Operations_x0020_IDB":"Function","Series_x0020_Operations_x0020_IDB":"Series","Approval_x0020_Number":"Approval Number","Sector_x0020_IDB":"Sector IDB","_dlc_DocIdUrl":"Document ID","Fiscal_x0020_Year_x0020_IDB":"Fiscal Year","TaxKeyword":"Tags","IDBDocs_x0020_Number":"IDBDocs Number","Editor":"Modified By","Division_x0020_or_x0020_Unit":"Division or Unit","Document_x0020_Language_x0020_IDB":"Document Language"};
var spPropertyBags = ["IDBProjectNumber","IDBSiteName","IDBSector","IDBFund","IDBSiteCoordinator","IDBProjectName","IDBOrgUnit","IDBOperationType","IDBOrganizationalFunction","IDBCountry","IDBSiteType","IDBCreatedOn","IDBCreatedby","IDBSiteDescription","IDBSubSector"];
var spViews = ['Merge Documents','Relink Documents','assetLibTemp'];
var ordinals = ["First","Second","Third","Fourth","Fifth","Sixth","Seventh","Eight"];
var ezDefaults = ["Disclosed","IIC"];
var ezPermissions = {"Contribute":2, "Full Control":4, "Read":1, "Design":3, "View Only":1, "Edit":2};
var listDefaults = [];
var listMembers = [];
var specialNames = {};
var structure = {};
var defaultsStructure = {};
var queue = [];
var queueFolders = [];
var queueLibraries = [];
var queueLibrariesDefaults = [];
var queueSecurity = [];
var SubsiteDepth = 1;
var ContentTypeDepth = 1;
var defaultsDepth = 1;
var securityDepth = 1;
var foldersDepth = 0;
var getMetadata = true;
var getStructure = true;
var getSecurity = true;
var getContentTypes = true;
var getDefaults = true;
var getViews = true;
var viewsCount = 0;
var columnsCount = 0;

var tableToExcel = (function () {
	var uri = 'data:application/vnd.ms-excel;base64,'
		, template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>'
		, base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
		, format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
	return function (table, name, wbname) {
		if (!table.nodeType) table = document.getElementById(table)
		var ctx = { worksheet: name || 'Worksheet', table: table.innerHTML }
		var link = document.createElement("A");
		link.href = uri + base64(format(template, ctx));
		link.download = wbname || 'Workbook.xls';
		link.target = '_blank';
		document.body.appendChild(link);
		link.click();
		document.body.removeChild(link);
	}
})();

$(function() {

	$(".ms-rte-layoutszone-outer").hide();
	$("#sideNavBox").hide();
	$("#DeltaTopNavigation").hide();
	$("#searchInputBox").hide();
	$("#pageTitle").find('a').text("Information Architecture Generator");
	$("#pageTitle").css("margin-top","20px");
	$("#contentBox").css("margin-left","20px");
	$("#contentBox").append('<div id="IAinfo" style="margin:0px auto;width:480px;overflow:none;"></div>');
	$("#IAinfo").append('<input type="text" placeholder="Enter the Site url" id="getStructureSite" style="width:300px;padding:6px;height:26px;"/>');
	$("#IAinfo").append('<input type="button" id="getStructure" value="Generate" style="margin-left:-4px;padding:7px 58px;background:#12888a;height:40px;border: none;color: white;font-size: 13px;cursor:pointer;"/>');
	$("#IAinfo").append('<div id="generatorSettings" style="display:none;margin-top:10px;padding-bottom:10px;border-bottom:1px solid #bdbdbd;"><label><input type="checkbox" id="generateStructure" checked>Site Structure</label>&nbsp;<label><input type="checkbox" id="generateSecurity" checked>Security</label>&nbsp;<label><input type="checkbox" id="generateContentTypes" checked>Content Types</label>&nbsp;<label><input type="checkbox" id="generateDefaults" checked>Defaults</label><br><label><input type="checkbox" id="generateMetadata" checked>Site Metadata</label>&nbsp;<label><input type="checkbox" id="generateViews" checked>Library Views</label></div>');
	$("#IAinfo").append('<div id="generatorSettingsLink" style="margin-top:5px;float:right;cursor:pointer;color:#0072c6;padding-right:5px;">Settings</div>');
	$("#IAinfo").append('<div id="progressMessages" style="margin-top:5px;margin-bottom:20px;max-width:420px;overflow-x:hidden"></div>');
	$("#contentBox").append('<div id="InfoOptions" style="display:none;margin-top:110px;margin-left:1px;"><div class="tabInfo active" id="showSiteStructure">Site Structure and Security</div><div class="tabInfo" style="float:left" id="showBagProperties">Site metadata</div><div class="tabInfo" style="float:left;display:none;" id="showViews">Views</div><a id="exportExcel" unselectable="on" class="ms-cui-ctl-medium " mscui:controltype="Button" role="button"><span unselectable="on" class="ms-cui-ctl-iconContainer"><span unselectable="on" class=" ms-cui-img-16by16 ms-cui-img-cont-float"><img unselectable="on" alt="" src="/_layouts/15/1033/images/formatmap16x16.png?rev=43" style="top: -243px; left: -55px;"></span></span><span unselectable="on" class="ms-cui-ctl-mediumlabel">Export to Excel</span></a><div style="clear:both;"></div></div>');
	$("#contentBox").append('<div id="displayResults" style="margin-top:0px;"></div>');
	
	$("#exportExcel").click(function() {
		var context = new SP.ClientContext.get_current();
		var web = context.get_web();
		var user = web.get_currentUser(); //must load this to access info.
		context.load(user);
		context.executeQueryAsync(function(){
			$(".currentRequester").text(" " + user.get_title());
			tableToExcel('siteStructure', 'Site Structure and Security', new Date().toISOString().slice(0, 10) + '_IA' + "_" + structure['Title']);
			if(getMetadata)
				tableToExcel('bagProperties', 'Bag Properties', structure['Title'] + '_Bag Properties');
		}, function(){console.log("Error getting current user");});
	});
	
	$("#generatorSettingsLink").click(function() {
		$("#generatorSettings").toggle();
	});
	
	$("#showSiteStructure").click(function() {
		$("#showSiteStructure").addClass("active");
		$("#showBagProperties").removeClass("active");
		$("#showViews").removeClass("active");
	    $("#bagProperties").hide();
		$(".libraryViews").hide();
		$("#siteStructure").show();
	});
	
	$("#showBagProperties").click(function() {
		$("#showSiteStructure").removeClass("active");
		$("#showViews").removeClass("active");
		$("#showBagProperties").addClass("active");
	    $("#bagProperties").show();
		$("#siteStructure").hide();
		$(".libraryViews").hide();
	});
	
	$("#showViews").click(function() {
		$("#showSiteStructure").removeClass("active");
		$("#showBagProperties").removeClass("active");
		$("#showViews").addClass("active");
		$("#bagProperties").hide();
		$("#siteStructure").hide();
		$(".libraryViews").show();
	});

	$('body').keypress(function (e) {
	  if (e.which == 13) {
		if($("#getStructureSite").val()!="")
			$('#getStructure').click();
		return false;
	  }
	});
	
	$("#generateContentTypes").click( function(){
		if($(this).is(':checked')) {
			$('#generateStructure').prop('checked',true);
		}
	});
	
	$("#generateSecurity").click( function(){
		if($(this).is(':checked')) {
			$('#generateStructure').prop('checked',true);
		}
	});
	
	$("#generateDefaults").click( function(){
		if($(this).is(':checked')) {
			$('#generateStructure').prop('checked',true);
		}
	});
	
	$("#generateStructure").click( function(){
		if(!$(this).is(':checked')) {
			$('#generateSecurity').prop('checked',false);
			$("#generateContentTypes").prop('checked',false);
			$("#generateDefaults").prop('checked',false);
		}
	});
	
	$("#getStructureSite").click(function() {
		$(this).select();
	});
	
	$("#getStructure").click(function() {
		if($("#getStructureSite").val()=="")
			alert("Enter the site URL");
		else{
			$("#InfoOptions").hide();
			$("#displayResults").html("");
			structure = {};
			defaultsStructure = {};
			structure['url'] = $("#getStructureSite").val();
			structure['libraries'] = {};
			queue = [];
			queueLibraries = [];
			queueFolders = [];
			specialNames = {};
			queueLibrariesDefaults = [];
			queueSecurity = [];
			listDefaults = [];
			listMembers = [];
			SubsiteDepth = 1;
			ContentTypeDepth = 1;
			defaultsDepth = 1;
			securityDepth = 1;
			foldersDepth = 0;
			getMetadata = true;
			getStructure = true;
			getSecurity = true;
			getContentTypes = true;
			getDefaults = true;
			getViews = true;
			
			if(!$('#generateMetadata').is(':checked'))
				getMetadata = false;
				
			if(!$('#generateStructure').is(':checked'))
				getStructure = false;
				
			if(!$('#generateSecurity').is(':checked'))
				getSecurity = false;
				
			if(!$('#generateContentTypes').is(':checked'))
				getContentTypes = false;
			
			if(!$("#generateDefaults").is(':checked'))
				getDefaults = false;
			
			if(!$('#generateStructure').is(':checked'))
				getStructure = false;
				
			if(!$('#generateViews').is(':checked'))
				getViews = false;
			
			queue.push({'url':structure['url'],'structure':structure['libraries']});
			$("#getStructureSite").prop('disabled', true);
			$("#getStructure").prop('disabled', true);
			$("#displayResults").append("<div style='width:390px;margin:200px auto;'><img src='/teams/ITE/Office365/eZShare/SiteAssets/hex-loader2.gif'></div>");
			//$("#gettingStructure").attr('src',"/teams/ITE/Office365/eZShare/SiteAssets/21.gif");
			SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function(){
				getAllWebs(
				function(allwebs){
					structure['subsites'] = {};
					for (var key in allwebs){
						var levels = allwebs[key]['structure'].split(";");
						if(SubsiteDepth<Math.ceil(levels.length/2))
							SubsiteDepth = Math.ceil(levels.length/2);
						var aux = structure;
						for (var j = 1; j < (levels.length); j++){
							if(j % 2 != 0){
								if('subsites' in aux){
									aux = aux["subsites"];
								}else{
									aux['subsites'] = {};
									aux = aux['subsites'];
								}	
							}else{
								if(levels[j] in aux){
									aux = aux[levels[j]];
								}else{
									aux[levels[j]] = {};
									aux = aux[levels[j]];
								}
							}
						}
						aux['url'] = allwebs[key]['url'];
						if(getStructure){
							aux['libraries'] = {};
							if(getSecurity){
								aux['Permissions'] = {};
								for(member in allwebs[key]["Permissions"]){
									aux['Permissions'][member] = allwebs[key]["Permissions"][member];
								}
							}
							
							queue.push({'url':allwebs[key]['url'],'structure':aux['libraries']});
						}
						
						if(getMetadata){
							aux['bagProperties'] = {};
							for(var prop in allwebs[key]['bagProperties']){
								aux['bagProperties'][prop] = allwebs[key]['bagProperties'][prop];
							}
						}
						
					}
					if(getStructure || getViews){
						$("#progressMessages").html("").append("<img src='/teams/ITE/Office365/eZShare/SiteAssets/loading.gif'/>&nbsp;Retrieving Libraries");
						retrieveAllLibraries();
						
					}else{
						retrieveRootProperties();
					}
					
				},
				function(sendera,args){
					alert(args.get_message());
					console.log("root failed");
					$("#getStructure").prop('disabled', false);
					$("#getStructureSite").prop('disabled', false);
					$("#displayResults").html("");
					$("#progressMessages").html("");
				});
				
			});
		}	
	});
	
	function getAllWebs(success,error){
	   $("#progressMessages").html("").append("<img src='/teams/ITE/Office365/eZShare/SiteAssets/loading.gif'/>&nbsp;Retrieving subsites");
	   var ctx = new SP.ClientContext(structure['url']);
	   var web = ctx.get_web();
	   var result = {};
	   var level = 0;
	   var initial = "root";
	   var getAllWebsInner = function(web,result,success,error,initial) 
	   {
		  level++;
		  var ctx = web.get_context();
		  var webs = web.getSubwebsForCurrentUser(null); 
		  ctx.load(webs,'Include(Title,Webs,Url,AllProperties,RoleAssignments.Include(Member,RoleDefinitionBindings))');
		  //ctx.load(webs, site => site.Include(Title,Webs,Url,AllProperties));
		  ctx.executeQueryAsync(
			function(){
				for(var i = 0; i < webs.get_count();i++){
					var web = webs.getItemAtIndex(i);
					
					// Changing the Message 
					$("#progressMessages").html("").append("<img src='/teams/ITE/Office365/eZShare/SiteAssets/loading.gif'/>&nbsp;" + web.get_title() + " - found");
					
					result[web.get_title()] = {};
					result[web.get_title()]['web'] = web;
					result[web.get_title()]['structure'] = initial + ";#S#;" +web.get_title();
					result[web.get_title()]['url'] = web.get_url();
					
					if(getMetadata){
						var allProps = web.get_allProperties().get_fieldValues();
						result[web.get_title()]['bagProperties'] = {}
						for (var prop in allProps) {
							if(prop.startsWith('IDB')){
								result[web.get_title()]['bagProperties'][prop] = allProps[prop];
							}
						}
					}
					
					if(getSecurity){
						var permissionsEnumerator = web.get_roleAssignments().getEnumerator();
			
						var permissions = {};
						while (permissionsEnumerator.moveNext()) {
							var permission = permissionsEnumerator.get_current();
							var rolesEnumerator = permission.get_roleDefinitionBindings().getEnumerator();
							while (rolesEnumerator.moveNext()) {
								var role = rolesEnumerator.get_current();
								if(spRoles.indexOf(role.get_name())==-1){
									permissions[permission.get_member().get_title()] = role.get_name();
									if($.inArray(permission.get_member().get_title(), listMembers) == -1)
										listMembers.push(permission.get_member().get_title());
								}
							}
						}	
						
						
						result[web.get_title()]['Permissions'] = {}
						for(var permission in permissions){
							result[web.get_title()]['Permissions'][permission] = permissions[permission];
						}
					}
					
					if(web.get_webs().get_count() > 0) {
					   getAllWebsInner(web,result,success,error,initial + ";#S#;" +web.get_title());
					} 
					
				}
				
				level--;
				if (level == 0 && success){
					success(result);  
				}
			},
			error);
	   };

	   getAllWebsInner(web,result,success,error);    
	}
	
	function retrieveAllLibraries() {
		if(queue.length > 0){		
			
			var actualSite = queue.pop();
			var ctx = new SP.ClientContext(actualSite['url']);
			web = ctx.get_web();
			lists = web.get_lists();
			
			//ctx.load(lists,'Include(Title,BaseType,Fields.Include(Title,InternalName,DefaultValue,TypeAsString),ContentTypes.Include(Name),RootFolder,RoleAssignments.Include(Member,RoleDefinitionBindings))');
			ctx.load(lists, 'Include(Title,BaseType,ContentTypes.Include(Name), Fields.Include(Title,InternalName,DefaultValue,TypeAsString), Views.Include(Title,ViewFields,DefaultView,Aggregations,RowLimit,ViewQuery))');
			ctx.executeQueryAsync(function(){onQuerySucceeded(actualSite['structure'],actualSite['url'])}, onQueryFailed);
		}else{
			if(getStructure){
				/*
				for(var i=0;i<queueLibraries.length;i++){
					console.log(queueLibraries[i]);
				}
				*/
				retrieveAllLibrariesInfo();
			}else{
				retrieveRootProperties();
			}
			
		}

	}

	function onQuerySucceeded(librariesStructure, url, sender, args) {
		var listEnumerator = lists.getEnumerator();

		while (listEnumerator.moveNext()) {
			var currentContentTypesDepth = 0;
			var oList = listEnumerator.get_current();
			if(oList.get_baseType() == 1){
				if(spLibraries.indexOf(oList.get_title())==-1){

					// Creating the Library in the general structure
					librariesStructure[oList.get_title()] = {}
					
					// Getting the Views
					librariesStructure[oList.get_title()]['views'] = {};
					var viewsEnumerator = oList.get_views().getEnumerator();
					var count = 0;
					
					while (viewsEnumerator.moveNext()){
						var view = viewsEnumerator.get_current();
						if(spViews.indexOf(view.get_title())==-1){
							librariesStructure[oList.get_title()]['views'][view.get_title()] = {};
							librariesStructure[oList.get_title()]['views'][view.get_title()]["default"] = view.get_defaultView();
							
							var totals = "";
							var xmlDoc = new DOMParser().parseFromString(view.get_aggregations(), 'text/xml');
							var items = xmlDoc.getElementsByTagName('FieldRef');
							for(var i = 0; i < items.length; i++) {
								if(totals != ""){
									totals += ";";
								}
								if(items[i].getAttribute('Name') in spInternalNames){
									totals += spInternalNames[items[i].getAttribute('Name')] + " - " + items[i].getAttribute('Type');
								}else{
									totals += items[i].getAttribute('Name') + " - " + items[i].getAttribute('Type');
								}
							}
							librariesStructure[oList.get_title()]['views'][view.get_title()]["totals"] = totals;
							
							librariesStructure[oList.get_title()]['views'][view.get_title()]["limit"] = view.get_rowLimit();
							
							var sort = "";
							xmlDoc = new DOMParser().parseFromString("<document>" + view.get_viewQuery() + "</document>", 'text/xml');
							items = xmlDoc.getElementsByTagName('OrderBy');
							var order = "";
							for(var i = 0; i < items.length; i++) {
								var childitems = items[i].childNodes;
								for(var j=0; j < childitems.length; j++){
									if(sort != ""){
										sort += ";";
									}
									if(childitems[j].getAttribute('Ascending')=="TRUE")
										order = 'Ascending';
									else
										order = 'Descending';
									if(childitems[j].getAttribute('Name') in spInternalNames){
										sort += spInternalNames[childitems[j].getAttribute('Name')] + " - " + order;
									}else{
										sort += childitems[j].getAttribute('Name') + " - " + order;
									}
								}
							}
							librariesStructure[oList.get_title()]['views'][view.get_title()]["sort"] = sort;
							
							items = xmlDoc.getElementsByTagName('GroupBy');
							var group = "";
							for(var i = 0; i < items.length; i++) {
								var childitems = items[i].childNodes;
								for(var j=0; j < childitems.length; j++){
									if(group != ""){
										group += ";";
									}
									if(childitems[j].getAttribute('Ascending')=="TRUE")
										order = 'Ascending';
									else
										order = 'Descending';
									if(childitems[j].getAttribute('Name') in spInternalNames){
										group += "Column '" + spInternalNames[childitems[j].getAttribute('Name')] + "' - show groups in " + order + " order";
									}else{
										group += "Column '" + childitems[j].getAttribute('Name') + "' - show groups in " + order + " order";
									}
								}
							}
							librariesStructure[oList.get_title()]['views'][view.get_title()]["group"] = group;
							
							items = xmlDoc.getElementsByTagName('Where');
							var filter = "";
							var operator = "";
							for(var i = 0; i < items.length; i++){
								if(filter != ""){
									filter += ";";
								}
								var childitems = items[i].childNodes;
								for(var j=0; j < childitems.length; j++){
									if(childitems[j].nodeName=="Or" || childitems[j].nodeName=="And"){
										if(childitems[j].childNodes[0].nodeName=="Eq"){
											operator = "=";
											filter += childitems[j].childNodes[0].childNodes[0].getAttribute('Name') + " " + operator + " ";
											if(childitems[j].childNodes[0].childNodes[1].getAttribute('Type')=="Boolean")
												if(childitems[j].childNodes[0].childNodes[1].childNodes[0].nodeValue == "1")
													filter += "True";
												else
													filter += "False";
											else
												filter += childitems[j].childNodes[0].childNodes[1].childNodes[0].nodeValue;
										}
										filter += " " + childitems[j].nodeName + " ";
										if(childitems[j].childNodes[1].nodeName=="Eq"){
											operator = "=";
											filter += childitems[j].childNodes[1].childNodes[0].getAttribute('Name') + " " + operator + " ";
											if(childitems[j].childNodes[1].childNodes[1].getAttribute('Type')=="Boolean")
												if(childitems[j].childNodes[1].childNodes[1].childNodes[0].nodeValue == "1")
													filter += "True";
												else
													filter += "False";
											else
												filter += childitems[j].childNodes[1].childNodes[1].childNodes[0].nodeValue;
										}
									}else{ 
										if(childitems[j].nodeName=="Eq"){
											if(childitems[j].childNodes[0].getAttribute('Name')=="Author"){
												if(childitems[j].childNodes[1].childNodes[0].nodeName=="UserID"){
													filter += "'Created By' is equal to [Me]";
												}
											}
										}
									}
								}
							}
							librariesStructure[oList.get_title()]['views'][view.get_title()]["filter"] = filter;
							
							var fieldsEnumerator = view.get_viewFields().getEnumerator();
						
							var fields = [];
							while (fieldsEnumerator.moveNext()){
								var field = fieldsEnumerator.get_current();
								if(field in spInternalNames)
									fields.push(spInternalNames[field]);
								else
									fields.push(field);
							}
							
							librariesStructure[oList.get_title()]['views'][view.get_title()]["columns"] = fields;
							
							if(columnsCount < fields.length)
								columnsCount = fields.length;
							
							count++;
						}
					}
					
					if(viewsCount < count){
						viewsCount = count;
					}
					
					if(getStructure){
						
						// Pushing the library to get the Defaults and Permissions later
						queueLibraries.push({'name':oList.get_title(),'structure':librariesStructure[oList.get_title()],"url":url});
						queueLibrariesDefaults.push({'name':oList.get_title(),'structure':librariesStructure[oList.get_title()],"url":url});
						
						if(getContentTypes){
							// Changing the Message 
							$("#progressMessages").html("").append("<img src='/teams/ITE/Office365/eZShare/SiteAssets/loading.gif'/>&nbsp;" + oList.get_title() + " - Retrieving Content Types");
						
							// Getting the content types of the library
							var contentTypes = [];
							var contentTypesEnumerator = (oList.get_contentTypes()).getEnumerator();

							while (contentTypesEnumerator.moveNext()) {
								var contentType = contentTypesEnumerator.get_current();
								if(contentType.get_name()!="Folder"){
									currentContentTypesDepth++;
									contentTypes.push(contentType.get_name());
								}
							}
							
							if(currentContentTypesDepth>ContentTypeDepth)
								ContentTypeDepth = currentContentTypesDepth;
							
							// Setting the Content types to the library in the general structure
							librariesStructure[oList.get_title()]['contentTypes'] = contentTypes;
						}
						
						if(getDefaults){
							// Getting the default values of the library
							var defaults = {};
							var fieldEnumerator = (oList.get_fields()).getEnumerator();

							while (fieldEnumerator.moveNext()) {
								var field = fieldEnumerator.get_current();
								if(spFields.indexOf(field.get_title())==-1){
									if(field.get_defaultValue()!=null & field.get_defaultValue()!=""){
										if($.inArray(field.get_title(), ezDefaults) == -1){
											if(field.get_typeAsString() == "Boolean"){
												if(field.get_defaultValue()==0)
													defaults[field.get_title()] = false;
												else
													defaults[field.get_title()] = true;	
											}else{
												var value = field.get_defaultValue();
												if(value.indexOf(";#")!=-1){
													value = value.split(";#")[1].split("|")[0];
												}
												defaults[field.get_title()] = value;
											}
										}
										if(field.get_title() in spInternalNames){
											if($.inArray(spInternalNames[field.get_title()], ezDefaults) == -1){
												if($.inArray(spInternalNames[field.get_title()], listDefaults) == -1){
													listDefaults.push(spInternalNames[field.get_title()]);
												}
											}
										}else{
											if($.inArray(field.get_title(), ezDefaults) == -1){
												if($.inArray(field.get_title(), listDefaults) == -1){
													listDefaults.push(field.get_title());
												}
											}
										}	
									}
								}
							}
							
							// Setting the Default Values to the library in the general structure
							librariesStructure[oList.get_title()]['defaults'] = defaults;
						}
					}
					
				}
				
			}
		}
		retrieveAllLibraries();
	}
		
	function onQueryFailed(sender, args) {
		/*alert('Request failed. ' + args.get_message() + 
			'\n' + args.get_stackTrace());*/
		console.log("error retrieveAllLibraries");
		retrieveAllLibraries();
	}
	
	function retrieveAllLibrariesInfo(){
		if(queueLibraries.length > 0){		
			
			var actualLibrary = queueLibraries.pop();
			var ctx = new SP.ClientContext(actualLibrary['url']);
			web = ctx.get_web();
			list = web.get_lists().getByTitle(actualLibrary['name']);

			//ctx.load(list,'RootFolder.Folders.Include(Name,ServerRelativeUrl,Folders,ListItemAllFields.RoleAssignments.Include(Member,RoleDefinitionBindings))');
			ctx.load(list,'RootFolder.Folders.Include(Name,ServerRelativeUrl,Folders)');
			ctx.load(list,'RootFolder.ServerRelativeUrl');
			//ctx.load(list,'RoleAssignments.Include(Member,RoleDefinitionBindings)');
			ctx.executeQueryAsync(function(){onQuerySucceededRetrieveAllLibrariesInfo(actualLibrary['structure'],actualLibrary['url'],actualLibrary['name'])}, onQueryFailedRetrieveAllLibrariesInfo);
		}else{
			retrieveFolders();
			//retrieveAllDefaults();
		}
	}
	
	function onQuerySucceededRetrieveAllLibrariesInfo(librariesStructure, url, name, sender, args) {
		
		$("#progressMessages").html("").append("<img src='/teams/ITE/Office365/eZShare/SiteAssets/loading.gif'/>&nbsp;" + name + " - Retrieving Folders");
		
		// Getting Folders
		var foldersEnumerator = list.get_rootFolder().get_folders().getEnumerator();
	
		var Folders = {};
		librariesStructure['Folders'] = {};
		librariesStructure['Permissions'] = {};
		
		queueSecurity.push({'structure':librariesStructure['Permissions'],'url':url,'name':name,'type':'Library'});
		
		while (foldersEnumerator.moveNext()) {
			if(foldersDepth < 1)
				foldersDepth = 1;
			var folder = foldersEnumerator.get_current();
			if(spFolders.indexOf(folder.get_name())==-1){
				librariesStructure['Folders'][folder.get_name()] = {};
				librariesStructure['Folders'][folder.get_name()]['url'] = location.protocol + '//' + location.hostname + folder.get_serverRelativeUrl();
				librariesStructure['Folders'][folder.get_name()]['defaults'] = {};
				librariesStructure['Folders'][folder.get_name()]['Permissions'] = {};
				
				defaultsStructure[librariesStructure['Folders'][folder.get_name()]['url']] = librariesStructure['Folders'][folder.get_name()]['defaults'];
				queueSecurity.push({'structure':librariesStructure['Folders'][folder.get_name()]['Permissions'],'url':url,'name':librariesStructure['Folders'][folder.get_name()]['url'],'type':'Folder'});
				
				var subFoldersEnumerator = folder.get_folders().getEnumerator();
				
				librariesStructure['Folders'][folder.get_name()]['Folders'] = {};
				
				while (subFoldersEnumerator.moveNext()) {
					var subFolder = subFoldersEnumerator.get_current();
					librariesStructure['Folders'][folder.get_name()]['Folders'][subFolder.get_name()] = {};
					queueFolders.push({'name':subFolder.get_name(),'structure': librariesStructure['Folders'][folder.get_name()]['Folders'][subFolder.get_name()],"url":url,"parent":name + "/" + folder.get_name(),"level":2});
				}
				
			}
		}
		
		librariesStructure['url'] = location.protocol + '//' + location.hostname + list.get_rootFolder().get_serverRelativeUrl();
		
		if(!("defaults" in librariesStructure)){
			librariesStructure['defaults'] = {};
		}
		defaultsStructure[librariesStructure['url']] = librariesStructure['defaults'];

		retrieveAllLibrariesInfo();
	}
		
	function onQueryFailedRetrieveAllLibrariesInfo(sender, args) {
		/*
		alert('Request failed. ' + args.get_message() + 
			'\n' + args.get_stackTrace());
		*/
		console.log("error onQueryFailedRetrieveAllLibrariesInfo");
		retrieveAllLibrariesInfo();
	}
	
	function retrieveFolders(){
		if(queueFolders.length > 0){		
			
			var actualFolder = queueFolders.pop();
			
			var ctx = new SP.ClientContext(actualFolder['url']);
			web = ctx.get_web();
			folder = web.getFolderByServerRelativeUrl(encodeURI(actualFolder['url'] + '/' + actualFolder['parent'] + '/' + actualFolder['name']));
			
			ctx.load(folder,'Name','ServerRelativeUrl','Folders');
			
			ctx.executeQueryAsync(function(){onQuerySucceededRetrieveFolders(actualFolder['structure'],actualFolder['url'],actualFolder['name'],actualFolder['parent'] + '/' + actualFolder['name'],actualFolder['level'])}, onQueryFailedRetrieveFolders);
			
		}else{
			retrieveAllDefaults();
		}
	}
	
	function onQuerySucceededRetrieveFolders(librariesStructure, url, name, parent, level, sender, args) {
		if(foldersDepth < level)
			foldersDepth = level;
		
		var subFoldersEnumerator = folder.get_folders().getEnumerator();
		
		librariesStructure['url'] = location.protocol + '//' + location.hostname + folder.get_serverRelativeUrl();
		librariesStructure['Folders'] = {};
		librariesStructure['defaults'] = {};
		librariesStructure['Permissions'] = {};
		
		defaultsStructure[librariesStructure['url']] = librariesStructure['defaults'];
		queueSecurity.push({'structure':librariesStructure['Permissions'],'url':url,'name':librariesStructure['url'],'type':'Folder'});
		
		while (subFoldersEnumerator.moveNext()) {
			var subFolder = subFoldersEnumerator.get_current();
			librariesStructure['Folders'][subFolder.get_name()] = {};
			queueFolders.push({'name':subFolder.get_name(),'structure': librariesStructure['Folders'][subFolder.get_name()],"url":url,"parent":parent,"level": level + 1});
		}
		
		retrieveFolders();
	}
		
	function onQueryFailedRetrieveFolders(sender, args) {
		alert('Request failed. ' + args.get_message() + 
			'\n' + args.get_stackTrace());
		console.log("error onQueryFailedRetrieveAllFolders");
		retrieveFolders();
	}
	
	function retrieveAllDefaults(){	
		
		if(getDefaults){
			var actualLibrary = queueLibrariesDefaults.pop();
			//console.log(actualLibrary);
			
			if(typeof actualLibrary === "undefined"){
				if(queueLibrariesDefaults.length > 0){
					retrieveAllDefaults();
				}else{
					retrieveRootProperties();
				}
			
			}else{
				$("#progressMessages").html("").append("<img src='/teams/ITE/Office365/eZShare/SiteAssets/loading.gif'/>&nbsp;" + actualLibrary['name'] + " - Retrieving Default Values");
				
				var auxName = actualLibrary['name'].replace(/,/g,"").replace(/\(/g,"").replace(/\)/g,"").replace(/-/g,"");
				
				specialNames[auxName] = actualLibrary['name'];
				
				$.ajax({
					url: actualLibrary['url'] + '/' + encodeURI(auxName) + '/Forms/client_LocationBasedDefaults.html',
					retry: 3,
					dataType: 'xml',
					success: function (data){
						$(data).find("a").each(function(){							
							var urlAux = location.protocol + '//' + location.hostname + decodeURIComponent($(this).attr("href"));
							
							if(urlAux in defaultsStructure){
								$(this).find("DefaultValue").each(function(){
									var value = $(this).text();
									if(value.indexOf(";#")!=-1){
										value = value.split(";#")[1].split("|")[0];
									}
									if($(this).attr("FieldName") in spInternalNames){
										if($.inArray(spInternalNames[$(this).attr("FieldName")], ezDefaults) == -1){
											defaultsStructure[urlAux][spInternalNames[$(this).attr("FieldName")]] = value;
											if($.inArray(spInternalNames[$(this).attr("FieldName")], listDefaults) == -1){
												listDefaults.push(spInternalNames[$(this).attr("FieldName")]);
											}
										}
									}else{
										if($.inArray($(this).attr("FieldName"), ezDefaults) == -1){
											defaultsStructure[urlAux][$(this).attr("FieldName")] = value;
											if($.inArray($(this).attr("FieldName"), listDefaults) == -1){
												listDefaults.push($(this).attr("FieldName"));
											}
										}
									}
								});
							}
						});
						if(queueLibrariesDefaults.length > 0){
							retrieveAllDefaults();
						}else{
							getSecurityElement();
						}
					},
					error: function(xhr, errorThrown ) {
							this.retry = this.retry - 1;
							if(this.retry > 0){
								if(this.retry > 1){
									this.url = actualLibrary['url'] + '/' + encodeURI(actualLibrary['name']) + '/Forms/client_LocationBasedDefaults.html';
									$.ajax(this);
								}else{
									var secondAux = auxName.replace(/\./g,"");
									//console.log(secondAux);
									this.url = actualLibrary['url'] + '/' + encodeURI(secondAux) + '/Forms/client_LocationBasedDefaults.html';
									$.ajax(this);
								}
							}
							else
								retrieveAllDefaults();
							return;      
					}
				});
			}
		}else{
			getSecurityElement();
		}
	}
	
	function getSecurityElement(){
		if(getSecurity){
			if(queueSecurity.length > 0){
				var actualLibrary = queueSecurity.pop();
				var ctx = new SP.ClientContext(actualLibrary['url']);
				web = ctx.get_web();
				if(actualLibrary['type']=="Library"){
					list = web.get_lists().getByTitle(actualLibrary['name']);
					ctx.load(list,'RoleAssignments.Include(Member,RoleDefinitionBindings)');
				}else{
					list = web.getFolderByServerRelativeUrl(actualLibrary['name']);
					ctx.load(list,'ListItemAllFields.RoleAssignments.Include(Member,RoleDefinitionBindings)');
				}
				
				ctx.executeQueryAsync(function(){onQuerySucceededGetSecurity(actualLibrary['structure'],actualLibrary['url'],actualLibrary['name'],actualLibrary['type'])}, onQueryFailedGetSecurity);
			}else{
				retrieveRootProperties();
			}
		}else{
			retrieveRootProperties();
		}
	}
	
	function onQuerySucceededGetSecurity(librariesStructure, url, name, type, sender, args) {
		if(getSecurity){
			$("#progressMessages").html("").append("<img src='/teams/ITE/Office365/eZShare/SiteAssets/loading.gif'/>&nbsp;" + name + " - Retrieving Permissions");
			
			// Getting the permissions
			if(type=="Library")
				var permissionsEnumerator = list.get_roleAssignments().getEnumerator();
			else
				var permissionsEnumerator = list.get_listItemAllFields().get_roleAssignments().getEnumerator();
			
			while (permissionsEnumerator.moveNext()) {
				var permission = permissionsEnumerator.get_current();
				var rolesEnumerator = permission.get_roleDefinitionBindings().getEnumerator();
				while (rolesEnumerator.moveNext()) {
					var role = rolesEnumerator.get_current();
					if(spRoles.indexOf(role.get_name())==-1){
						librariesStructure[permission.get_member().get_title()] = role.get_name();
						if($.inArray(permission.get_member().get_title(), listMembers) == -1)
							listMembers.push(permission.get_member().get_title());
					}
				}
			}
		}
		
		getSecurityElement();
	}
		
	function onQueryFailedGetSecurity(sender, args) {
		/*
		alert('Request failed. ' + args.get_message() + 
			'\n' + args.get_stackTrace());
		*/
		console.log("error onQueryFailedGetSecurity: " + args.get_message() + '\n' + args.get_stackTrace());
		getSecurityElement();
	}
	
	function retrieveRootProperties(){
		var ctx = new SP.ClientContext(structure['url']);
		web = ctx.get_web();
		props =  web.get_allProperties();
		webPermissions = web.get_roleAssignments();
		
		ctx.load(web,'Title');
		if(getMetadata)
			ctx.load(props); 
		ctx.load(webPermissions,'Include(Member,RoleDefinitionBindings)'); 
		ctx.executeQueryAsync(onQuerySucceededRetrieveRootProperties, onQueryFailedRetrieveRootProperties);
	}
	
	function onQuerySucceededRetrieveRootProperties(){
	
		$("#progressMessages").html("").append("<img src='/teams/ITE/Office365/eZShare/SiteAssets/loading.gif'/>&nbsp; Retrieving information of the Site Collection");

		structure['Title'] = web.get_title();
		if(getMetadata){
			var allProps = props.get_fieldValues();
			structure['bagProperties'] = {};
			for (var prop in allProps) {
				if(prop.startsWith('IDB')){
					structure['bagProperties'][prop] = allProps[prop];
				}
			}
		}
		
		if(getSecurity){
			var permissionsEnumerator = webPermissions.getEnumerator();
			
			var permissions = {};
			while (permissionsEnumerator.moveNext()) {
				var permission = permissionsEnumerator.get_current();
				var rolesEnumerator = permission.get_roleDefinitionBindings().getEnumerator();
				while (rolesEnumerator.moveNext()) {
					var role = rolesEnumerator.get_current();
					if(spRoles.indexOf(role.get_name())==-1){
						permissions[permission.get_member().get_title()] = role.get_name();
						if($.inArray(permission.get_member().get_title(), listMembers) == -1)
							listMembers.push(permission.get_member().get_title());
					}
				}
			}
		}
		
		defaultsDepth = listDefaults.length;
		securityDepth = listMembers.length;
		
		console.log(structure);
		console.log('finished');
		
		createTableHtml();
		
	}
	
	function onQueryFailedRetrieveRootProperties(sender, args){
		/*alert('Request failed. ' + args.get_message() + 
			'\n' + args.get_stackTrace());*/
		console.log('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
		createTableHtml();
	}
	
	function renderSecurityHeader(renderGroups){
		var table = "";
		if(getSecurity){
			if(renderGroups){
				if(listMembers.length>0){
					for(var j=0; j< listMembers.length;j++){
						if(j==listMembers.length-1)
							table += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;' align='center' class='rotate'><div><span>" + listMembers[j] + "</span></div></td>";
						else
							table += "<td style='background:#0B3861;color:#fff;border:0;' align='center' class='rotate'><div><span>" + listMembers[j] + "</span></div></td>";
					}
				}else{
					table += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;'></td>";
				}
			}else{
				for(var i =0; i < securityDepth;i++){
					if(i==securityDepth-1)
						table += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;'></td>";
					else
						table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
				}
			}
		}
		return table;
	}
	
	function renderSecurityBody(structureObject){
		var table = "";
		if(getSecurity){
			if(structureObject != null){
				if("Permissions" in structureObject){
					if(Object.keys(structureObject["Permissions"]).length > 0){
						for(var j=0;j<listMembers.length;j++){
							if(j==listMembers.length-1){
								if(listMembers[j] in structureObject["Permissions"])
									table += "<td style='color:#1c1c1c;max-width:150px;border-right: 2px solid #0B3861;' align='center'>" + ezPermissions[structureObject["Permissions"][listMembers[j]]] + "</td>";
								else
									table += "<td style='color:#1c1c1c;background:#E6E6E6;border-right: 2px solid #0B3861;' align='center'>0</td>";
							}else{
								if(listMembers[j] in structureObject["Permissions"]){
									table += "<td style='color:#1c1c1c;max-width:150px;' align='center'>" + ezPermissions[structureObject["Permissions"][listMembers[j]]] + "</td>";
								}
								else{
									table += "<td style='color:#1c1c1c;background:#E6E6E6;' align='center'>0</td>";
								}
							}
						}	
					}else{
						table += "<td style='color:#1c1c1c;background:#F6CECE;border-right: 2px solid #0B3861;' align='center' colspan='" + listMembers.length + "'>NO ACCESS</td>";
					}
				}else{
					table += "<td style='color:#1c1c1c;background:#F6CECE;border-right: 2px solid #0B3861;' align='center' colspan='" + listMembers.length + "'> NO ACCESS </td>";
				}
			}else{
				for(var i =0; i < securityDepth;i++){
					if(i==securityDepth-1)
						table += "<td style='color:#1c1c1c;background:#BDBDBD;border-right: 2px solid #0B3861;'></td>";
					else
						table += "<td style='color:#1c1c1c;background:#BDBDBD;'></td>";
				}
			}
		}
		return table;
	}
	
	function renderContentTypesHeader(renderOrder){
		var table = "";
		if(getContentTypes){
			if(renderOrder){
				for(var i =0; i < ContentTypeDepth;i++){
					if(i==0)
						table += "<td style='background:#0B3861;color:#fff;border:0;' align='center'>Default</td>";
					else
						if(i==ContentTypeDepth-1)
							table += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;' align='center'>" + ordinals[i] + "</td>";
						else
							table += "<td style='background:#0B3861;color:#fff;border:0;' align='center'>" + ordinals[i] + "</td>"
				}
			}else{
				for(var i =0; i < ContentTypeDepth;i++){
					if(i==ContentTypeDepth-1)
						table += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;'></td>";
					else
						table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
				}
			}
		}
		return table;
	}
	
	function renderContentTypesBody(structureObject){
		var table = "";
		if(getContentTypes){
			if(structureObject != null){
				if("contentTypes" in structureObject){
					for(var i=0;i<structureObject["contentTypes"].length;i++){
						if(i==structureObject["contentTypes"].length-1 & structureObject["contentTypes"].length==ContentTypeDepth)
								table += "<td style='color:#1c1c1c;border-right: 2px solid #0B3861;' padding='5'>" + structureObject["contentTypes"][i] + "</td>";
							else
								table += "<td style='color:#1c1c1c;max-width:100px;' padding='5'>" + structureObject["contentTypes"][i] + "</td>";
					}
					if(structureObject["contentTypes"].length<ContentTypeDepth){
						for(var i=0; i<(ContentTypeDepth-structureObject["contentTypes"].length);i++){
							if(i==ContentTypeDepth-structureObject["contentTypes"].length-1)
								table += "<td style='color:#1c1c1c;border-right: 2px solid #0B3861;'></td>";
							else
								table += "<td style='color:#1c1c1c;'></td>";
						}
					}
				}else{
					for(var i =0; i < ContentTypeDepth;i++){
						if(i==ContentTypeDepth-1)
							table += "<td style='color:#1c1c1c;border-right: 2px solid #0B3861;'></td>";
						else
							table += "<td style='color:#1c1c1c;'></td>";
					}
				}
			}else{
				for(var i =0; i < ContentTypeDepth;i++){
					if(i==ContentTypeDepth-1)
						table += "<td style='color:#1c1c1c;background:#BDBDBD;border-right: 2px solid #0B3861;'></td>";
					else
						table += "<td style='color:#1c1c1c;background:#BDBDBD;'></td>";
				}
			}
		}
		return table;
	}
	
	function renderDefaultsHeader(renderFields){
		var table = "";
		if(getDefaults){
			if(renderFields){
				if(defaultsDepth>0){
					for(var j=0; j< listDefaults.length;j++){
						table += "<td style='background:#0B3861;color:#fff;border:0;' align='center'>" + listDefaults[j] + "</td>";
					}
				}else{
					table += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;'></td>";
				}
			}else{
				if(defaultsDepth>0){
					for(var i =0; i < listDefaults.length;i++){
						table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
					}
				}else{
					table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
				}
			}
		}
		return table;
	}
	
	function renderDefaultsBody(structureObject,isFolder,parentStructure){
		var table = "";
		if(getDefaults){
			if(structureObject != null){
				if(isFolder){
					if("defaults" in structureObject){
						for(var j=0;j<listDefaults.length;j++){
							if(listDefaults[j] in structureObject["defaults"])
								table += "<td style='color:#1c1c1c;max-width:150px;'>" + structureObject["defaults"][listDefaults[j]] + "</td>";
							else{
								if(listDefaults[j] in parentStructure["defaults"])
									table += "<td style='color:#1c1c1c;max-width:150px;'>" + parentStructure["defaults"][listDefaults[j]] + "</td>";
								else
									table += "<td style='color:#1c1c1c;'></td>";
							}
						}	
					}else{
						if(defaultsDepth>0){
							for(var j=0;j<listDefaults.length;j++){
								if(listDefaults[j] in parentStructure["defaults"])
									table += "<td style='color:#1c1c1c;max-width:150px;'>" + parentStructure["defaults"][listDefaults[j]] + "</td>";
								else
									table += "<td style='color:#1c1c1c;'></td>";
							}
						}else{
							table += "<td style='color:#1c1c1c;'></td>";
						}
					}
				}else{
					if("defaults" in structureObject){
						if(defaultsDepth>0){
							for(var j=0;j<listDefaults.length;j++){
								if(listDefaults[j] in structureObject["defaults"])
									table += "<td style='color:#1c1c1c;max-width:150px;'>" + structureObject["defaults"][listDefaults[j]] + "</td>";
								else
									table += "<td style='color:#1c1c1c;'></td>";
							}
						}else{
							table += "<td style='color:#1c1c1c;'></td>";
						}
					}
				}
			}else{
				if(defaultsDepth>0){
					for(var i =0; i < listDefaults.length;i++){
						table += "<td style='background:#BDBDBD;color:#fff;'></td>";
					}
				}else{
					table += "<td style='background:#BDBDBD;color:#fff;'></td>";
				}
			}
		}
		return table;
	}
	
	function createTableHtml(){

		$("#progressMessages").html("").append("<img src='/teams/ITE/Office365/eZShare/SiteAssets/loading.gif'/>&nbsp; Rendering Information Architecture");
	
		var css = "<style>";
		css += "table#siteStructure, #siteStructure th, #siteStructure td, table#bagProperties, #bagProperties th, #bagProperties td, table.libraryViews, .libraryViews th, .libraryViews td{border: 0.2px solid #ddd;} #siteStructure th, #siteStructure td, #bagProperties th, #bagProperties td {padding: 5px;}";
		css +="td.rotate{height:60px;white-space: nowrap;font-size:10px;} td.rotate > div {transform:translate(0px, 10px) rotate(270deg);width: 25px;} td.rotate > div > span {}";
		css +="#InfoOptions a{margin-top:10px;margin-rigth:0px;float:right;cursor:pointer;} .tabInfo{float:left;padding:10px;border-top:3px solid #E6E6E6; cursor:pointer;margin-right:5px;} .tabInfo.active, .tabInfo.active:hover{border-top:3px solid #0B3861;font-weight:bold;} .tabInfo:hover{border-top:3px solid #0B3861;cursor:pointer;}";
		css += "</style>";
		
		var table = "<table id='siteStructure' cellspacing='0'>";
		var tableProperties = "";
		var tableViews = "";
		
		table += "<thead>";
		
		table += "<tr>";
			table += "<td style='display:none;' colspan='" + (3 + SubsiteDepth) + securityDepth + ContentTypeDepth + defaultsDepth + "'><h1 style='font-weight:400;color:#1c1c1c;'>ezShare Information Architecture: " + structure["Title"] + "</h1></td>";
		table += "</tr>";
		table += "<tr>";
			table += "<td style='display:none;' colspan='2' align='right'>Generated on: </td><td style='display:none;' align='left' colspan='" + (1 + SubsiteDepth) + securityDepth + ContentTypeDepth + defaultsDepth + "'> " + new Date().toISOString().slice(0, 10) + " </td>";
		table += "</tr>";
		table += "<tr>";
			table += "<td style='display:none;' colspan='2' align='right'>Generated by: </td><td style='display:none;' class='currentRequester' align='left' colspan='" + (1 + SubsiteDepth) + securityDepth + ContentTypeDepth + defaultsDepth + "'>Test</td>";
		table += "</tr>";
		table += "<tr><td style='display:none;' colspan='" + (3 + SubsiteDepth) + securityDepth + ContentTypeDepth + defaultsDepth + "'> </td></tr>";
		
		// First row
		table += "<tr>";
			table += "<td colspan='" + (2 + SubsiteDepth + foldersDepth) + "' class='first' style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;' padding='10'><h1 style='color:#fff;font-weight:400;'>Site Structure</h1></td>";
			if(getSecurity){
				table += "<td colspan='" + securityDepth + "' style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;' padding='10' align='center'><h1 style='color:#fff;font-weight:400;'>Security</h1></td>";
			}
			if(getContentTypes){
				table += "<td colspan='" + ContentTypeDepth + "' style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;min-width:250px;' padding='10' align='center'><h1 style='color:#fff;font-weight:400;'>Content Types</h1></td>";
			}
			if(getDefaults){
				table += "<td colspan='" + listDefaults.length + "' style='background:#0B3861;color:#fff;border:0;' padding='10' align='center'><h1 style='color:#fff;font-weight:400;'>Defaults</h1></td>";
			}
		table += "</tr>";
		table += "</table>";

		$("#displayResults").html("").append(css + table);
		
		if(getMetadata){
			tableProperties = "<table style='display:none;' id='bagProperties' cellspacing='0'>";
			tableProperties += "<tr>";
				tableProperties += "<td style='display:none;' colspan='" + spPropertyBags.length + "'><h1 style='font-weight:400;color:#1c1c1c;'>ezShare Site metadata: " + structure["Title"] + "</h1></td>";
			tableProperties += "</tr>";
			tableProperties += "<tr>";
				tableProperties += "<td style='display:none;' colspan='2' align='right'>Generated on: </td><td style='display:none;' align='left' colspan='" + (spPropertyBags.length - 2) + "'> " + new Date().toISOString().slice(0, 10) + " </td>";
			tableProperties += "</tr>";
			tableProperties += "<tr>";
				tableProperties += "<td style='display:none;' colspan='2' align='right'>Generated by: </td><td style='display:none;' class='currentRequester' align='left' colspan='" + (spPropertyBags.length - 2) + "'>Test</td>";
			tableProperties += "</tr>";
			tableProperties += "<tr><td style='display:none;' colspan='" + spPropertyBags.length + "'> </td></tr>";
			
			// First row Bag Properties
			tableProperties += "<tr>";
			for(var i=0; i < spPropertyBags.length;i++){
				if(spPropertyBags[i] == "IDBSiteDescription"){
					tableProperties += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;min-width:300px;' padding='10' align='center'><h3 style='color:#fff;font-weight:400;'>" + spPropertyBags[i].substring(3,spPropertyBags[i].length); + "</h3></td>";
				}else{
					tableProperties += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;' padding='10' align='center'><h3 style='color:#fff;font-weight:400;'>" + spPropertyBags[i].substring(3,spPropertyBags[i].length); + "</h3></td>";
				}
			}
			tableProperties += "</tr>";
			tableProperties += "</table>";
			$("#displayResults").append(tableProperties);
		}
		
		if(getViews){
			var id_aux = (structure["Title"]).replace(/\s/g,"_").replace(/,/g,"_");
			tableViews = "<table style='display:none;' id='libraryViews_" + id_aux + "' class='libraryViews' cellspacing='0'>";
			if("libraries" in structure){
				tableViews += "<tr>";
					tableViews += "<td style='color:#084B8A;font-weight:bold;' colspan='2'>Site</td><td align='center' colspan='" + (Object.keys(structure["libraries"]).length + 1) + "'>" + structure["Title"] + "</td>";
				tableViews += "</tr>";
				tableViews += "<tr>";
					tableViews += "<td style='color:#084B8A;font-weight:bold;' colspan='2'>Library</td>";
				tableViews += "</tr>";
				for(var i=0;i<viewsCount;i++){
					tableViews += "<tr class='viewName_" + (i+1) + "'>";
						tableViews += "<td style='color:#084B8A;font-weight:bold;' colspan='2'>View Name</td>";
					tableViews += "</tr>";
					tableViews += "<tr>";
						tableViews += "<th rowspan='6' style='color:#084B8A;font-weight:bold;'>View Specification</th><td>Default</td>";
					tableViews += "</tr>";
					tableViews += "<tr><td>Sort</td></tr>";
					tableViews += "<tr><td>Filter</td></tr>";
					tableViews += "<tr><td>Group</td></tr>";
					tableViews += "<tr><td>Item Limit</td></tr>";
					tableViews += "<tr><td>Totals</td></tr>";
					tableViews += "<tr>";
						tableViews += "<td style='color:#084B8A;font-weight:bold;' colspan='2'>Columns</td>";
					tableViews += "</tr>";
				}
			}
			tableViews += "</table>";
			$("#displayResults").append(tableViews);
		}
		
		
		table = "";
		// Site collection row
		table += "<tr>";
			//table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
			table += "<td class='first'  style='background:#0B3861;color:#fff;border:0;min-width:80px;' padding='5' align='center'>Site Collection</td>";
			for(var i =0; i < SubsiteDepth + 1 + foldersDepth;i++){
				if(i==SubsiteDepth + 1 + foldersDepth -1)
					table += "<td class='first' style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;min-width:80px;'></td>";
				else
					table += "<td class='first' style='background:#0B3861;color:#fff;border:0;min-width:80px;'></td>";
				
			}
			if(getSecurity){
				table += "<td colspan='" + securityDepth + "' style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;' padding='10' align='center'><b>Levels:</b> No Access - 0, Read-Only - 1, Contribute - 2; Design - 3, Full control - 4</td>";
			}
			table += renderContentTypesHeader(false);
			table += renderDefaultsHeader(false);
		table += "</tr>";
		
		$("#siteStructure").append(table);
		
		table = "";
		// Root Site row
		table += "<tr>";
			//table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
			table += "<td style='background:#0B3861;color:#fff;border:0;'></td><td style='background:#0B3861;color:#fff;border:0;' padding='5' align='center'>Root Site</td>";
			for(var i =0; i < SubsiteDepth - 1 + 1 + foldersDepth;i++){
				if(i==SubsiteDepth - 1 + 1 + foldersDepth -1)
					table += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;'></td>";
				else
					table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
			}
			table += renderSecurityHeader(false);
			table += renderContentTypesHeader(false);
			table += renderDefaultsHeader(false);
		table += "</tr>";
		
		$("#siteStructure").append(table);
		
		table = "";
		// Subsites Rows
		for(var i =0; i < SubsiteDepth-1;i++){
			table += "<tr>";
				//table += "<td style='background:#0B3861;color:#fff;border:0;'>No.</td>";
				for(var j=0;j<2+i;j++){
					table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
				}
				table += "<td style='background:#0B3861;color:#fff;border:0;' padding='5' align='center'>Sub-Site L" + (i+1) + "</td>";
				for(var j =0; j < SubsiteDepth-(i+2) + 1 + foldersDepth;j++){
					if(j==SubsiteDepth-(i+2) + 1 + foldersDepth -1)
						table += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;'></td>";
					else
						table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
				}
				table += renderSecurityHeader(false); 
				table += renderContentTypesHeader(false);
				table += renderDefaultsHeader(false);
			table += "</tr>";
		}
		
		$("#siteStructure").append(table);
		
		table = "";
		// Library Row
		table += "<tr>";
			//table += "<td style='background:#0B3861;color:#fff;border:0;'>No.</td>";
			for(var i=0;i<SubsiteDepth-1+2;i++){
				table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
			}
			table += "<td style='background:#0B3861;color:#fff;border:0;' padding='5' align='center'>Library</td>"
			for(var i=0;i<foldersDepth;i++){
				if(i == foldersDepth-1)
					table += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;'></td>";
				else
					table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
			}
			table += renderSecurityHeader(false);
			table += renderContentTypesHeader(false);
			table += renderDefaultsHeader(false);
		table += "</tr>";
		
		$("#siteStructure").append(table);
		
		table = "";
		// Folder Row
		for(var index = 0; index < foldersDepth; index++){
			table += "<tr>";
				//table += "<td style='background:#0B3861;color:#fff;border:0;'>No.</td>";
				for(var i=0;i<SubsiteDepth+2;i++){
					table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
				}
				for (var subindex = 0; subindex < index; subindex++){
					table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
				}
				if(index == foldersDepth-1){
					table += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;' padding='5' align='center'>Folder</td>";
					table += renderSecurityHeader(true);
					table += renderContentTypesHeader(true);
					table += renderDefaultsHeader(true);
				}else{
					table += "<td style='background:#0B3861;color:#fff;border:0;' padding='5' align='center'>Folder</td>";
					for (var subindex = index+1; subindex < foldersDepth; subindex++){
						if(subindex == foldersDepth - 1)
							table += "<td style='background:#0B3861;color:#fff;border:0;border-right: 2px solid #ddd;'></td>";
						else
							table += "<td style='background:#0B3861;color:#fff;border:0;'></td>";
					}
					table += renderSecurityHeader(false);
					table += renderContentTypesHeader(false);
					table += renderDefaultsHeader(false);
				}
				
			table += "</tr>";
		}
		
		table += "</thead>";
		table += "<body>";
		
		$("#siteStructure").append(table);
		
		// Site Collection name
		table = "";
		table += "<tr>";
			//table += "<td style='color:#1c1c1c;'>1</td>";
			if("Title" in structure){
				table += "<td style='color:#1c1c1c;border-right: 2px solid #0B3861;' colspan='" + (SubsiteDepth+1+foldersDepth+1) + "'>" + structure["Title"] + "</td>";
			}else{
				table += "<td style='color:#1c1c1c;border-right: 2px solid #0B3861;' colspan='" + (SubsiteDepth+1+foldersDepth+1) + "'>Site Collection Name (NO ACCESS)</td>";
			}
			table += renderSecurityBody(null);
			table += renderContentTypesBody(null);
			table += renderDefaultsBody(null);
		table += "</tr>";
		
		$("#siteStructure").append(table);
		
		// Root site name
		table = "";
		table += "<tr>";
			//table += "<td style='color:#1c1c1c;'>1.1</td>";
			table += "<td style='color:#1c1c1c;'></td>";
			if("Title" in structure){
				table += "<td style='color:#1c1c1c;'><a href='" + structure["url"] + "' target='_blank'>" + structure["Title"] + "</a></td>";
			}else{
				table += "<td style='color:#1c1c1c;'>Root Site Name(NO ACCESS)</td>";
			}
			for(var i=0;i<SubsiteDepth-1+1+foldersDepth;i++){
				if(i==SubsiteDepth-1+1+foldersDepth-1)
					table += "<td style='color:#fff;border-right: 2px solid #0B3861;'></td>";
				else
					table += "<td style='color:#1c1c1c;'></td>";
			}
			table += renderSecurityBody(structure);
			table += renderContentTypesBody(null);
			table += renderDefaultsBody(null);
		table += "</tr>";
		
		$("#siteStructure").append(table);
		
		// Bag Properties for root Site
		tableProperties = "";
		if("bagProperties" in structure && Object.keys(structure["bagProperties"]).length > 0){
			tableProperties += "<tr>";
			for(var i=0; i < spPropertyBags.length;i++){
				if(spPropertyBags[i] in structure["bagProperties"])
					tableProperties += "<td style='color:#1c1c1c;' padding='10' align='center'>" + structure["bagProperties"][spPropertyBags[i]] + "</td>";
				else
					tableProperties += "<td style='color:#1c1c1c;' padding='10' align='center'></td>";
			}
			tableProperties += "</tr>";
			tableProperties += "</table>";

			$("#bagProperties").append(tableProperties);
		}
		
		// Looping through the Structure
		table = "";
		for(var key in structure){
			if(key == "libraries"){
				table += displayLibraries(structure["libraries"],(structure["Title"]).replace(/\s/g,"_").replace(/,/g,"_"));
			}else if(key == "subsites"){
				table += displaySubsites(structure["subsites"],0,"1.1.0");
			}
		}

		table += "</body>";
		
		$("#siteStructure").append(table);
		
		$("#siteStructure").wrap('<div id="table_container"></div>').after('<div id="bottom_anchor"></div>');
	
		$("#getStructure").prop('disabled', false);
		$("#getStructureSite").prop('disabled', false);
		//$("#gettingStructure").attr("src","");
		//$("#progressMessages").html("").append('<font color="green">Process finished</font>');
		
		$("#progressMessages").html("");
		
		$("#showSiteStructure").removeClass("active");
		$("#showBagProperties").removeClass("active");
		$("#showViews").removeClass("active");
		
		$("#showSiteStructure").hide();
		$("#showBagProperties").hide();
		$("#showViews").hide();
		
		if(getStructure){
			$("#showSiteStructure").show();
			$("#showSiteStructure").click();
		}
		
		if(getMetadata){
			$("#showBagProperties").show();
			if(!getStructure){
				$("#showBagProperties").click();
			}
		}
		
		if(getViews){
			$("#showViews").show();
			if(!getStructure && !getMetadata){
				$("#showViews").click();
			}
		}
		
		$("#InfoOptions").show();

	}
	
	function displayFolders(currentStructure,parentStructure,level){
		var table = "";
		if("Folders" in currentStructure){
			for(var folder in currentStructure["Folders"]){
				table += "<tr>";
				//table += "<td style='color:#1c1c1c;'>" + numeration + "</td>";
				for(var i=0;i<SubsiteDepth-1+2+1+level;i++){
					table += "<td style='color:#1c1c1c;'></td>";
				}
				if(foldersDepth == 1 || level == foldersDepth-1)
					table += "<td style='color:#1c1c1c;border-right: 2px solid #0B3861;max-width:150px;'><a href='" + currentStructure["Folders"][folder]['url'] + "' target='_blank'>" + folder + "</a></td>";
				else
					table += "<td style='color:#1c1c1c;max-width:150px;'><a href='" + currentStructure["Folders"][folder]['url'] + "' target='_blank'>" + folder + "</a></td>";
				for(var i=0;i<foldersDepth-1-level;i++){
					if(i == foldersDepth - 1 - level - 1)
						table += "<td style='color:#1c1c1c;border-right: 2px solid #0B3861;'></td>";
					else
						table += "<td style='color:#1c1c1c;'></td>";
				}
				table += renderSecurityBody(currentStructure["Folders"][folder]);
				table += renderContentTypesBody(parentStructure);
				table += renderDefaultsBody(currentStructure["Folders"][folder],true,parentStructure);
				table += "</tr>";
				table += displayFolders(currentStructure["Folders"][folder],parentStructure,level+1);
			}
		}
		return table;
	}
	
	function displayLibraries(currentStructure,parent){
		var table = "";
		for(var library in currentStructure){
			table += "<tr>";
				//table += "<td style='color:#1c1c1c;'>" + numeration + "</td>";
				for(var i=0;i<SubsiteDepth+1;i++){
					table += "<td style='color:#1c1c1c;'></td>";
				}
				table += "<td style='color:#1c1c1c;max-width:150px;'><a href='" + currentStructure[library]['url'] + "' target='_blank'>" + library + "</a></td>";
				for(var i=0;i<foldersDepth;i++){
					if(i==foldersDepth-1){
						table += "<td style='color:#1c1c1c;border-right: 2px solid #0B3861;'></td>";
					}else{
						table += "<td style='color:#1c1c1c;'></td>";
					}
				}
				table += renderSecurityBody(currentStructure[library]);
				table += renderContentTypesBody(currentStructure[library]);
				table += renderDefaultsBody(currentStructure[library]);
			table += "</tr>";
			
			table += displayFolders(currentStructure[library],currentStructure[library],0);
			
			if(getViews){
				$("#libraryViews_" + parent).find('tr:nth-child(2)').append('<td align="center">' + library + '</td>');
				var viewRow = 1;
				for(var view in currentStructure[library]["views"]){
					$("#libraryViews_" + parent).find(".viewName_" + viewRow).append('<td>' + view + '</td>');
					$("#libraryViews_" + parent).find(".viewName_" + viewRow).next("tr").append('<td>' + currentStructure[library]["views"][view]["default"] + '</td>');
					$("#libraryViews_" + parent).find(".viewName_" + viewRow).next("tr").next("tr").append('<td>' + currentStructure[library]["views"][view]["sort"] + '</td>');
					$("#libraryViews_" + parent).find(".viewName_" + viewRow).next("tr").next("tr").next("tr").append('<td>' + currentStructure[library]["views"][view]["filter"] + '</td>');
					$("#libraryViews_" + parent).find(".viewName_" + viewRow).next("tr").next("tr").next("tr").next("tr").append('<td>' + currentStructure[library]["views"][view]["group"] + '</td>');
					$("#libraryViews_" + parent).find(".viewName_" + viewRow).next("tr").next("tr").next("tr").next("tr").next("tr").append('<td>' + currentStructure[library]["views"][view]["limit"] + '</td>');
					$("#libraryViews_" + parent).find(".viewName_" + viewRow).next("tr").next("tr").next("tr").next("tr").next("tr").next("tr").append('<td>' + currentStructure[library]["views"][view]["totals"] + '</td>');
					viewRow++;
				}
				if(viewRow < viewsCount){
					for(var i = viewRow; i<viewsCount+1; i++){
						$("#libraryViews_" + parent).find(".viewName_" + i).append('<td style="background:#E6E6E6;"></td>');
						$("#libraryViews_" + parent).find(".viewName_" + i).next("tr").append('<td style="background:#E6E6E6;"></td>');
						$("#libraryViews_" + parent).find(".viewName_" + i).next("tr").next("tr").append('<td style="background:#E6E6E6;"></td>');
						$("#libraryViews_" + parent).find(".viewName_" + i).next("tr").next("tr").next("tr").append('<td style="background:#E6E6E6;"></td>');
						$("#libraryViews_" + parent).find(".viewName_" + i).next("tr").next("tr").next("tr").next("tr").append('<td style="background:#E6E6E6;"></td>');
						$("#libraryViews_" + parent).find(".viewName_" + i).next("tr").next("tr").next("tr").next("tr").next("tr").append('<td style="background:#E6E6E6;"></td>');
						$("#libraryViews_" + parent).find(".viewName_" + i).next("tr").next("tr").next("tr").next("tr").next("tr").next("tr").append('<td style="background:#E6E6E6;"></td>');
					}
				}
			}
			
		}
		return table;
	}
	
	function displaySubsites(currentStructure,level,numeration){
		var table = "";
		var currentNumeracion = numeration;
		for(var subsite in currentStructure){
			currentNumeracion = currentNumeracion.split(".");
			currentNumeracion[currentNumeracion.length-1] = parseInt(currentNumeracion[currentNumeracion.length-1])+1;
			currentNumeracion = currentNumeracion.join(".");
			table += "<tr>";
				//table += "<td style='color:#1c1c1c;'>" + currentNumeracion + "</td>";
				table += "<td style='color:#1c1c1c;'></td><td style='color:#1c1c1c;'></td>";
				for (var i =0;i<level;i++){
					table += "<td style='color:#1c1c1c;'></td>";
				}
				table += "<td style='color:#1c1c1c;max-width:150px;'><a href='" + currentStructure[subsite]['url'] + "' target='_blank'>" + subsite + "</a></td>";
				for(var i=0;i<SubsiteDepth-2-level;i++){
					table += "<td style='color:#1c1c1c;'></td>";
				}
				table += "<td style='color:#1c1c1c;'></td>";
				for(var i=0;i<foldersDepth;i++){
					if(i == foldersDepth - 1)
						table += "<td style='color:#1c1c1c;border-right: 2px solid #0B3861;'></td>";
					else
						table += "<td style='color:#1c1c1c;'></td>";
				}
				table += renderSecurityBody(currentStructure[subsite]);
				table += renderContentTypesBody(null);
				table += renderDefaultsBody(null);
			table += "</tr>";
			if("libraries" in currentStructure[subsite]){
				if(getViews){
					var id_aux = subsite.replace(/\s/g,"_").replace(/,/g,"_").replace(/'/g, '').replace(/[\/|\\]/g,"_").replace(/\(/g,"_").replace(/\)/g,"_");
					var tableViews = "<table style='display:none;' id='libraryViews_" + id_aux + "' class='libraryViews' cellspacing='0'>";
					tableViews += "<tr>";
						tableViews += "<td style='color:#084B8A;font-weight:bold;' colspan='2'>Site</td><td align='center' colspan='" + (Object.keys(currentStructure[subsite]['libraries']).length) + "'>" + subsite + "</td>";
					tableViews += "</tr>";
					tableViews += "<tr>";
						tableViews += "<td style='color:#084B8A;font-weight:bold;' colspan='2'>Library</td>";
					tableViews += "</tr>";
					for(var i=0;i<viewsCount;i++){
						tableViews += "<tr class='viewName_" + (i+1) + "'>";
							tableViews += "<td style='color:#084B8A;font-weight:bold;' colspan='2'>View Name</td>";
						tableViews += "</tr>";
						tableViews += "<tr>";
							tableViews += "<th rowspan='6' style='color:#084B8A;font-weight:bold;'>View Specification</th><td>Default</td>";
						tableViews += "</tr>";
						tableViews += "<tr><td>Sort</td></tr>";
						tableViews += "<tr><td>Filter</td></tr>";
						tableViews += "<tr><td>Group</td></tr>";
						tableViews += "<tr><td>Item Limit</td></tr>";
						tableViews += "<tr><td>Totals</td></tr>";
						tableViews += "<tr>";
							tableViews += "<td style='color:#084B8A;font-weight:bold;' colspan='2'>Columns</td>";
						tableViews += "</tr>";
					}
					tableViews += "</table>";
					$("#displayResults").append(tableViews);
				}

				table += displayLibraries(currentStructure[subsite]["libraries"],subsite.replace(/\s/g,"_").replace(/,/g,"_").replace(/'/g, '').replace(/[\/|\\]/g,"_").replace(/\(/g,"_").replace(/\)/g,"_"));
			}
			
			tableProperties = "";
			if("bagProperties" in currentStructure[subsite] && Object.keys(currentStructure[subsite]["bagProperties"]).length > 0){
				tableProperties += "<tr>";
				for(var i=0; i < spPropertyBags.length;i++){
					if(spPropertyBags[i] in currentStructure[subsite]["bagProperties"])
						tableProperties += "<td style='color:#1c1c1c;' padding='10' align='center'>" + currentStructure[subsite]["bagProperties"][spPropertyBags[i]] + "</td>";
					else
						tableProperties += "<td style='color:#1c1c1c;' padding='10' align='center'></td>";
				}
				tableProperties += "</tr>";
				tableProperties += "</table>";

				$("#bagProperties").append(tableProperties);
			}
			
			if("subsites" in currentStructure[subsite]){
				table += displaySubsites(currentStructure[subsite]["subsites"],level+1,currentNumeracion+".0");
			}
		}
		return table;
	}

});