script_directory = File.dirname(__FILE__)
require File.join(script_directory,"Nx.jar")
java_import "com.nuix.nx.NuixConnection"
java_import "com.nuix.nx.LookAndFeelHelper"
java_import "com.nuix.nx.dialogs.ChoiceDialog"
java_import "com.nuix.nx.dialogs.TabbedCustomDialog"
java_import "com.nuix.nx.dialogs.CommonDialogs"
java_import "com.nuix.nx.dialogs.ProgressDialog"
java_import "com.nuix.nx.digest.DigestHelper"
java_import "com.nuix.nx.controls.models.Choice"

LookAndFeelHelper.setWindowsIfMetal
NuixConnection.setUtilities($utilities)
NuixConnection.setCurrentNuixVersion(NUIX_VERSION)

require "csv"

def escape_tag_for_search(tag)
	return tag
		.gsub("\\","\\\\\\") #Escape \
		.gsub("?","\\?") #Escape ?
		.gsub("\"","\\\"") #Escape "
end

def escape_tag_for_fs(tag)
	result = tag.gsub(/[\/\\:\?\"<>\|]+/,"_")
	result = result.gsub(/\*/,"")
	return result
end

email_format_choices = [
	"pst",
	"nsf",
]

dialog = TabbedCustomDialog.new("Export By Tag")

tag_choices = $current_case.getAllTags.map{|tag| Choice.new(tag)}
input_tab = dialog.addTab("input_tab","Tags")
if $current_selected_items.nil? == false && $current_selected_items.size > 0
	input_tab.appendHeader("Selected Items: #{$current_selected_items.size}")
end
input_tab.appendCheckBox("include_nested","Include nested tags",true)
input_tab.appendRadioButton("input_from_csv","Use tags from CSV","input_group",false)
input_tab.appendOpenFileChooser("input_csv","Input CSV","CSV","csv")
input_tab.enabledOnlyWhenChecked("input_csv","input_from_csv")
input_tab.appendRadioButton("input_from_selection","Use selected tags","input_group",true)
input_tab.appendChoiceTable("selected_tags","Selected Tags",tag_choices)
input_tab.enabledOnlyWhenChecked("selected_tags","input_from_selection")

options_tab = dialog.addTab("options_tab","Options")
options_tab.appendComboBox("email_format","Email Format",email_format_choices)
options_tab.appendDirectoryChooser("output_directory","Output Directory")
options_tab.appendSaveFileChooser("output_csv","Output CSV","CSV","csv")
# Production set creation is feature restricted, additionally the method to enable/disable it
# throws error if you don't have that feature
if $utilities.getLicence.hasFeature("PRODUCTION_SET")
	options_tab.appendCheckBox("create_production_set","Create default production set",false)
end
options_tab.appendHeader("Scope Query (blank means no additional scoping)")
options_tab.appendTextArea("scope_query","","")

workers_tab = dialog.addTab("workers_tab","Worker Settings")
workers_tab.appendLocalWorkerSettings("worker_settings")

dialog.validateBeforeClosing do |values|
	if values["output_directory"].strip.empty?
		CommonDialogs.showWarning("Please provide a value for 'Output Directory'")
		next false
	end

	if values["input_from_csv"]
		if values["input_csv"].strip.empty?
			CommonDialogs.showWarning("Please provide a value for 'Input CSV'")
			next false
		end

		if java.io.File.new(values["input_csv"]).exists == false
			CommonDialogs.showWarning("Path provided for 'Input CSV' does not exist.")
			next false
		end
	end

	if values["input_from_selection"] && values["selected_tags"].size < 1
		CommonDialogs.showWarning("Please select at least one tag")
		next false
	end

	if values["output_csv"].strip.empty?
		CommonDialogs.showWarning("Please provide a value for 'Output CSV'")
		next false
	end

	next true
end

dialog.display
if dialog.getDialogResult == true
	values = dialog.toMap
	last_progress = Time.now
	data = []

	ProgressDialog.forBlock do |pd|
		pd.setTitle("Export By Tag")
		pd.setAbortButtonVisible(false)

		if values["input_from_csv"]
			pd.setMainStatusAndLogIt("Reading CSV...")
			#Expected columns:
			# Tag
			# Sub Directory (Optional)
			CSV.foreach(values["input_csv"],{:headers => :first_row}) do |row|
				entry = {
					:tag => row[0],
				}
				if row.size > 1
					entry[:dir] = row[1]
				end
				data << entry
			end
			pd.logMessage("Read #{data.size} entries from CSV")
		end

		if values["input_from_selection"]
			data = values["selected_tags"].map{|tag| {:tag => tag} }
		end

		j_output_csv_dir = java.io.File.new(values["output_csv"]).getParentFile
		if !j_output_csv_dir.exists
			j_output_csv_dir.mkdirs
		end

		CSV.open(values["output_csv"],"w:utf-8") do |csv|
			#Write out headers
			csv << [
				"Tag",
				"Hits",
				"AuditedSize",
				"StartTime",
				"FinishTime",
				"Elapsed",
				"ExportDirectory",
				"LooseFiles",
			]

			pd.setMainProgress(0,data.size)

			data.each_with_index do |entry,entry_index|
				pd.setMainStatus("Exporting #{entry_index+1}/#{data.size}: #{entry[:tag]}")
				pd.logMessage("=== Exporting #{entry_index+1}/#{data.size}: #{entry[:tag]} ===")
				pd.setMainProgress(entry_index+1)
				
				export_start_time = Time.now
				tag = entry[:tag]
				if values["include_nested"]
					tag = "#{tag}*"
				end
				dir = entry[:dir] || tag
				dir = escape_tag_for_fs(dir)

				if !values["scope_query"].nil? && ! values["scope_query"].strip.empty?
					pd.logMessage("Scope Query: #{values["scope_query"]}")
					query = "(#{values["scope_query"]}) AND (tag:\"#{escape_tag_for_search(tag)}\")"
				else
					query = "tag:\"#{escape_tag_for_search(tag)}\""
				end

				export_directory = "#{values["output_directory"]}\\#{dir}"
				loose_files = 0

				pd.logMessage("Exporting: #{tag}")
				pd.logMessage("Subdirectory: #{dir}")
				pd.logMessage("Export Directory: #{export_directory}")
				pd.logMessage("Query: #{query}")

				pd.setSubStatusAndLogIt("Searching...")
				items = $current_case.search(query)
				pd.logMessage("Hit Items: #{items.size}")

				if $current_selected_items.nil? == false && $current_selected_items.size > 0
					items = $utilities.getItemUtility.intersection(items,$current_selected_items)
					pd.logMessage("Hits Items from Selected Items: #{items.size}")
				end

				pd.setSubStatusAndLogIt("Calculating total audited size...")
				total_audited_size = items.map{|i|i.getAuditedSize||0}.reject{|s|s<0}.reduce(0,:+)
				pd.logMessage("Total Audited Size: #{total_audited_size}")

				if items.size < 1
					pd.logMessage("No items to export")
					export_finish_time = Time.now
					elapsed = Time.at(export_finish_time - export_start_time).gmtime.strftime("%H:%M:%S")
					#Output CSV record
					csv << [
						tag,
						items.size,
						total_audited_size,
						export_start_time,
						export_finish_time,
						elapsed,
						export_directory,
						loose_files,
					]
				else
					pd.setSubStatusAndLogIt("Creating Exporter...")
					exporter = $utilities.createBatchExporter(export_directory)

					pd.setSubStatusAndLogIt("Configuring Exported Products...")
					exporter.addProduct("native",{
						:naming => "guid",
						:path => "NATIVE",
						:mailFormat => values["email_format"],
						:includeAttachments => true,
					})

					pd.setSubStatusAndLogIt("Settings Traversal Options...")
					exporter.setTraversalOptions({
						:strategy => "items",
						:sortOrder => "position",
					})

					# Disable default behavior of creating a production set
					if $utilities.getLicence.hasFeature("PRODUCTION_SET")
						unless values["create_production_set"] == true
							exporter.setNumberingOptions({"createProductionSet" => false})
						end
					end

					parallel_settings = values["worker_settings"]

					if !ENV_JAVA["nuix.processing.sharedTempDirectory"].nil? && !ENV_JAVA["nuix.processing.sharedTempDirectory"].empty?
						pd.logMessage("Setting worker shared temp directory using argument '-Dnuix.processing.sharedTempDirectory'")
						parallel_settings[:workerTemp] = ENV_JAVA["nuix.processing.sharedTempDirectory"]
					end

					exporter.setParallelProcessingSettings(parallel_settings)

					pd.setSubProgress(0,items.size)
					exporter.whenItemEventOccurs do |event|
						if (Time.now - last_progress) > 1
							pd.setSubStatus("Exporting, Stage: #{event.getStage}")
							pd.setSubProgress(event.getStageCount)
							last_progress = Time.now
						end
					end

					pd.setSubStatusAndLogIt("Beginning Export...")
					exporter.exportItems(items)

					export_finish_time = Time.now
					elapsed = Time.at(export_finish_time - export_start_time).gmtime.strftime("%H:%M:%S")
					pd.logMessage("Export Finished in #{elapsed}")

					pd.setSubStatusAndLogIt("Renaming...")
					# TODO: might look at finding a better way to do this, including scan only once
					all_exported = Dir.glob("#{export_directory}/NATIVE/**/*").reject{|p|java.io.File.new(p).isDirectory}
					mailstores = Dir.glob("#{export_directory}/NATIVE/**/*.#{values["email_format"]}")
					loose_files = all_exported.size - mailstores.size

					mailstores.each do |mail_store|
						#We determine the target name to rename each mailstore to.  Note we check for
						#collisions with files already present and sequence them "_1", "_2", etc
						target_name = "#{export_directory}/NATIVE/#{dir}.#{values["email_format"]}"
						sequence = 1
						while File.exists?(target_name)
							target_name = "#{export_directory}/NATIVE/#{dir}_#{sequence}.#{values["email_format"]}"
							sequence += 1
						end
						#Rename
						File.rename(mail_store,target_name)
					end

					#Cleanup empty directories after moving around mail stores
					pd.setSubStatusAndLogIt("Performing empty sub directory cleanup...")
					natives_root_directory = java.io.File.new("#{export_directory}\\NATIVE\\")
					native_root_subdirectories = natives_root_directory.listFiles.select{|f|f.isDirectory}
					native_root_subdirectories.each do |sub_directory|
						sub_directory_contents_count = sub_directory.list.size
						if sub_directory_contents_count > 0
							pd.logMessage("'#{sub_directory.getName}' contains #{sub_directory_contents_count} files and/or directories, skipping deletion")
						else
							#pd.logMessage("'#{sub_directory.getName}' is empty, deleting...")
							sub_directory.delete
						end
					end

					#Output CSV record
					csv << [
						tag,
						items.size,
						total_audited_size,
						export_start_time,
						export_finish_time,
						elapsed,
						export_directory,
						loose_files,
					]
				end				
			end
		end
		pd.setCompleted
	end
end