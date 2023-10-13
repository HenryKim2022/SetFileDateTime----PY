from datetime import datetime
from docx import Document
from pytz import timezone
import bin


# Create filedate object
Target_File = "E:\DEL\Doc1.docx"
File_Date = bin.File(Target_File)

# Get file date
File_Date.get()
# "01.01.2023 12:00"; "3:30PM 2001/02/02"; "3rd March 2002 20:00:30"
XCreated = "10.13.2023 12:00PM"
XModified = "10.13.2023 02:00PM"
XAccessed = datetime.now()


Tf = Target_File.lower()
if (Tf.endswith(".xlsx")):
 	pass
elif (Tf.endswith(".docx")):
	# Open the Word document
	document = Document(Target_File)

	# Convert 'content_created' & 'date_last_saved' to datetime object
	content_created = datetime.strptime(XCreated, "%m.%d.%Y %I:%M%p")
	content_created = timezone('Asia/Jakarta').localize(content_created)
	date_last_saved = datetime.strptime(XModified, "%m.%d.%Y %I:%M%p")
	date_last_saved = timezone('Asia/Jakarta').localize(date_last_saved)

	# Get the core properties of the document
	core_properties = document.core_properties

	# Update the "Content Created" and "Date Last Saved" properties
	core_properties.created = content_created.astimezone(timezone('UTC'))
	core_properties.modified = date_last_saved.astimezone(timezone('UTC'))


	from docx import *
	from docx.oxml import *

	# Calculate the editing time duration
	editing_duration = date_last_saved - content_created
	# Convert the duration to hours, minutes, and seconds
	duration_hours = editing_duration.total_seconds() // 3600
	duration_minutes = (editing_duration.total_seconds() % 3600) // 60
	duration_seconds = editing_duration.total_seconds() % 60
	# Format the duration as a string
	duration_string = f"{int(duration_hours):02d}:{int(duration_minutes):02d}:{int(duration_seconds):02d}"


	# Get the core properties of the document
	core_properties = document.core_properties

	# Update the "Content Created" and "Date Last Saved" properties
	core_properties.created = content_created.astimezone(timezone('UTC'))
	core_properties.modified = date_last_saved.astimezone(timezone('UTC'))


	# Save the modified Word document
	document.save(Target_File)






	# Save the modified Word document
	core_properties.revision = int(core_properties.revision) + 1
	output_file = Target_File
	document.save(output_file)


	### FILE META: Set file created, mod, and access
	File_Date.set(
		created = XCreated,
		modified = XModified,
		accessed = XAccessed
	)




		# # Calculate the editing time duration
	# editing_duration = date_last_saved - content_created
	# # Convert the duration to hours, minutes, and seconds
	# duration_hours = editing_duration.total_seconds() // 3600
	# duration_minutes = (editing_duration.total_seconds() % 3600) // 60
	# duration_seconds = editing_duration.total_seconds() % 60
	# # Format the duration as a string
	# duration_string = f"{int(duration_hours):02d}:{int(duration_minutes):02d}:{int(duration_seconds):02d}"
	# # Set the "Total Editing Time" property
	# core_properties.total_editing_time = str(duration_string)
	