from pathlib import Path
import pandas as pd

# Set this to your Excel directory.
folder = r"C:\Users\gregor.pobegen\OneDrive - Österreichischer Bergrettungsdienst Landesorganisation Kärnten\Dateien von IT - Bergrettung - OS_Villach\Beteiligung"
types = ["Einsätze", "Übungen", "Dienste", "Übungstour"] # Sicherheitstage, 

criteria = { # mindestens
	"Einsätze": 5,
	"Übungen": 6,
	"Dienste": 2,
	"Übungstour": 1,
}

# scan subfolder to see what years are available
base_path = Path(folder)
available_years = sorted([int(p.name) for p in base_path.iterdir() if p.is_dir()])
print(f"Available years: {available_years}")

# now make a summary file for every year
for year in available_years:
	year_folder = base_path / str(year)
	available_files = sorted([p.name for p in year_folder.glob("*.xlsx")])
	print(f"\nAvailable files in {year_folder.name}: {available_files}")

	# read out the data sources for the last available year and put all into one dataframe, where "Anzahl" is replaced by the original file name
	data = pd.DataFrame()
	for file_name in available_files:
		file_path = year_folder / file_name
		column_name = file_name.split(".")[0]
		xls = pd.ExcelFile(file_path)
		for sheet_name in xls.sheet_names:
			df = pd.read_excel(xls, sheet_name=sheet_name)
			df.set_index(["Ausweisnummer", "Vorname", "Nachname"], inplace=True)
			# ensure that the index is unique, otherwise we cannot merge the dataframes
			if not df.index.is_unique:
				print(f"Warning: Index is not unique in {file_name} - {sheet_name}. Duplicates will be dropped.")
				df = df[~df.index.duplicated(keep="first")]
			# replace "Anzahl" with the original file name
			df = df.rename(columns={"Anzahl": column_name})
			# remove all colums except "Ausweisnummer", "Vorname", "Nachname"
			df = df[[column_name]]
			# ensure that the column is integer
			df[column_name] = df[column_name].astype(int)
			# add to the main dataframe to the correct person
			data = pd.concat([data, df], axis=1)

	# sort the columns and the data by the types
	# make sure the data has all the types, if not, add them with missing data
	for t in types:
		if t not in data.columns:
			data[t] = None
	data = data[types + [col for col in data.columns if col not in types]]
	data = data.sort_values(by=types, ascending=False)
	# save that file for that year into the parent folder
	output_file = base_path / f"Beteiligung {year_folder.name}.xlsx"
	data.to_excel(output_file)



# now summarize every year and show the total for each type, and save that into a new file
summary_data = pd.DataFrame()
for year in available_years[::-1]:  # reverse order to have the latest year first
	year_folder = base_path / str(year)
	summary_file = base_path / f"Beteiligung {year_folder.name}.xlsx"
	if not summary_file.exists():
		print(f"Warning: Summary file for {year} does not exist. Skipping.")
		continue
	df = pd.read_excel(summary_file, "Sheet1")
	df.set_index(["Ausweisnummer", "Vorname", "Nachname"], inplace=True)
	# combine the data and use MultiIndex to keep the year as a column
	# first add the multiindex
	df.columns = pd.MultiIndex.from_product([[year], df.columns])
	# then add the data to the summary dataframe, if the summary dataframe is empty, just assign it, otherwise concat
	summary_data = pd.concat([summary_data, df], axis=1)

# save the summary data to a new file
summary_output_file = base_path / "Beteiligung Zusammenfassung.xlsx"
# use xlsxwriter to format the output file, so that the first row is bold and the columns are auto-sized and the values are marked red if the criteria is not reached
with pd.ExcelWriter(summary_output_file, engine='xlsxwriter') as writer:
	summary_data.to_excel(writer, sheet_name="Zusammenfassung")
	workbook = writer.book
	worksheet = writer.sheets["Zusammenfassung"]
	# format the first row as bold and centered
	format_bold_centered = workbook.add_format({'bold': True, 'align': 'center'})
	worksheet.set_row(0, None, format_bold_centered)
	# mark the values in red if the criteria is not reached, i.e. if the value is less than 5 for Einsätze, less than 6 for Übungen, less than 3 for Dienste and less than 2 for Übungstour
	red_format = workbook.add_format({'bg_color': 'red'})
	green_format = workbook.add_format({'bg_color': 'green'})
	number_of_rows = len(summary_data)+2
	number_of_cols = len(available_years)*len(available_files)+2
	for criterium in criteria:
		for year in available_years:
			# index the year and the criterium
			col_index_year = 3 + list(summary_data.columns.get_level_values(0)).index(year)
			col_index = col_index_year + list(summary_data.columns.get_level_values(1)).index(criterium)
			worksheet.conditional_format(3, col_index, number_of_rows, col_index, # first row, first col, last row, last col
			{'type':     'cell',
				'criteria': 'less than',
				'value':    criteria[criterium],
				'format':   red_format}
				)
			# all others are green
			worksheet.conditional_format(3, col_index, number_of_rows, col_index, # first row, first col, last row, last col
			{'type':     'cell',
				'criteria': 'greater than or equal to',
				'value':    criteria[criterium],
				'format':   green_format}
				)
			# print the criteria right to the tabel for documentation
			worksheet.write(2, col_index, f"Min: {criteria[criterium]}")

