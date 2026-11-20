from pathlib import Path
import pandas as pd

# Set this to your Excel directory.
folder = r"C:\Users\gregor.pobegen\OneDrive - Österreichischer Bergrettungsdienst Landesorganisation Kärnten\Dateien von IT - Bergrettung - OS_Villach\Beteiligung"
types = ["Einsätze", "Übungen", "Dienste", "Übungstour"] # Sicherheitstage, 
summary_output_file_name = "Beteiligung Zusammenfassung.xlsx"
criteria = { # mindestens
	"Einsätze": 5,
	"Übungen": 6,
	"Dienste": 0,
	"Übungstour": 0,
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

# determine the Ampel (traffic light) for every participant and every year:
#   green:  criteria met in that year
#   orange: criteria not met in that year, but met in the previous year
#   red:    criteria not met in that year and not met in the previous year
# for the first year there is no previous year, so it can only be green or orange
summary_years = sorted(summary_data.columns.get_level_values(0).unique())
ampel_by_year = pd.DataFrame(index=summary_data.index)
previous_met = pd.Series(True, index=summary_data.index)  # first year: benefit of the doubt
for year in summary_years:
	year_df = summary_data.xs(year, axis=1, level=0, drop_level=True)
	# missing values mean no participation, i.e. 0
	met = year_df[list(criteria)].fillna(0).ge(pd.Series(criteria)).all(axis=1)
	status = pd.Series("red", index=summary_data.index, dtype="object")
	status[~met & previous_met] = "orange"
	status[met] = "green"
	ampel_by_year[year] = status
	previous_met = met
ampel_by_year = ampel_by_year[summary_years[::-1]]  # latest year first, like the summary
latest_year = summary_years[-1] if summary_years else None
status_by_person = ampel_by_year[latest_year] if latest_year is not None else pd.Series("red", index=summary_data.index, dtype="object")

ampel_labels = {"green": "grün", "orange": "orange", "red": "rot"}
ampel_colors = {"green": "#90EE90", "orange": "#FFD580", "red": "#FF7F7F"}

# save the summary data to a new file
# every year is a block (Ampel first, then the types) with a border around it and an empty column in between
summary_output_file = base_path / summary_output_file_name
with pd.ExcelWriter(summary_output_file, engine='xlsxwriter') as writer:
	workbook = writer.book
	worksheet = workbook.add_worksheet("Zusammenfassung")

	# xlsxwriter needs one format object per combination of properties, so cache them
	format_cache = {}
	def fmt(**props):
		key = tuple(sorted(props.items()))
		if key not in format_cache:
			format_cache[key] = workbook.add_format(props)
		return format_cache[key]

	def write_value(sheet, row, col, value, cell_format):
		if pd.isna(value):
			sheet.write_blank(row, col, None, cell_format)
		else:
			sheet.write(row, col, value, cell_format)

	header_rows = 3  # row 0: year, row 1: column names, row 2: minimum criteria
	last_row = header_rows + len(summary_data) - 1
	index_names = list(summary_data.index.names)

	# names, colored by the Ampel of the latest year
	for col, name in enumerate(index_names):
		worksheet.write(header_rows - 1, col, name, fmt(bold=True))
	for row, (person, status) in enumerate(zip(summary_data.index, status_by_person.values), start=header_rows):
		for col, value in enumerate(person):
			write_value(worksheet, row, col, value, fmt(bg_color=ampel_colors[status]))
	worksheet.set_column(0, len(index_names) - 1, 15)

	col = len(index_names)
	for year in summary_years[::-1]:  # latest year first
		year_df = summary_data.xs(year, axis=1, level=0, drop_level=True)
		block_columns = ["Ampel"] + list(year_df.columns)
		first_col, last_col = col, col + len(block_columns) - 1

		def border(row, col):
			b = {}
			if row == 0: b["top"] = 2
			if row == last_row: b["bottom"] = 2
			if col == first_col: b["left"] = 2
			if col == last_col: b["right"] = 2
			return b

		# header: year (merged over the block), column names, minimum criteria
		worksheet.merge_range(0, first_col, 0, last_col, year, fmt(bold=True, align="center", top=2, left=2, right=2))
		for offset, column_name in enumerate(block_columns):
			c = first_col + offset
			worksheet.write(1, c, column_name, fmt(bold=True, align="center", **border(1, c)))
			min_text = f"Min: {criteria[column_name]}" if column_name in criteria else None
			write_value(worksheet, 2, c, min_text, fmt(align="center", **border(2, c)))

		# data
		for row_offset, person in enumerate(summary_data.index):
			row = header_rows + row_offset
			status = ampel_by_year.at[person, year]
			worksheet.write(row, first_col, ampel_labels[status], fmt(bg_color=ampel_colors[status], align="center", **border(row, first_col)))
			for offset, column_name in enumerate(year_df.columns, start=1):
				c = first_col + offset
				value = year_df.at[person, column_name]
				# missing values mean no participation, i.e. 0
				value = 0 if pd.isna(value) else int(value)
				props = border(row, c)
				props["align"] = "center"
				if column_name in criteria:
					props["bg_color"] = "green" if value >= criteria[column_name] else "red"
				worksheet.write(row, c, value, fmt(**props))

		worksheet.set_column(first_col, last_col, 11)
		worksheet.set_column(last_col + 1, last_col + 1, 3)  # empty column between the years
		col = last_col + 2

	worksheet.freeze_panes(header_rows, len(index_names))

	# separate sheet with only the Ampel over the years (latest year first)
	ampel_by_year.map(ampel_labels.get).to_excel(writer, sheet_name="Ampel")
	ampel_sheet = writer.sheets["Ampel"]
	for col_offset, year in enumerate(ampel_by_year.columns):
		for row, status in enumerate(ampel_by_year[year].values, start=1):
			ampel_sheet.write(row, len(index_names) + col_offset, ampel_labels[status], fmt(bg_color=ampel_colors[status], align="center"))
	ampel_sheet.set_column(0, len(index_names) - 1, 15)
