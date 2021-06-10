"""
#######################################################################################################################
###											PSSE Contingency Test													###
###		Script to deal with the importing and processing of files													###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JI7867- EirGrid - Capital Project 966																###
###																													###
#######################################################################################################################
"""

import logging
import pandas as pd
import os
import string

import optimisation.constants as constants


def colnum_string(n):
	string = ""
	while n > 0:
		n, remainder = divmod(n - 1, 26)
		string = chr(65 + remainder) + string
	return string


def colstring_number(col):
	num = 0
	for c in col:
		if c in string.ascii_letters:
			num = num * 26 + (ord(c.upper()) - ord('A')) + 1
	return num


class ImportContingencies:

	def __init__(self, pth):
		"""
			Initialises the class and imports the complete set of contingencies
			Contingencies are contained in an excel workbook split across a different worksheet to identify
			circuits, transformers 2 winding, transformers 3 winding, busbars
		:param str pth: Path to the file that needs importing
		"""
		self.circuits = pd.DataFrame()
		self.tx2 = pd.DataFrame()
		self.tx3 = pd.DataFrame()
		self.busbars = pd.DataFrame()
		self.fixed_shunts = pd.DataFrame()
		self.switched_shunts = pd.DataFrame()
		self.contingency_names = set()

		self.pth = pth
		self.c = constants.Excel

		self.logger = logging.getLogger(constants.Logging.logger_name)

		self.import_workbook()

	def import_workbook(self):
		"""
			Function to import the workbook and deal with all the processing
		:return:
		"""
		# xl = pd.ExcelFile(self.pth)
		# Get circuit contingencies
		self.circuits = pd.read_excel(
			self.pth, sheet_name=self.c.circuit, header=0,
			names=self.c.columns[self.c.circuit]
		)
		self.circuits[self.c.id] = self.circuits[self.c.id].apply(lambda x: '{0:.0f}'.format(x) if isinstance(x, float) else x)

		# Get 2 winding transformer contingencies
		self.tx2 = pd.read_excel(self.pth, sheet_name=self.c.tx2, header=0,
								 names=self.c.columns[self.c.tx2])
		self.tx2[self.c.id] = self.tx2[self.c.id].apply(lambda x: '{0:.0f}'.format(x) if isinstance(x, float) else x)

		# Get 3 winding transformer contingencies
		self.tx3 = pd.read_excel(self.pth, sheet_name=self.c.tx3, header=0,
								 names=self.c.columns[self.c.tx3])
		self.tx3[self.c.id] = self.tx3[self.c.id].apply(lambda x: '{0:.0f}'.format(x) if isinstance(x, float) else x)

		# Get busbar contingencies
		self.busbars = pd.read_excel(self.pth, sheet_name=self.c.busbars, header=0,
									 names=self.c.columns[self.c.busbars])

		# Get shunt contingencies
		self.fixed_shunts = pd.read_excel(self.pth, sheet_name=self.c.fixed_shunts, header=0,
										  names=self.c.columns[self.c.fixed_shunts])
		self.fixed_shunts[self.c.id] = self.fixed_shunts[self.c.id].apply(lambda x: '{0:.0f}'.format(x) if isinstance(x, float) else x)

		self.switched_shunts = pd.read_excel(self.pth, sheet_name=self.c.switched_shunts, header=0,
											 names=self.c.columns[self.c.switched_shunts])
		self.switched_shunts[self.c.id] = self.switched_shunts[self.c.id].apply(lambda x: '{0:.0f}'.format(x) if isinstance(x, float) else x)

		# Extract unique list of all contingencies being considered
		all_cont_names = [x[constants.Contingency.header] for x in [
			self.circuits,
			self.tx2,
			self.tx3,
			self.busbars,
			self.fixed_shunts,
			self.switched_shunts
		]]

		# Combine into a single series and then extract using set to drop duplicates sorted into alphabetic order
		self.contingency_names = sorted(set(pd.concat(all_cont_names)))

		if not self.contingency_names:
			self.logger.error(('The contingencies workbook: {} is empty and therefore there'
							   'are not continencies to test.  User entry error!')
							  .format(self.pth))
			raise ValueError('Empty Contingencies Workbook')

	def group_contingencies_by_name(self, cont_name):
		"""
			Groups the contingencies into a dictionary based on contingency name where
			the dictionary contains a subset of the main DataFrames
		:return (pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()),
				(circuit_data, tx2_data, tx3_data, busbar_data):
		"""
		column_header = constants.Contingency.header

		circuit_data = self.circuits[self.circuits[column_header] == cont_name]
		tx2_data = self.tx2[self.tx2[column_header] == cont_name]
		tx3_data = self.tx3[self.tx3[column_header] == cont_name]
		busbars = self.busbars[self.busbars[column_header] == cont_name]
		fixed_shunts = self.fixed_shunts[self.fixed_shunts[column_header] == cont_name]
		switched_shunts = self.switched_shunts[self.switched_shunts[column_header] == cont_name]

		return circuit_data, tx2_data, tx3_data, busbars, fixed_shunts, switched_shunts


class ExportResults:
	"""
		Class to deal with the processing and exporting of results to excel
	"""

	def __init__(self, pth_workbook, results, convergence):
		"""
			Initialise
		:param str pth_workbook: Full path for workbook to be exported
		:param collections.OrderedDict results: Dictionary in the format {sheet:pd.DataFrame} of results to be exported
		:param pd.DataFrame convergence: Dictionary in the format {contingency:(message, convergence)} of results to be exported
		"""

		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.pth = pth_workbook

		if os.path.exists(self.pth):
			self.logger.info('Target workbook path {} already exists and will be overwritten'.format(self.pth))
			try:
				os.remove(self.pth)
			except WindowsError:
				self.logger.critical('Unable to overwrite file {} since it is currently being used, please close'
									 .format(self.pth))
				raise WindowsError('The process cannot access the file because it is being used by another process: {}'
								   .format(self.pth))

		# Write all data to same workbook on different sheets
		with pd.ExcelWriter(path=self.pth, engine='xlsxwriter') as excel_writer:
			# Set format for conditional formatting
			book = excel_writer.book
			format_error = book.add_format(constants.Excel.cell_format_error)
			format_change = book.add_format(constants.Excel.cell_format_change)

			# Write convergence DataFrame
			convergence.to_excel(excel_writer, sheet_name=constants.Excel.convergence)

			for sht, df in results.iteritems():
				df.sort_index(axis=0, ascending=True, inplace=True)
				df.to_excel(excel_writer, sheet_name=sht)
				# Also writes copy of data in transposed state so can filter for non-compliance
				# Does not include conditional formatting
				df.T.to_excel(excel_writer, sheet_name='{}_T'.format(sht))

				# Get conditional formatting when DataFrame isn't empty
				if not df.empty:
					data_range, criteria = self.get_conditional_formatting(
						df=df, cell_format_error=format_error, cell_format_change=format_change
					)
					if data_range:
						worksheet = excel_writer.sheets[sht]
						worksheet.conditional_format(data_range, criteria)

	def get_conditional_formatting(self, df, cell_format_error, cell_format_change):
		"""
			Function returns the conditional formatting that needs to be added
		:param pd.DataFrame df: DataFrame that contains the data that should have conditional formatting added
		:param xlsxwriter.format cell_format: format class to be used for the conditional formatting
		:return tuple conditional_formatting:
		"""
		# Get the start column
		start_col = colnum_string(df.columns.get_loc(constants.Contingency.bc) + 2)
		end_col = colnum_string(len(df.columns) + 1)
		start_row = df.columns.nlevels + 1
		end_row = len(df.index)

		# Determine reference column
		if constants.Contingency.rate_for_checking in df.columns:
			# Circuit loading
			ref_column = colnum_string(df.columns.get_loc(constants.Contingency.rate_for_checking) + 2)
			criteria = '=and({}{}<>"",${}{}>{},{}{}>${}{})'.format(start_col, start_row,
																   ref_column, start_row,
																   constants.EirGridThresholds.rating_threshold,
																   start_col, start_row, ref_column, start_row)
			# Select the required cell formatting for this type of conditional formatting
			selected_format = cell_format_error
		elif constants.Busbars.lower_limit in df.columns:
			# Voltage range checking
			ref_column1 = colnum_string(df.columns.get_loc(constants.Busbars.lower_limit) + 2)
			ref_column2 = colnum_string(df.columns.get_loc(constants.Busbars.upper_limit) + 2)
			criteria = '=and({}{}<>"", or({}{}<${}{}, {}{}>${}{}))'.format(start_col, start_row,
																		   start_col, start_row, ref_column1, start_row,
																		   start_col, start_row, ref_column2, start_row)
			# Select the required cell formatting for this type of conditional formatting
			selected_format = cell_format_error
		elif constants.Contingency.v_step_lbl in df.index:
			# No need to include the BASE CASE so move 1 column along
			start_col = colnum_string(colstring_number(start_col) + 1)
			# Voltage step limit check
			ref_row = df.index.get_loc(constants.Contingency.v_step_lbl) + 2
			ref_column = colnum_string(df.columns.get_loc(constants.Contingency.bc) + 2)
			criteria = '=and({}{}<>"",{}{}>{}${})'.format(start_col, start_row,
														  start_col, start_row, start_col, ref_row)
			# Select the required cell formatting for this type of conditional formatting
			selected_format = cell_format_error
			end_row = ref_row - 2
		else:
			# The remainder relate to status changes
			ref_column = colnum_string(df.columns.get_loc(constants.Contingency.bc) + 2)
			criteria = 'and({}{}<>"",{}{}<>${}{})'.format(start_col, start_row,
														  start_col, start_row, ref_column, start_row)
			# Select the required cell formatting for this type of conditional formatting
			selected_format = cell_format_change
			end_row = len(df.index) + 1

		data_range = '{}{}:{}{}'.format(start_col, start_row, end_col, end_row)
		conditional_formatting = {'type': 'formula',
								  'criteria': criteria,
								  'format': selected_format}
		return data_range, conditional_formatting


def busbars_to_consider(pth_busbar_list):
	"""
		Function returns a list of busbar numbers to be considered
	:param str pth_busbar_list:  Path to file which contains list of busbars to consider
	:return (pd.DataFrame, pd.Index), (df_busbars, busbars_to_keep):  List of busbars to be kept
	:return:
	"""

	# Get list of busbars to be plotted
	df_busbars = pd.read_excel(pth_busbar_list, sheet_name='Busbars', index_col=0, header=0)

	# Reduce dataframe to only be those which are labelled as keep
	df_busbars = df_busbars.loc[df_busbars['Include'] == 1]

	# Sort into order for plotting
	df_busbars.sort_values(by=['Nominal', 'Plot Name'], ascending=[False, True], inplace=True)

	# Extract index for busbars that should be plotted
	busbars_to_keep = df_busbars.index

	return df_busbars, busbars_to_keep
