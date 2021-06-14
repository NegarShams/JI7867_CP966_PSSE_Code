"""
#######################################################################################################################
###											psse.py																	###
###		Deals with all interactions relating to PSSE																###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK 	 		###
###		project Jk7261 WPD Virtual Statcom project																	###
###																													###
#######################################################################################################################
"""

# Project specific imports
import optimisation.constants as constants

# Generic python package imports
import sys
import os
import logging
import numpy as np
import pandas as pd


class InitialisePsspy:
	"""
		Class to deal with the initialising of PSSE by checking the correct directory is being referenced and has been
		added to the system path and then attempts to initialise it
	"""
	def __init__(self, psse_version=constants.PSSE.version):
		"""
			Intialises the paths and checks that import psspy works
		:param int psse_version: (optional=34)
		"""

		self.psse = False

		# Get PSSE path
		self.psse_py_path, self.psse_os_path = constants.PSSE().get_psse_path(psse_version=psse_version)
		# Add to system path if not already there
		if self.psse_py_path not in sys.path:
			sys.path.append(self.psse_py_path)

		if self.psse_os_path not in os.environ['PATH']:
			os.environ['PATH'] += ';{}'.format(self.psse_os_path)

		if self.psse_py_path not in os.environ['PATH']:
			os.environ['PATH'] += ';{}'.format(self.psse_py_path)

		global psspy
		global pssarrays
		try:
			import psspy
			self.psspy = psspy
			import pssarrays
			self.pssarrays = pssarrays
		except ImportError:
			self.psspy = None
			self.pssarrays = None

	def initialise_psse(self):
		"""
			Initialise PSSE
		:return bool self.psse: True / False depending on success of initialising PSSE
		"""
		if self.psse is True:
			pass
		else:
			error_code = self.psspy.psseinit()

			if error_code != 0:
				self.psse = False
				raise RuntimeError('Unable to initialise PSSE, error code {} returned'.format(error_code))
			else:
				self.psse = True
				# Disable screen output based on PSSE constants
				self.change_output(destination=constants.PSSE.output[constants.DEBUG_MODE])

		return self.psse

	def change_output(self, destination):
		"""
			Function disables the reporting output from PSSE
		:param int destination:  Target destination, default is to disable which sets it to 6
		:return None:
		"""
		print('PSSE output set to: {}'.format(destination))

		# Disables all PSSE output
		_ = self.psspy.report_output(islct=destination)
		_ = self.psspy.progress_output(islct=destination)
		_ = self.psspy.alert_output(islct=destination)
		_ = self.psspy.prompt_output(islct=destination)

		return None


class BranchData:
	"""
		Branch data
	"""
	def __init__(self, flag, tx=False, sid=-1):
		"""

		:param int flag: 2 for branches, 6 for two-winding transformers
		:param bool tx: (optional=False) - If set to True then this BranchData is for 2 winding transformers and
											therefore different functions used for switching status
		"""
		# Initialise empty variables
		self.df_status = pd.DataFrame()
		self.df_loading = pd.DataFrame()

		# constants
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.c = constants.Branches

		# So know whether refering to a transformer or a circuit, transformers return additional data
		self.tx = tx

		# Function for changing status
		if self.tx:
			self.func_switch = psspy.two_winding_chng_4
			# If wrong flag value provided for this type of data retrieval then update
			if flag not in (5, 6):
				new_flag = 6
				self.logger.warning(('The class <BranchData> has been called with tx set to {} and the flag set to {}. '
									 'These are not compatible and so instead the flag has been set to {}')
									.format(tx, flag, new_flag))
				flag = new_flag
		else:
			# Errors when using psspy.branch_chng_3
			self.func_switch = psspy.branch_chng
			# If wrong flag value provided for this type of data retrieval then update
			if flag not in (1, 2):
				new_flag = 2
				self.logger.warning(('The class <BranchData> has been called with tx set to {} and the flag set to {}. '
									 'These are not compatible and so instead the flag has been set to {}')
									.format(tx, flag, new_flag))
				flag = new_flag
		self.flag = flag
		self.sid = sid

		# Update DataFrame from PSSE case
		self.update()

	def update(self, cont_name=None):
		"""
			Updates data from SAV case
		:param str cont_name:  Contingency name for these results
		"""
		# Declare functions
		func_int = psspy.abrnint
		func_char = psspy.abrnchar
		func_real = psspy.abrnreal

		# Function input constants
		entry = 2  # Double entry
		ties = 3 # Include interior and exterior branches

		# Retrieve data from PSSE
		ierr_int, iarray = func_int(
			sid=self.sid,
			flag=self.flag,
			entry=entry,
			ties=ties,
			string=(self.c.from_bus,
					self.c.to_bus,
					self.c.status))
		ierr_char, carray = func_char(
			sid=self.sid,
			flag=self.flag,
			entry=entry,
			ties=ties,
			string=(self.c.id,))
		ierr_real, rarray = func_real(
			sid=self.sid,
			flag=self.flag,
			entry=entry,
			ties=ties,
			string=(self.c.ratea, self.c.rateb, self.c.ratec, self.c.loading))
		if ierr_int > 0 or ierr_char > 0 or ierr_real > 0:
			self.logger.error(('Unable to retrieve the branch data from the SAV case and the following error codes '
							   'were returned: {}, {} and {} from functions <{}>, <{}> and <{}>')
							  .format(ierr_int, ierr_char, ierr_real,
									  func_int.__name__, func_char.__name__, func_real.__name__))
			raise ValueError('Error importing data from PSSE SAV case')

		# Combine data into single list of lists
		# Populate the complete DataFrame, overwriting any new values
		data = iarray + carray + rarray
		# Column headers initially in same order as data but then reordered to something more useful for exporting
		# in case needed
		initial_columns = [self.c.from_bus, self.c.to_bus, self.c.status, self.c.id,
						   self.c.ratea, self.c.rateb, self.c.ratec, self.c.loading]
		# Transposed so columns in correct location and then columns reordered to something more suitable
		complete_df = pd.DataFrame(data).transpose()
		complete_df.columns = initial_columns
		# Remove empty string in ID column (PSSE will return '1' as '1 ' but will accept '1' as an input
		complete_df[self.c.id] = complete_df[self.c.id].str.strip()

		if cont_name is None:
			# Extract data for status of branch
			columns_for_status = [self.c.from_bus, self.c.to_bus, self.c.id, self.c.status]
			self.df_status = complete_df[columns_for_status]
			# Add column for base case
			self.df_status[constants.Contingency.bc] = self.df_status[self.c.status]

			# Extract data for the circuit loading and specified ratings
			columns_for_rating = [self.c.from_bus, self.c.to_bus, self.c.id, self.c.ratea, self.c.rateb,
								  self.c.ratec, self.c.loading]
			self.df_loading = complete_df[columns_for_rating]
			# Add column for base case
			self.df_loading[constants.Contingency.bc] = self.df_loading[self.c.loading]

		else:
			# Add new column to DataFrame with the status of the circuit during this contingency
			self.df_status[cont_name] = complete_df[self.c.status]
			self.df_loading[cont_name] = complete_df[self.c.loading]

	def change_state(self, asset=pd.Series(), restore=False, restore_all=False):
		"""
			Switches the status of the branch for a specific contingency and updates the branch status
		:param pd.Series asset: (optional=pd.Series()) Branch whose status is to be changed
		:param bool restore: (optional=False) - if set to True then will restore single branch back to original value
		:param bool restore_all: (optional=False) - if set to True then will restore every branch back to original value
		:return bool success: Whether successfully switched circuit or not
		"""
		# check if cont_name already in DataFrame and if not add it populated with the initial status values
		# #if cont_name not in self.df_status.columns:
		# #	self.df_status[cont_name] = self.df_status[self.c.status]
		success = False

		if restore_all:
			# Loop through and restore every asset to its original status
			for i, row in self.df_status.iterrows():
				# TODO: This may be quite slow so could be useful to find out current status and only switch those
				# TODO: which are actually different

				# TODO: Currently will loop through every asset since they are duplicated in direction, should detect this
				bus1 = row[self.c.from_bus]
				bus2 = row[self.c.to_bus]
				ckt = row[self.c.id]
				status = row[self.c.status]
				success = self.switch(bus1, bus2, ckt, status)

		else:
			bus1 = asset[self.c.from_bus]
			bus2 = asset[self.c.to_bus]
			ckt = asset[self.c.id]
			if restore:
				# Get original status value
				status = self.df_status.loc[
					(self.df_status[self.c.from_bus] == bus1) &
					(self.df_status[self.c.to_bus] == bus2) &
					(self.df_status[self.c.id] == ckt),
					self.c.status].item()
			else:
				status = asset[self.c.status]

			# Just change status of selected asset updating of column in DataFrame comes once all updated
			success = self.switch(
				bus1=bus1,
				bus2=bus2,
				ckt=ckt,
				status=status
			)

		return success

	def switch(self, bus1, bus2, ckt, status):
		"""
			Actually switches the branch
		:param bus1: From busbar number
		:param bus2: To busbar number
		:param ckt: ID for circuit
		:param status: New status
		:return bool success:
		"""
		ierr = self.func_switch(i=int(bus1),
								j=int(bus2),
								ckt=str(ckt),
								intgar1=int(status))
		# If a transformer then the function to edit will return additional data along with the error code.
		# This needs to be extracted before error checking
		if self.tx:
			ierr = ierr[0]

		if ierr > 0:
			success = False
			self.logger.error(('Error code {} raised when calling function {} for branch {}-{}-{} to '
							   'change status to {}')
							  .format(ierr, self.func_switch.__name__, int(bus1), int(bus2), str(ckt), status))
		else:
			success = True
		return success

	def check_compliance(self, cont_names):
		"""
			Function inserts a row at the top of the DataFrame which confirms whether the
			loadings are all withing limits
		:param list cont_names:  List of all the contigencies and therefore those which should be considered for the
								validation
		:return None:
		"""
		# Get list of transformers which do not have any winding data
		items_with_no_rating = self.df_loading[self.df_loading[constants.Contingency.rate_for_checking] <=
											   constants.EirGridThresholds.rating_threshold]
		if not items_with_no_rating.empty:
			if self.tx:
				asset = '2 winding transformer'
			else:
				asset = 'circuit'
			msg0 = (('The following {} have a rating of less than or equal to {:.2f} MVA and have '
					 'therefore been ignored from checking for compliance')
				.format(asset, constants.EirGridThresholds.rating_threshold))
			msg1 = '\n'.join(['\t - {} connected between busbars {} and {} with ID {}'
							 .format(asset.capitalize(), row[self.c.from_bus], row[self.c.to_bus], row[self.c.id])
							  for _, row in items_with_no_rating.iterrows()])
			self.logger.warning('{}\n{}'.format(msg0, msg1))

		# Steady state validation
		compliance = dict()

		# Loop through each contingency and check if all steady state circuit loadings are within limits
		df_loading = self.df_loading
		df_loading.loc[:, 'Threshold'] = constants.EirGridThresholds.rating_threshold

		for cont in cont_names:
			# Check to see if loading is less than rating where rating is > the EirGrid threshold value
			condition = ((self.df_loading[cont] < self.df_loading[constants.Contingency.rate_for_checking]) |
						 (constants.EirGridThresholds.rating_threshold ==
						  df_loading[constants.Contingency.rate_for_checking]))

			# Check compliant at all busbars
			compliance[cont] = all(np.where(condition, True, False))

		# Update DataFrame with new row to show whether compliant
		for key, value in compliance.iteritems():
			self.df_loading.loc[constants.Contingency.compliant, key] = value

		# Return slice from DataFrame showing whether this dataframe is compliant
		if self.tx:
			col_title = constants.Excel.tx2_loading
		else:
			col_title = constants.Excel.circuit_loading
		df_compliance = pd.DataFrame(index=cont_names, columns=(col_title,))
		df_compliance[col_title] = self.df_loading.loc[constants.Contingency.compliant,
													   cont_names]

		return df_compliance


class ShuntData:
	"""
		Shunt data
	"""
	def __init__(self, flag=4, sid=-1, fixed=True):
		"""
		:param int flag: 4 for all fixed bus shunts
		"""
		# Initialise empty variables
		self.df_status = pd.DataFrame()

		# constants
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.c = constants.Shunts

		# Whether a fixed shunt or a switched shunt these are the functions for switching them
		self.fixed = fixed
		if self.fixed:
			self.func_int = psspy.afxshuntint
			self.func_char = psspy.afxshuntchar
			self.func_switch = psspy.shunt_chng
		else:
			self.func_int = psspy.aswshint
			self.func_char = psspy.aswshchar
			self.func_switch = psspy.switched_shunt_chng_3
		self.flag = flag
		self.sid = sid

		# Update DataFrame from PSSE case
		self.update()

	def update(self, cont_name=None):
		"""
			Updates data from SAV case
		:param str cont_name:  Contingency name for these results
		"""

		# Retrieve data from PSSE
		ierr_int, iarray = self.func_int(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus,
					self.c.status))
		if self.fixed:
			str_identifier = self.c.id
		else:
			str_identifier = 'NAME'
		ierr_char, carray = self.func_char(
			sid=self.sid,
			flag=self.flag,
			string=(str_identifier,))

		if ierr_int > 0 or ierr_char > 0:
			if self.fixed:
				msg = 'fixed'
			else:
				msg = 'switched'
			self.logger.error(('Unable to retrieve the {} shunt data from the SAV case and the following error codes '
							   'were returned: {} and {} from functions <{}> and <{}>')
							  .format(msg,
									  ierr_int, ierr_char,
									  self.func_int.__name__, self.func_char.__name__))
			raise ValueError('Error importing data from PSSE SAV case')

		# Combine data into single list of lists
		# Populate the complete DataFrame, overwriting any new values
		data = iarray + carray
		# Column headers initially in same order as data but then reordered to something more useful for exporting
		# in case needed
		initial_columns = [self.c.bus, self.c.status, self.c.id]
		# Transposed so columns in correct location and then columns reordered to something more suitable
		complete_df = pd.DataFrame(data).transpose()
		complete_df.columns = initial_columns
		# Remove empty string in ID column (PSSE will return '1' as '1 ' but will accept '1' as an input
		if not complete_df.empty:
			complete_df[self.c.id] = complete_df[self.c.id].str.strip()
		else:
			return

		if cont_name is None:
			# Extract data for status of branch
			columns_for_status = [self.c.bus, self.c.id, self.c.status]
			self.df_status = complete_df[columns_for_status]
			# Add column for base case results
			self.df_status[constants.Contingency.bc] = self.df_status[self.c.status]

		else:
			# Add new column to DataFrame with the status of the circuit during this contingency
			self.df_status[cont_name] = complete_df[self.c.status]

	def change_state(self, asset=pd.Series(),
					 restore=False, restore_all=False):
		"""
			Switches the status of the branch for a specific contingency and updates the branch status
		:param pd.Series asset: (optional=pd.Series()) Branch whose status is to be changed
		:param bool restore: (optional=False) - if set to True then will restore single branch back to original value
		:param bool restore_all: (optional=False) - if set to True then will restore every branch back to original value
		:return bool success: Whether change of state worked correctly or not
		"""
		# check if cont_name already in DataFrame and if not add it populated with the initial status values
		# #if cont_name not in self.df_status.columns:
		# #	self.df_status[cont_name] = self.df_status[self.c.status]
		success = False

		if restore_all:
			# Loop through and restore every asset to its original status
			for i, row in self.df_status.iterrows():
				# TODO: This may be quite slow so could be useful to find out current status and only switch those
				# TODO: which are actually different

				# TODO: Currently will loop through every asset since they are duplicated in direction, should detect this
				bus = row[self.c.status]
				ckt = row[self.c.id]
				status = row[self.c.status]
				success = self.switch(bus, ckt, status)

		else:
			bus = asset[self.c.bus]
			ckt = asset[self.c.id]
			if restore:
				# Get original status value
				status = self.df_status.loc[
					(self.df_status[self.c.bus] == bus) &
					(self.df_status[self.c.id] == ckt),
					self.c.status].item()
			else:
				status = asset[self.c.status]

			# Just change status of selected asset updating of column in DataFrame comes once all updated
			success = self.switch(
				bus=bus,
				ckt=ckt,
				status=status
			)

		return success

	def switch(self, bus, ckt, status):
		"""
			Actually switches the branch
		:param bus: Busbar number
		:param ckt: ID for circuit
		:param status: New status
		:return int new_status:
		"""
		if self.fixed:
			ierr = self.func_switch(i=int(bus),
									id=str(ckt),
									intgar1=int(status))
		else:
			ierr = self.func_switch(i=int(bus),
									intgar11=int(status))

		# Check no errors have occurs
		if ierr > 0:
			success = False
			if self.fixed:
				msg = 'fixed'
			else:
				msg = 'switched'
			self.logger.error(('Error code {} raised when calling function {} for {} shunt at busbar {} with '
							   'ID {} to change status to {}')
							  .format(ierr, self.func_switch.__name__, msg, bus, ckt, status))
		else:
			success = True
		return success


class Tx3WndData:
	"""
		For the 3 phase transformers ratings are determined on a winding basis rather than total transformer.  It
		should generally be winding 1 that is the limiting factor
	"""
	def __init__(self, flag=3, sid=-1):
		"""
		:param int flag: 3 for all windings of 3 winding transformers
		"""
		# Initialise empty DataFrames which will contain the loading data for each contingency
		self.df_status = pd.DataFrame()
		self.df_loading = pd.DataFrame()

		# constants
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.c = constants.Tx3

		self.flag = flag
		self.sid = sid

		# Update DataFrame from PSSE case
		self.update()

	def update(self, cont_name=None):
		"""
			Updates data from SAV case
		:param str cont_name:  If provided will add the updated data to an existing dataframe
		"""
		# Declare functions
		func_int = psspy.awndint
		func_char = psspy.awndchar
		func_real = psspy.awndreal

		# Function input constants
		entry = 1
		ties = 3

		# Retrieve data from PSSE
		ierr_int, iarray = func_int(
			sid=self.sid,
			flag=self.flag,
			entry=entry,
			ties=ties,
			string=(self.c.this_winding,
					self.c.wind1,
					self.c.wind2,
					self.c.wind3,
					self.c.status))
		ierr_char, carray = func_char(
			sid=self.sid,
			flag=self.flag,
			entry=entry,
			ties=ties,
			string=(self.c.id,))
		ierr_real, rarray = func_real(
			sid=self.sid,
			flag=self.flag,
			entry=entry,
			ties=ties,
			string=(self.c.ratea,
					self.c.rateb,
					self.c.ratec,
					self.c.loading))
		if ierr_int > 0 or ierr_char > 0 or ierr_real > 0:
			self.logger.error(('Unable to retrieve the area data from the SAV case and the following error codes '
							   'were returned: {}, {} and {} from functions <{}>, <{}> and <{}>')
							  .format(ierr_int, ierr_char, ierr_real,
									  func_int.__name__, func_char.__name__, func_real.__name__))
			raise ValueError('Error importing data from PSSE SAV case')

		# Combine data into single list of lists
		data = iarray + carray + rarray
		# Column headers initially in same order as data but then reordered to something more useful for exporting
		# in case needed
		initial_columns = [self.c.this_winding, self.c.wind1, self.c.wind2, self.c.wind3,
						   self.c.status, self.c.id, self.c.ratea, self.c.rateb, self.c.ratec,
						   self.c.loading]

		# Transposed so columns in correct location and then columns reordered to something more suitable
		combined_df = pd.DataFrame(data).transpose()
		combined_df.columns = initial_columns

		# Remove empty string in ID column (PSSE will return '1' as '1 ' but will accept '1' as an input
		combined_df[self.c.id] = combined_df[self.c.id].str.strip()

		# Populate the complete DataFrame, overwriting any new values
		if cont_name is None:
			status_columns = [self.c.this_winding, self.c.wind1, self.c.wind2, self.c.wind3,
							  self.c.id, self.c.status]
			loading_columns = [self.c.this_winding, self.c.wind1, self.c.wind2, self.c.wind3,
							   self.c.id, self.c.ratea, self.c.rateb, self.c.ratec,
							   self.c.loading]

			# Transposed so columns in correct location and then columns reordered to something more suitable
			self.df_status = combined_df[status_columns]
			self.df_loading = combined_df[loading_columns]
			# Add column for base case
			self.df_status[constants.Contingency.bc] = self.df_status[self.c.status]
			self.df_loading[constants.Contingency.bc] = self.df_loading[self.c.loading]
		else:
			# Add column of status for this contingency
			self.df_status[cont_name] = combined_df[self.c.status]
			self.df_loading[cont_name] = combined_df[self.c.loading]

	def check_compliance(self, cont_names):
		"""
			Function inserts a row at the top of the DataFrame which confirms whether the
			loadings are all withing limits
		:param list cont_names:  List of all the contigencies and therefore those which should be considered for the
								validation
		:return None:
		"""
		# Get list of transformers which do not have any winding data
		items_with_no_rating = self.df_loading[self.df_loading[constants.Contingency.rate_for_checking] <=
											   constants.EirGridThresholds.rating_threshold]
		if not items_with_no_rating.empty:
			msg0 = (('The following 3 winding transformer have a rating of less than or equal to {:.2f} MVA and have '
					 'therefore been ignored from checking for compliance')
					.format(constants.EirGridThresholds.rating_threshold))
			msg1 = '\n'.join(['\t - Transformer connected between busbars {}, {}, {} with ID {}'
							  .format(row[self.c.wind1], row[self.c.wind2], row[self.c.wind3], row[self.c.id])
							  for _, row in items_with_no_rating.iterrows()])
			self.logger.warning('{}\n{}'.format(msg0, msg1))

		# Steady state validation
		compliance = dict()
		# Loop through each contingency and check if all steady state voltages are within limits
		for cont in cont_names:
			# Check to see if loading is less than rating where rating is > the EirGrid threshold value
			condition = ((self.df_loading[cont] < self.df_loading[constants.Contingency.rate_for_checking]) |
						 (self.df_loading[constants.Contingency.rate_for_checking] <=
						  constants.EirGridThresholds.rating_threshold))

			# Check compliant at all busbars
			compliance[cont] = all(np.where(condition, True, False))

		# Update DataFrame with new row to show whether compliant
		for key, value in compliance.iteritems():
			self.df_loading.loc[constants.Contingency.compliant, key] = value

		# Return slice from DataFrame showing whether each contingency is compliant
		col_title = constants.Excel.tx3_loading
		df_compliance = pd.DataFrame(index=cont_names, columns=(col_title,))
		df_compliance[col_title] = self.df_loading.loc[constants.Contingency.compliant,
													   cont_names]

		return df_compliance


class Tx3Data:
	"""
		Tx3 Data
	"""

	def __init__(self, flag=2, sid=-1):
		"""
		:param int flag: 2 for all 3 winding transformers
		"""
		# Initialise empty variables
		self.df = pd.DataFrame()

		# constants
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.c = constants.Tx3

		# Function for changing status (errors when using three_wnd_imped_chng_4)
		self.func_switch = psspy.three_wnd_imped_chng_3
		self.flag = flag
		self.sid = sid

		# Update DataFrame from PSSE case
		self.update()

	def update(self, cont_name=None):
		"""
			Updates data from SAV case
		:param str cont_name:  If provided will add the updated data to an existing dataframe
		"""
		# Declare functions
		func_int = psspy.atr3int
		func_char = psspy.atr3char

		# Function input constants
		entry = 1
		ties = 3

		# Retrieve data from PSSE
		ierr_int, iarray = func_int(
			sid=self.sid,
			flag=self.flag,
			ties=ties,
			entry=entry,
			string=(self.c.wind1,
					self.c.wind2,
					self.c.wind3,
					self.c.status))
		ierr_char, carray = func_char(
			sid=self.sid,
			flag=self.flag,
			entry=entry,
			ties=ties,
			string=(self.c.id,))
		if ierr_int > 0 or ierr_char > 0:
			self.logger.error(('Unable to retrieve the area data from the SAV case and the following error codes '
							   'were returned: {} and {} from functions <{}> and <{}>')
							  .format(ierr_int, ierr_char, func_int.__name__, func_char.__name__))
			raise ValueError('Error importing data from PSSE SAV case')

		# Combine data into single list of lists
		data = iarray + carray
		# Column headers initially in same order as data but then reordered to something more useful for exporting
		# in case needed
		initial_columns = [self.c.wind1, self.c.wind2, self.c.wind3, self.c.status, self.c.id]

		# Transposed so columns in correct location and then columns reordered to something more suitable
		combined_df = pd.DataFrame(data).transpose()
		combined_df.columns = initial_columns

		# Remove empty string in ID column (PSSE will return '1' as '1 ' but will accept '1' as an input
		combined_df[self.c.id] = combined_df[self.c.id].str.strip()

		# Populate the complete DataFrame, overwriting any new values
		if cont_name is None:
			# Column ordering for exported dataframe
			status_columns = [self.c.wind1, self.c.wind2, self.c.wind3, self.c.id, self.c.status]
			self.df = combined_df[status_columns]
			# Add column for base case results
			self.df[constants.Contingency.bc] = self.df[self.c.status]
		else:
			# Add column of status for this contingency
			self.df[cont_name] = combined_df[self.c.status]

	def change_state(self, asset=pd.Series(), restore=False, restore_all=False):
		"""
			Switches the status of the branch for a specific contingency and updates the branch status
		:param pd.Series asset: (optional=pd.Series()) 3 Winding transformer whose status is to be changed
		:param bool restore: (optional=False) - if set to True then will restore single tx3 back to original value
		:param bool restore_all: (optional=False) - if set to True then will restore every tx3 back to original value
		:return bool success:  Whether load flow run successfully or not
		"""
		success = False

		if restore_all:
			# Loop through and restore every branch to its original status
			for i, row in self.df.iterrows():
				# TODO: This may be quite slow so could be useful to find out current status and only switch those
				# TODO: which are actually different

				# TODO: Currently will loop through every branch since they are duplicated in direction, should detect this
				bus1 = row[self.c.wind1]
				bus2 = row[self.c.wind2]
				bus3 = row[self.c.wind3]
				ckt = row[self.c.id]
				status = row[self.c.status]
				success = self.switch(bus1, bus2, bus3, ckt, status)

		else:
			bus1 = asset[self.c.wind1]
			bus2 = asset[self.c.wind2]
			bus3 = asset[self.c.wind3]
			ckt = asset[self.c.id]

			if restore:
				# Status value is original value for this item
				status = self.df.loc[
					(self.df[self.c.wind1] == bus1) &
					(self.df[self.c.wind2] == bus2) &
					(self.df[self.c.wind3] == bus3) &
					(self.df[self.c.id] == ckt),
					self.c.status].item()
			else:
				# otherwise use value from supplied series
				status = asset[self.c.status]

			# Just change status of selected branch and DataFrame with the new status
			# TODO: This may return an error if entries in the wrong order and should be captured
			success = self.switch(
				bus1=bus1,
				bus2=bus2,
				bus3=bus3,
				ckt=ckt,
				status=status
			)

		return success

	def switch(self, bus1, bus2, bus3, ckt, status):
		"""
			Actually switches the branch
		:return bool success:
		"""
		ierr, _ = self.func_switch(i=int(bus1),
								   j=int(bus2),
								   k=int(bus3),
								   ckt=str(ckt),
								   intgar8=int(status))
		if ierr > 0:
			success = False
			self.logger.error(('Error code {} raised when calling function {} for branch {}-{}-{}-{} to '
							   'change status to {}')
							  .format(ierr, self.func_switch.__name__, bus1, bus2, bus3, ckt, status))
		else:
			success = True
		return success


class BusData:
	"""
		Stores busbar data
	"""

	def __init__(self, flag=2, sid=-1):
		"""

		:param int flag: (optional=2) - Includes in and out of service busbars
		:param int sid: (optional=-1) - Allows customer region to be defined
		"""
		# DataFrames populated with type and voltages for each study
		# Index of dataframe is busbar number as an integer
		self.df_state = pd.DataFrame()
		self.df_voltage_steady = pd.DataFrame()
		self.df_voltage_step = pd.DataFrame()

		# Populated with list of contingency names where voltages exceeded
		self.voltages_exceeded_steady = list()
		self.voltages_exceeded_step = list()

		# constants
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.c = constants.Busbars
		self.func_switch = psspy.bus_chng_3

		self.flag = flag
		self.sid = sid
		self.update()

	def get_voltages(self):
		"""
			Returns the voltage data from the latest load flow
		:return (list, list), (nominal, voltage):
		"""
		func_real = psspy.abusreal
		ierr_real, rarray = func_real(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.nominal, self.c.voltage))

		if ierr_real>0:
			self.logger.critical(('Unable to retrieve the busbar voltage data from the SAV case and PSSE returned '
								  'the following error code {} from the function <{}>')
								 .format(ierr_real, func_real.__name__)
								 )
			raise SyntaxError('Error importing data from PSSE SAV case')

		return rarray[0], rarray[1]

	def update(self, cont_name=None, step=False):
		"""
			Updates busbar data from SAV case
		:param str cont_name: (optional=None) - If a name is provided then this is used for the column header
		:param bool step: (optional=False) - If a step change voltage then update associated data frame
		"""
		# Declare functions
		func_int = psspy.abusint
		func_char = psspy.abuschar

		# Retrieve data from PSSE
		ierr_int, iarray = func_int(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus, self.c.state))
		ierr_char, carray = func_char(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus_name,))

		if ierr_int > 0 or ierr_char > 0:
			self.logger.critical(('Unable to retrieve the busbar type codes from the SAV case and PSSE returned the '
								  'following error codes {} and {} from the functions <{}> and <{}>')
								 .format(ierr_int, ierr_char, func_int.__name__, func_char.__name__)
							)
			raise SyntaxError('Error importing data from PSSE SAV case')

		nominal, voltage = self.get_voltages()

		# Combine data into single list of lists
		data = iarray + [nominal] + [voltage] + carray
		# Column headers initially in same order as data but then reordered to something more useful for exporting
		# in case needed
		initial_columns = [self.c.bus, self.c.state, self.c.nominal, self.c.voltage, self.c.bus_name]

		state_columns = [self.c.bus, self.c.bus_name, self.c.nominal, self.c.state]
		voltage_columns = [self.c.bus, self.c.bus_name, self.c.nominal, self.c.voltage]

		# Transposed so columns in correct location and then columns reordered to something more suitable
		df = pd.DataFrame(data).transpose()
		df.columns = initial_columns
		df.index = df[self.c.bus]

		if cont_name is None:
			# Since not a contingency populate all columns
			self.df_state = df[state_columns]
			self.df_voltage_steady = df[voltage_columns]
			self.df_voltage_step = df[voltage_columns]

			# Insert steady_state voltage limits
			self.add_voltage_limits(df=self.df_voltage_steady)

			# Add column for base case results
			self.df_state[constants.Contingency.bc] = self.df_state[self.c.state]
			self.df_voltage_steady[constants.Contingency.bc] = self.df_voltage_steady[self.c.voltage]
			self.df_voltage_step[constants.Contingency.bc] = self.df_voltage_step[self.c.voltage]

		else:
			# Update status of busbar
			self.df_state[cont_name] = df[self.c.state]

			# State change stored when changing status
			# #if step:
			# #	# If this is run for a step change voltage then update voltage values
			# #	self.df_voltage_step[cont_name] = df_status[self.c.voltage]
			# #else:
			# #	# If this is run for a steady state voltage then update voltage values
			# #	self.df_voltage_steady[cont_name] = df_status[self.c.voltage]

	def add_voltage_limits(self, df):
		"""
			Function will insert the upper and lower voltage limits for each busbar into the DataFrame
		:param pd.DataFrame df: DataFrame for which voltage limits should be inserted into (must contain busbar nominal
				voltage under the column [c.nominal]
		"""

		# Get voltage limits from constants which are returned as a dictionary
		v_limits = constants.EirGridThresholds.steady_state_limits

		# Convert to a DataFrame
		df_limits = pd.concat([pd.DataFrame(data=v_limits.keys()),
							   pd.DataFrame(data=v_limits.values())],
							  axis=1)
		c = constants.Busbars
		df_limits.columns = [c.nominal_lower, c.nominal_upper, c.lower_limit, c.upper_limit]

		# Add limits into busbars voltage DataFrame
		for i, row in df_limits.iterrows():
			# Add lower limit values
			df.loc[
				(df[c.nominal] > row[c.nominal_lower]) &
				(df[c.nominal] <= row[c.nominal_upper]),
				c.lower_limit] = row[c.lower_limit]
			# Add upper limit values
			df.loc[
				(df[c.nominal] > row[c.nominal_lower]) &
				(df[c.nominal] <= row[c.nominal_upper]),
				c.upper_limit] = row[c.upper_limit]
		return None

	def change_state(self, asset=pd.Series(), restore=False, restore_all=False):
		"""
			Changes the status of the busbar
		:param pd.Series asset: (optional=pd.Series()) Busbar whose state is to be changed
		:param bool restore: (optional=False) - if set to True then will restore this busbar back to original value
		:param bool restore_all: (optional=False) - if set to True then will restore every busbar back to original value
		:return bool success:  True / False on whether able to change state
		"""
		success = False

		if restore_all:
			# Loop through and restore every branch to its original status
			for i, row in self.df_state.iterrows():
				# TODO: This may be quite slow so could be useful to find out current status and only switch those
				# TODO: which are actually different

				# TODO: Currently will loop through every branch since they are duplicated in direction, should detect this
				bus = row[self.c.bus]
				state = row[self.c.state]
				success = self.switch(bus, state)

		else:
			bus = asset[self.c.bus]
			if restore:
				state = self.df_state.loc[
					(self.df_state[self.c.bus] == bus),
					self.c.state]
			else:
				state = asset[self.c.state]

			# Just change status of selected branch and DataFrame with the new status
			# TODO: This may return an error if entries in the wrong order
			success = self.switch(
				bus=bus,
				state=state
			)

		return success

	def switch(self, bus, state):
		"""
			Actually switches the busbar
		:return bool success:
		"""
		ierr = self.func_switch(i=int(bus),
								   intgar1=int(state))
		if ierr > 0:
			success = False
			self.logger.error(('Error code {} raised when calling function {} for busbar {} to '
							   'change state to {}')
							  .format(ierr, self.func_switch.__name__, bus, state))
		else:
			success = True
		return success

	def add_islanded_busbars(self, buses, cont_name):
		"""
			Function to change the status of the busbars associated with this contingency to reflect their out of
			service status
		:param pd.DataFrame() buses:  DataFrame containing the busbars which were islanded and the new state
		:param str cont_name: Name of contingency that this relates to
		:return:
		"""
		# check if cont_name already in DataFrame and if not add it populated with the initial status values
		self.df_state[cont_name] = buses[self.c.state]

	def update_voltages(self, cont_name, voltage_step, non_convergence=False):
		"""
			Updates the DataFrame which represents the step_change voltage (i.e. run without a tap changer enabled)
		:param str cont_name: Name of contingency which this data relates to
		:param bool non_convergence: If non-convergence then sets a new value which represents the value to be populated
		:param bool voltage_step: If True then updated self.df_voltage_step, if False then updates self.df_voltage_steady
		:return None:
		"""
		# If non-convergence then set every value to this value
		if non_convergence:
			self.logger.debug('Non convergence = {}'.format(non_convergence))
			# Assumes non-convergent is the same in all cases
			latest_voltages = np.nan
		else:
			# Retrieves the latest voltage values from the PSSE case
			_, latest_voltages = self.get_voltages()

		# Updates the relevant DataFrame depending on the study being done
		if voltage_step:
			# Convert value to a step change rather than actual voltage
			self.df_voltage_step[cont_name] = abs(latest_voltages-self.df_voltage_step[self.c.voltage])
		else:
			self.df_voltage_steady[cont_name] = latest_voltages

		return None

	def check_compliance(self, cont_names, voltage_step_limit=None):
		"""
			Function inserts a row at the top of the DataFrame which confirms whether the
			voltages are all withing limits
		:param list cont_names:  List of all the contigencies and therefore those which should be considered for the
								validation
		:param float voltage_step_limit:  (optional=None) if value is provided then needs to be maximum change
								between pre- and post-contingency voltage step with limit being dependant on whether
								relates to a voltage step change.  If no value provided then just looks at
								steady state
		:return pd.DataFrame contingency_compliance:  Returns a DataFrame with index=Contingency, column=type and
														True/False for each contingency to determine whether it was
														compliant to this test
		"""
		# Return slice from DataFrame showing whether this dataframe is compliant
		# Steady state validation
		compliance = dict()

		rows_to_ignore = [constants.Contingency.v_step_lbl, constants.Contingency.compliant]

		# Only check for steady state voltages if no limit provided
		if voltage_step_limit is None:
			# Empty DataFrame that will be populated with compliance data
			df_compliance = pd.DataFrame(index=cont_names, columns=(constants.Excel.voltage_steady,))
			# Loop through each contingency and check if all steady state voltages are within limits
			for cont in cont_names:
				# Condition for check
				condition = ((self.df_voltage_steady[cont] <= self.df_voltage_steady[self.c.upper_limit]) &
							(self.df_voltage_steady[cont] >= self.df_voltage_steady[self.c.lower_limit]))
				# Check compliant at all busbars
				compliance[cont] = all(np.where(condition, True, False))

			# Update DataFrame with new row to show whether compliant
			for key, value in compliance.iteritems():
				self.df_voltage_steady.loc[constants.Contingency.compliant, key] = value

			df_compliance[constants.Excel.voltage_steady] = self.df_voltage_steady.loc[constants.Contingency.compliant,
																					   cont_names]
		else:
			# Empty DataFrame that will be populated with compliance data
			df_compliance = pd.DataFrame(index=cont_names, columns=(constants.Excel.voltage_step,))

			# Step change validation
			compliance = dict()
			# #compliance[constants.Contingency.bc] = True
			# Skips base case since base case must be compliant
			for cont in cont_names:
				# Removes rows that shouldn't be considered in comparison
				bad_df = self.df_voltage_step.index.isin(rows_to_ignore)
				df_volt_step = self.df_voltage_step[~bad_df]
				# Check whether compliant with step change for contingency
				# #condition = ((self.df_voltage_step[cont] <= voltage_step_limit) |
				# #			 (self.df_voltage_step[cont] == False))
				condition = df_volt_step[cont] <= voltage_step_limit

				# #condition = (abs(self.df_voltage_step[cont] - self.df_voltage_step[self.c.voltage])
				# #			 < voltage_step_limit)
				# Check compliant at all busbars
				compliance[cont] = all(np.where(condition, True, False))

			# Update DataFrame with new row to show whether compliant
			for key, value in compliance.iteritems():
				# Set the compliant flag for this contingency
				self.df_voltage_step.loc[constants.Contingency.compliant, key] = value
				# Add in label that identifies the limit that applied to this contingency
				self.df_voltage_step.loc[constants.Contingency.v_step_lbl, key] = voltage_step_limit

			df_compliance[constants.Excel.voltage_step] = self.df_voltage_step.loc[
				constants.Contingency.compliant, cont_names
			]

		return df_compliance

	def check_within_limits(self, cont_name, busbars_to_ignore=tuple()):
		"""
			Function updates voltages and then returns True / False on whether within acceptable limits for steady state
		:param str cont_name:  Name of contingency to be checked
		:param tuple busbars_to_ignore:  Tuple of all the busbars which should be ignored for the analysis
		:return (bool, int) (within_limits, target_bus):  True / False on whether all within limits and busbar to target
		"""
		# Update all busbar details for this contingency
		self.update_voltages(cont_name=cont_name, voltage_step=False)

		# Get subset of Dataframe which excludes the busbars in the ignore list and just the values for this contingency
		df_subset = self.df_voltage_steady.loc[
			~self.df_voltage_steady[self.c.bus].isin(busbars_to_ignore),
			(cont_name, self.c.upper_limit)
		]
		# Confirm voltages are within limits
		within_limits = (df_subset[cont_name] <= df_subset[self.c.upper_limit]).all()

		target_bus = df_subset.idxmax()[cont_name]
		max_voltage = df_subset.max()[cont_name]

		return within_limits, target_bus, max_voltage


class MachineData:
	"""
		Stores machine data
	"""

	def __init__(self, flag=4, sid=-1):
		"""

		:param int flag: (optional=2) - Includes in and out of service machines
		:param int sid: (optional=-1) - Allows customer region to be defined
		"""
		# DataFrames populated with type and voltages for each study
		# Index of dataframe is busbar number as an integer
		self.df_loading = pd.DataFrame()

		# Constants
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.c = constants.Machines

		self.flag = flag
		self.sid = sid

		# Update DataFrame from PSSE case
		self.update()

	def update(self, cont_name=None):
		"""
			Updates machine data from SAV case
		:param str cont_name: (optional=None) - If a name is provided then this is used for the column header
		"""
		# Declare functions
		func_int = psspy.amachint
		func_char = psspy.amachchar
		func_real = psspy.amachreal

		# Retrieve data from PSSE
		ierr_int, iarray = func_int(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus, self.c.state))
		ierr_char, carray = func_char(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.id,))
		ierr_real, rarray = func_real(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.qgen,))

		if ierr_int > 0 or ierr_char > 0 and ierr_real > 0:
			self.logger.critical((
				'Unable to retrieve the machine data from the SAV case and PSSE returned the following error '
				'codes {}, {} and {} from the functions <{}>, <{}> and <{}>'
			).format(ierr_int, ierr_char, ierr_real,func_int.__name__, func_char.__name__, func_real.__name__))
			raise SyntaxError('Error importing data from PSSE SAV case')

		# Combine data into single list of lists
		data = iarray + carray + rarray

		# Column headers initially in same order as data but then reordered to something more useful for exporting
		# in case needed
		initial_columns = [self.c.bus, self.c.state, self.c.id, self.c.qgen]

		# Transposed so columns in correct location and then columns reordered to something more suitable
		df = pd.DataFrame(data).transpose()
		df.columns = initial_columns

		if cont_name is None:
			# Since not a contingency populate all columns
			self.df_loading = df[initial_columns]

			# Add column for base case results
			self.df_loading[constants.Contingency.bc] = self.df_loading[self.c.qgen]

		else:
			# Update status of busbar
			self.df_loading[cont_name] = df[self.c.qgen]

	def change_target(self, bus_num, target=1.0, target_bus=0):
		"""
			Function adjusts the machine target voltage as requested
		:param int bus_num: Number of busbar machine is connected to
		:param float target:  Voltage which machine should target
		:param int target_bus:  Busbar to target for voltage control (optional = 0)
		:return None:
		"""

		func = psspy.plant_chng

		ierr = func(
			i=bus_num,
			intgar1=target_bus,
			realar1=target
		)

		if ierr > 0:
			success = False
			self.logger.error((
					'Error code {} raised when calling function {} for machine {} to change target voltage to {}'
				).format(ierr, func.__name__, bus_num, target))
		else:
			success = True
		return success

	def change_output(self, bus_num, machine_id, q_target=0.0):
		"""
			Function adjusts the machine reactive power output as requested
		:param int bus_num: Number of busbar machine is connected to
		:param str machine_id:  ID of machine
		:param float q_target:  Reactive power dispatch value to use
		:return None:
		"""

		func = psspy.machine_chng_2

		ierr = func(
			i=bus_num,
			id=machine_id,
			realar2=q_target,
			realar3=q_target,
			realar4=q_target
		)

		if ierr > 0:
			success = False
			self.logger.error(
				'Error code {} raised when calling function {} for machine {} to change reactive power to {:.2f} Mvar'.format(
					ierr, func.__name__, bus_num, q_target)
			)
		else:
			success = True
		return success


class Contingency:
	"""
		Will contain the details of the specific contingency event
	"""
	def __init__(self, circuits, tx2, tx3, busbars, fixed_shunts, switched_shunts, name, busbars_to_ignore):
		"""
			Initialise generator class instance
		:param pd.DataFrame circuits: Branches associated with this contingency
		:param pd.DataFrame tx2: Two winding transformers associated with this contingency
		:param pd.DataFrame tx3: Three winding transformers associated with this contingency
		:param pd.DataFrame busbars: Islanded busbars associated with this contingency
		:param pd.DataFrame fixed_shunts: Shunts associated with this contingency
		:param pd.DataFrame switched_shunts: Shunts associated with this contingency
		:param tuple busbars_to_ignore:  Tuple of busbars which should be ignored from the analysis
		:param str name: Name of contingency
		"""

		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.circuits = circuits
		self.tx2 = tx2
		self.tx3 = tx3
		self.busbars = busbars
		self.fixed_shunts = fixed_shunts
		self.switched_shunts = switched_shunts
		self.name = name

		# Boolean set to True once contingency is setup and set to False when not.
		self.setup_correctly = False
		# Set to True if contingency is convergent.  First entry reflect
		self.convergent_v_step = False
		self.convergent_v_steady = False

		self.busbars_to_ignore = busbars_to_ignore

		# Function checks that the dataframe inputs are of the expected format
		self.check_input_data()

		# Determine whether contingency is purely just voltage control switching
		if all([self.circuits.empty, self.tx2.empty, self.tx3.empty, self.busbars.empty]):
			self.voltage_control_contingency=True
		else:
			self.voltage_control_contingency = False

	@property
	def convergence_message(self):
		"""
			Returns whether the contingency was convergent or not
		:return str:
		"""
		if not self.convergent_v_step and not self.convergent_v_steady:
			msg = constants.Contingency.non_convergent
		elif self.convergent_v_steady and self.convergent_v_step:
			msg = constants.Contingency.convergent
		elif self.convergent_v_steady and not self.convergent_v_step:
			msg = constants.Contingency.non_convergent_vstep
		elif self.convergent_v_step and not self.convergent_v_steady:
			msg = constants.Contingency.non_convergent_vsteady
		else:
			msg = constants.Contingency.error
			self.logger.error(
				(
					'Unexpected condition reached when trying to determine convergence for contingency {}'
					' and so have assumed {}').format(self.name, constants.Contingency.non_convergent)
			)
		return msg

	@property
	def convergent(self):
		"""
			Returns whether the contingency was convergent or not
		:return bool:
		"""
		if self.convergent_v_step and self.convergent_v_steady:
			return True
		else:
			return False

	def check_input_data(self):
		"""
			Function ensures that the input data is in the correct format
		:return:
		"""
		for df in (self.circuits, self.tx2, self.tx3, self.fixed_shunts, self.switched_shunts):
			# Convert input data for circuit ID to a string for all non-empty dataframes
			if not df.empty:
				# #df_status[constants.Branches.id].apply(str)
				df[constants.Branches.id] = df[constants.Branches.id].astype('str')

	def setup_contingency(
			self, circuit_data, tx2_data, tx3_data, tx3_wind_data, bus_data, fixed_shunt_data,
			switched_shunt_data, restore=False
	):
		"""
			Switch elements associated with this contingency
		:param BranchData circuit_data:  Class containing all the circuit data for the PSSE model
		:param BranchData tx2_data:  Class containing all the 2 winding transformer data for the PSSE model
		:param Tx3Data tx3_data:  Class containing all the 3 winding transformer data for the PSSE model
		:param Tx3WndData tx3_wind_data:  Class containing all the 3 winding transformer windings from the PSSE model
		:param BusData bus_data:  Class containing all the Busbar data for the PSSE model
		:param ShuntData fixed_shunt_data:  Class containing all the Shunt data for the PSSE model
		:param ShuntData switched_shunt_data:  Class containing all the Shunt data for the PSSE model
		:param bool restore: (optional=False) - Set to True and will set every branch in this contingency back to
			initial value
		:return None:
		"""

		# Skip contingency setup if BASE_CASE
		if self.name == constants.Contingency.bc:
			self.setup_correctly = True
			self.logger.debug(
				'Contingency {} is the base case and therefore no switching actions have taken place'.format(self.name)
			)
			return None

		# Iterate through each circuit and change status
		success = []
		for i, row in self.circuits.iterrows():
			success.append(circuit_data.change_state(asset=row, restore=restore))

		# Iterate through each two winding transformer and change status
		for i, row in self.tx2.iterrows():
			success.append(tx2_data.change_state(asset=row, restore=restore))

		# Iterate through each three winding transformer and change status
		for i, row in self.tx3.iterrows():
			success.append(tx3_data.change_state(asset=row, restore=restore))

		# Iterate through each busbar and change state
		for i, row in self.busbars.iterrows():
			success.append(bus_data.change_state(asset=row, restore=restore))

		# Iterate through each fixed shunts and change state
		for i, row in self.fixed_shunts.iterrows():
			success.append(fixed_shunt_data.change_state(asset=row, restore=restore))

		# Iterate through each switched shunt and change state
		for i, row in self.switched_shunts.iterrows():
			success.append(switched_shunt_data.change_state(asset=row, restore=restore))

		if restore:
			self.setup_correctly = False
		else:
			for data in (
					circuit_data, tx2_data, tx3_data, tx3_wind_data, bus_data, switched_shunt_data, fixed_shunt_data
			):
				data.update(cont_name=self.name)
			self.setup_correctly = all(success)

		if not self.setup_correctly:
			self.logger.warning(
				'An error occured when setting up contingency {} and so it has been skipped from further studies'.format(
					self.name)
			)

		return None

	def test_contingency(
			self, psse, bus_data, circuit_data, tx2_data, tx3_wind_data, machine_data, adjust_reactive
	):
		"""
			Runs load flow, checks convergent and updates busbar data with new loads
		:param PsseControl psse:  Handle to the psse control class for running functions which will impact it
		:param BusData bus_data:  Handle to the busbar data which is updated with new status values
		:param BranchData circuit_data:
		:param BranchData tx2_data:
		:param Tx3WndData tx3_wind_data:
		:param MachineData machine_data:
		:param bool adjust_reactive:  Determines whether reactive power should be adjusted or not
		:return bool success:  Returns True / False on whether the contingency ran successfully
		"""
		# Check contingency has been setup correctly
		if not self.setup_correctly:
			self.logger.warning(
				'The contingency {} has not been setup correctly and therefore cannot be tested.'.format(self.name)
			)
			bus_data.update_voltages(cont_name=self.name, voltage_step=False, non_convergence=True)
			bus_data.update_voltages(cont_name=self.name, voltage_step=True, non_convergence=True)
			return False

		# First run is with locked_taps to get the step-change voltage values, second run is with moving taps to get
		# the steady-state voltages.

		# If using to establish reactive compensation requirement initially set machine output to 0 Mvar
		if constants.ReactiveCompensationLimits.target_shunts:
			for _, machine in constants.ReactiveCompensationLimits.target_machines.iteritems():
				machine_data.change_output(bus_num=machine[0], machine_id=machine[1])

		# Run a load flow and check for convergence along with any islanded busbars
		convergent_load_flow, islanded_buses = psse.run_load_flow(lock_taps=True)

		# If any busbars islanded then update bus_data before repeating load flow study
		if not islanded_buses.empty:
			bus_data.add_islanded_busbars(buses=islanded_buses, cont_name=self.name)

			# Repeat load flow to confirm now convergent and no errors
			convergent_load_flow, islanded_buses = psse.run_load_flow(lock_taps=True)
			if not islanded_buses.empty:
				raise SyntaxError('Have managed to still have islanded busbars on second run so issue with script')

		if not convergent_load_flow:
			self.logger.info(
				(
					'Load flow for contingency {} is not convergent with fixed taps.  No busbar voltage or '
					'circuit loading values will be returned'
				).format(self.name))
			self.convergent_v_step = False

			# Update busbar data with non-convergence flag to leave as blank
			bus_data.update_voltages(cont_name=self.name, voltage_step=True, non_convergence=True)
			# Continues only in case a convergent load flow is reached on next step
		else:
			bus_data.update_voltages(cont_name=self.name, voltage_step=True)
			self.convergent_v_step = True

		# Run load flow again to obtain steady_state voltages
		# Updated the relevant aspects of bus_data for this contingency with the change in status
		convergent_load_flow, _ = psse.run_load_flow(lock_taps=False)
		if not convergent_load_flow:
			# Run with flat start and then re-run with non-flat start
			convergent_load_flow, _ = psse.run_load_flow(flat_start=True, lock_taps=False)
			if not convergent_load_flow:
				self.logger.error(
					'Unable to get convergent load flow with flat start for contingency {}'.format(self.name)
				)
			else:
				convergent_load_flow, _ = psse.run_load_flow(flat_start=False, lock_taps=False)

		if not convergent_load_flow and not adjust_reactive:
			self.logger.error(
				(
					'Unable to get load flow for contingency {} to converge fully and so no voltages / busbar details '
					'will be reliably obtained.'
				).format(self.name)
			)
			self.convergent_v_steady = False

			# Update busbar data with non-convergence flag to leave as blank
			bus_data.update_voltages(cont_name=self.name, voltage_step=False, non_convergence=True)
		else:
			if not convergent_load_flow:
				self.logger.info(
					(
						'Contingency {} not convergent without reactive compensation but will now attempt to determine '
						'reactive compensation necessary to make convergent'
					).format(self.name)
				)

			# Confirm if voltages all within limits otherwise reduce target set-point until voltages within limits at
			# remote busbars.
			if adjust_reactive:
				self.check_voltage_adjust_machines(psse=psse, bus_data=bus_data, machine_data=machine_data)

			self.convergent_v_steady = True
			bus_data.update_voltages(cont_name=self.name, voltage_step=False)
			circuit_data.update(cont_name=self.name)
			tx2_data.update(cont_name=self.name)
			tx3_wind_data.update(cont_name=self.name)
			machine_data.update(cont_name=self.name)

		return True

	def check_voltage_adjust_machines(self, psse, bus_data, machine_data):
		"""
			Function will check if voltages are all within limits and if not it will increase the reactive power at
			the designated machines until all voltages remain within limits
		:param PsseControl psse:
		:param BusData bus_data:
		:param MachineData machine_data:
		:return:
		"""

		c = constants.ReactiveCompensationLimits
		iter_count = 0
		# Used to keep track of the target voltages and busbars that have already been used
		targets_log = dict()

		# Check if all steady state voltages within limits
		within_limits, target_bus, max_voltage = bus_data.check_within_limits(cont_name=self.name, busbars_to_ignore=self.busbars_to_ignore)
		target_voltage = c.vmax_target
		target_q = c.q_min

		# Reset target_bus to local busbar
		if not c.test_target_bus:
			target_bus = 0

		# Loop round gradually increasing shunt reactive power values until either within limits or max_iterations
		# exceeded
		while not within_limits and target_voltage > c.vmin_target and target_q > c.q_max:
			# Determine which busbar to target if using machines
			if target_bus in targets_log.keys():
				target_voltage = targets_log[target_bus] - c.v_step
			else:
				target_voltage = c.vmax_target

			# Calculate next reactive power level to use
			target_q += c.q_step

			# Update dictionary
			targets_log[target_bus] = target_voltage

			# Iterate through each machine and change values
			for _, machine in c.target_machines.iteritems():
				if c.target_shunts:
					machine_data.change_output(bus_num=machine[0], machine_id=machine[1], q_target=target_q)
				else:
					machine_data.change_target(bus_num=machine[0], target=target_voltage, target_bus=target_bus)

			# Increment iter_count
			iter_count += 1

			# Run load flow and check if within_limits
			# Run a load flow and check for convergence along with any islanded busbars
			convergent, _ = psse.run_load_flow(lock_taps=False)
			if not convergent:
				convergent, _ = psse.run_load_flow(flat_start=True)
				if convergent:
					_, _ = psse.run_load_flow(flat_start=False)
				elif c.target_shunts:
					self.logger.info('\tNon-convergent for contingency {} with Q = {:.2f} Mvar'.format(self.name, target_q))
					within_limits = False
					continue
				else:
					self.logger.info(
						'\tNon-convergent for contingency {} with target_v = {:.3f} at bus {} p.u.'.format(
							self.name, target_voltage, target_bus)
					)
					within_limits = False
					continue
			within_limits, target_bus, max_voltage = bus_data.check_within_limits(
				cont_name=self.name, busbars_to_ignore=self.busbars_to_ignore
			)

			# Reset target_bus to local busbar
			if not c.test_target_bus:
				target_bus = 0

		return None






class PsseControl:
	"""
		Class to obtain and store the PSSE data
	"""

	def __init__(self):
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.sav = str()
		self.sav_name = str()
		self.sid = -1

	def load_data_case(self, pth_sav=None):
		"""
			Load the study case that PSSE should be working with
		:param str pth_sav:  (optional=None) Full path to SAV case that should be loaded
							if blank then it will reload previous
		:return None:
		"""
		try:
			func = psspy.case
		except NameError:
			self.logger.debug('PSSE has not been initialised when trying to load save case, therefore initialised now')
			success = InitialisePsspy().initialise_psse()
			if success:
				func = psspy.case
			else:
				self.logger.critical('Unable to initialise PSSE')
				raise ImportError('PSSE Initialisation Error')

		# Allows case to be reloaded
		if pth_sav is None:
			pth_sav = self.sav
		else:
			# Store the sav case path and name of the file
			self.sav = pth_sav
			self.sav_name, _ = os.path.splitext(os.path.basename(pth_sav))

		# Load case file
		ierr = func(sfile=pth_sav)
		if ierr > 0:
			self.logger.critical(('Unable to load PSSE Saved Case file:  {}.\n'
								  'PSSE returned the error code {} from function {}')
								 .format(pth_sav, ierr, func.__name__))
			raise ValueError('Unable to Load PSSE Case')

		# Set the PSSE load flow tolerances to ensure all studies done with same parameters
		self.set_load_flow_tolerances()

		return None

	def set_load_flow_tolerances(self):
		"""
			Function sets the tolerances for when performing Load Flow studies
		:return None:
		"""
		# Function for setting PSSE solution parameters
		func = psspy.solution_parameters_4

		ierr = func(intgar2=constants.PSSE.max_iterations,
					realar6=constants.PSSE.mw_mvar_tolerance)

		if ierr > 0:
			self.logger.warning(('Unable to set the max iteration limit in PSSE to {}, the model may'
								 'strugle to converge but will continue anyway.  PSSE returned the '
								 'error code {} from function <{}>')
								.format(constants.PSSE.max_iterations, ierr, func.__name__))
		return None

	def run_load_flow(self, flat_start=False, lock_taps=False):
		"""
			Function to run a load flow on the psse model for the contingency, if it is not possible will
			report the errors that have occurred
		:param bool flat_start: (optional=False) Whether to carry out a Flat Start calculation
		:param bool lock_taps: (optional=False)
		:return (bool, pd.DataFrame) (convergent, islanded_busbars):
			Returns True / False based on convergent load flow existing
			If islanded busbars then disconnects them and returns details of all the islanded busbars in a DataFrame
		"""
		# Function declarations
		if flat_start:
			# If a flat start has been requested then must use the "Fixed Slope Decoupled Newton-Raphson Power Flow Equations"
			func = psspy.fdns
		else:
			# If flat start has not been requested then use "Newton Raphson Power Flow Calculation"
			func = psspy.fnsl

		if lock_taps:
			tap_changing = 0
		else:
			tap_changing = 1

		# Run loadflow with screen output controlled
		# TODO: Define these in constants
		ierr = func(
			options1=tap_changing,  # Tap changer stepping enabled
			options2=0,  # Don't enable tie line flows
			options3=0,  # Phase shifting adjustment disabled
			options4=1,  # DC tap adjustment enabled
			options5=tap_changing,  # Include switched shunt adjustment
			options6=flat_start,  # Flat start depends on status of <flat_start> input
			options7=0,  # Apply VAR limits immediately
			options8=0)  # Non divergent solution

		# Error checking
		if ierr == 1 or ierr == 5:
			# 1 = invalid OPTIONS value
			# 5 = prerequisite requirements for API are not met
			self.logger.critical('Script error, invalid options value or API prerequisites are not met')
			raise SyntaxError('SCRIPT error, invalid options value or API prerequisites not met')
		elif ierr == 2:
			# generators are converted
			self.logger.critical(
				'Generators have been converted, you must reload SAV case or reverse the conversion of the generators')
			raise IOError('The generators are converted and therefore it is not possible to run loadflow')
		elif ierr == 3 or ierr == 4:
			# buses in island(s) without a swing bus; use activity TREE
			islanded_busbars = self.get_islanded_busbars()
			return False, islanded_busbars
		elif ierr > 0:
			# Capture future errors potential from a change in PSSE API
			self.logger.critical('UNKNOWN ERROR')
			raise SyntaxError('UNKNOWN ERROR')

		# Check whether load flow was convergent
		convergent = self.check_convergent_load_flow()

		return convergent, pd.DataFrame()

	def check_convergent_load_flow(self):
		"""
			Function to check if the previous load flow was convergent
		:return bool conergent:  True if convergent and false if not
		"""
		error = psspy.solved()
		if error == 0:
			convergent = True
		elif error in (1, 2, 3, 5):
			self.logger.debug(
				'Non-convergent load flow due to a non-convergent case with error code {}'.format(error))
			convergent = False
		else:
			self.logger.error(
				'Non-convergent load flow due to script error or user input with error code {}'.format(error))
			convergent = False

		return convergent

	def get_islanded_busbars(self):
		"""Function to retrieve any islanded busbars in PSSE
			the <psspy.tree> function does not return lists of busbars so instead need to use a function that
			actually switches out the islanded busbars using the <psspy.island> switch out all islanded busbars and
			then compare the changes to produce list of in service busbars.  This is a destructive function and
			something has to be done to restore the model to the pre-contingency situation afterwards.
			Source: https://psspy.org/psse-help-forum/question/700/get-buses-numbers-when-calling-tree-function/

		:return pd.DataFrame busbars: DataFrame of the busbars that have been disconnected as part of this contingency
		"""
		# Get list of all busbars which are in service
		# busbars_initial = handle to class <bus_data> before any changes have been made
		busbars_initial = BusData(flag=1).df_state

		# Run ISLAND function to trip out any in-service branches connected to type 4 buses and disconnects islands
		#   that don't contain a swing bus
		func = psspy.island
		ierr = func()
		if ierr == -1:
			self.logger.critical('There are no in-service buses remaining after taking this contingency')
			raise IOError('There are no in-service buses remaining after taking this contingency')

		# Get list of all busbars which are now in service
		busbars_final = BusData(flag=1).df_state

		# Compared those which were in service initially and those which are in service now and return those which
		# are updated DataFrame to only show those which are not duplicated
		df_disconnected_busbars = pd.concat([busbars_initial, busbars_final])
		df_disconnected_busbars.drop_duplicates(subset=constants.Busbars.bus, keep=False, inplace=True)

		return df_disconnected_busbars

	def define_bus_subsystem(self, busbars=tuple()):
		"""
			Function to define the bus subsystem
		:param tuple busbars:  List of busbars to be included in bus subsystem
		:return None:
		"""
		# psspy function
		func = psspy.bsys

		sid = constants.PSSE.sid

		min_voltage = min([min(x) for x in constants.EirGridThresholds.steady_state_limits.keys()])
		max_voltage = max([max(x) for x in constants.EirGridThresholds.steady_state_limits.keys()])

		# #area = constants.PSSE.area
		# #zone = constants.PSSE.zone
		# #if area is None and zone is None:
		# #	self.sid = -1
		# #	return None
		# #elif area is not None:
		# #	# Produce subsystem based on area alone
		# #	ierr = func(sid=sid,
		# #				areas=[area])
		# #elif zone is not None:
		# #	# Produce subsystem based on zone alone
		# #	ierr = func(sid=sid,
		# #				zones=[zone])
		# #else:
		# #	# Produce a subsystem based on area and zone
		# #	ierr = func(sid=sid,
		# #				areas=[area],
		# #				zones=[zone])
		# If busbars have been provided as an input then define the bus subsystem
		if busbars:
			ierr = func(
				sid=sid, usekv=1, basekv1=min_voltage, basekv2=max_voltage,
				numbus=len(busbars), buses=busbars
			)
		else:
			ierr = func(sid=sid, usekv=1, basekv1=min_voltage, basekv2=max_voltage)

		if ierr == 1:
			self.logger.critical(
				(
					'Error defining bus subsystem since sid number {} is deemed invalid by PSSE. '
					'PSSE returned the error code {} for the function {}'
				).format(sid, ierr, func.__name__))
			raise SyntaxError('Error in PSSE sid value')
		else:
			self.sid = sid

		return None

	def get_fault_currents(self):
		"""
			Using PSSARRAYS module to return fault currents
		:return pd.DataFrame() df:  DataFrame of the ik'' fault currents for every busbar in this study case
		"""
		# Obtain the Ik'' fault current for every busbar in the system
		rlst = pssarrays.iecs_currents(all=1, flt3ph=1)

		# Convert results into a dictionary
		results = {bus: abs(rlst.flt3ph[i].ia1) for i, bus in enumerate(rlst.fltbus)}
		df = pd.DataFrame.from_dict(data=results, orient='index', columns=[self.sav_name])

		return df


