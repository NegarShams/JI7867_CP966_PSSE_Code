"""
#######################################################################################################################
###											PSSE Contingency Test													###
###		Script loops through a list of contingencies, switching out the relevant elements and then runs load flows  ###
###		to determine whether any of the nodes will exceed voltage limits											###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JI7867- EirGrid - Capital Project 966																###
###																													###
#######################################################################################################################
"""

import os
import pandas as pd

# Set to True to run in debug mode and therefore collect all output to window
DEBUG_MODE = False


class EirGridThresholds:
	"""
		Constants specific to the WPD study
	"""
	# voltage_limits are declared as a dictionary in the format
	# {(lower_voltage,upper_voltage):(pu_limit_lower,pu_limit_upper)} where
	# <lower_voltage> and <upper_voltage> represent the extremes over which
	# the voltage <pu_limit_lower> and <pu_limit_upper> applies
	# These are based on the post-contingency steady state limits provided in
	# EirGrid, "Transmission System Security and Planning Standards"
	# 380 used since that is base voltage in PSSE
	steady_state_limits = {
		(109, 111.0): (99.0/110.0, 120.0/110.0),
		(219.0, 221.0): (200.0/220.0, 240.0/220.0),
		(250.0, 276.0): (250.0/275.0, 303.0/275.0),
		(379.0, 401.0): (360.0/380.0, 410.0/380.0)
	}

	reactor_step_change_limit = 0.03
	cont_step_change_limit = 0.1

	# This is a threshold value, circuits with ratings less than this are reported and ignored
	rating_threshold = 0

	def __init__(self):
		""" Purely added to avoid error message"""
		pass


class ReactiveCompensationLimits:
	""" Constants associated with finding the optimum reactive power limits """

	# Maximum number of iterations to perform before failing the contingency
	max_iterations = 100

	# Reactive compensation step size (when tuning with shunts)
	qvar_step = 5.0

	# Minimum target voltage (when tuning with machines)
	vmax_target = 1.078947
	vmin_target = 0.948
	v_step = (vmax_target-vmin_target)/max_iterations

	# Limits for reactive compensation levels to allow testing within limits
	q_max = -1000.0
	q_min = 0.0
	q_step = -5.0

	target_machines = {
		'DNS4': (2204, 'SH'),
		'WOO4': (5464, 'SH')
	}

	# Use shunts
	target_shunts = False

	# Set to True and will adjust reactive power to target reducing value at remote busbar
	test_target_bus = False

	def __init__(self):
		pass


class PSSE:
	"""
		Class to hold all of the constants associated with PSSE initialisation
	"""
	# default version = 33
	version = 33

	# Setting on whether PSSE should output results based on whether operating in DEBUG_MODE or not
	output = {True: 1, False: 6}

	# Maximum number of iterations for a Newton Raphson load flow (default = 20)
	max_iterations = 100
	# Tolerance for mismatch in MW/Mvar (default = 0.1)
	mw_mvar_tolerance = 2.0

	sid = 1

	def __init__(self):
		self.psse_py_path = str()
		self.psse_os_path = str()

	def get_psse_path(self, psse_version, reset=False):
		"""
			Function returns the PSSE path specific to this version of psse
		:param int psse_version:
		:param bool reset: (optional=False) - If set to True then class is reset with a new psse_version
		:return str self.psse_path:
		"""
		if self.psse_py_path and self.psse_os_path and not reset:
			return self.psse_py_path, self.psse_os_path

		if 'PROGRAMFILES(X86)' in os.environ:
			program_files_directory = r'C:\Program Files (x86)\PTI'
		else:
			program_files_directory = r'C:\Program Files\PTI'

		# TODO: Data checking of input version
		psse_paths = {
			32: 'PSSE32\PSSBIN',
			33: 'PSSE33\PSSBIN',
			34: 'PSSE34\PSSPY27'
		}
		os_paths = {
			32: 'PSSE32\PSSBIN',
			33: 'PSSE33\PSSBIN',
			34: 'PSSE34\PSSBIN'
		}
		self.psse_py_path = os.path.join(program_files_directory, psse_paths[psse_version])
		self.psse_os_path = os.path.join(program_files_directory, os_paths[psse_version])
		return self.psse_py_path, self.psse_os_path

	def find_psspy(self, start_directory=r'C:'):
		"""
			Function to search entire directory and find PSSE installation

		"""
		# Clear variables
		self.psse_py_path = str()
		self.psse_os_path = str()

		psspy_to_find = "psspy.pyc"
		psse_to_find = "psse.bat"
		for root, dirs, files in os.walk(start_directory):  # Walks through all subdirectories searching for file
			if psspy_to_find in files:
				self.psse_py_path = root
			elif psse_to_find in files:
				self.psse_os_path = root

			if self.psse_py_path and self.psse_os_path:
				break

		return self.psse_py_path, self.psse_os_path


class Contingency:
	bc = 'BASE_CASE'
	final = 'FINAL'

	compliant = 'Compliant'
	v_step_lbl = 'Voltage Step Limit'
	rate_for_checking = 'RATEA'

	# Message that is displayed to indicate whether study was convergent
	# Key reflects the number of True statements
	non_convergent = 'Non-Convergent'
	convergent = 'Convergent'
	non_convergent_vstep = 'Non-Convergent Step Change'
	non_convergent_vsteady = 'Non-Convergent Steady State'
	error = 'ERROR'

	# Label used to identify all the contingencies by name
	header = 'CONTINGENCY'

	def __init__(self):
		pass


class Busbars:
	bus = 'NUMBER'
	state = 'TYPE'
	nominal = 'BASE'
	voltage = 'PU'
	bus_name = 'EXNAME'

	# Labels used for columns to determine voltage limits
	nominal_lower = 'NOM_LOWER'
	nominal_upper = 'NOM_UPPER'
	lower_limit = 'LOWER_LIMIT'
	upper_limit = 'UPPER_LIMIT'

	# Number stored in DataFrame if an error has occured
	error_id = -1

	def __init__(self):
		pass


class Machines:
	bus = 'NUMBER'
	state = 'STATUS'
	id = 'ID'
	qgen = 'QGEN'


	# Number stored in DataFrame if an error has occured
	error_id = -1

	def __init__(self):
		pass


class Branches:
	from_bus = 'FROMNUMBER'
	to_bus = 'TONUMBER'
	id = 'ID'
	status = 'STATUS'

	# This returns that maximum loading from either end of the circuit
	loading = 'MVA'
	ratea = 'RATEA'
	rateb = 'RATEB'
	ratec = 'RATEC'

	# Number stored in DataFrame if an error has occured
	error_id = -1

	def __init__(self):
		pass


class Shunts:
	bus = 'NUMBER'
	id = 'ID'
	status = 'STATUS'

	# Number stored in DataFrame if an error has occured
	error_id = -1

	def __init__(self):
		pass


class Tx3:
	wind1 = 'WIND1NUMBER'
	wind2 = 'WIND2NUMBER'
	wind3 = 'WIND3NUMBER'
	this_winding = 'WNDBUSNUMBER'
	status = 'STATUS'
	id = 'ID'

	# This returns that maximum loading from either end of the circuit
	loading = 'MVA'
	ratea = 'RATEA'
	rateb = 'RATEB'
	ratec = 'RATEC'

	# Number stored in DataFrame if an error has occured
	error_id = -1

	def __init__(self):
		pass


class Logging:
	"""
		Log file names to use
	"""
	logger_name = ''
	debug = 'DEBUG'
	progress = 'INFO'
	error = 'ERROR'
	extension = '.log'

	def __init__(self):
		"""
			Just included to avoid Pycharm error message
		"""
		pass


class Excel:
	""" Constants associated with inputs from excel """
	circuit = 'Circuits'
	tx2 = '2 Winding'
	tx3 = '3 Winding'
	busbars = 'Busbars'
	fixed_shunts = 'Fixed Shunts'
	switched_shunts = 'Switched Shunts'
	machine_data = 'Machine Data'

	status = 'Status'
	loading = 'Loading'
	circuit_status = '{}_{}'.format(circuit, status)
	circuit_loading = '{}_{}'.format(circuit, loading)
	tx2_status = '{}_{}'.format(tx2, status)
	tx2_loading = '{}_{}'.format(tx2, loading)
	tx3_wind_status = '{}_{}'.format(tx3, status)
	tx3_wind_loading = '{}_{}'.format(tx3, loading)

	comment = 'COMMENT'

	voltage_steady = 'V Steady'
	voltage_step = 'V Step'

	convergence = 'Convergent'
	message = 'Description'

	tx2_loading = '2 Winding Loading'
	tx3_loading = '3 Winding Loading'
	circuit_loading = 'Circuit Loading'

	start_row = 0

	id = 'ID'

	# This is the column names that will be applied and is irrelevant of what is already in the workboook
	columns = {
		circuit: [Contingency.header, Branches.from_bus, Branches.to_bus, Branches.id, Branches.status, comment],
		tx2: [Contingency.header, Branches.from_bus, Branches.to_bus, Branches.id, Branches.status, comment],
		tx3: [Contingency.header, Tx3.wind1, Tx3.wind2, Tx3.wind3, Tx3.id, Tx3.status, comment],
		busbars: [Contingency.header, Busbars.bus, Busbars.state, comment],
		fixed_shunts: [Contingency.header, Shunts.bus, Shunts.id, Shunts.status, comment],
		switched_shunts: [Contingency.header, Shunts.bus, Shunts.id, Shunts.status, comment]
	}

	# Cell format used for conditional formatting of cells
	cell_format_error = {'bold': False,
						 'bg_color': 'red'}
	cell_format_change = {'bold': False,
						 'bg_color': 'green'}

	def __init__(self):
		pass
