"""
#######################################################################################################################
###											PSSE Contingency Test													###
###		Script loops through a list of contingencies, switching out the relevant elements and then runs load flows  ###
###		to determine whether any of the nodes will exceed voltage limits.  Further development as part of JI7867	###
###		in that the will also adjust the reactive compensation to maintain voltages within acceptable equipment		###
###		limits anywhere on the system.																				###
###																													###
###		Full details of the methodology applied by this script can be found in JI7867-6-0							###
###																													###
###		The constants detailed below will need to be updated for the specific user running these studies			###
###																													###
###		To produce the outputs as shown in JI7867-6-0 the script data_visulisation.py needs to be run				###																													###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JI7867- EirGrid - Capital Project 966																###
###																													###
#######################################################################################################################
"""

import collections
import logging
import os
import time
import pandas as pd

import optimisation
import optimisation.constants as constants

# Constants for running studies
# The following directory needs to be defined and provides a reference where all of the .SAV case files will be stored
project_directory = (
	r'C:\Users\david\Power Systems Consultants Inc\Jobs - JI7867 - Cable Integration Studies for Capital Project 966'
	r'\5 Working Docs\Phase B')

# Paths where all of the PSSe SAV cases are stored and the following naming convention has taken place
# 	BC = without CP966
# 	CP966 = With CP966
# 	RC = With fixed reactive compensation for validation
pth_EirGrid_SAV = [
	os.path.join(project_directory, 'PSSE Models', 'snv_high wind_30snv_mec_33_v1(BC).sav'),
	os.path.join(project_directory, 'PSSE Models', 'snv_low wind_30snv_mec_33_v2(BC).sav'),
	os.path.join(project_directory, 'PSSE Models', 'WP_high wind_30w_mec_33_v1(BC).sav'),
	os.path.join(project_directory, 'PSSE Models', 'snv_high wind_30snv_mec_33_v1(CP966).sav'),
	os.path.join(project_directory, 'PSSE Models', 'snv_low wind_30snv_mec_33_v2(CP966).sav'),
	os.path.join(project_directory, 'PSSE Models', 'WP_high wind_30w_mec_33_v1(CP966).sav'),
	os.path.join(project_directory, 'PSSE Models', 'snv_high wind_30snv_mec_33_v1(CP966)_RC.sav'),
	os.path.join(project_directory, 'PSSE Models', 'snv_low wind_30snv_mec_33_v2(CP966)_RC.sav'),
	os.path.join(project_directory, 'PSSE Models', 'WP_high wind_30w_mec_33_v1(CP966)_RC.sav'),
	os.path.join(project_directory, 'PSSE Models', 'WP_high wind_30w_mec_33_v1(CP966)_RC2.sav')
]

# Which studies to consider adjusting reactive power for voltage limits, False means no tuning of the reactive
# compensation will be considered and True means tuning of the reactive compensation will be carried out to maintain
# voltages within acceptable limits.
adjust_reactive_comp = [
	False,
	False,
	False,
	True,
	True,
	True,
	False,
	False,
	False,
	False
]

# Path where a spreadsheet detailing the contingencies are stored (and empty example is stored in the Python package)
pth_Contingencies = os.path.join(project_directory, 'Contingencies.xlsx')

# Results paths where results from each of the study case will be stored
pth_results = [
	os.path.join(project_directory, 'Results_SVHW(BC).xlsx'),
	os.path.join(project_directory, 'Results_SVLW(BC).xlsx'),
	os.path.join(project_directory, 'Results_WPHW(BC).xlsx'),
	os.path.join(project_directory, 'Results_SVHW(CP966).xlsx'),
	os.path.join(project_directory, 'Results_SVLW(CP966).xlsx'),
	os.path.join(project_directory, 'Results_WPHW(CP966).xlsx'),
	os.path.join(project_directory, 'Results_SVHW(CP966)_RC.xlsx'),
	os.path.join(project_directory, 'Results_SVLW(CP966)_RC.xlsx'),
	os.path.join(project_directory, 'Results_WPHW(CP966)_RC.xlsx'),
	os.path.join(project_directory, 'Results_WPHW(CP966)_RC2.xlsx')
]

# Populate this list with the specific study cases from the lists above that you wish to include for analysis.  I.e. if
# only one number is included in the list then studies will only be run for that single study case.
selector = [9]

# Contains details of the busbars to include and these are then used to focus the area of the analysis for voltage
# compliance checking.
pth_busbar_list = os.path.join(project_directory, 'Model_Review.xlsx')


def main(cont_workbook, psse_sav_case, target_workbook, adjust_reactive, pth_busbars=str()):
	"""
		Main function
	:param str cont_workbook: Path to workbook which contains contingency details
	:param str psse_sav_case:  Path to psse SAV case
	:param str target_workbook: Path where results should be saved
	:param bool adjust_reactive:  Whether the reactive compensation should be adjusted to maintain voltages or not
	:param str pth_busbars:  Path to file where a list of busbars are located
	:return str target_workbook:  Path to excel file created as part of study
	"""

	# Load PSSE case and get details
	logger = logging.getLogger(constants.Logging.logger_name)

	# Get list of busbars to include (if path provided as an input)
	if pth_busbars:
		df_busbars, busbars_to_consider = optimisation.file_handling.busbars_to_consider(pth_busbar_list=pth_busbars)
	else:
		busbars_to_consider = tuple()

	# Initialise and load PSSE case
	logger.info('Loading PSSE case {} and checking convergent'.format(psse_sav_case))
	psse_case = optimisation.psse.PsseControl()
	psse_case.load_data_case(pth_sav=psse_sav_case)

	# Create a bus subsystem within PSSE to only include results for specific busbars
	psse_case.define_bus_subsystem(busbars=tuple(busbars_to_consider))

	# Check have an initially convergent load flow
	load_flow_success, islanded_busbars = psse_case.run_load_flow()
	if not load_flow_success or not islanded_busbars.empty:
		if not islanded_busbars.empty:
			msg0 = (
				'The loaded PSSE case {} has the following busbars islanded in the base case:'.format(psse_sav_case)
			)
			msg1 = '\n'.join([
				'\t - Busbar: {}'.format(bus) for bus in islanded_busbars[optimisation.constants.Busbars.bus]
			])
			msg = '{}\n{}'.format(msg0, msg1)
			logger.critical(msg)
			raise ValueError('Check PSSE SAV Case - Islanded busbars')
		if not load_flow_success:
			logger.critical('The base PSSE SAV case {} is non-convergent'.format(psse_sav_case))
			raise ValueError('Check PSSE SAV Case - Non-Convergent')
	else:
		logger.info('Initial load flow study convergent')

	# Obtain data for all elements in initial conditions (prior to any contingencies or balancing of reactive
	# compensation)
	bus_data = optimisation.psse.BusData(sid=psse_case.sid)
	machine_data = optimisation.psse.MachineData(sid=psse_case.sid)
	circuit_data = optimisation.psse.BranchData(flag=2, sid=psse_case.sid)
	tx2_data = optimisation.psse.BranchData(flag=6, tx=True, sid=psse_case.sid)
	tx3_data = optimisation.psse.Tx3Data(sid=psse_case.sid)
	tx3_wind_data = optimisation.psse.Tx3WndData(sid=psse_case.sid)
	fixed_shunt_data = optimisation.psse.ShuntData(fixed=True, sid=psse_case.sid)
	switched_shunt_data = optimisation.psse.ShuntData(fixed=False, sid=psse_case.sid)

	# Import workbook of contingency details identifying those elements which need to be switched out / switched in
	logger.info('Importing details of all contingencies from workbook {}'.format(cont_workbook))
	contingency_data = optimisation.file_handling.ImportContingencies(pth=cont_workbook)
	contingency_circuits_details = collections.OrderedDict()

	# Process imported contingencies to identify associated elements in PSSE case and assign to a Contingency class
	for cont_name in contingency_data.contingency_names:
		logger.debug('Processing contingency: {}'.format(cont_name))

		# Check if contingency has any busbars which should be excluded when adjusting the reactive power
		if pth_busbars:
			# noinspection PyUnboundLocalVariable
			busbars_to_ignore = tuple(df_busbars[df_busbars['Contingency'] == cont_name].index)
		else:
			busbars_to_ignore = tuple()

		# Setup contingency by switching out the elements
		circuits, tx2, tx3, busbars, fixed_shunts, switched_shunts = contingency_data.group_contingencies_by_name(
			cont_name=cont_name)
		contingency = optimisation.psse.Contingency(
			circuits=circuits, tx2=tx2, tx3=tx3, busbars=busbars, fixed_shunts=fixed_shunts,
			switched_shunts=switched_shunts, name=cont_name, busbars_to_ignore=busbars_to_ignore)
		contingency_circuits_details[cont_name] = contingency
	logger.info('All contingencies from workbook {} imported'.format(cont_workbook))

	# Empty DataFrame to use for populating details of whether the contingency was convergent, circuit / shunt switching
	# used to keep track of whether shunt or normal asset outage for analysis against voltage limits.
	circuit_switching = list()
	shunt_switching = list()
	contingency_convergence = pd.DataFrame(
		index=contingency_circuits_details.keys(), columns=(constants.Excel.message, constants.Excel.convergence)
	)
	# Add Base case message
	contingency_convergence.loc[constants.Contingency.bc, constants.Excel.message] = constants.Contingency.convergent
	contingency_convergence.loc[constants.Contingency.bc, constants.Excel.convergence] = True

	# Loop through all contingencies, apply outage, run load flow, check for reactive compensation requirements
	for name, contingency in contingency_circuits_details.iteritems():

		logger.info('Testing contingency {}'.format(name))
		if contingency.voltage_control_contingency:
			shunt_switching.append(name)
		else:
			circuit_switching.append(name)

		contingency.setup_contingency(
			circuit_data=circuit_data,
			tx2_data=tx2_data,
			tx3_data=tx3_data,
			bus_data=bus_data,
			fixed_shunt_data=fixed_shunt_data,
			switched_shunt_data=switched_shunt_data,
			tx3_wind_data=tx3_wind_data
		)
		contingency.test_contingency(
			psse=psse_case, bus_data=bus_data, circuit_data=circuit_data, tx2_data=tx2_data,
			tx3_wind_data=tx3_wind_data, machine_data=machine_data, adjust_reactive=adjust_reactive
		)

		# Rather than restoring it is quicker to just reload the SAV case
		# It also avoids a potential error where circuits are not necessarily switched back in
		psse_case.load_data_case()
		# Update DataFrame with contingency convergence details
		contingency_convergence.loc[name, constants.Excel.message] = contingency.convergence_message
		contingency_convergence.loc[name, constants.Excel.convergence] = contingency.convergent

	# Loop through all results data and confirm compliance
	compliance = list()

	# Have to check voltage step
	df_compliance_cont_step = bus_data.check_compliance(
		cont_names=circuit_switching, voltage_step_limit=constants.EirGridThresholds.cont_step_change_limit)

	df_compliance_reactor_step = bus_data.check_compliance(
		cont_names=shunt_switching, voltage_step_limit=constants.EirGridThresholds.reactor_step_change_limit)

	# Combine compliance for voltage step due to contingency or reactor switching and add to list of all compliance
	# flags
	df_compliance_step_change = pd.concat([df_compliance_cont_step, df_compliance_reactor_step], axis=0)
	df_compliance_step_change.loc[constants.Contingency.bc] = True
	compliance.append(df_compliance_step_change)

	# Combine all contingency names into a single list
	all_cont_names = circuit_switching + shunt_switching
	for data_set in (bus_data, circuit_data, tx2_data, tx3_wind_data):
		compliance.append(data_set.check_compliance(cont_names=all_cont_names))

	# Combine compliance results into a single dataset
	contingency_compliance = pd.concat([contingency_convergence] + compliance, axis=1, sort=True)

	# Data to write is formatted into a dictionary with the associated name to use
	c = optimisation.constants.Excel
	results_data = collections.OrderedDict()
	results_data[c.voltage_steady] = bus_data.df_voltage_steady
	results_data[c.voltage_step] = bus_data.df_voltage_step
	results_data[c.busbars] = bus_data.df_state
	results_data[c.circuit_status] = circuit_data.df_status
	results_data[c.circuit_loading] = circuit_data.df_loading
	results_data[c.tx2_status] = tx2_data.df_status
	results_data[c.tx2_loading] = tx2_data.df_loading
	results_data[c.tx3] = tx3_data.df
	results_data[c.tx3_wind_status] = tx3_wind_data.df_status
	results_data[c.tx3_wind_loading] = tx3_wind_data.df_loading
	results_data[c.fixed_shunts] = fixed_shunt_data.df_status
	results_data[c.switched_shunts] = switched_shunt_data.df_status
	results_data[c.machine_data] = machine_data.df_loading

	# Export all results to the relevant results workbook
	optimisation.file_handling.ExportResults(
		pth_workbook=target_workbook, results=results_data, convergence=contingency_compliance
	)

	return target_workbook


if __name__ == '__main__':
	# Starting point
	# Get time stamp for tracking progress
	t0 = time.time()
	# Get location for log files to be stored
	script_path = os.path.dirname(os.path.abspath(__file__))
	log_path = os.path.join(script_path, 'Logs')
	uid = 'Contingencies_{}'.format(time.strftime('%Y%m%d_%H%M%S'))
	log_cls = optimisation.Logger(pth_logs=log_path, uid=uid)
	local_logger = log_cls.logger
	local_logger.info('Study Started')

	# Iterates through each of the SAV case / results files provided
	for i in selector:
		t1 = time.time()
		pth_sav = pth_EirGrid_SAV[i]
		pth_res = pth_results[i]
		local_logger.info('Processing SAV case {}'.format(pth_sav))
		try:
			pth_excel = main(
				cont_workbook=pth_Contingencies, psse_sav_case=pth_sav, target_workbook=pth_res,
				pth_busbars=pth_busbar_list,
				adjust_reactive=adjust_reactive_comp[i]
			)
			local_logger.info(
				'SAV case {} completed in {:.2f}, results saved in {}'.format(pth_sav, time.time()-t1, pth_res)
			)
		except ValueError:
			local_logger.error('Contingency analysis for the SAV case {} could not be completed'.format(pth_sav))

	local_logger.info('Study completed in {:.2f} seconds'.format(time.time()-t0))
