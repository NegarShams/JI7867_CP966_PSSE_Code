import unittest
import os
import sys
import time
import pandas as pd
import pickle

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_LOGS = os.path.join(TESTS_DIR, 'logs')

SAV_CASE_COMPLETE = os.path.join(TESTS_DIR, 'test_sav_complete.sav')
SAV_CASE_ISLANDED = os.path.join(TESTS_DIR, 'test_sav_islanded.sav')

TEST_PICKLE_FILE = os.path.join(TESTS_DIR, 'test_sav_islanded.pkl')


two_up = os.path.abspath(os.path.join(TESTS_DIR, '../..'))
sys.path.append(two_up)

import optimisation
import optimisation.psse as TestModule
import optimisation.constants as constants

SKIP_SLOW_TESTS = False
DELETE_LOG_FILES = True

# These constants are used to return the environment back to its original format
# for testing the PSSPY import functions
original_sys = sys.path
original_environ = os.environ['PATH']

# ----- UNIT TESTS -----
@unittest.skipIf(SKIP_SLOW_TESTS, 'PSSPY import testing skipped since slow to run')
class TestPsseInitialise(unittest.TestCase):
	"""
		Functions to check that PSSE import and initialisation is possible
	"""
	def test_psse32_psspy_import_fail(self):
		"""
			Test that PSSE version 32 cannot be initialised because it is not installed
		:return:
		"""
		sys.path = original_sys
		os.environ['PATH'] = original_environ
		self.psse = TestModule.InitialisePsspy(psse_version=32)
		self.assertIsNone(self.psse.psspy)

	def test_psse33_psspy_import_success(self):
		"""
			Test that PSSE version 33 can be initialised
		:return:
		"""
		self.psse = TestModule.InitialisePsspy(psse_version=33)
		self.assertIsNotNone(self.psse.psspy)

	def test_psse34_psspy_import_success(self):
		"""
			Test that PSSE version 34 can be initialised
		:return:
		"""
		self.psse = TestModule.InitialisePsspy(psse_version=34)
		self.assertIsNotNone(self.psse.psspy)

		# Initialise psse
		status = self.psse.initialise_psse()
		self.assertTrue(status)

	def tearDown(self):
		"""
			Tidy up by removing variables and paths that are not necessary
		:return:
		"""
		sys.path.remove(self.psse.psse_py_path)
		os.environ['PATH'] = os.environ['PATH'].strip(self.psse.psse_os_path)
		os.environ['PATH'] = os.environ['PATH'].strip(self.psse.psse_py_path)

@unittest.skipIf(SKIP_SLOW_TESTS, 'PSSE initialisation skipped since slow to run')
class TestPsseControl(unittest.TestCase):
	"""
		Unit test for loading of SAV case file and subsequent operations
	"""
	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		cls.logger = optimisation.Logger(pth_logs=TEST_LOGS, uid='TestPsseControl', debug=True)
		cls.psse = TestModule.PsseControl()
		cls.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

	def test_load_case_error(self):
		psse = TestModule.PsseControl()
		with self.assertRaises(ValueError):
			psse.load_data_case()

	def test_load_flow_full(self):
		load_flow_success, df = self.psse.run_load_flow()
		self.assertTrue(load_flow_success)
		self.assertTrue(df.empty)

	def test_load_flow_flatstart(self):
		load_flow_success, df = self.psse.run_load_flow(flat_start=True)
		self.assertTrue(load_flow_success)
		self.assertTrue(df.empty)

	def test_load_flow_locked_taps(self):
		load_flow_success, df = self.psse.run_load_flow(lock_taps=True)
		self.assertTrue(load_flow_success)
		self.assertTrue(df.empty)

	def test_load_flow_islands(self):
		psse = TestModule.PsseControl()
		psse.load_data_case(pth_sav=SAV_CASE_ISLANDED)
		load_flow_success, df = self.psse.run_load_flow()
		self.assertFalse(load_flow_success)
		self.assertFalse(df.empty)
		print(df)

	@classmethod
	def tearDownClass(cls):
		# Delete log files created by logger
		if DELETE_LOG_FILES:
			paths = [cls.logger.pth_debug_log,
					cls.logger.pth_progress_log,
					cls.logger.pth_error_log]
			del cls.logger
			for pth in paths:
				if os.path.exists(pth):
					os.remove(pth)

@unittest.skipIf(SKIP_SLOW_TESTS, 'PSSE taking of contingencies skipped since slow to run')
class TestApplyPsseContingencies(unittest.TestCase):
	"""
		Script tests that contingencies can be applied without any issues
		the contingencies are manually constructed for the testing.
	"""
	@classmethod
	def setUpClass(cls):
		c = constants
		# Populate some dummy data for contingency circuits
		circuit_data = [['Test Circuit', 400, 401, 1, 0]]
		circuit_cols = [c.Contingency.header,
						c.Branches.from_bus,
						c.Branches.to_bus,
						c.Branches.id,
						c.Branches.status]
		cls.cont_circuit = TestModule.Contingency(circuits=pd.DataFrame(data=circuit_data,
																		columns=circuit_cols),
												  tx2=pd.DataFrame(),
												  tx3=pd.DataFrame(),
												  busbars=pd.DataFrame(),
												  fixed_shunts=pd.DataFrame(),
												  switched_shunts=pd.DataFrame(),
												  name='Circuit Test')

		# Populate some dummy data for contingency 3 winding transformers
		tx3_data = [['Test Transformer', 402, 202, 10001, 1, 0]]
		tx3_cols = [c.Contingency.header,
						c.Tx3.wind1,
						c.Tx3.wind2,
						c.Tx3.wind3,
						c.Branches.id,
						c.Branches.status]
		cls.cont_tx3 = TestModule.Contingency(circuits=pd.DataFrame(),
											  tx2=pd.DataFrame(),
											  tx3=pd.DataFrame(data=tx3_data,
															   columns=tx3_cols),
											  busbars=pd.DataFrame(),
											  fixed_shunts=pd.DataFrame(),
											  switched_shunts=pd.DataFrame(),
											  name='Tx3 Test')

		# Populate some dummy data for contingency circuits in reverse order
		circuit_data2 = [['Test Circuit', 401, 400, 1, 0]]
		circuit_cols = [c.Contingency.header,
						c.Branches.from_bus,
						c.Branches.to_bus,
						c.Branches.id,
						c.Branches.status]
		cls.cont_circuit2 = TestModule.Contingency(circuits=pd.DataFrame(data=circuit_data2,
																		 columns=circuit_cols),
												   tx2=pd.DataFrame(),
												   tx3=pd.DataFrame(),
												   busbars=pd.DataFrame(),
												   fixed_shunts=pd.DataFrame(),
												   switched_shunts=pd.DataFrame(),
												   name='Circuit Test2')

		# Initialise logger
		cls.logger = optimisation.Logger(pth_logs=TEST_LOGS, uid='TestApplyContingencies', debug=True)

		# Initialise and load PSSE SAV case
		cls.psse = TestModule.PsseControl()

	def test_contingency_not_ready(self):
			"""
				Tests that running of this contingency is not possible since it has not
				yet been setup
			:return:
			"""
			self.assertFalse(self.cont_circuit.setup_correctly)
			self.assertFalse(self.cont_circuit.test_contingency(psse=None, bus_data=pd.DataFrame()))

	def test_confirm_input_data_checking_correct(self):
		"""
			Confirms that the data has been converted to the correct format
		:return:
		"""
		branch_ids = self.cont_circuit.circuits.loc(axis=1)[constants.Branches.id].tolist()
		data_type_is_string = [type(x) is str for x in branch_ids]
		self.assertTrue(all(data_type_is_string))

	def test_switch_out_circuit(self):
		# Load PSSE sav case
		self.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

		# Obtain technical data from PSSE
		circuit_data = TestModule.BranchData(flag=2)
		tx2_data = TestModule.BranchData(flag=6, tx=True)
		tx3_data = TestModule.Tx3Data()
		bus_data = TestModule.BusData()
		tx3_wind_data = optimisation.psse.Tx3WndData()
		fixed_shunt_data = optimisation.psse.ShuntData(fixed=True)
		switched_shunt_data = optimisation.psse.ShuntData(fixed=False)


		self.cont_circuit.setup_contingency(circuit_data=circuit_data, tx2_data=tx2_data,
											tx3_data=tx3_data, tx3_wind_data=tx3_wind_data,
											fixed_shunt_data=fixed_shunt_data,
											switched_shunt_data=switched_shunt_data,
											bus_data=bus_data)

		# Confirm that setup correctly flag has been set
		self.assertTrue(self.cont_circuit.setup_correctly)
		self.assertTrue(self.cont_circuit.test_contingency(psse=self.psse, bus_data=bus_data))

		# Reset contingency back to original values
		self.cont_circuit.setup_contingency(circuit_data=circuit_data, tx2_data=tx2_data,
											tx3_data=tx3_data, tx3_wind_data=tx3_wind_data,
											fixed_shunt_data=fixed_shunt_data,
											switched_shunt_data=switched_shunt_data,
											bus_data=bus_data,
											restore=True)
		# Confirm that setup correctly flag has been cleared
		self.assertFalse(self.cont_circuit.setup_correctly)

	def test_switch_out_tx3(self):
		# Load PSSE sav case
		self.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

		# Obtain technical data from PSSE
		circuit_data = TestModule.BranchData(flag=2)
		tx2_data = TestModule.BranchData(flag=6, tx=True)
		tx3_data = TestModule.Tx3Data()
		bus_data = TestModule.BusData()
		tx3_wind_data = optimisation.psse.Tx3WndData()
		fixed_shunt_data = optimisation.psse.ShuntData(fixed=True)
		switched_shunt_data = optimisation.psse.ShuntData(fixed=False)

		self.cont_tx3.setup_contingency(circuit_data=circuit_data, tx2_data=tx2_data,
										tx3_data=tx3_data, tx3_wind_data=tx3_wind_data,
										fixed_shunt_data=fixed_shunt_data,
										switched_shunt_data=switched_shunt_data,
										bus_data=bus_data)

		# Confirm that setup correctly flag has been set
		self.assertTrue(self.cont_tx3.setup_correctly)
		self.assertTrue(self.cont_tx3.test_contingency(psse=self.psse, bus_data=bus_data))

		# Reset contingency back to original values
		self.cont_tx3.setup_contingency(circuit_data=circuit_data, tx2_data=tx2_data,
										tx3_data=tx3_data, tx3_wind_data=tx3_wind_data,
										fixed_shunt_data=fixed_shunt_data,
										switched_shunt_data=switched_shunt_data,
										bus_data=bus_data, restore=True)
		# Confirm that setup correctly flag has been cleared
		self.assertFalse(self.cont_tx3.setup_correctly)

	def test_switch_out_circuit_reversed(self):
		# Load PSSE sav case
		self.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

		# Obtain technical data from PSSE
		circuit_data = TestModule.BranchData(flag=2)
		tx2_data = TestModule.BranchData(flag=6, tx=True)
		tx3_data = TestModule.Tx3Data()
		bus_data = TestModule.BusData()
		tx3_wind_data = optimisation.psse.Tx3WndData()
		fixed_shunt_data = optimisation.psse.ShuntData(fixed=True)
		switched_shunt_data = optimisation.psse.ShuntData(fixed=False)

		self.cont_circuit2.setup_contingency(circuit_data=circuit_data, tx2_data=tx2_data,
											 tx3_data=tx3_data, tx3_wind_data=tx3_wind_data,
											 fixed_shunt_data=fixed_shunt_data,
											 switched_shunt_data=switched_shunt_data,
											 bus_data=bus_data)

		# Confirm that setup correctly flag has been set
		self.assertTrue(self.cont_circuit2.setup_correctly)
		self.assertTrue(self.cont_circuit2.test_contingency(psse=self.psse, bus_data=bus_data))

		# Reset contingency back to original values
		self.cont_circuit2.setup_contingency(circuit_data=circuit_data, tx2_data=tx2_data,
											 tx3_data=tx3_data, tx3_wind_data=tx3_wind_data,
											 fixed_shunt_data=fixed_shunt_data,
											 switched_shunt_data=switched_shunt_data,
											 bus_data=bus_data, restore=True)
		# Confirm that setup correctly flag has been cleared
		self.assertFalse(self.cont_circuit2.setup_correctly)
		print(circuit_data)

	@classmethod
	def tearDownClass(cls):
		# Delete log files created by logger once completed but only if success
		if DELETE_LOG_FILES:
			paths = [cls.logger.pth_debug_log,
					 cls.logger.pth_progress_log,
					 cls.logger.pth_error_log]
			del cls.logger
			for pth in paths:
				if os.path.exists(pth):
					os.remove(pth)


class TestPickle(unittest.TestCase):
	@classmethod
	def setUpClass(cls):
		cls.logger = optimisation.Logger(pth_logs=TEST_LOGS, uid='TestApplyContingencies', debug=True)

	def test_pickle(self):
		# Initialise and load PSSE SAV case
		psse = TestModule.PsseControl()
		psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

		circuit_data = TestModule.BranchData(flag=2).df_status

		with open(TEST_PICKLE_FILE, 'wb') as f:
			pickle.dump(circuit_data, f, pickle.HIGHEST_PROTOCOL)

	@classmethod
	def tearDownClass(cls):
		# Delete log files created by logger once completed but only if success
		if DELETE_LOG_FILES:
			paths = [cls.logger.pth_debug_log,
					 cls.logger.pth_progress_log,
					 cls.logger.pth_error_log]
			del cls.logger
			for pth in paths:
				if os.path.exists(pth):
					os.remove(pth)



if __name__ == '__main__':
	unittest.main()