"""
#######################################################################################################################
###											PSSE Contingency Data Visualisation										###
###		Script will house functions to produce graphs to visualise the raw data produced by the Contingency Testing ###
###		scripts 																									###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JI7867 - EirGrid - Capital Project 966																###
###																													###
#######################################################################################################################
"""


import matplotlib.pyplot as plt
import matplotlib.cbook as cbook
import collections
from matplotlib.ticker import FormatStrFormatter
import numpy as np
import pandas as pd
import os
import optimisation.file_handling as file_handling

# Constants / Filenames
project_directory = (
	r'C:\Users\david\Power Systems Consultants Inc\Jobs - JI7867 - Cable Integration Studies for Capital Project 966'
	r'\5 Working Docs\Phase B')

# List and selection of raw data files
selector = [3, 4, 5]
source_files = [
	os.path.join(project_directory, 'Results_SVHW(BC).xlsx'),
	os.path.join(project_directory, 'Results_SVLW(BC).xlsx'),
	os.path.join(project_directory, 'Results_WPHW(BC).xlsx'),
	os.path.join(project_directory, 'Results_SVHW(CP966)2.xlsx'),
	os.path.join(project_directory, 'Results_SVLW(CP966)2.xlsx'),
	os.path.join(project_directory, 'Results_WPHW(CP966)2.xlsx'),
	os.path.join(project_directory, 'Results_SVHW(CP966)_RC.xlsx'),
	os.path.join(project_directory, 'Results_SVLW(CP966)_RC.xlsx'),
	os.path.join(project_directory, 'Results_WPHW(CP966)_RC.xlsx'),
	os.path.join(project_directory, 'Results_SVLW(CP966)_RC2.xlsx')
]

pth_busbar_list = os.path.join(project_directory, 'Model_Review.xlsx')
sht = 'V Steady'


def produce_plots_voltage(source_file):
	"""
		Produces plots based on type of output
	:param source_file:
	:return:
	"""
	# Figure Name
	file_name, _ = os.path.splitext(source_file)
	fig_name = file_name + '_Voltage.png'

	# Columns and indexes in raw data that do not need to be considered in plot
	cols_to_drop = ['NUMBER.1', 'EXNAME', 'BASE', 'PU', 'LOWER_LIMIT', 'UPPER_LIMIT']
	index_to_drop = ['Compliant']

	df_busbars, busbars_to_keep = file_handling.busbars_to_consider(pth_busbar_list=pth_busbar_list)

	# Get results from contingency tool
	df = pd.read_excel(source_file, sheet_name=sht, index_col=0)
	df.drop(index=index_to_drop, inplace=True)

	# Where based on a 380kV nominal adjust to be based on 400kV nominal
	df.loc[df['BASE'] == 380.0] = df[df.select_dtypes(include=['number']).columns] * (380.0/400.0)

	for bus, contingency in df_busbars['Contingency'].iteritems():
		if not pd.isna(contingency):
			df.loc[bus, contingency] = np.nan

	# Extract voltage upper and lower threshold
	# l_threshold = df['LOWER_LIMIT']
	# u_threshold = df['UPPER_LIMIT']
	thresholds = df.loc[:, ['LOWER_LIMIT', 'UPPER_LIMIT']]

	# Tidy / drop unnecessary entries from DataFrame
	df.drop(columns=cols_to_drop, inplace=True)
	df.dropna(axis=1, inplace=True)
	# Reduce list to only include those that exist in this list
	busbars_to_keep = [int(x) for x in busbars_to_keep if x in df.index]

	# Produce figure and axis
	fig2 = plt.figure(figsize=(28, 15), dpi=300)
	ax2 = fig2.gca()

	# X axis values
	x2 = df.loc[busbars_to_keep, :].index

	# Calculate box positions for boundary limits
	stats = collections.OrderedDict()
	for x in x2:
		# noinspection PyUnresolvedReferences
		stats[x] = cbook.boxplot_stats(thresholds.loc[x, :].values, labels=[x])[0]
		stats[x]['q1'] = thresholds.loc[x, 'LOWER_LIMIT']
		stats[x]['q3'] = thresholds.loc[x, 'UPPER_LIMIT']

	# #error_points = [1.0 for x in x_categories]

	# Produce violin plot of busbar voltages
	# Y axis values containing all busbar voltages for each contingency of selected busbars
	y2 = df.loc[x2, :].T.values
	vp = ax2.violinplot(y2, showmedians=False, showextrema=True, widths=0.7)
	# Adjust violin plot colours
	for pc in vp['bodies']:
		pc.set_facecolor('green')
	vp['cbars'].set_linewidth(1)
	vp['cmaxes'].set_linewidth(1)
	vp['cmins'].set_linewidth(1)

	_ = ax2.bxp(
		stats.values(), showcaps=False, medianprops={'linewidth': 0}, boxprops={'linewidth': 1},
		whiskerprops={'linewidth': 0}, widths=0.8
	)

	# Get labels to give for each busbar
	x_categories = df_busbars.loc[busbars_to_keep, 'Plot Name']
	x_categories = [str(x) for x in x_categories]

	# Adjust graph properties for better presentation
	ax2.set_xticks(np.arange(1, len(x2) + 1))
	ax2.set_xticklabels(x_categories)
	plt.xticks(rotation='vertical', fontsize=14)
	plt.yticks(np.linspace(0.89, 1.11, 23), fontsize=14)
	plt.ylim((0.89, 1.11))
	ax2.yaxis.set_major_formatter(FormatStrFormatter('%.2f'))
	plt.grid(True, which='major', linestyle='--', linewidth=1, color='lightgrey')
	plt.ylabel('Steady State Voltage (p.u.)', fontsize=14)
	plt.tight_layout()

	# Save figure
	plt.savefig(fig_name, dpi='figure')
	plt.close()


if __name__ == '__main__':
	for i in selector:
		res_file = source_files[i]

		if os.path.isfile(res_file):
			produce_plots_voltage(source_file=res_file)
		else:
			print('File <{}> does not exist'.format(res_file))
