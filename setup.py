from setuptools import setup 

setup( 
	name='chromedriver_manager', 
	version='0.1', 
	description='Manage downloads of chromedrivers for Windows machines', 
	author='Ash', 
	packages=['chromedriver_manager'], 
	install_requires=[ 
		'requests==2.32.3', 
		'pywin32==306', 
	], 
) 
