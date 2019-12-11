from setuptools import setup, find_packages

setup(name='ndb_scheduler',
      version='0.1',
      description='A tool to schedule days for the Naturskolan Database',
      license='MIT',
      packages=find_packages(),
      include_package_data=True,
      install_requires=[
        'Click',
        'openpyxl',
      ],
      entry_points={
        'console_scripts': [
            'extract_visits_and_fritids=ndb_scheduler.commands:extract_visits_and_fritids',

        ],
      },
      zip_safe=False)
