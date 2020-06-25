from distutils.core import setup
setup(
  name='sdtables',
  packages=['sdtables'],
  version='1.0.3',
  license='MIT',
  description='sdtables (schema data tables) is a module providing convenient wrapper functions for working with tabulated from various sources including MS Excel',
  author='Richard Cunningham',
  author_email='cunningr@gmail.com',
  url='https://github.com/cunningr/sdtables',
  download_url='https://github.com/cunningr/sdtables/archive/1.0.0.zip',
  keywords=['Excel', 'tables', 'schema'],
  install_requires=[
          'openpyxl==3.0.3',
          'jsonschema',
      ],
  classifiers=[
    'Development Status :: 4 - Beta',
    'Intended Audience :: Developers',
    'Topic :: Software Development :: Build Tools',
    'License :: OSI Approved :: MIT License',
    'Programming Language :: Python :: 3.6'
  ],
)
