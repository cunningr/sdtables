from sdtables.sdtables import SdTables
import yaml

agent = """
description: "Defines agents"
properties:
  agentId:
    type: number
  agentName:
    type: ["string", "null"]
    description: 'Agent hostname'
  agentType:
    type: ["string", "null"]
  countryId:
    type: ["string", "null"]
  enabled:
    type: ["string", "null"]
  location:
    type: ["string", "null"]
  network:
    type: ["string", "null"]
    pattern: '.*/\d+'
  prefix:
    type: ["string", "null"]
    format: "ipv4"
  agentState:
    type: ["string", "null"]
  hostname:
    type: ["string", "null"]
  tagAll:
    type: ["string", "null"]
    enum: ['X', '']
    default: 'X'
  tagCustomZ:
    enum: ['X', '']
    type: ["string", "null"]
"""

# We can add an empty table based on the schema definition with data validators
test_schema = yaml.load(agent, Loader=yaml.SafeLoader)
tables = SdTables()
tables.add_xlsx_table_from_schema('test.schema', test_schema, worksheet_name='TestSheet')

# Or we can a tables with data and apply the schema defined validators
data = [
    {'agentId': 4495, 'agentName': 'Chicago, IL (Trial)', 'agentType': 'Cloud', 'countryId': 'US', 'enabled': None, 'location': 'Chicago Area', 'network': "10.0.0.0/24", 'prefix': '1.1.1.1', 'agentState': None, 'hostname': None, 'tagAll': 'X', 'tagCustomZ': 'X'},
    {'agentId': 4497, 'agentName': 'Ashburn, VA (Trial)', 'agentType': 'Cloud', 'countryId': 'US', 'enabled': None, 'location': 'Ashburn Area', 'network': None, 'prefix': None, 'agentState': None, 'hostname': None, 'tagAll': 'X', 'tagCustomZ': 'X'},
    {'agentId': 4500, 'agentName': 'Sydney, Australia (Trial)', 'agentType': 'Cloud', 'countryId': 'AU', 'enabled': None, 'location': 'New South Wales, Australia', 'network': None, 'prefix': None, 'agentState': None, 'hostname': None, 'tagAll': 'X', 'tagCustomZ': 'X'},
    {'agentId': 4503, 'agentName': 'Hong Kong (Trial)', 'agentType': 'Cloud', 'countryId': 'HK', 'enabled': None, 'location': 'Hong Kong', 'network': None, 'prefix': None, 'agentState': None, 'hostname': None, 'tagAll': 'X', 'tagCustomZ': 'X'}
]

tables.add_xlsx_table_from_schema('test.schema2', test_schema, data=data, worksheet_name='TestSheet')

# Save the xlsx workbook
tables.save_xlsx('sdtables_example2')

