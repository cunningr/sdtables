from sdtables.sdtables import SdTables

# Add tables to xlsx
worksheet_name='customSheetName'
name = 'table.name.1'
data = [
    {'presence': 'present', 'dhcpServerIps': None, 'dnsServerIps': None, 'gateways': None, 'ipPoolCidr': '172.31.0.0/16', 'ipPoolName': 'IPP_Automation_Underlay'},
    {'dhcpServerIps': None, 'dnsServerIps': None, 'gateways': None, 'ipPoolCidr': '172.16.0.0/16', 'ipPoolName': 'IPP_Overlay', 'presence': 'present'},
    {'dhcpServerIps': None, 'dnsServerIps': None, 'gateways': None, 'ipPoolCidr': 'fd12:3456::/36', 'ipPoolName': 'IPP_IPv6', 'presence': 'present'}
]
tables = SdTables()
tables.add_xlsx_table_from_data(name, data, worksheet_name=worksheet_name)

tables.save_xlsx('sdtables_example')
