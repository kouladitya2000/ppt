import pandas as pd
import json

# Load the Excel file
excel_file_path = "/path_to_your_excel_file.xlsx"  # Replace with the actual path to your Excel file
excel_data = pd.read_excel(excel_file_path)

# Function to convert Excel data to JSON format
def convert_to_json(df):
    output = []
    
    # Iterate through each row in the DataFrame
    for _, row in df.iterrows():
        instance_data = {}
        
        # Populate instance data from the Excel file
        instance_data['Environment'] = row['Environment']  # Custom mapping for the 'Environment' column
        instance_data['SubscriptionName'] = row['SubscriptionName']
        instance_data['ResourceGroupName'] = row['ResourceGroupName']
        instance_data['Instance-type'] = row['Instance-type']
        
        # Nested structure for instance-creation
        instance_data['instance-creation'] = {
            "name-label": row['Name-label'],
            "requestorID": row['RequestorID'],
            "requestorEmail": row['RequestorEmail'],
            "locationCode": row['LocationCode'],
            "billingCode": row['BillingCode'],
            "service": row['Service'],
            "osVersionName": row['OSVersionName'],
            "osImageType": row['OSImageType'],
            "machineType": row['MachineType'],
            "zone": row.get('Zone', "")  # If zone is not always present, set default as an empty string
        }
        
        # Additional disks data, if present
        instance_data['additional-disks'] = [
            {
                "additionalDiskName": row.get('additionalDiskName', ""),
                "additionalDiskSize": row.get('additionalDiskSize', ""),
                "additionalDiskType": row.get('additionalDiskType', ""),
                "additionalDiskLun": row.get('additionalDiskLun', "")
            }
        ]
        
        # Windows-only data
        instance_data['windows-only'] = {
            "owner": row.get('windows-only.owner', ""),
            "approvers1": row.get('windows-only.approvers1', ""),
            "approvers2": row.get('windows-only.approvers2', ""),
            "BusinessUnit": row.get('windows-only.BusinessUnit', ""),
            "SupportEntity": row.get('windows-only.SupportEntity', "")
        }
        
        # Linux-only data
        instance_data['linux-only'] = {
            "centrify": row.get('linux-only.centrify', "")
        }
        
        # Network-optional data
        instance_data['network-optional'] = {
            "vnetName": row.get('network-optional.vnetName', ""),
            "vnetResourceGroupName": row.get('network-optional.vnetResourceGroupName', ""),
            "subnetName": row.get('network-optional.subnetName', ""),
            "ipAddress": row.get('network-optional.ipAddress', ""),
            "proxyServer": row.get('network-optional.proxyServer', ""),
            "proxyPort": row.get('network-optional.proxyPort', "")
        }

        # Instance deletion details, if available
        instance_data['instance-deletion'] = {
            "demise": row.get('instance-deletion.demise', ""),
            "CrNumber": row.get('instance-deletion.CrNumber', "")
        }
        
        # Add the row's data to the output list
        output.append(instance_data)
    
    return output

# Convert the Excel data to JSON format
json_output = convert_to_json(excel_data)

# Save the JSON data to a file
output_file_path = "output.json"
with open(output_file_path, "w") as json_file:
    json.dump(json_output, json_file, indent=4)

print(f"Excel data has been successfully converted to JSON and saved as {output_file_path}.")
