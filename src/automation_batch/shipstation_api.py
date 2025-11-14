#API Calls
from base64 import *
import requests
import json
import re

#v2Key xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ

def findList(data, target_value, path=""):
    matches = []

    if isinstance(data, dict):
        for key, value in data.items():
            new_path = f"{path}.{key}" if path else key
            if value == target_value:
                matches.append(new_path)
            elif isinstance(value, (dict, list)):
                matches.extend(findList(value, target_value, new_path))

    elif isinstance(data, list):
        for i, item in enumerate(data):
            new_path = f"{path}[{i}]"
            if item == target_value:
                matches.append(new_path)
            elif isinstance(item, (dict, list)):
                matches.extend(findList(item, target_value, new_path))

    return matches

def basePaths(paths):
    """Trim off the last property (after the final '.') from each path."""
    base_paths = set()
    for p in paths:
        if '.' in p:
            base_paths.add(p.rsplit('.', 1)[0])  # split from the right, once
    return base_paths

def getValue(data, path):
    """Navigate nested dict/list using a string path like 'a.b[0].c'."""
    current = data
    # Split on dots but keep parts with [indexes]
    parts = path.split('.')
    for part in parts:
        # Extract the base key and any list indices (e.g. "rates[2][3]")
        match = re.match(r"([^\[]+)(\[.*\])?", part)
        if not match:
            raise ValueError(f"Invalid path part: {part}")
        key = match.group(1)
        indices = match.group(2)

        # Go into the dict
        current = current[key]

        # Handle one or more [number] parts
        if indices:
            for idx in re.findall(r"\[(\d+)\]", indices):
                current = current[int(idx)]
    return current

def getOrder(orderNum):#This works for awaiting shipping orders. External Shipment ID is Order Number
    headers = {"api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"}
    url = "https://api.shipstation.com/v2/shipments"
    
    query = {
        "shipment_status": "pending",#swap to pending when live, label_purchased if testing with shipped products
        "external_shipment_id": orderNum
    }

    response = requests.get(url, headers=headers, params=query)
    data = response.json()

    #filename = "S:/Workstation - DavidK/Code/Automation_Art/Json Outputs/getOrder.json"
    #with open(filename, 'w') as f:
    #    json.dump(data, f, indent=4)

    order = data["shipments"][0]["shipment_id"]
    #carrier = data["shipments"][0]["carrier_id"]
    #shipment = data["shipments"][0]["service_code"]
    #package = data["shipments"][0]["packages"][0]["package_code"]

    #if package == 'thick_envelope':
    #    package = 'package'

    #print(str(order)+": "+str(carrier)+": "+str(shipment))
    return order#, carrier, shipment, package

def getRateId(shipmentid,carrierid,shipment,package):
    headers = {
    "Content-Type": "application/json",
    "api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"
    }
    url = "https://api.shipstation.com/v2/rates"

    payload = {
        "shipment_id": shipmentid,
        "rate_options": {
            "carrier_ids": [
            carrierid
            ]
        }
    }
    
    response = requests.post(url, json=payload, headers=headers)
    data = response.json()

    filename = "S:/Workstation - DavidK/Code/Automation_Art/Json Outputs/getRates.json"
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

    shipmentMatch = findList(data,shipment)
    baseShipment = basePaths(shipmentMatch)

    if len(shipmentMatch) > 2:
        print('Multiple Shipments with Carrier')
        packageMatch = findList(data,package)
        basePackage = basePaths(packageMatch)
        common = baseShipment & basePackage
        common = next(iter(common))
        common = common+".rate_id"
        #print(common)
    else: 
        print('One Shipment with Carrier')
        common = next(iter(baseShipment))
        common = common+".rate_id"
        #print(common)

    rateId = getValue(data,common)
    return rateId

def getCarriers():
    headers = {"api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"}
    url = "https://api.shipstation.com/v2/carriers"

    query = {}

    response = requests.get(url, headers=headers, params=query)
    data = response.json()

    #filename = "S:/Workstation - DavidK/Code/Automation_Art/Json Outputs/getCarriers.json"
    #with open(filename, 'w') as f:
    #    json.dump(data, f, indent=4)

def makeBatch(batchNum,shipmentId,rateId,warehouseId):
    url = "https://api.shipstation.com/v2/batches"

    payload = {
    "external_batch_id": batchNum,
    "batch_notes": "TEST BATCH WITH MULTIPLE",
    "shipment_ids": shipmentId,
    "rate_ids": rateId,
    "warehouse_ids": warehouseId
    }

    headers = {
    "Content-Type": "application/json",
    "api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"
    }

    response = requests.post(url, json=payload, headers=headers)
    data = response.json()

    filename = "S:/Workstation - DavidK/Code/Automation_Art/Json Outputs/makeBatch.json"
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

def makeEmptyBatch(externalId):
    url = "https://api.shipstation.com/v2/batches"

    payload = {
    "external_batch_id": '"'+externalId+'"',
    "batch_notes": "Nov 12 2025"
    }

    headers = {
    "Content-Type": "application/json",
    "api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"
    }

    response = requests.post(url, json=payload, headers=headers)
    print(response.status_code)
    data = response.json()

    filename = "S:/Workstation - DavidK/Code/Automation_Art/Json Outputs/makeBatch.json"
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

    batchId = data["batch_id"]
    shstId = data["batch_number"]#shipstation id
    return batchId, shstId

def addToBatch(batchId,shipmentId):
    url = "https://api.shipstation.com/v2/batches/"+batchId+"/add"

    payload = {
        "shipment_ids": [shipmentId]#shipmentId,
        #"rate_ids": [rateId],#rateId,
        #"warehouse_ids": warehouseId
    }

    headers = {
    "Content-Type": "application/json",
    "api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"
    }

    response = requests.post(url, json=payload, headers=headers)
    #print(response.status_code)
    #data = response.json()

    #filename = "S:/Workstation - DavidK/Code/Automation_Art/Json Outputs/makeBatch.json"
    #with open(filename, 'w') as f:
    #    json.dump(data, f, indent=4)

def RATETEST():
    headers = {
    "Content-Type": "application/json",
    "api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"
    }
    url = "https://api.shipstation.com/v2/rates"

    payload = {
        "shipment_id": "se-1326980720",
        "rate_options": {
            "carrier_ids": ["se-67891"],
            "package_type": ["fedex_envelope_onerate"],
            "service_code": ["fedex_2day"]
        }
    }
    
    response = requests.post(url, json=payload, headers=headers)
    data = response.json()
    print(response.status_code)

    filename = "S:/Workstation - DavidK/Code/Automation_Art/Json Outputs/FedEx.json"
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

def BATCHTEST(batchId,shipmentId):
    url = "https://api.shipstation.com/v2/batches/"+batchId+"/add"

    payload = {
        "shipment_ids": [shipmentId]
    }

    headers = {
    "Content-Type": "application/json",
    "api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"
    }

    response = requests.post(url, json=payload, headers=headers)
    #data = response.json()
    print(response.status_code)

    #filename = "S:/Workstation - DavidK/Code/Automation_Art/Json Outputs/FedEx.json"
    #with open(filename, 'w') as f:
    #    json.dump(data, f, indent=4)

def getbatchbyid():
    url = "https://api.shipstation.com/v2/batches/external_batch_id/FedExTestBatchNov42025"

    headers = {"api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"}

    response = requests.get(url, headers=headers)
    data = response.json()
    print(response.status_code)

    filename = "S:/Workstation - DavidK/Code/Automation_Art/Json Outputs/FedEx.json"
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

def main():
    #shipmentId, carrier, shipment, package = getOrder(7117854)
    #rateId = getRateId(shipmentId,carrier,shipment,package)#put this variable in when getting from getRates(): shipment
    #print(shipmentId+": "+rateId)
    #getCarriers()

    #batchId, shstid = makeEmptyBatch('FedExTestBatch-Nov42025-Test2')
    #print(batchId)

    #getbatchbyid()

    shipmentId, carrier, shipment, package = getOrder(7118467)
    #rateId = getRateId(shipmentId,carrier,shipment,package)
    BATCHTEST("se-309493174",shipmentId)
    
    #RATETEST()

    print('done')

#main()