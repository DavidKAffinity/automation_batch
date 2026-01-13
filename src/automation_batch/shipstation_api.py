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
    #Navigate nested dict/list using a string path like 'a.b[0].c'
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

def requestJson(response,title):#do after requests.post
    print(response.status_code)
    data = response.json()
    filename = "./src/automation_batch/jsonOutputs/"+title+".json"
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

def getOrder(orderNum):#This works for awaiting shipping orders. External Shipment ID is Order Number
    headers = {"api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"}
    url = "https://api.shipstation.com/v2/shipments"
    
    query = {
        "shipment_status": "pending",#swap to pending when live, label_purchased if testing with shipped products
        "external_shipment_id": orderNum
    }

    response = requests.get(url, headers=headers, params=query)
    requestJson(response,'order')
    data = response.json()

    order = data["shipments"][0]["shipment_id"]
    return order

def makeEmptyBatch(externalId,batchNotes):
    url = "https://api.shipstation.com/v2/batches"

    payload = {
    "external_batch_id": '"'+externalId+'"',
    "batch_notes": '"'+batchNotes+'"'
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
    return batchId

def addToBatch(batchId,shipmentId):
    url = "https://api.shipstation.com/v2/batches/"+batchId+"/add"

    payload = {
        "shipment_ids": [shipmentId]
    }

    headers = {
    "Content-Type": "application/json",
    "api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"
    }

    #response = requests.post(url, json=payload, headers=headers)
    requests.post(url, json=payload, headers=headers)

def listByTag(response,targetTag):
    shipmentList = []

    for shipment in response.get("shipments",[]):
        tags = shipment.get("tags", [])

        if any(tag.get("name") == targetTag for tag in tags):
            shipmentList.append(shipment.get("shipment_id"))

    print(f"Found {len(shipmentList)} matching shipments.")
    return shipmentList

def listShipments():
    headers = {"api-key": "xFLRkudJ/+cRC+jDUzSiB31fe6E4Vy1m1M1sootREdQ"}
    url = "https://api.shipstation.com/v2/shipments"
    
    query = {
        "shipment_status": "pending",#swap to pending when live, label_purchased if testing with shipped products
        "page_size": 100000,
    }

    response = requests.get(url, headers=headers, params=query)
    requestJson(response,'list')
    data = response.json()
    return data

def main():
    #response = listShipments()
    #shipmentList = listByTag(response,'A) Needs to print')
    #print(shipmentList)

    getOrder(7149797)

    print('done')

main()