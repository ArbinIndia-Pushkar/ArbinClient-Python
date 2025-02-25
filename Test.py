from pythonnet import load
load("coreclr")
import clr
import os
import sys
import time


"""
Section 0:
Adding reference to the DLL files: ArbinClient.DLL and ArbinDataModel.DLL
The DLL files are stored inside the folder 'DLL'
"""
newpath = os.path.dirname(os.path.abspath(__file__)) + '\\DLL'
sys.path.append(newpath)

classDLL = clr.AddReference('ArbinClient')
classDLL = clr.AddReference('ArbinDataModel')



"""
Section 1:
Importing required classes from the respective namespaces inside the DLL files
"""
from ArbinClient.Core import ArbinClient
from ArbinClient.Core import CreateArbinClientArgs
from Arbin.Library.DataModel.RequestInformation import *
from Arbin.Library.DataModel import *



"""
Section 2:
Set up
"""
_client = ArbinClient()

""" Connects to ArbinClient """
args = CreateArbinClientArgs()
args.IPAddress = "127.0.0.1"
args.UserName = "admin"
args.Password = "000000"
args.Timeout = 5000

nErrorCode = -1
Client, nErrorCode = ArbinClient.CreateArbinClient(args)

if nErrorCode != 0:
    print(f"Error Code: {nErrorCode}. Failed to create ArbinClient.")
    sys.exit(1)

print(f"Error Code: {nErrorCode}")
print(f"Connected: {Client.IsConnected()}")

if not Client.IsConnected():
    print("Failed to connect to the client.")
    sys.exit(1)



"""
Section 3:
Implementing APIs
"""

""" Gets ArbinClient Version """
def GetArbinClientVersion():
    print("ArbinClient Version: " + _client.GetArbinClientVersion())
    print("-----------------------------------")
    display_menu()


""" Unsubscribe to Monitor Data """
def UnSubscribeToMonitorDataPython():
    # Clear the console before printing new data
    os.system('cls' if os.name == 'nt' else 'clear')
    send = _client.UnSubscribeMonitorData()
    print(f"Unsubscription Request Sent: {send}")
    print("-----------------------------------")
    display_menu()



""" Subscribe to Monitor Data """
feedback = SubscribeMonitorDataFDBK()

def M_client_OnSubscribeMonitorData(feedback1):
    global feedback
    feedback = feedback1

def OnLog(args):
    print(args.Message)

Client.OnLog += OnLog
Client.OnSubscribeMonitorData += M_client_OnSubscribeMonitorData

def SubscribeToMonitorDataPython(_client):
    if not _client.IsConnected():
        print("Client is not connected. Exiting.")
        sys.exit(1)
    
    subArgs = SubscribeMonitorDataArgs()
    send = _client.SubscribeMonitorData(subArgs)
    print(f"Subscription Request Sent: {send}")
    
    global feedback

    while True:

        # Clear the console before printing new data
        os.system('cls' if os.name == 'nt' else 'clear')
        
        # Prepare the result data for printing
        result = []
        for channel in feedback.ChannelMonitorDatas:
            result.append(f"\tChannelID = {channel.ChannelID}")
            result.append(f"\tStatus = {channel.Status}")
            result.append(f"\tCommunicationFailure = {channel.CommunicationFailure}")
            result.append(f"\tScheduleName = {channel.ScheduleName}")
            result.append(f"\tTestObjectName = {channel.TestObjectName}")
            result.append(f"\tCANBMSName = {channel.CANBMSName}")
            result.append(f"\tSMBName = {channel.SMBName}")
            result.append(f"\tChartName = {channel.ChartName}")
            result.append(f"\tTestName = {channel.TestName}")
            result.append(f"\tExitCondition = {channel.ExitCondition}")
            result.append(f"\tStepAndCycle = {channel.StepAndCycle}")
            result.append(f"\tStepID = {channel.StepID}")
            result.append(f"\tSubStepID = {channel.SubStepID}")
            result.append(f"\tCycleID = {channel.CycleID}")
            result.append(f"\tBarcode = {channel.Barcode}")
            result.append(f"\tMasterChannelID = {channel.MasterChannelID}")
            result.append(f"\tTestTime = {channel.TestTime}")
            result.append(f"\tStepTime = {channel.StepTime}")
            result.append(f"\tVoltage = {channel.Voltage}")
            result.append(f"\tCurrent = {channel.Current}")
            result.append(f"\tPower = {channel.Power}")

            # Optional: add more channel details here

        # Join all lines and print them to the console
        print("\n".join(result))
        
        # Sleep for 3 seconds before printing again
        time.sleep(3)      


"""
Section 4:
Main Program
"""

""" Main Menu """
def display_menu():
    print("Please Select")
    print("0. Get ArbinClient Version")
    print("1. Subscribe To Monitor Data")
    choice = int(input("Enter your choice: "))

    """ Switch Case"""
    if choice == 0:
        GetArbinClientVersion()
    elif choice == 1:
        SubscribeToMonitorDataPython(Client)


print("-----------------------------------")
display_menu()

