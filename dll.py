import ctypes
from EHLLAPI import Emulator
# Load DLL into memory.
hapi = Emulator()

if(hapi.connect()==0):
    print("Connected!")
    # hapi.send_keys("NMTS")
    # hapi.send_keys("@T")
    # hapi.send_keys("NMTS")
    # hapi.send_keys("@E")
    # hapi.send_keys("@E")
    #cls
    #print(hapi.set_cursor(753))
    #print(hapi.get_field(453,20))
   
    # hapi.send_keys("140220@O@O@T")
    # hapi.copy_str_to_field("720",654)
    # hapi.send_keys("@T@E")
    #hapi.send_keys("@<@<")
    #print(hapi.get_cursor())
    #print(hapi.search_str("WORD",0))
    #print(hapi.lock_kb())
    #print(hapi.find_field_length("T ",453))
    #print(hapi._wait())

    # try:
    #     hapi.search_str("WORLD",0)
    # except:
    #     print("CODE 24")
    feildStr = str(hapi.get_field(453,10),encoding = "utf-8").strip()

    print(feildStr)
    print(len(feildStr))

    if(hapi.disconnect()==0):
        print("Disconnected")  