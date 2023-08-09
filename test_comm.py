import pyvisa
import time 
# Define device addresses
DEVICE_ADDRESS = 'GPIB0::1::INSTR'  # Change to the proper chamber address
DMM_ADDRESS = 'GPIB0::22::INSTR'  # Change to the proper multimeter address
SECOND_DMM_ADDRESS = 'GPIB0::11::INSTR'  # Change to the proper address for the second DMM

# Open resources
rm = pyvisa.ResourceManager()
device = rm.open_resource(DEVICE_ADDRESS)
dmm = rm.open_resource(DMM_ADDRESS)
second_dmm = rm.open_resource(SECOND_DMM_ADDRESS)
time.sleep(10)
# Perform a basic communication test
print("Chamber ID:", device.query("*IDN?"))
print("DMM ID:", dmm.query("*IDN?"))
print("Second DMM ID:", second_dmm.query("*IDN?"))
time.sleep(10)

# Close resources
device.close()
dmm.close()
second_dmm.close()
