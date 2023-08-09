import pyvisa

def convert_temperature(temp_format):
    # Convert the temperature format (multiple of 10) to actual temperature
    temp = float(temp_format) / 10.0
    return temp

def convert_temperature_format(temp):
    # Convert the actual temperature to temperature format (multiple of 10)
    temp_format = str(int(float(temp) * 10))
    return temp_format

rm = pyvisa.ResourceManager()
device = rm.open_resource('GPIB0::1::INSTR')  # Replace 'GPIB0::1::INSTR' with the appropriate device address

# Configure device properties
device.baud_rate = 19200
device.data_bits = 8
device.parity = pyvisa.constants.Parity.none
device.stop_bits = pyvisa.constants.StopBits.one
device.timeout = 10  # Set a longer timeout duration

def send_query_command(command):
    # Send the command to the device and read the response
    retries = 1
    while retries > 0:
        try:
            response = device.query(command + '\r')
            return response
        except pyvisa.VisaIOError as e:
            retries -= 1
            print(f"An error occurred while sending the command, please try again")
    print("Failed to communicate with the device. Please check the connection and try again.")
    return None

print('Enter your command below.\r\n1. Read actual chamber temperature\n2. Read set point temperature\n3. Set the temperature')

while True:
    user_input = input(">> ")

    if user_input == 'exit':
        device.close()
        break
    elif user_input == '1':
        response = send_query_command('R 100,1')
        if response:
            temperature = convert_temperature(response)
            temperature_str = str(temperature) + " C"
            print("Current Chamber Temperature:", temperature_str)
    elif user_input == '2':
        response = send_query_command('R 300,1')
        if response:
            temperature = convert_temperature(response)
            temperature_str = str(temperature) + " C"
            print("Set Point Temperature:", temperature_str)
    elif user_input == '3':
        temp = input("Enter temperature value to set to: ")
        try:
            temp_format = convert_temperature_format(temp)
            command = 'W 300,' + temp_format
            response = device.write(command)
            temperature_str = str(temp) + " C"
            ("Temperature set to", temperature_str)
        except ValueError:
            print("Invalid temperature value. Please try again.")
    else:
        print("Invalid command. Please enter a valid option (1, 2, or 3).")
