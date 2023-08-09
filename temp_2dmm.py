import pyvisa
import time
import os
import logging
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference
from openpyxl.chart.trendline import Trendline
from openpyxl.utils import get_column_letter

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Define device addresses
DEVICE_ADDRESS = 'GPIB0::1::INSTR'  # Change to the proper chamber address
DMM_ADDRESS = 'GPIB0::22::INSTR'  # Change to the proper multimeter address
SECOND_DMM_ADDRESS = 'GPIB0::11::INSTR'  # Change to the proper address for the second DMM
DIRECTORY = "C:/Users/HPflaste/OneDrive/"  # Change to the proper directory

# Open Environmental Chamber
rm = pyvisa.ResourceManager()
device = rm.open_resource(DEVICE_ADDRESS)
dmm = rm.open_resource(DMM_ADDRESS)
second_dmm = rm.open_resource(SECOND_DMM_ADDRESS)

def configure_dmm(dmm_resource):
    dmm_resource.baud_rate = 19200
    dmm_resource.data_bits = 8
    dmm_resource.parity = pyvisa.constants.Parity.none
    dmm_resource.stop_bits = pyvisa.constants.StopBits.one
    dmm_resource.timeout = 6000  # Timeout duration
    # Add more configurations if needed for other DMM settings

# In the main code:
# Open Environmental Chamber
rm = pyvisa.ResourceManager()
device = rm.open_resource(DEVICE_ADDRESS)

# Open and configure DMMs
dmm = rm.open_resource(DMM_ADDRESS)
second_dmm = rm.open_resource(SECOND_DMM_ADDRESS)
configure_dmm(dmm)
configure_dmm(second_dmm)


# Configure device properties
device.baud_rate = 19200
device.data_bits = 8
device.parity = pyvisa.constants.Parity.none
device.stop_bits = pyvisa.constants.StopBits.one
device.timeout = 100  # Timeout duration


def convert_temperature(temp_format):
    # Convert the temperature format (multiple of 10) to actual temperature
    temp = float(temp_format) / 10.0
    return temp

def convert_temperature_format(temp):
    # Convert the actual temperature to temperature format (multiple of 10)
    temp_format = str(int(float(temp) * 10))
    return temp_format

def measure_voltage():
    # Configure the DMM for voltage measurement
    logger.info("Measuring voltage...")
    time.sleep(3)
    dmm.write("INIT")  # Initiate measurement
    second_dmm.write("INIT")
    voltage = abs(float(dmm.query("FETCH?")) * 1000.0)  # Multiply by 1000 to convert to millivolts
    voltage2 = abs(float(second_dmm.query("FETCH?")))
    time.sleep(2)
    return voltage, voltage2

def read_actual_temperature(max_retries=70):
    for attempt in range(max_retries):
        try:
            logger.info('Waiting for temperature to stabilize...')
            time.sleep(60)
            response_1 = device.query('R 100,1')
            time.sleep(5)
            actual_temp = convert_temperature(response_1)
            logger.info('[{}] Actual temp: {}'.format(time.strftime("%H:%M:%S"), actual_temp))
            return actual_temp

        except:
            logger.error("VisaIOError occurred")
            if attempt < max_retries - 1:
                logger.info("Retrying the communication...")
                time.sleep(5)
            else:
                logger.error("Failed to read temperature after %d attempts.", max_retries)
                raise

from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series

def save_data_to_excel(data, directory):
    # Save the data to an XLSX file using the openpyxl library
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Temperature (C)"
    sheet["B1"] = "Voltage1 (mV)"
    sheet["C1"] = "Voltage2 (mV)"

    row_number = 2
    voltage_data_dmm1 = {}  # Dictionary for DMM1 voltage values
    voltage_data_dmm2 = {}  # Dictionary for DMM2 voltage values

    for temp, voltages in data.items():
        sheet[f"A{row_number}"] = temp
        sheet[f"B{row_number}"] = float(voltages[0])  # Voltage from the first DMM
        sheet[f"C{row_number}"] = float(voltages[1])  # Voltage from the second DMM
        voltage_data_dmm1[temp] = float(voltages[0])
        voltage_data_dmm2[temp] = float(voltages[1])
        row_number += 1

    # Create a scatter chart for DMM1 voltage
    chart_dmm1 = ScatterChart()
    chart_dmm1.title = "Temperature vs. Voltage (DMM1)"
    chart_dmm1.x_axis.title = "Temperature(C)"
    chart_dmm1.y_axis.title = "Voltage (mV)"
    chart_dmm1.legend = None

    x_values_dmm1 = Reference(sheet, min_col=1, min_row=2, max_row=row_number - 1)
    y_values_dmm1 = Reference(sheet, min_col=2, min_row=1, max_row=row_number - 1)

    series_dmm1 = Series(y_values_dmm1, x_values_dmm1, title_from_data=True)

    chart_dmm1.series.append(series_dmm1)

    s_dmm1 = chart_dmm1.series[0]
    s_dmm1.marker.symbol = "dot"
    s_dmm1.marker.graphicalProperties.solidFill = "000000"  # Marker filling
    s_dmm1.marker.graphicalProperties.line.solidFill = "000000"  # Marker outline

    line_dmm1 = chart_dmm1.series[0]
    line_dmm1.trendline = Trendline(trendlineType='linear', dispRSqr=True, dispEq=True)
    line_dmm1.graphicalProperties.line.noFill = True
    line_dmm1.legend = None

    sheet.add_chart(chart_dmm1, "E2")

    # Create a scatter chart for DMM2 voltage
    chart_dmm2 = ScatterChart()
    chart_dmm2.title = "Temperature vs. Voltage (DMM2)"
    chart_dmm2.x_axis.title = "Temperature(C)"
    chart_dmm2.y_axis.title = "Voltage (V)"
    chart_dmm2.legend = None

    x_values_dmm2 = Reference(sheet, min_col=1, min_row=2, max_row=row_number - 1)
    y_values_dmm2 = Reference(sheet, min_col=3, min_row=1, max_row=row_number - 1)

    series_dmm2 = Series(y_values_dmm2, x_values_dmm2, title_from_data=True)

    chart_dmm2.series.append(series_dmm2)

    s_dmm2 = chart_dmm2.series[0]
    s_dmm2.marker.symbol = "x"
    s_dmm2.marker.graphicalProperties.solidFill = "FF0000"  # Marker filling (red color)
    s_dmm2.marker.graphicalProperties.line.solidFill = "FF0000"  # Marker outline (red color)

    line_dmm2 = chart_dmm2.series[0]
    line_dmm2.trendline = Trendline(trendlineType='linear', dispRSqr=True, dispEq=True)
    line_dmm2.graphicalProperties.line.noFill = True
    line_dmm2.legend = None

    sheet.add_chart(chart_dmm2, "E17")

    date_string = time.strftime("%Y-%m-%d_%H%M")
    file_path = os.path.join(directory, f"temp_voltage_part_{date_string}.xlsx")
    workbook.save(file_path)

    logger.info("Data saved to %s", file_path)
    time.sleep(2)


def run_temperature_voltage_test():
    # Prompt user for start, end, and increment temperatures
    start_temp = float(input("Enter start temperature value: "))
    end_temp = float(input("Enter end temperature value: "))
    increment = float(input("Enter degree to increment by: "))
    #Format Temperatures
    start_temp_format = convert_temperature_format(start_temp)
    end_temp_format = convert_temperature_format(end_temp)

    command = 'W 300,' + start_temp_format
    #send command
    logger.info("Setting temperature: %s", command)
    device.write(command + '\r')
    time.sleep(5)
    temperature_str = str(start_temp) + " C"
    logger.info("Temperature set to %s", temperature_str)
    time.sleep(5)

    #Create dictionary to store values
    data = {}

    while start_temp <= end_temp:
        temperature_stabilized = False  # Flag variable to control the inner loop

        while not temperature_stabilized:
            try:
                #Call function to wait for temperature stabalization
                actual_temp = read_actual_temperature()
                #Stop waiting for stabalization once in range (can change the range to how you see fit)
                if abs(actual_temp - float(start_temp)) <= 0.3:
                    logger.info('Temperature has stabilized.')
                    voltage, voltage2  = measure_voltage()
                    time.sleep(3)

                    logger.info("Voltage: %s mV", voltage)
                    logger.info("Voltage 2: %s V", voltage2)
                    logger.info("Temperature: %s C", actual_temp)
                    time.sleep(5)
                    #record temp and voltage into dictionary
                    data[actual_temp] = float(voltage),float(voltage2)
                    temperature_stabilized = True  # Set the flag to exit the inner loop

            except pyvisa.VisaIOError as e:
                logger.error("Failed to stabilize temperature: %s", str(e))
                break
            except Exception as e:
                logger.error("An error occurred: %s", str(e))
                break

        if start_temp == end_temp:
            logger.info("End temperature reached: %s C", end_temp)
            break

        start_temp += increment  # Increment temperature
        logger.info('New temp: %s', start_temp)
        start_temp_format = convert_temperature_format(start_temp)
        command = 'W 300,' + start_temp_format
        response = device.write(command + '\r')
        time.sleep(3)
        temperature_str = str(start_temp) + " C"
        logger.info("Temperature set to %s", temperature_str)
        time.sleep(5)

    logger.info(data)
    time.sleep(4)

    save_data_to_excel(data, DIRECTORY)
    time.sleep(3)

if __name__ == "__main__":
    try:
        run_temperature_voltage_test()  # Call the main function to run the temperature-voltage test
    except pyvisa.VisaIOError as e:
        logger.error("VisaIOError occurred: %s", str(e))  # Log an error message if a VisaIOError exception occurs
    except Exception as e:
        logger.error("An error occurred: %s", str(e))  # Log an error message if any other exception occurs
