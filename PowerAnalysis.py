# Ver 2
# Ismael Banihamed
# do not daisy chain
import xlsxwriter
import os
import time
import Tkinter, tkFileDialog
import visa
import numpy
import matplotlib.pyplot as plt
import threading
from Tkinter import *
from decimal import *
from functools import partial

NUM_POWER_SUPPLIES = 0  # represents number of power supplies in use
NUM_CHANNELS = 3  # represents number of channels in each power supply

rm = visa.ResourceManager()
print(rm.list_resources())

period = 0
seconds = 0
curr_time = 0
row = 1
running = False  # Global flag
start_time = time.strftime("%m-%d_%Hh%Mm")
device = []
wait_time = 1000
curr = []


def _refresh():
    ListBoxDevices.delete(0, 'end')
    com_port_tuple = rm.list_resources()
    com_port_list = list(com_port_tuple)
    for i in range(len(com_port_list)):
        ListBoxDevices.insert(i, com_port_list[i])


def _connect():
    global NUM_POWER_SUPPLIES
    global NUM_CHANNELS
    global max_curr
    global temp_max
    global max_curr_entry
    global time_entry
    global channel_voltage
    global channel_current
    global channel
    global channel_chkbtn
    global curr
    global currents
    global ps
    global ax
    del device[:]
    root.geometry("1100x900")
    i = 0
    for idx in ListBoxDevices.curselection():
        device.append("")
        device[i] = rm.open_resource(ListBoxDevices.get(idx))
        device[i].query('*IDN?')  # returns the ID of the power supply to verify connection
        device[i].write('FORM ASCII')  # set the format to single precision floating point
        data_format = device[i].query('FORM?')  # returns the format being used
        print(data_format)
        print(device[i])
        i += 1
    NUM_POWER_SUPPLIES = len(device)
    start_btn.config(state="normal")  # enable all buttons after connected to instrument
    stop_btn.config(state="normal")

    # allocate space for arrays
    channel = [[0 for x in range(NUM_CHANNELS)] for y in range(NUM_POWER_SUPPLIES)]
    channel_chkbtn = [[0 for x in range(NUM_CHANNELS)] for y in range(NUM_POWER_SUPPLIES)]
    max_curr_entry = [[0 for x in range(NUM_CHANNELS)] for y in range(NUM_POWER_SUPPLIES)]
    time_entry = [[0 for x in range(NUM_CHANNELS)] for y in range(NUM_POWER_SUPPLIES)]
    channel_voltage = [[0 for x in range(NUM_CHANNELS)] for y in range(NUM_POWER_SUPPLIES)]
    channel_current = [[0 for x in range(NUM_CHANNELS)] for y in range(NUM_POWER_SUPPLIES)]
    max_curr = [[0 for x in range(NUM_CHANNELS)] for y in range(NUM_POWER_SUPPLIES)]
    temp_max = [[0 for x in range(NUM_CHANNELS)] for y in range(NUM_POWER_SUPPLIES)]
    curr = [[0 for x in range(NUM_CHANNELS)] for y in range(NUM_POWER_SUPPLIES)]
    ax = [[0 for x in range(NUM_CHANNELS)] for y in range(NUM_POWER_SUPPLIES)]
    currents = [0 for x in range(NUM_POWER_SUPPLIES)]

    for i in range(NUM_POWER_SUPPLIES):
        Tkinter.Label(root, text='Power Supply ' + str(i + 1)).grid(row=i * 5, column=1)
        Tkinter.Label(root, text="Channel Voltage").grid(row=i * 5, column=8, sticky='e')
        Tkinter.Label(root, text="Channel Current").grid(row=i * 5, column=10, sticky='e')
        Tkinter.Label(root, text="Max Current (A)").grid(row=i * 5, column=4, sticky='e')
        Tkinter.Label(root, text="Time (s)").grid(row=i * 5, column=6, sticky='e')
        for j in range(NUM_CHANNELS):
            channel[i][j] = Tkinter.IntVar()
            channel_chkbtn[i][j] = Tkinter.Checkbutton()
            max_curr_entry[i][j] = Tkinter.IntVar()
            time_entry[i][j] = Tkinter.Checkbutton()
            power_action = partial(power_toggle, i, j)
            channel_chkbtn[i][j] = Tkinter.Checkbutton(root, text="channel " + str(j + 1), variable=channel[i][j], command=power_action).grid(row=i * 5 + 1 + j, column=1)
            max_curr_entry[i][j] = Tkinter.Entry(root)
            max_curr_entry[i][j].config(state='normal', background='red')
            max_curr_entry[i][j].grid(row=i * 5 + 1 + j, column=4, sticky='e')
            time_entry[i][j] = Tkinter.Entry(root)
            time_entry[i][j].config(state='normal', background='red')
            time_entry[i][j].grid(row=i * 5 + 1 + j, column=6)
            channel_voltage[i][j] = Tkinter.Entry(root)
            channel_voltage[i][j].config(state='normal', background='red')
            channel_voltage[i][j].grid(row=i * 5 + 1 + j, column=8)
            channel_current[i][j] = Tkinter.Entry(root)
            channel_current[i][j].config(state='normal', background='red')
            channel_current[i][j].grid(row=i * 5 + 1 + j, column=10)


def power_toggle(i, j):
    if channel[i][j].get() == 1:
        if channel_voltage[i][j].index("end") == 0 or channel_current[i][j].index("end") == 0:
            channel_voltage[i][j].config(state='normal')
            channel_current[i][j].config(state='normal')
            channel_voltage[i][j].config(background='yellow')
            channel_current[i][j].config(background='yellow')
            max_curr_entry[i][j].config(background='yellow')
            time_entry[i][j].config(background='yellow')
        else:
            device[i].write("VOLT " + channel_voltage[i][j].get() + ",(@" + str(j + 1) + ")")
            device[i].write("CURR:RANG " + channel_current[i][j].get() + ",(@" + str(j + 1) + ")")
            device[i].write("OUTP ON, (@" + str(j + 1) + ")")
            channel_voltage[i][j].config(state='normal')
            channel_current[i][j].config(state='normal')
            channel_voltage[i][j].config(background='green')
            channel_current[i][j].config(background='green')
            max_curr_entry[i][j].config(background='green')
            time_entry[i][j].config(background='green')
    else:
        device[i].write("OUTP OFF, (@" + str(j + 1) + ")")
        channel_voltage[i][j].config(state='normal')
        channel_current[i][j].config(state='normal')
        channel_voltage[i][j].config(background='red')
        channel_current[i][j].config(background='red')
        max_curr_entry[i][j].config(background='red')
        time_entry[i][j].config(background='red')


def _start():
    """Enable scanning by setting the global flag to True."""
    global running
    global start_time
    global workbook
    global worksheet
    global curr_time
    global period
    global start_time
    global wait_time
    global row
    global freq
    global file_path
    global fig

    row = 1
    curr_time = 0
    start_time = time.strftime("%m-%d_%Hh%Mm")
    status.config(state='normal')  # set state text block to "polling"
    status.delete(0, 'end')
    status.insert(10, "polling")
    status.config(foreground='green')

    freq = int(frequency.get())
    wait_time = freq
    period = Decimal(1) / Decimal(freq)  # use inverse of frequency to get period
    # might need to change time interval into user specified span in order to not lose any data
    instrument_t_int = 'SENS:SWE:TINT ' + str(period) + ',(@1:3)'  # time interval between measurements for device
    instrument_num_points = 'SENS:SWE:POIN ' + str(freq) + ',(@1:3)'  # number of values to be measured by device
    print(instrument_t_int)  # print all commands sent to device
    print(instrument_num_points)  # print all commands sent to device

    fig = plt.figure()

    # open up excel book for writing data to
    filename_str = str(start_time) + "_PowerAnalysis.xlsx"  # create file name using current time stamp
    workbook = xlsxwriter.Workbook(filename_str)
    worksheet = workbook.add_worksheet()  # create file and workbook

    file_path = str(os.getcwd) + filename_str
    # write headers for data in file
    header = ["Time (s)"]
    #cell_format = workbook.add_format()
    #cell_format.set_bg_color('blue')
    k = 1
    for i in range(NUM_POWER_SUPPLIES):
        # write to power supplies the data requirements
        device[i].write(instrument_t_int)
        device[i].write(instrument_num_points)
        for j in range(NUM_CHANNELS):
            header.append("channel " + str(j + 1))
            max_curr_entry[i][j].delete(0, 'end')
            time_entry[i][j].delete(0, 'end')
            ax = plt.subplot(NUM_CHANNELS, NUM_POWER_SUPPLIES, k)
            ax.set_title("Power Supply " + str(i + 1) + ", Channel " + str(j + 1))
            k += 1

    for col, data in enumerate(header):
        worksheet.write(0, col, data)

    time_stamp.config(state="normal")
    time_stamp.delete(0, 'end')
    time_stamp.insert(10, str(time.strftime("%I:%M:%S")))  # write time stamp to start time text field
    stop_btn.config(state="active")
    frequency.config(state="disabled")
    running = True


def _scanning():
    global max_curr
    global seconds  # function is called every second
    global temp_max
    global curr_time
    global row
    global time_list
    global ax

    threads = []
    if running:  # Only do this if the Stop button has not been clicked
        seconds += 1  # increment seconds counter
        print("iteration" + str(seconds))
        time_list = numpy.arange(curr_time, curr_time + 1, period)
        worksheet.write_column(row, 0, time_list)
        curr_time += 1
        col = 0
        for i in range(NUM_POWER_SUPPLIES):
            t = threading.Thread(target=data_pull, args=(i, row, col))
            col += 3
            threads.append(t)
            t.start()
        row += int(frequency.get())
        if seconds > 1:
            k = 1
            for i in range(NUM_POWER_SUPPLIES):
                for j in range(NUM_CHANNELS):
                    ax = plt.subplot(NUM_POWER_SUPPLIES, NUM_CHANNELS, k)
                    ax.set_xticks(range(seconds - 10, seconds))
                    if seconds > 10:
                        ax.set_xlim(seconds - 10, seconds)
                    if temp_max[i][j] == max_curr[i][j]:
                        ax.plot(time_list, curr[i][j], color='red')
                    else:
                        ax.plot(time_list, curr[i][j], color='blue')
                    ax.set_title("Power Supply " + str(i + 1) + ", Channel " + str(j + 1))
                    ax.set_ylabel("Current (Amps)")
                    ax.set_xlabel("Time (seconds)")
                    k += 1
            plt.pause(0.01)
    root.after(1000, _scanning)


def data_pull(i, row, col):
    currents[i] = device[i].query('MEAS:ARR:CURR? (@1:3)')  # returns csv list of current measures from ch 1-3
    currents[i] = map(float, currents[i].split(','))
    curr[i][0] = currents[i][0:freq]
    curr[i][1] = currents[i][freq:freq*2]
    curr[i][2] = currents[i][freq*2::]
    for j in range(NUM_CHANNELS):
        worksheet.write_column(row, col + j + 1, curr[i][j])
        temp_max[i][j] = max(curr[i][j])
        if float(temp_max[i][j]) > float(max_curr[i][j]):
            max_curr[i][j] = temp_max[i][j]
            max_curr_entry[i][j].delete(0, "end")
            max_curr_entry[i][j].insert(10, str(float(temp_max[i][j])))
            time_entry[i][j].delete(0, "end")
            time_entry[i][j].insert(10, str(seconds))
    return


def _stop():
    """Stop scanning by setting the global flag to False."""
    global running
    global seconds
    running = False
    seconds = 0
    workbook.close()

    stop_btn.config(state="disabled")
    start_btn.config(state="active")
    frequency.config(state="normal")
    status.config(state='normal')
    status.delete(0, 'end')
    status.insert(10, "done")
    status.config(state="disabled")


def create_windowlist():
    global ListBoxDevices
    window = Tkinter.Toplevel(root)
    ok_btn = Tkinter.Button(window, text="Ok", command=_connect)
    ok_btn.grid(row=2, column=1)
    ListBoxDevices = Listbox(window, selectmode='multiple')
    ListBoxDevices.grid(row=1, column=1)
    refresh_btn = Tkinter.Button(window, text="-------Refresh------", command=_refresh)
    refresh_btn.grid(row=0, column=1)
    refresh_btn.config(state="active")


# GUI formatting***********************************************************************************************
root = Tkinter.Tk()  # all code below here deal with the construction of GUI
root.title("Current Analysis")

app = Tkinter.Frame(root)
app.grid()

Tkinter.Label(root, text="Frequency").grid(row=2, column=11, sticky='e')
Tkinter.Label(root, text="Start Time").grid(row=1, column=11, sticky='e')
Tkinter.Label(root, text="State").grid(row=3, column=11, sticky='e')
frequency = Tkinter.Entry(root)
status = Tkinter.Entry(root)
time_stamp = Tkinter.Entry(root)
status.config(state='normal', background='black', foreground='yellow')
time_stamp.config(state='normal', background='black', foreground='green')
frequency.config(background='black', foreground='blue')
time_stamp.insert(10, str("00:00:00"))
frequency.insert(5, "Hz")
status.insert(10, "ready")
frequency.grid(row=2, column=12)
time_stamp.grid(row=1, column=12)
status.grid(row=3, column=12)

start_btn = Tkinter.Button(root, text="Start", command=_start)
start_btn.config(state="normal", foreground='dark green')
stop_btn = Tkinter.Button(root, text="Stop", command=_stop)
stop_btn.config(state="normal", foreground='red')
connect_btn = Tkinter.Button(root, text="Connect", command=create_windowlist)

start_btn.grid(row=1, column=13)
stop_btn.grid(row=2, column=13)
connect_btn.grid(row=9, column=12)

root.geometry("300x150")
root.after(1000, _scanning)  # After 1 second, call scanning
root.mainloop()
